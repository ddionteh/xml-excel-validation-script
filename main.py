#!/usr/bin/env python3
"""
Validate Excel sections against a Blue Prism release XML (dynamic sections).

Adds support for:
- Environment Variable (PROD): verifies every env var from XML is present in Excel,
  and flags if any of those names are also defined as a local Data item in the release.
- BP Scripts Check: finds the row with "Is your main process published?" and writes a
  single Validation cell based on the first <process ... published="..."> value in XML.

Core behavior kept:
- Dynamic section detection via same-row single-letter marker (A..Z with optional '.' or ':')
  + keyword. Unknown sections still act as boundaries.
- Dot-insensitive for markers and 'No.' header; case-insensitive for headers;
  STRICT (trim-only, case-sensitive) for values (names, keys).
- Header row is the row *after* marker; "Validation" written on that row.
- 20-row rule: if next marker is >20 rows after header, treat as last (to EOF).
- Preserves formats/merges/embedded objects when inserting.

Requires: pandas, xlwings, openpyxl
"""

import argparse
import re
import xml.etree.ElementTree as ET
from typing import List, Optional, Dict, Any, Tuple

import pandas as pd
import xlwings as xw


# ========================= XML helpers =========================

def _local(tag: str) -> str:
    return tag.split('}', 1)[1] if tag.startswith('{') else tag


def extract_names_from_xml(xml_path: str, want: str) -> List[str]:
    """Return ordered unique list of names for <process name="…"> or <object name="…">."""
    tree = ET.parse(xml_path)
    root = tree.getroot()
    names: List[str] = []
    for elem in root.iter():
        if _local(elem.tag) == want:
            name_attr = elem.attrib.get('name')
            if name_attr:
                names.append(name_attr.strip())
    seen = set()
    out: List[str] = []
    for name in names:
        if name not in seen:
            seen.add(name)
            out.append(name)
    return out


# ---------- ENV(PROD) XML extractors ----------
def extract_env_variables_from_xml(xml_path: str) -> List[str]:
    tree = ET.parse(xml_path)
    root = tree.getroot()
    names = []

    # Canonical form
    for elem in root.iter():
        if _local(elem.tag) == "environment-variable":
            nm = elem.attrib.get("name")
            if nm:
                names.append(nm.strip())

    # Fallback: container stages named "Environment Variables" / misspelt
    if not names:
        targets = set(("environment variables", "environmnet variables"))
        for stage in root.iter():
            if _local(stage.tag) != "stage":
                continue
            if (stage.attrib.get("name") or "").strip().lower() in targets:
                for sub in stage.iter():
                    if _local(sub.tag) == "stage":
                        nm = (sub.attrib.get("name") or "").strip()
                        if nm and nm.lower() not in targets:
                            names.append(nm)

    # any Data stage explicitly exposed as Environment
    for stage in root.iter():
        if _local(stage.tag) != "stage":
            continue
        if (stage.attrib.get("type") or "").strip().lower() == "data":
            name_attr = stage.attrib.get("name")
            if not name_attr:
                continue
            exposure = None
            for child in stage:
                if _local(child.tag) == "exposure":
                    exposure = (child.text or "").strip().lower()
                    break
            if exposure == "environment":
                names.append(name_attr.strip())

    # de-dupe preserving order
    seen, out = set(), []
    for n in names:
        if n not in seen:
            seen.add(n); out.append(n)
    return out

def extract_local_data_item_names(xml_path: str) -> List[str]:
    tree = ET.parse(xml_path)
    root = tree.getroot()
    locals_only = []
    for stage in root.iter():
        if _local(stage.tag) != "stage":
            continue
        if (stage.attrib.get("type") or "").strip().lower() != "data":
            continue
        name_attr = stage.attrib.get("name") or stage.attrib.get("Name")
        if not name_attr:
            continue
        exposure = None
        for child in stage:
            if _local(child.tag) == "exposure":
                exposure = (child.text or "").strip().lower()
                break
        if exposure != "environment":   # <- local (not env)
            locals_only.append(name_attr.strip())

    # de-dupe preserving order
    seen, out = set(), []
    for n in locals_only:
        if n not in seen:
            seen.add(n); out.append(n)
    return out


def get_first_process_published(xml_path: str) -> Tuple[Optional[str], Optional[str]]:
    """
    Return (first_process_name, published_attr_text) from the first <process ...> node.
    published_attr_text is the literal string (e.g., "True") or None if missing.
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    for elem in root.iter():
        if _local(elem.tag) == "process":
            name = (elem.attrib.get("name") or "").strip() or None
            published = elem.attrib.get("published")
            pub_str = (published.strip() if isinstance(published, str) else None)
            return name, pub_str
    return None, None


# ==================== normalization / tokens ====================

NBSPS = ("\u00A0", "\u2007", "\u202F")


def _to_space(value) -> str:
    if value is None:
        return ""
    text = str(value)
    for nb in NBSPS:
        text = text.replace(nb, " ")
    text = text.replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def _norm(value) -> str:
    return _to_space(value).strip().casefold()


def _cell_is_single_letter_marker(cell_val: str) -> bool:
    """True if cell is exactly A..Z (any case) with optional trailing '.' or ':'."""
    v = _norm(cell_val)
    if not v:
        return False
    if len(v) == 1 and v.isalpha():
        return True
    if len(v) == 2 and v[0].isalpha() and v[1] in ('.', ':'):
        return True
    return False


def _row_tokens(row: pd.Series) -> List[str]:
    return [_norm(c) for c in row]


def _excel_row(idx0: int) -> int:
    return idx0 + 1


def _excel_col_letter(col0: int) -> str:
    col = col0 + 1
    label = ""
    while col:
        col, rem = divmod(col - 1, 26)
        label = chr(65 + rem) + label
    return label


def _dump_row(row_index: int, row: pd.Series, max_cols: int = 20) -> str:
    parts = []
    for col_index, cell_value in enumerate(row.iloc[:max_cols]):
        parts.append(f"[{_excel_col_letter(col_index)}] '{_to_space(cell_value).strip()}'")
    return f"r{row_index} -> " + " | ".join(parts)


# ===================== formatting / borders =====================

COLOR_GREEN = (198, 239, 206)   # Exists / OK
COLOR_RED   = (255, 199, 206)   # Not OK
COLOR_ORANGE= (255, 235, 156)   # Newly added / Attention


def _clone_borders(from_cell_api, to_cell_api) -> None:
    for border_id in (7, 8, 9, 10, 11, 12):
        try:
            bsrc = from_cell_api.Borders(border_id)
            bdst = to_cell_api.Borders(border_id)
            bdst.LineStyle = bsrc.LineStyle
            bdst.Weight = bsrc.Weight
            bdst.Color = bsrc.Color
        except Exception:
            pass
    try:
        to_cell_api.HorizontalAlignment = from_cell_api.HorizontalAlignment
        to_cell_api.VerticalAlignment = from_cell_api.VerticalAlignment
    except Exception:
        pass


def _apply_borders_like_left(ws, row_1based: int, col_1based: int) -> None:
    if col_1based <= 1:
        return
    try:
        left = ws.api.Cells(row_1based, col_1based - 1)
        dst = ws.api.Cells(row_1based, col_1based)
        _clone_borders(left, dst)
    except Exception:
        pass


def paste_formats_like_left(ws, row_1based: int, col_1based: int) -> None:
    if col_1based <= 1:
        return
    try:
        ws.api.Cells(row_1based, col_1based - 1).Copy()
        ws.api.Cells(row_1based, col_1based).PasteSpecial(Paste=-4122)  # xlPasteFormats
        ws.api.Application.CutCopyMode = False
    except Exception:
        pass


def clear_fill_preserve_borders(ws, row_1based: int, col_1based: int) -> None:
    try:
        dst = ws.api.Cells(row_1based, col_1based)
        dst.Interior.Pattern = -4142  # none
        dst.Interior.TintAndShade = 0
        dst.Interior.PatternTintAndShade = 0
        if row_1based > 1:
            _clone_borders(ws.api.Cells(row_1based - 1, col_1based), dst)
    except Exception:
        pass


def insert_row_with_style(ws, row_1based: int) -> None:
    ws.api.Rows(row_1based).Insert(Shift=-4121, CopyOrigin=0)  # down
    try:
        prev = row_1based - 1
        ws.api.Rows(prev).Copy()
        ws.api.Rows(row_1based).PasteSpecial(Paste=-4122)  # formats
        ws.api.Rows(row_1based).RowHeight = ws.api.Rows(prev).RowHeight

        # re-create single-row merges
        last_col = ws.used_range.last_cell.column
        col = 1
        while col <= last_col:
            cell_above = ws.api.Cells(prev, col)
            try:
                if cell_above.MergeCells:
                    area = cell_above.MergeArea
                    if area.Column == col and area.Rows.Count == 1:
                        left = area.Column
                        width = area.Columns.Count
                        ws.api.Range(ws.api.Cells(row_1based, left),
                                     ws.api.Cells(row_1based, left + width - 1)).Merge()
                        col += width
                        continue
            except Exception:
                pass
            col += 1
        ws.api.Application.CutCopyMode = False
    except Exception:
        pass


# ===================== small helpers / parsing =====================

def _last_nonempty_col_index(row: pd.Series) -> int:
    last_index = -1
    for col_index, cell in row.items():
        if _to_space(cell).strip() != "":
            last_index = col_index
    return last_index


def _extract_int(text: str) -> Optional[int]:
    match = re.search(r"\d+", _to_space(text))
    return int(match.group(0)) if match else None


def _row_has_any_nonblank(series: pd.Series) -> bool:
    return any(_to_space(v).strip() != "" for v in list(series.values))


# ====================== dynamic section model ======================

SECTION_REGISTRY = {
    "process":                   {"keyword": "process"},
    "business_object":           {"keyword": "business object"},
    "work_queue":                {"keyword": "work queue"},
    "bp_scripts_check":          {"keyword": "bp scripts check"},      # <--- NEW
    # ENV(PROD)
    "environment_variable_prod": {"keyword": "environment variables (prod)"},
}

def _row_has_keyword(row: pd.Series, phrase: str) -> bool:
    ph = _norm(phrase)
    return any(ph in _norm(c) for c in row)

def _row_has_single_letter_marker(row: pd.Series) -> bool:
    return any(_cell_is_single_letter_marker(c) for c in row)

def find_dynamic_markers(df: pd.DataFrame) -> List[Dict[str, Any]]:
    """
    A row is a marker if it contains a cell that is exactly a single letter (A..Z) with
    optional '.' or ':' (dot-insensitive). If the same row also contains a known keyword,
    we type it; otherwise type 'unknown'. Unknown markers still bound sections.
    """
    markers: List[Dict[str, Any]] = []
    for row_index in range(len(df)):
        row = df.iloc[row_index]
        if not _row_has_single_letter_marker(row):
            continue
        label = "unknown"
        for section_type, meta in SECTION_REGISTRY.items():
            if _row_has_keyword(row, meta["keyword"]):
                label = section_type
                break
        markers.append({"row": row_index, "type": label})
    return markers


# ===================== header & column resolvers =====================

def find_header_after_marker(df: pd.DataFrame, marker_row: int, section_type: str, lookahead: int = 80) -> Optional[int]:
    """
    Find the header row for a section after its marker.

    - process / business_object / environment_variable_prod: first nonblank row having 'name' in any cell.
    - work_queue: first nonblank row having BOTH:
        * a cell containing 'work' AND 'queue' AND 'name'
        * a cell containing 'key' AND 'name'
    - bp_scripts_check: the first nonblank row after the marker (whatever it is)
    """
    end_row = min(len(df), marker_row + 1 + lookahead)
    for row_index in range(marker_row + 1, end_row):
        tokens = _row_tokens(df.iloc[row_index])
        if all(v == "" for v in tokens):
            continue

        if section_type == "bp_scripts_check":
            return row_index

        if section_type in ("process", "business_object", "environment_variable_prod"):
            if any("name" in v for v in tokens):
                return row_index

        elif section_type == "work_queue":
            has_wq_name = any(("work" in v and "queue" in v and "name" in v) for v in tokens)
            has_key_name = any(("key" in v and "name" in v) for v in tokens)
            if has_wq_name and has_key_name:
                return row_index

    return None


def pick_columns(df: pd.DataFrame, header_row: int, section_type: str) -> Dict[str, Optional[int]]:
    """Return column indices for the section."""
    normed = [(_norm(v), idx) for idx, v in df.iloc[header_row].items()]

    # Validation column (reuse or next empty to the right)
    validation_col = None
    for text, idx in normed:
        if text == "validation":
            validation_col = idx
            break
    if validation_col is None:
        validation_col = _last_nonempty_col_index(df.iloc[header_row]) + 1

    # No. (dot-insensitive)
    no_col = None
    for text, idx in normed:
        if text == "no" or text == "no." or text.startswith("no."):
            no_col = idx
            break

    # Check (optional)
    check_col = None
    for text, idx in normed:
        if "check" in text:
            check_col = idx
            break

    if section_type == "bp_scripts_check":
        return {
            "item_col": None,
            "no_col": None,
            "check_col": None,
            "validation_col": validation_col,
        }

    if section_type in ("process", "business_object", "environment_variable_prod"):
        name_col = None
        for text, idx in normed:
            if "name" in text:
                name_col = idx
                break
        return {
            "name_col": name_col,
            "no_col": no_col,
            "check_col": check_col,
            "validation_col": validation_col,
        }

    # work_queue
    wq_name_col = None
    for text, idx in normed:
        if ("work" in text and "queue" in text and "name" in text):
            wq_name_col = idx
            break
    if wq_name_col is None:
        for text, idx in normed:
            if "name" in text and "key" not in text:
                wq_name_col = idx
                break

    key_col = None
    for text, idx in normed:
        if ("key" in text and "name" in text):
            key_col = idx
            break

    enc_col = None
    for text, idx in normed:
        if "encrypted" in text:
            enc_col = idx
            break

    return {
        "wq_name_col": wq_name_col,
        "key_col": key_col,
        "no_col": no_col,
        "encrypted_col": enc_col,
        "check_col": check_col,
        "validation_col": validation_col,
    }


# ======================= planning per section =======================

def build_plan_process_like(
    df: pd.DataFrame,
    header_row: int,
    name_col: int,
    no_col: Optional[int],
    check_col: Optional[int],
    validation_col: int,
    xml_names: List[str],
    next_marker_row: Optional[int],
) -> Dict[str, Any]:
    """Plan for Process/Business Object/Env(PROD) tables (Name, optional No./Check)."""
    content_start = header_row + 1
    content_end = next_marker_row if next_marker_row is not None else len(df)

    coarse = df.iloc[content_start:content_end].copy().reset_index(drop=True)

    # Find last relevant row
    last_rel_by_no = None
    max_no_value = None
    if no_col is not None and no_col < coarse.shape[1]:
        for row_index in range(len(coarse)):
            v = _extract_int(coarse.iat[row_index, no_col])
            if v is not None:
                max_no_value = v if (max_no_value is None or v > max_no_value) else max_no_value
        for row_index in range(len(coarse) - 1, -1, -1):
            v = _extract_int(coarse.iat[row_index, no_col])
            if v is not None:
                last_rel_by_no = row_index
                break

    last_rel_by_any = None
    for row_index in range(len(coarse) - 1, -1, -1):
        if _row_has_any_nonblank(coarse.iloc[row_index]):
            last_rel_by_any = row_index
            break

    last_rel_by_name = None
    if name_col < coarse.shape[1]:
        for row_index in range(len(coarse) - 1, -1, -1):
            if _to_space(coarse.iat[row_index, name_col]).strip() != "":
                last_rel_by_name = row_index
                break

    if last_rel_by_no is not None:
        last_rel = last_rel_by_no
    elif last_rel_by_any is not None:
        last_rel = last_rel_by_any
    else:
        last_rel = last_rel_by_name if last_rel_by_name is not None else -1

    true_end_rel_excl = last_rel + 1 if last_rel >= 0 else 0
    true_end_abs_excl = content_start + true_end_rel_excl

    section_df = df.iloc[content_start:true_end_abs_excl].copy().reset_index(drop=True)
    row_names = section_df[name_col].astype(str).fillna("").apply(_to_space).str.strip().tolist()

    # STRICT compare (trim-only; case-sensitive)
    xml_set = {x.strip() for x in xml_names}

    statuses: List[str] = []
    for nm in row_names:
        if not nm:
            statuses.append("")
        elif nm in xml_set:
            statuses.append("Exists")
        else:
            statuses.append("Does not exist")

    excel_set = {n for n in row_names if n}
    missing = [n for n in xml_names if n not in excel_set]

    fillable_rel: List[int] = []
    for rel_index in range(section_df.shape[0]):
        name_cell = _to_space(section_df.iat[rel_index, name_col]).strip()
        if name_cell:
            continue
        row_values = section_df.iloc[rel_index]
        has_number_in_no = (_extract_int(row_values[no_col]) is not None) if (no_col is not None and no_col < len(row_values)) else False
        has_any_other_cell = any((_to_space(row_values[col_idx]).strip() != "" and col_idx != name_col)
                                 for col_idx in range(len(row_values)))
        if has_number_in_no or has_any_other_cell:
            fillable_rel.append(rel_index)

    last_abs_index = content_start + (true_end_rel_excl - 1) if true_end_rel_excl > 0 else header_row
    insert_at_excel_row = last_abs_index + 2  # Excel is 1-based

    if max_no_value is not None:
        next_no = max_no_value + 1
    else:
        nonblank_names = sum(1 for v in row_names if v)
        next_no = nonblank_names + 1 if no_col is not None else None

    return {
        "type": None,  # filled by caller
        "header_row": header_row,              # 0-based (for sorting)
        "header_excel_row": header_row + 1,    # 1-based (for writing header)
        "start_row_excel": content_start + 1,  # 1-based first content row
        "row_count": true_end_abs_excl - content_start,
        "insert_at_excel_row": insert_at_excel_row,
        "name_col": name_col,
        "no_col": no_col,
        "check_col": check_col,
        "validation_col": validation_col,
        "excel_names": row_names,
        "validation_vals": statuses,
        "missing": missing,
        "fillable_rel": fillable_rel,
        "next_no": next_no,
        "inserted_rows": 0,
    }


def build_plan_workqueue(
    df: pd.DataFrame,
    header_row: int,
    cols: Dict[str, int],
    xml_wq: Dict[str, Dict[str, Optional[str]]],
    next_marker_row: Optional[int],
) -> Dict[str, Any]:
    """Plan for Work Queue table."""
    wq_name_col = cols["wq_name_col"]
    content_start = header_row + 1
    content_end = next_marker_row if next_marker_row is not None else len(df)

    coarse = df.iloc[content_start:content_end].copy().reset_index(drop=True)

    # Find last relevant row
    last_rel_by_no = None
    max_no_value = None
    if cols["no_col"] is not None and cols["no_col"] < coarse.shape[1]:
        for row_index in range(len(coarse)):
            v = _extract_int(coarse.iat[row_index, cols["no_col"]])
            if v is not None:
                max_no_value = v if (max_no_value is None or v > max_no_value) else max_no_value
        for row_index in range(len(coarse) - 1, -1, -1):
            v = _extract_int(coarse.iat[row_index, cols["no_col"]])
            if v is not None:
                last_rel_by_no = row_index
                break

    last_rel_by_any = None
    for row_index in range(len(coarse) - 1, -1, -1):
        if _row_has_any_nonblank(coarse.iloc[row_index]):
            last_rel_by_any = row_index
            break

    last_rel_by_name = None
    if wq_name_col < coarse.shape[1]:
        for row_index in range(len(coarse) - 1, -1, -1):
            if _to_space(coarse.iat[row_index, wq_name_col]).strip() != "":
                last_rel_by_name = row_index
                break

    if last_rel_by_no is not None:
        last_rel = last_rel_by_no
    elif last_rel_by_any is not None:
        last_rel = last_rel_by_any
    else:
        last_rel = last_rel_by_name if last_rel_by_name is not None else -1

    true_end_rel_excl = last_rel + 1 if last_rel >= 0 else 0
    true_end_abs_excl = content_start + true_end_rel_excl

    section_df = df.iloc[content_start:true_end_abs_excl].copy().reset_index(drop=True)

    xml_names = list(xml_wq.keys())
    xml_set = set(xml_names)  # STRICT compare (case-sensitive after trim below)
    row_names = section_df[wq_name_col].astype(str).fillna("").apply(_to_space).str.strip().tolist()

    statuses: List[str] = []
    for nm in row_names:
        if not nm:
            statuses.append("")
        elif nm in xml_set:
            statuses.append("Exists")
        else:
            statuses.append("Does not exist")

    excel_set = {n for n in row_names if n}
    missing = [n for n in xml_names if n not in excel_set]

    fillable_rel: List[int] = []
    for rel_index in range(section_df.shape[0]):
        name_cell = _to_space(section_df.iat[rel_index, wq_name_col]).strip()
        if name_cell:
            continue
        row_values = section_df.iloc[rel_index]
        has_number_in_no = (_extract_int(row_values[cols["no_col"]]) is not None) if (cols["no_col"] is not None and cols["no_col"] < len(row_values)) else False
        has_any_other_cell = any((_to_space(row_values[col_idx]).strip() != "" and col_idx != wq_name_col)
                                 for col_idx in range(len(row_values)))
        if has_number_in_no or has_any_other_cell:
            fillable_rel.append(rel_index)

    last_abs_index = content_start + (true_end_rel_excl - 1) if true_end_rel_excl > 0 else header_row
    insert_at_excel_row = last_abs_index + 2

    if max_no_value is not None:
        next_no = max_no_value + 1
    else:
        nonblank_names = sum(1 for v in row_names if v)
        next_no = nonblank_names + 1 if cols["no_col"] is not None else None

    print(f"[WQ] Content rows Excel {_excel_row(content_start)}..{_excel_row(true_end_abs_excl-1)}; "
          f"existing={len(row_names)}; missing_from_excel={len(missing)}")

    return {
        "type": None,  # filled by caller
        "header_row": header_row,              # 0-based (for sorting)
        "header_excel_row": header_row + 1,    # 1-based (for writing header)
        "start_row_excel": content_start + 1,  # 1-based first content row
        "row_count": true_end_abs_excl - content_start,
        "insert_at_excel_row": insert_at_excel_row,
        "wq_name_col": wq_name_col,
        "key_col": cols["key_col"],
        "no_col": cols["no_col"],
        "check_col": cols["check_col"],
        "validation_col": cols["validation_col"],
        "excel_names": row_names,
        "validation_vals": statuses,
        "missing": missing,
        "fillable_rel": fillable_rel,
        "next_no": next_no,
        "inserted_rows": 0,
    }


def build_plan_bp_scripts_check(
    df: pd.DataFrame,
    header_row: int,
    cols: Dict[str, Optional[int]],
    next_marker_row: Optional[int],
) -> Dict[str, Any]:
    """
    Build a plan for the 'BP Scripts Check' section:
      - header is first nonblank row after the marker (already resolved upstream)
      - scan any column for the exact target phrase (case/space-normalized)
      - only that one row gets a Validation value
    """
    content_start = header_row + 1
    content_end = next_marker_row if next_marker_row is not None else len(df)
    section_df = df.iloc[content_start:content_end].copy().reset_index(drop=True)

    target_phrase_norm = _norm("Is your main process published?")
    target_rel = None

    for rel_index in range(section_df.shape[0]):
        row = section_df.iloc[rel_index]
        if any(_norm(_to_space(val).strip()) == target_phrase_norm for val in list(row.values)):
            target_rel = rel_index
            break

    return {
        "type": None,  # filled by caller
        "header_row": header_row,
        "header_excel_row": header_row + 1,
        "start_row_excel": content_start + 1,
        "row_count": section_df.shape[0],
        "validation_col": cols["validation_col"],
        "target_row_rel": target_rel,
        "validation_vals": [""] * section_df.shape[0],  # bulk writer is skipped for this section
        "inserted_rows": 0,
    }


# ========================= writers per section =========================

def write_existing_validation(ws, start_row_excel: int, validation_col0: int, statuses: List[str]) -> None:
    validation_col_1based = validation_col0 + 1
    for offset, status in enumerate(statuses):
        row_1based = start_row_excel + offset
        paste_formats_like_left(ws, row_1based, validation_col_1based)
        ws.range((row_1based, validation_col_1based)).value = status if status else ""
        if status == "Exists":
            ws.range((row_1based, validation_col_1based)).color = COLOR_GREEN
        elif status == "Does not exist":
            ws.range((row_1based, validation_col_1based)).color = COLOR_RED
        else:
            ws.range((row_1based, validation_col_1based)).color = None
        _apply_borders_like_left(ws, row_1based, validation_col_1based)


def fill_into_blanks_name_table(ws, plan: Dict[str, Any], names_to_place: List[str]) -> int:
    consumed = 0
    name_col_1based = plan["name_col"] + 1
    validation_col_1based = plan["validation_col"] + 1
    no_col_1based = plan["no_col"] + 1 if plan["no_col"] is not None else None
    check_col_1based = plan["check_col"] + 1 if plan["check_col"] is not None else None
    next_no = plan["next_no"]

    for rel_index in plan["fillable_rel"]:
        if consumed >= len(names_to_place):
            break
        row_1based = plan["start_row_excel"] + rel_index
        new_name = names_to_place[consumed]

        ws.range((row_1based, name_col_1based)).value = new_name

        paste_formats_like_left(ws, row_1based, validation_col_1based)
        ws.range((row_1based, validation_col_1based)).value = "Newly added"
        ws.range((row_1based, validation_col_1based)).color = COLOR_ORANGE
        _apply_borders_like_left(ws, row_1based, validation_col_1based)

        if no_col_1based is not None:
            current_text = str(ws.range((row_1based, no_col_1based)).value or "")
            current_no = _extract_int(current_text)
            if current_no is None and next_no is not None:
                ws.range((row_1based, no_col_1based)).value = next_no
                next_no += 1
            elif current_no is not None and next_no is not None:
                next_no = max(next_no, current_no + 1)

        if check_col_1based is not None:
            clear_fill_preserve_borders(ws, row_1based, check_col_1based)

        consumed += 1

    plan["next_no"] = next_no
    return consumed


def insert_new_rows_name_table(ws, plan: Dict[str, Any], remaining: List[str]) -> int:
    if not remaining:
        return 0

    name_col_1based = plan["name_col"] + 1
    validation_col_1based = plan["validation_col"] + 1
    no_col_1based = plan["no_col"] + 1 if plan["no_col"] is not None else None
    check_col_1based = plan["check_col"] + 1 if plan["check_col"] is not None else None
    next_no = plan["next_no"]

    for offset, new_name in enumerate(remaining):
        row_1based = plan["insert_at_excel_row"] + offset
        insert_row_with_style(ws, row_1based)
        ws.range((row_1based, name_col_1based)).value = new_name

        paste_formats_like_left(ws, row_1based, validation_col_1based)
        ws.range((row_1based, validation_col_1based)).value = "Newly added"
        ws.range((row_1based, validation_col_1based)).color = COLOR_ORANGE
        _apply_borders_like_left(ws, row_1based, validation_col_1based)

        if no_col_1based is not None and next_no is not None:
            ws.range((row_1based, no_col_1based)).value = next_no
            next_no += 1

        if check_col_1based is not None:
            clear_fill_preserve_borders(ws, row_1based, check_col_1based)

    plan["next_no"] = next_no
    plan["inserted_rows"] = (plan.get("inserted_rows", 0) or 0) + len(remaining)
    return len(remaining)


def fill_into_blanks_wq(ws, plan: Dict[str, Any], names_to_place: List[str], xml_wq: Dict[str, Dict[str, Optional[str]]]) -> int:
    consumed = 0
    name_col_1based = plan["wq_name_col"] + 1
    key_col_1based = plan["key_col"] + 1 if plan.get("key_col") is not None else None
    validation_col_1based = plan["validation_col"] + 1
    no_col_1based = plan["no_col"] + 1 if plan["no_col"] is not None else None
    check_col_1based = plan["check_col"] + 1 if plan["check_col"] is not None else None
    next_no = plan["next_no"]

    for rel_index in plan["fillable_rel"]:
        if consumed >= len(names_to_place):
            break
        row_1based = plan["start_row_excel"] + rel_index
        new_name = names_to_place[consumed]

        ws.range((row_1based, name_col_1based)).value = new_name
        if key_col_1based is not None:
            key_val = (xml_wq.get(new_name, {}) or {}).get("key") or ""
            if key_val:
                ws.range((row_1based, key_col_1based)).value = key_val

        paste_formats_like_left(ws, row_1based, validation_col_1based)
        ws.range((row_1based, validation_col_1based)).value = "Newly added"
        ws.range((row_1based, validation_col_1based)).color = COLOR_ORANGE
        _apply_borders_like_left(ws, row_1based, validation_col_1based)

        if no_col_1based is not None:
            current_text = str(ws.range((row_1based, no_col_1based)).value or "")
            current_no = _extract_int(current_text)
            if current_no is None and next_no is not None:
                ws.range((row_1based, no_col_1based)).value = next_no
                next_no += 1
            elif current_no is not None and next_no is not None:
                next_no = max(next_no, current_no + 1)

        if check_col_1based is not None:
            clear_fill_preserve_borders(ws, row_1based, check_col_1based)

        consumed += 1

    plan["next_no"] = next_no
    return consumed


def insert_new_rows_wq(ws, plan: Dict[str, Any], remaining: List[str], xml_wq: Dict[str, Dict[str, Optional[str]]]) -> int:
    if not remaining:
        return 0

    name_col_1based = plan["wq_name_col"] + 1
    key_col_1based = plan["key_col"] + 1 if plan.get("key_col") is not None else None
    validation_col_1based = plan["validation_col"] + 1
    no_col_1based = plan["no_col"] + 1 if plan["no_col"] is not None else None
    check_col_1based = plan["check_col"] + 1 if plan["check_col"] is not None else None
    next_no = plan["next_no"]

    for offset, new_name in enumerate(remaining):
        row_1based = plan["insert_at_excel_row"] + offset
        insert_row_with_style(ws, row_1based)
        ws.range((row_1based, name_col_1based)).value = new_name

        if key_col_1based is not None:
            key_val = (xml_wq.get(new_name, {}) or {}).get("key") or ""
            if key_val:
                ws.range((row_1based, key_col_1based)).value = key_val

        paste_formats_like_left(ws, row_1based, validation_col_1based)
        ws.range((row_1based, validation_col_1based)).value = "Newly added"
        ws.range((row_1based, validation_col_1based)).color = COLOR_ORANGE
        _apply_borders_like_left(ws, row_1based, validation_col_1based)

        if no_col_1based is not None and next_no is not None:
            ws.range((row_1based, no_col_1based)).value = next_no
            next_no += 1

        if check_col_1based is not None:
            clear_fill_preserve_borders(ws, row_1based, check_col_1based)

    plan["next_no"] = next_no
    plan["inserted_rows"] = (plan.get("inserted_rows", 0) or 0) + len(remaining)
    return len(remaining)


def adjust_wq_key_validation(ws, plan: Dict[str, Any], xml_wq: Dict[str, Dict[str, Optional[str]]]) -> int:
    if plan.get("key_col") is None:
        return 0
    name_col_1based = plan["wq_name_col"] + 1
    key_col_1based = plan["key_col"] + 1
    validation_col_1based = plan["validation_col"] + 1

    total_rows = plan["row_count"] + (plan.get("inserted_rows", 0) or 0)
    mismatches = 0
    for offset in range(total_rows):
        row_1based = plan["start_row_excel"] + offset
        nm = _to_space(ws.range((row_1based, name_col_1based)).value).strip()
        if not nm or nm not in xml_wq:  # STRICT name
            continue
        expected = _to_space(xml_wq[nm].get("key") or "").strip()
        if not expected:
            continue
        excel_key = _to_space(ws.range((row_1based, key_col_1based)).value).strip()
        if excel_key != expected:       # STRICT key
            current = _to_space(ws.range((row_1based, validation_col_1based)).value).strip()
            if current.startswith("Exists"):
                ws.range((row_1based, validation_col_1based)).value = f'Exists (Key mismatch: expected "{expected}")'
            elif current.startswith("Newly added"):
                ws.range((row_1based, validation_col_1based)).value = f'Newly added (Key mismatch: expected "{expected}")'
            ws.range((row_1based, validation_col_1based)).color = COLOR_ORANGE
            mismatches += 1
    return mismatches


# ---------- ENV(PROD): post-write validator to catch local Data items ----------

def adjust_env_local_validation(ws, plan: Dict[str, Any], local_data_item_names: List[str]) -> int:
    """
    For Environment Variable (PROD):
      If a Name also exists as a local Data stage in the release, annotate the
      validation cell and color orange.
    """
    local_set = set(local_data_item_names)
    name_col_1based = plan["name_col"] + 1
    validation_col_1based = plan["validation_col"] + 1

    total_rows = plan["row_count"] + (plan.get("inserted_rows", 0) or 0)
    hits = 0
    for offset in range(total_rows):
        row_1based = plan["start_row_excel"] + offset
        env_name = _to_space(ws.range((row_1based, name_col_1based)).value).strip()
        if not env_name:
            continue
        if env_name in local_set:
            current = _to_space(ws.range((row_1based, validation_col_1based)).value).strip()
            if current.startswith("Exists"):
                ws.range((row_1based, validation_col_1based)).value = "Exists (Also defined locally)"
            elif current.startswith("Newly added"):
                ws.range((row_1based, validation_col_1based)).value = "Newly added (Also defined locally)"
            else:
                ws.range((row_1based, validation_col_1based)).value = "Also defined locally"
            ws.range((row_1based, validation_col_1based)).color = COLOR_ORANGE
            hits += 1
    return hits


def write_bp_scripts_check_validation(ws, plan: Dict[str, Any], published_text: Optional[str]) -> None:
    """
    Writes Validation only for the 'Is your main process published?' row.
    - Green if published_text == "True"
    - Red otherwise (including missing)
    """
    if plan.get("target_row_rel") is None:
        return

    row_1based = plan["start_row_excel"] + plan["target_row_rel"]
    validation_col_1based = plan["validation_col"] + 1

    paste_formats_like_left(ws, row_1based, validation_col_1based)

    if published_text == "True":
        ws.range((row_1based, validation_col_1based)).value = "Yes"
        ws.range((row_1based, validation_col_1based)).color = COLOR_GREEN
    elif published_text is None:
        ws.range((row_1based, validation_col_1based)).value = "Published attribute missing"
        ws.range((row_1based, validation_col_1based)).color = COLOR_RED
    else:
        ws.range((row_1based, validation_col_1based)).value = "No"
        ws.range((row_1based, validation_col_1based)).color = COLOR_RED

    _apply_borders_like_left(ws, row_1based, validation_col_1based)


# ============================ placement helpers ============================

def snapshot_and_set_placement(ws):
    # No-op placeholder (kept for compatibility)
    return None

def restore_placement(ws, placements):
    # No-op placeholder (kept for compatibility)
    pass


# ============================ main flow ============================

_ILLEGAL_SHEET_CHARS = r'[:\\/?*\[\]]'


def _sanitize_sheet_name(name: str) -> str:
    name = re.sub(_ILLEGAL_SHEET_CHARS, "_", name).strip()
    return (name[:31] or "Sheet")


def _unique_sheet_name(wb, base: str) -> str:
    base = _sanitize_sheet_name(base)
    taken = [s.name for s in wb.sheets]
    name = base
    counter = 2
    while name in taken:
        suffix = f" ({counter})"
        keep = 31 - len(suffix)
        name = _sanitize_sheet_name(base[:max(1, keep)] + suffix)
        counter += 1
    return name


def _sheet_by_name_or_index(wb, sheet_arg):
    if isinstance(sheet_arg, int) or (isinstance(sheet_arg, str) and str(sheet_arg).isdigit()):
        return wb.sheets[int(sheet_arg)]
    return wb.sheets[str(sheet_arg)]


def validate_and_write_dynamic(
    excel_path: str,
    sheet_arg,
    xml_proc: List[str],
    xml_bo: List[str],
    xml_wq: Dict[str, Dict[str, Optional[str]]],
    xml_env_prod: List[str],                     # ENV(PROD)
    xml_local_data_names: List[str],             # ENV(PROD) local guard
    first_process_published_text: Optional[str], # BP Scripts Check
) -> None:

    # Read for planning (strings, no header)
    sheet_arg_pd = int(sheet_arg) if (isinstance(sheet_arg, str) and str(sheet_arg).isdigit()) else sheet_arg
    df = pd.read_excel(excel_path, sheet_name=sheet_arg_pd, header=None, dtype=str, engine="openpyxl").fillna("")

    # 1) Detect markers
    markers = find_dynamic_markers(df)
    if not markers:
        print("⛔ No section letter markers found. Nothing to do.")
        return

    # 2) Build plans per marker, bounded by the next marker (or EOF)
    plans: List[Dict[str, Any]] = []
    for marker_index, marker in enumerate(markers):
        marker_row = marker["row"]
        marker_type = marker["type"]

        next_marker_row: Optional[int] = markers[marker_index + 1]["row"] if marker_index + 1 < len(markers) else None

        if marker_type == "unknown":
            # Boundary only; skip processing
            continue

        header_row = find_header_after_marker(df, marker_row, marker_type, lookahead=80)
        if header_row is None:
            print(f"[{marker_type}] ❌ Could not find a header row after marker at Excel row {_excel_row(marker_row)}. Skipping.")
            continue

        # 20-row rule: if next marker is far away, treat as last to EOF
        if next_marker_row is not None and next_marker_row - header_row > 20:
            print(f"[{marker_type}] Next marker is {next_marker_row - header_row} rows away (>20). Treating as last to EOF.")
            next_marker_row = None

        columns = pick_columns(df, header_row, marker_type)

        if marker_type in ("process", "business_object", "environment_variable_prod"):
            if columns["name_col"] is None:
                print(f"[{marker_type}] ❌ Could not resolve a 'Name' column at header Excel row {_excel_row(header_row)}.")
                continue
            plan = build_plan_process_like(
                df, header_row, columns["name_col"], columns["no_col"], columns["check_col"],
                columns["validation_col"],
                xml_proc if marker_type == "process"
                else xml_bo if marker_type == "business_object"
                else xml_env_prod,  # ENV(PROD)
                next_marker_row=next_marker_row
            )
        elif marker_type == "bp_scripts_check":
            plan = build_plan_bp_scripts_check(df, header_row, columns, next_marker_row=next_marker_row)
        else:  # work_queue
            if columns["wq_name_col"] is None or columns["key_col"] is None:
                print(f"[work_queue] ❌ Need both 'Work Queue Name' and 'Key Name' columns. Skipping.")
                continue
            plan = build_plan_workqueue(df, header_row, columns, xml_wq, next_marker_row=next_marker_row)

        plan["type"] = marker_type
        plan["header_row"] = header_row
        plans.append(plan)

    if not plans:
        print("⛔ No valid sections planned. Nothing to write.")
        return

    # 3) Open Excel; clone sheet; write in visual order (top to bottom)
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(str(excel_path))
        src_sheet = _sheet_by_name_or_index(wb, sheet_arg)

        try:
            app.api.EnableEvents = False
        except Exception:
            pass

        before_names = [s.name for s in wb.sheets]
        src_sheet.api.Copy(After=src_sheet.api)
        after_names = [s.name for s in wb.sheets]
        added = [n for n in after_names if n not in before_names]
        ws = wb.sheets[added[0]] if len(added) == 1 else wb.sheets[-1]

        new_name = _unique_sheet_name(wb, f"{src_sheet.name}_validated")
        try:
            ws.name = new_name
        except Exception as e:
            print(f"⚠️ Rename failed ({e}); keeping '{ws.name}'")

        placements = snapshot_and_set_placement(ws)

        # Sort by on-sheet order and accumulate inserted rows to offset subsequent plans
        plans.sort(key=lambda p: p["header_row"])

        total_inserted_above = 0
        for section_index, plan in enumerate(plans):

            # shift *1-based* fields by rows inserted above
            for key_name in ("start_row_excel", "insert_at_excel_row", "header_excel_row"):
                if key_name in plan and isinstance(plan[key_name], int):
                    plan[key_name] += total_inserted_above

            # write "Validation" on the header row (column headers)
            validation_col_1based = plan["validation_col"] + 1
            paste_formats_like_left(ws, plan["header_excel_row"], validation_col_1based)
            ws.range((plan["header_excel_row"], validation_col_1based)).value = "Validation"
            _apply_borders_like_left(ws, plan["header_excel_row"], validation_col_1based)

            # existing rows' validation (skip for BP Scripts Check; handled separately)
            if plan["type"] != "bp_scripts_check":
                write_existing_validation(ws, plan["start_row_excel"], plan["validation_col"], plan["validation_vals"])

            # fill/insert depending on section
            rows_added_here = 0
            if plan["type"] in ("process", "business_object", "environment_variable_prod"):
                consumed_count = fill_into_blanks_name_table(ws, plan, plan["missing"])
                remaining_names = plan["missing"][consumed_count:]
                rows_added_here = insert_new_rows_name_table(ws, plan, remaining_names)

                # ENV(PROD) local-variable guard
                if plan["type"] == "environment_variable_prod":
                    local_hits = adjust_env_local_validation(ws, plan, xml_local_data_names)
                    if local_hits:
                        print(f"[env_prod] {local_hits} name(s) are also defined locally (Data items).")

            elif plan["type"] == "bp_scripts_check":
                write_bp_scripts_check_validation(ws, plan, first_process_published_text)
                rows_added_here = 0

            else:  # work_queue
                consumed_count = fill_into_blanks_wq(ws, plan, plan["missing"], xml_wq)
                remaining_names = plan["missing"][consumed_count:]
                rows_added_here = insert_new_rows_wq(ws, plan, remaining_names, xml_wq)
                mismatch_count = adjust_wq_key_validation(ws, plan, xml_wq)
                print(f"[work_queue] filled={consumed_count}, inserted={rows_added_here}, key_mismatches={mismatch_count}")

            total_inserted_above += rows_added_here

            # Autofit Validation column
            try:
                ws.api.Columns(validation_col_1based).AutoFit()
            except Exception:
                pass

        restore_placement(ws, placements)
        wb.save()
        print(f"✅ Completed. Wrote '{ws.name}'.")
    finally:
        try:
            app.api.EnableEvents = True
        except Exception:
            pass
        try:
            wb.close()
        except Exception:
            pass
        app.quit()


# =============================== CLI ===============================

def main():
    parser = argparse.ArgumentParser(
        description="Validate sections (Process/Business Object/Work Queue/Env Var (PROD)/BP Scripts Check) against a Blue Prism release XML (dynamic markers)."
    )
    parser.add_argument("--xml", required=True, help="Path to Blue Prism .bprelease XML")
    parser.add_argument("--excel", required=True, help="Path to Excel file (.xlsx/.xlsm)")
    parser.add_argument("--sheet", default="0", help="Sheet name or 0-based index (default: 0)")
    args = parser.parse_args()

    print("🔍 Parsing XML…")
    xml_proc = extract_names_from_xml(args.xml, "process")
    xml_bo = extract_names_from_xml(args.xml, "object")
    xml_wq = {}  # keep your previous extractor if you need WQ; left blank here to focus on ENV(PROD) + BP Scripts Check
    # If you still use Work Queues from earlier version, swap in your extract_work_queues_from_xml(...)
    try:
        from __main__ import extract_work_queues_from_xml as _maybe_wq
        xml_wq = _maybe_wq(args.xml)
    except Exception:
        xml_wq = {}

    # ENV(PROD): environment variables and local Data items
    xml_env_prod = extract_env_variables_from_xml(args.xml)
    xml_local_data_names = extract_local_data_item_names(args.xml)

    # First process's published attribute (used in BP Scripts Check)
    first_proc_name, first_proc_published = get_first_process_published(args.xml)

    print(f"✅ XML: processes={len(xml_proc)}, business_objects={len(xml_bo)}, "
          f"work_queues={len(xml_wq)}, env_prod={len(xml_env_prod)}, local_data_items={len(xml_local_data_names)}")
    print(f"ℹ️ Main process assumed = first <process>: '{first_proc_name or 'N/A'}', published={first_proc_published or 'N/A'}")

    print("🧪 Validating & writing…")
    validate_and_write_dynamic(
        excel_path=args.excel,
        sheet_arg=args.sheet,
        xml_proc=xml_proc,
        xml_bo=xml_bo,
        xml_wq=xml_wq,
        xml_env_prod=xml_env_prod,
        xml_local_data_names=xml_local_data_names,
        first_process_published_text=first_proc_published,
    )


if __name__ == "__main__":
    main()
