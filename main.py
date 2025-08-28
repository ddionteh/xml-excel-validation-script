#!/usr/bin/env python3
"""
Validate Excel sections against a Blue Prism release XML (dynamic sections).

Key points
----------
- Process section: inserts THREE columns before Validation:
    1) Published status
    2) Hard coded values
    3) Exception Types
  (then writes Validation to the right)

- Published status: writes the process's published value. Robust reader:
  prefers <process published="..."> (any case), else falls back to child
  elements like <published>true</published> / <is-published>true</is-published>.

- Exception Types: only values CONTAINING "Exception". Allowed:
  "System Exception", "Business Exception". Others flagged as
  "Found unknown exception types: …". If none contain "Exception", leave blank.

- BP Scripts Check: "Is your main process published?" is YES only if
  every process that HAS a published value evaluates to True (case-insensitive).

- Environment Variables (PROD): via Data stages with <exposure>Environment</exposure>
  (preferred) + fallback “Environment Variables” / “Environmnet Variables” block,
  with a structure dump for inspection.

- Work Queue header detection: phrase “Work Queue Name” (and “Key Name”),
  not split-word matching.

- Prevents vertical merge bleed: unmerges vertically-merged cells in each section
  before writing, and unmerges any target cell before setting a value.

- Headers wrapped + min width applied to the new columns; values wrap and rows auto-fit.

Requires: pandas, xlwings, openpyxl
"""

import argparse
import csv
import re
import xml.etree.ElementTree as ET
from typing import List, Optional, Dict, Any, Tuple, Set

import pandas as pd
import xlwings as xw
from contextlib import contextmanager


# ========================= XML helpers =========================

def _local(tag: str) -> str:
    return tag.split('}', 1)[1] if tag.startswith('{') else tag


def extract_names_from_xml(xml_path: str, want: str) -> List[str]:
    tree = ET.parse(xml_path)
    root = tree.getroot()
    names: List[str] = []
    for elem in root.iter():
        if _local(elem.tag).casefold() == want.casefold():   # <- case-insensitive
            name_attr = elem.attrib.get('name') or elem.attrib.get('Name')
            if name_attr:
                names.append(name_attr.strip())
    # ordered de-dupe
    seen = set()
    out: List[str] = []
    for name in names:
        if name not in seen:
            seen.add(name); out.append(name)
    return out


def _read_published_value(proc_elem: ET.Element) -> Optional[str]:
    """
    Prefer <process published="..."> (any case), else fall back to child nodes:
      <published>...</published>, <is-published>...</is-published>, <ispublished>...</ispublished>
    Return the literal string found (normalized to 'True'/'False' for common booleans), else None.
    """
    # 1) attribute (accept different casings)
    for attr_name in ("published", "Published", "PUBLISHED"):
        if attr_name in proc_elem.attrib:
            raw = (proc_elem.attrib.get(attr_name) or "").strip()
            break
    else:
        raw = None

    # 2) child nodes if attribute missing/empty
    if not raw:
        for child in proc_elem.iter():
            tag = _local(child.tag).casefold()
            if tag in {"published", "is-published", "ispublished"}:
                raw = (child.text or "").strip()
                if raw:
                    break

    if not raw:
        return None

    # normalize common boolean strings but keep a readable literal
    low = raw.casefold()
    if low in {"true", "1", "yes", "y"}:
        return "True"
    if low in {"false", "0", "no", "n"}:
        return "False"
    return raw  # unusual literal e.g. "TRUE " or "t" — keep as-is

def get_process_metadata(xml_path: str) -> Dict[str, Dict[str, Any]]:
    """
    Returns:
      { "<process name>": {"published": "True"/"False"/<literal>/None, "exception_types": set([...])}, ...}
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    meta: Dict[str, Dict[str, Any]] = {}

    for proc in root.iter():
        if _local(proc.tag).casefold() != "process":   # <- case-insensitive
            continue
        name = (proc.attrib.get("name") or proc.attrib.get("Name") or "").strip()
        if not name:
            continue

        published = _read_published_value(proc)

        types: Set[str] = set()
        for elem in proc.iter():
            if _local(elem.tag).casefold() == "exception":   # <- case-insensitive
                t = elem.attrib.get("type")
                if t:
                    types.add(t.strip())

        # Merge with existing entry without losing a real published value
        if name in meta:
            prev = meta[name]
            keep_pub = prev.get("published")
            if keep_pub is None and published is not None:
                keep_pub = published  # only upgrade None -> real value
            meta[name] = {
                "published": keep_pub,
                "exception_types": (prev.get("exception_types") or set()) | types,
            }
        else:
            meta[name] = {"published": published, "exception_types": types}

    return meta


@contextmanager
def excel_perf_mode(app):
    api = app.api
    prev = {}
    for prop in ("ScreenUpdating", "DisplayAlerts", "EnableEvents", "Calculation"):
        try:
            prev[prop] = getattr(api, prop)
        except Exception:
            prev[prop] = None
    try:
        try: api.ScreenUpdating = False
        except Exception: pass
        try: api.DisplayAlerts = False
        except Exception: pass
        try: api.EnableEvents = False
        except Exception: pass
        try: api.Calculation = -4135   # xlCalculationManual
        except Exception: pass
        yield
    finally:
        try:
            if prev.get("Calculation") is not None: api.Calculation = prev["Calculation"]
            if prev.get("ScreenUpdating") is not None: api.ScreenUpdating = prev["ScreenUpdating"]
            if prev.get("DisplayAlerts") is not None: api.DisplayAlerts = prev["DisplayAlerts"]
            if prev.get("EnableEvents") is not None: api.EnableEvents = prev["EnableEvents"]
        except Exception:
            pass


# ---------- Environment variables via exposure / block ----------

def extract_env_variables_from_xml(xml_path: str) -> List[str]:
    tree = ET.parse(xml_path)
    root = tree.getroot()
    names: List[str] = []

    # Preferred: Data stage with exposure=Environment
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
        if exposure == "environment":
            names.append(name_attr.strip())

    # Fallback container/block
    targets = {"environment variables", "environmnet variables"}
    for stage in root.iter():
        if _local(stage.tag) != "stage":
            continue
        stage_name = (stage.attrib.get("name") or "").strip()
        if stage_name.casefold() in targets:
            print(f"[env_debug] Found block '{stage_name}' (type={stage.attrib.get('type')})")
            child_count = 0
            for sub in stage:
                if _local(sub.tag) != "stage":
                    continue
                child_count += 1
                sub_name = (sub.attrib.get("name") or sub.attrib.get("Name") or "").strip()
                sub_type = (sub.attrib.get("type") or "").strip()
                exposure = None
                for gg in sub:
                    if _local(gg.tag) == "exposure":
                        exposure = (gg.text or "").strip()
                        break
                print(f"  - child stage: name='{sub_name}', type='{sub_type}', exposure='{exposure or ''}'")
                if sub_name and sub_name.casefold() not in targets:
                    names.append(sub_name)
            print(f"[env_debug] Block children counted: {child_count}")

    # de-dupe
    seen, out = set(), []
    for n in names:
        if n not in seen:
            seen.add(n); out.append(n)
    return out


def extract_local_data_item_names(xml_path: str) -> List[str]:
    tree = ET.parse(xml_path)
    root = tree.getroot()
    locals_only: List[str] = []
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
        if exposure != "environment":
            locals_only.append(name_attr.strip())

    seen, out = set(), []
    for n in locals_only:
        if n not in seen:
            seen.add(n); out.append(n)
    return out


# ---------- Hardcoded literal finder (heuristic, optional) ----------

_TEXTY_TAGS = {"initialvalue", "expression", "text", "sql", "note",
               "url", "path", "filename", "address", "value"}
_TEXTY_ATTRS = {"value", "expression", "text", "sql", "url", "address", "path",
                "filename", "server", "database", "username", "password"}

def _looks_hardcoded(val: str, min_len: int = 3) -> bool:
    t = re.sub(r"\s+", " ", (val or "")).strip()
    if not t:
        return False
    low = t.casefold()
    if low in {"true", "false", "yes", "no", "null", "none"}:
        return False
    if re.fullmatch(r"\[[^\]]+\]", t):
        return False
    if re.fullmatch(r"[-+]?\d+(\.\d+)?", t):
        return False
    if any(len(q.strip()) >= min_len for q in re.findall(r'"([^"]+)"', t)):
        return True
    if "://" in t or re.search(r"[\\/]|\w+@\w+", t):
        return True
    if len(t) >= max(8, min_len) and re.search(r"[A-Za-z]", t) and re.search(r"\d", t):
        return True
    if len(t) >= min_len and " " in t and not re.fullmatch(r"[A-Za-z0-9_ ]+", t):
        return True
    return len(t) >= 20

def find_hardcoded_literals(xml_path: str, min_len: int = 3) -> List[Dict[str, str]]:
    tree = ET.parse(xml_path)
    root = tree.getroot()
    parent_map = {c: p for p in root.iter() for c in p}

    def _context(elem: ET.Element) -> str:
        p = elem
        while p is not None:
            if _local(p.tag) == "stage":
                typ = p.attrib.get("type") or "?"
                nm = p.attrib.get("name") or p.attrib.get("Name") or "?"
                return f"Stage:{typ}/{nm}"
            p = parent_map.get(p)
        p = elem
        while p is not None:
            loc = _local(p.tag)
            if loc in {"process", "object"}:
                nm = p.attrib.get("name") or "?"
                return f"{loc.capitalize()}:{nm}"
            p = parent_map.get(p)
        return _local(elem.tag)

    hits: List[Dict[str, str]] = []
    for elem in root.iter():
        loc = _local(elem.tag)
        txt = (elem.text or "").strip()
        if txt and loc.lower() in _TEXTY_TAGS and _looks_hardcoded(txt, min_len=min_len):
            hits.append({"where": _context(elem), "tag_or_attr": f"<{loc}>", "value": txt})
        for an, av in list(elem.attrib.items()):
            if an.lower() in _TEXTY_ATTRS and av and _looks_hardcoded(av, min_len=min_len):
                hits.append({"where": _context(elem), "tag_or_attr": f"@{an}", "value": av})
    return hits


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


# ===================== formatting / borders =====================

COLOR_GREEN = (198, 239, 206)
COLOR_RED   = (255, 199, 206)
COLOR_ORANGE= (255, 235, 156)

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
        dst.Interior.Pattern = -4142
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
        # re-create single-row merges (not vertical)
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


# ====================== dynamic section model ======================

SECTION_REGISTRY = {
    "process":                   {"keyword": "process"},
    "business_object":           {"keyword": "business object"},
    "work_queue":                {"keyword": "work queue"},
    "bp_scripts_check":          {"keyword": "bp scripts check"},
    # ENV(PROD)
    "environment_variable_prod": {"keyword": "environment variables (prod)"},
}

def _row_has_keyword(row: pd.Series, phrase: str) -> bool:
    ph = _norm(phrase)
    return any(ph in _norm(c) for c in row)

def _row_has_single_letter_marker(row: pd.Series) -> bool:
    return any(_cell_is_single_letter_marker(c) for c in row)

def find_dynamic_markers(df: pd.DataFrame) -> List[Dict[str, Any]]:
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
            has_wq_name = any(("work queue name" in v) for v in tokens)
            has_key_name = any(("key name" in v) for v in tokens)
            if has_wq_name and has_key_name:
                return row_index
    return None


def pick_columns(df: pd.DataFrame, header_row: int, section_type: str) -> Dict[str, Optional[int]]:
    normed = [(_norm(v), idx) for idx, v in df.iloc[header_row].items()]

    # Validation column (reuse or next empty to the right)
    validation_col = None
    for text, idx in normed:
        if text == "validation":
            validation_col = idx; break
    if validation_col is None:
        validation_col = _last_nonempty_col_index(df.iloc[header_row]) + 1

    # No. (dot-insensitive)
    no_col = None
    for text, idx in normed:
        if text == "no" or text == "no." or text.startswith("no."):
            no_col = idx; break

    # Check (optional)
    check_col = None
    for text, idx in normed:
        if "check" in text:
            check_col = idx; break

    if section_type == "bp_scripts_check":
        return {"item_col": None, "no_col": None, "check_col": None, "validation_col": validation_col}

    if section_type in ("process", "business_object", "environment_variable_prod"):
        name_col = None
        for text, idx in normed:
            if "name" in text:
                name_col = idx; break
        return {"name_col": name_col, "no_col": no_col, "check_col": check_col, "validation_col": validation_col}

    # work_queue (phrase-based)
    wq_name_col = None
    for text, idx in normed:
        if "work queue name" in text:
            wq_name_col = idx; break
    key_col = None
    for text, idx in normed:
        if "key name" in text:
            key_col = idx; break
    if wq_name_col is None:
        for text, idx in normed:
            if text == "name":  # last-resort fallback
                wq_name_col = idx; break

    return {
        "wq_name_col": wq_name_col,
        "key_col": key_col,
        "no_col": no_col,
        "encrypted_col": None,
        "check_col": check_col,
        "validation_col": validation_col,
    }


# ======================= planning per section =======================

def _row_has_any_nonblank(series: pd.Series) -> bool:
    return any(_to_space(v).strip() != "" for v in list(series.values))

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
                last_rel_by_no = row_index; break

    last_rel_by_any = None
    for row_index in range(len(coarse) - 1, -1, -1):
        if _row_has_any_nonblank(coarse.iloc[row_index]):
            last_rel_by_any = row_index; break

    last_rel_by_name = None
    if name_col < coarse.shape[1]:
        for row_index in range(len(coarse) - 1, -1, -1):
            if _to_space(coarse.iat[row_index, name_col]).strip() != "":
                last_rel_by_name = row_index; break

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
        "type": None,
        "header_row": header_row,
        "header_excel_row": header_row + 1,
        "start_row_excel": content_start + 1,
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
    wq_name_col = cols["wq_name_col"]
    content_start = header_row + 1
    content_end = next_marker_row if next_marker_row is not None else len(df)

    coarse = df.iloc[content_start:content_end].copy().reset_index(drop=True)

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
                last_rel_by_no = row_index; break

    last_rel_by_any = None
    for row_index in range(len(coarse) - 1, -1, -1):
        if _row_has_any_nonblank(coarse.iloc[row_index]):
            last_rel_by_any = row_index; break

    last_rel_by_name = None
    if wq_name_col < coarse.shape[1]:
        for row_index in range(len(coarse) - 1, -1, -1):
            if _to_space(coarse.iat[row_index, wq_name_col]).strip() != "":
                last_rel_by_name = row_index; break

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
    xml_set = set(xml_names)
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
        "type": None,
        "header_row": header_row,
        "header_excel_row": header_row + 1,
        "start_row_excel": content_start + 1,
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
    content_start = header_row + 1
    content_end = next_marker_row if next_marker_row is not None else len(df)
    section_df = df.iloc[content_start:content_end].copy().reset_index(drop=True)

    target_phrase_norm = _norm("Is your main process published?")
    target_rel = None

    for rel_index in range(section_df.shape[0]):
        row = section_df.iloc[rel_index]
        if any(_norm(_to_space(val).strip()) == target_phrase_norm for val in list(row.values)):
            target_rel = rel_index; break

    return {
        "type": None,
        "header_row": header_row,
        "header_excel_row": header_row + 1,
        "start_row_excel": content_start + 1,
        "row_count": section_df.shape[0],
        "validation_col": cols["validation_col"],
        "target_row_rel": target_rel,
        "validation_vals": [""] * section_df.shape[0],
        "inserted_rows": 0,
    }


# ========================= writers & utilities =========================

def _unmerge_cell_if_rowspan(ws, row_1based: int, col_1based: int) -> None:
    """If the cell belongs to a MergeArea spanning multiple rows, unmerge it."""
    try:
        cell = ws.api.Cells(row_1based, col_1based)
        if cell.MergeCells:
            area = cell.MergeArea
            if int(area.Rows.Count) > 1:
                area.UnMerge()
    except Exception:
        pass

def _unmerge_verticals_in_section(ws, start_row_1based: int, end_row_1based: int, col_1based_list: List[int]) -> None:
    for r in range(start_row_1based, end_row_1based + 1):
        for c in col_1based_list:
            _unmerge_cell_if_rowspan(ws, r, c)

def _ensure_min_col_width(ws, cols: List[int], min_width: float = 18.0) -> None:
    for c in cols:
        try:
            col = ws.api.Columns(c)
            col.AutoFit()
            if float(col.ColumnWidth) < min_width:
                col.ColumnWidth = min_width
        except Exception:
            pass

def _wrap_columns_and_autofit_rows(ws, start_row_1based: int, end_row_1based: int, cols: List[int]) -> None:
    if start_row_1based > end_row_1based:
        return
    for c in cols:
        try:
            rng = ws.api.Range(ws.api.Cells(start_row_1based, c), ws.api.Cells(end_row_1based, c))
            rng.WrapText = True
        except Exception:
            pass
    try:
        ws.api.Rows(f"{start_row_1based}:{end_row_1based}").AutoFit()
    except Exception:
        pass

def _apply_alignment_like_header_for_rows(ws, header_row_1based: int, start_row_1based: int, end_row_1based: int, cols: List[int]) -> None:
    if start_row_1based > end_row_1based:
        return
    for c in cols:
        try:
            hdr = ws.api.Cells(header_row_1based, c)
            rng = ws.api.Range(ws.api.Cells(start_row_1based, c), ws.api.Cells(end_row_1based, c))
            rng.HorizontalAlignment = hdr.HorizontalAlignment
            rng.VerticalAlignment = hdr.VerticalAlignment
        except Exception:
            pass

def _header_index_map(ws, header_row_1based: int) -> Dict[str, int]:
    last_col = ws.used_range.last_cell.column
    vals = ws.range((header_row_1based, 1), (header_row_1based, last_col)).value
    if isinstance(vals, list) and len(vals) and isinstance(vals[0], list):
        vals = vals[0]
    out = {}
    for i, v in enumerate(vals or [], start=1):
        key = _norm(v)
        if key and key not in out:
            out[key] = i
    return out

def write_existing_validation(ws, start_row_excel: int, validation_col0: int, statuses: List[str]) -> None:
    validation_col_1based = validation_col0 + 1
    for offset, status in enumerate(statuses):
        row_1based = start_row_excel + offset
        _unmerge_cell_if_rowspan(ws, row_1based, validation_col_1based)
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

        for col1 in [name_col_1based, validation_col_1based, no_col_1based or 0, check_col_1based or 0]:
            if col1:
                _unmerge_cell_if_rowspan(ws, row_1based, col1)

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

        for col1 in [name_col_1based, validation_col_1based, no_col_1based or 0, check_col_1based or 0]:
            if col1:
                _unmerge_cell_if_rowspan(ws, row_1based, col1)

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

        for col1 in [name_col_1based, validation_col_1based, key_col_1based or 0, no_col_1based or 0, check_col_1based or 0]:
            if col1:
                _unmerge_cell_if_rowspan(ws, row_1based, col1)

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

        for col1 in [name_col_1based, validation_col_1based, key_col_1based or 0, no_col_1based or 0, check_col_1based or 0]:
            if col1:
                _unmerge_cell_if_rowspan(ws, row_1based, col1)

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
        if not nm or nm not in xml_wq:
            continue
        expected = _to_space(xml_wq[nm].get("key") or "").strip()
        if not expected:
            continue
        excel_key = _to_space(ws.range((row_1based, key_col_1based)).value).strip()
        if excel_key != expected:
            current = _to_space(ws.range((row_1based, validation_col_1based)).value).strip()
            if current.startswith("Exists"):
                ws.range((row_1based, validation_col_1based)).value = f'Exists (Key mismatch: expected "{expected}")'
            elif current.startswith("Newly added"):
                ws.range((row_1based, validation_col_1based)).value = f'Newly added (Key mismatch: expected "{expected}")'
            ws.range((row_1based, validation_col_1based)).color = COLOR_ORANGE
            mismatches += 1
    return mismatches


# ---------- ENV(PROD) post-write validator ----------

def adjust_env_local_validation(ws, plan: Dict[str, Any], local_data_item_names: List[str]) -> int:
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


def write_bp_scripts_check_validation(ws, plan: Dict[str, Any], all_processes_published: bool) -> None:
    if plan.get("target_row_rel") is None:
        return
    row_1based = plan["start_row_excel"] + plan["target_row_rel"]
    validation_col_1based = plan["validation_col"] + 1
    paste_formats_like_left(ws, row_1based, validation_col_1based)
    if all_processes_published:
        ws.range((row_1based, validation_col_1based)).value = "Yes"
        ws.range((row_1based, validation_col_1based)).color = COLOR_GREEN
    else:
        ws.range((row_1based, validation_col_1based)).value = "No"
        ws.range((row_1based, validation_col_1based)).color = COLOR_RED
    _apply_borders_like_left(ws, row_1based, validation_col_1based)


# ---------- Process extras (insert columns, write values) ----------

_ALLOWED_EXC = {"System Exception", "Business Exception"}

def _insert_process_extra_columns(ws, plan: Dict[str, Any]) -> None:
    """
    Physically insert 3 columns BEFORE validation, then shift 'validation_col' right by 3.
    """
    vcol_1 = plan["validation_col"] + 1
    # Insert three columns at the SAME index so they end up immediately before the old Validation
    for _ in range(3):
        try:
            ws.api.Columns(vcol_1).Insert(Shift=-4161)  # xlShiftToRight
        except Exception:
            # Fallback: use Range.EntireColumn
            ws.api.Cells(1, vcol_1).EntireColumn.Insert(Shift=-4161)

    # Update plan indices
    base0 = plan["validation_col"]
    plan["published_col"]  = base0
    plan["hardcoded_col"]  = base0 + 1
    plan["exception_col"]  = base0 + 2
    plan["validation_col"] = base0 + 3  # moved to the right

def _write_process_extra_headers(ws, plan: Dict[str, Any]) -> None:
    row = plan["header_excel_row"]
    # headers
    labels = [("Published status", plan["published_col"] + 1),
              ("Hard coded values", plan["hardcoded_col"] + 1),
              ("Exception Types",   plan["exception_col"] + 1)]
    for text, col1 in labels:
        paste_formats_like_left(ws, row, col1)
        ws.range((row, col1)).value = text
        _apply_borders_like_left(ws, row, col1)
        try:
            ws.api.Cells(row, col1).WrapText = True
        except Exception:
            pass

def write_process_extras(ws, plan: Dict[str, Any], proc_meta: Dict[str, Dict[str, Any]]) -> None:
    # Build a casefold lookup for fallback
    meta_ci = { (k or "").casefold(): v for k, v in (proc_meta or {}).items() }

    name_col_1 = plan["name_col"] + 1
    pub_col_1  = plan["published_col"] + 1
    hard_col_1 = plan["hardcoded_col"] + 1
    exc_col_1  = plan["exception_col"] + 1

    total_rows = plan["row_count"] + (plan.get("inserted_rows", 0) or 0)

    # Before writing, unmerge any vertical merges in the process section across all relevant columns
    cols_to_unmerge = [name_col_1, pub_col_1, hard_col_1, exc_col_1, plan["validation_col"] + 1]
    if plan.get("no_col") is not None:    cols_to_unmerge.append(plan["no_col"] + 1)
    if plan.get("check_col") is not None: cols_to_unmerge.append(plan["check_col"] + 1)
    _unmerge_verticals_in_section(ws, plan["start_row_excel"], plan["start_row_excel"] + total_rows - 1, cols_to_unmerge)

    for offset in range(total_rows):
        row = plan["start_row_excel"] + offset
        nm = _to_space(ws.range((row, name_col_1)).value).strip()

        for col1 in (pub_col_1, hard_col_1, exc_col_1):
            paste_formats_like_left(ws, row, col1)
            ws.range((row, col1)).value = ""
            _apply_borders_like_left(ws, row, col1)

        if not nm:
            continue

        meta = proc_meta.get(nm)
        if meta is None:
            meta = meta_ci.get(nm.casefold())

        # Published status (write literal string if present)
        pub = (meta or {}).get("published")
        if pub is not None:
            ws.range((row, pub_col_1)).value = pub

        # Hard coded values: (placeholder) left empty

        # Exception types
        raw_types = set((meta or {}).get("exception_types") or [])
        filtered = [t for t in sorted(raw_types) if "exception" in t.casefold()]
        if filtered:
            unknown = [t for t in filtered if t not in _ALLOWED_EXC]
            if unknown:
                ws.range((row, exc_col_1)).value = f"Found unknown exception types: {', '.join(unknown)}"
            else:
                ws.range((row, exc_col_1)).value = ", ".join(filtered)


# ============================ placement helpers ============================

XL_MOVE_AND_SIZE = 1
XL_MOVE          = 2
XL_FREE_FLOATING = 3

def _iter_shapes_with_placement(ws):
    try:
        shapes = ws.api.Shapes
        count = int(shapes.Count)
        for i in range(1, count + 1):
            shp = shapes.Item(i)
            name = str(getattr(shp, "Name", f"Shape{i}"))
            yield ("Shape", name, shp)
    except Exception:
        pass

    try:
        charts = ws.api.ChartObjects()
        count = int(charts.Count)
        for i in range(1, count + 1):
            co = charts.Item(i)
            name = str(getattr(co, "Name", f"Chart{i}"))
            yield ("ChartObject", name, co)
    except Exception:
        pass

    try:
        oles = ws.api.OLEObjects()
        count = int(oles.Count)
        for i in range(1, count + 1):
            ole = oles.Item(i)
            name = str(getattr(ole, "Name", f"OLE{i}"))
            yield ("OLEObject", name, ole)
    except Exception:
        pass


def snapshot_and_set_placement(ws):
    snapshots = []
    for kind, name, obj in _iter_shapes_with_placement(ws):
        try:
            placement = int(obj.Placement)
        except Exception:
            continue
        snapshots.append({"kind": kind, "name": name, "placement": placement})
        try:
            obj.Placement = XL_MOVE
        except Exception:
            pass
    return snapshots


def restore_placement(ws, snapshots):
    if not snapshots:
        return
    lookup = {(s["kind"], s["name"]): s["placement"] for s in snapshots}
    for kind, name, obj in _iter_shapes_with_placement(ws):
        key = (kind, name)
        if key in lookup:
            try:
                obj.Placement = int(lookup[key])
            except Exception:
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


def _all_processes_published(proc_meta: Dict[str, Dict[str, Any]]) -> bool:
    found_any = False
    failing = []
    for name, info in (proc_meta or {}).items():
        pub = info.get("published")
        if pub is None:
            continue
        found_any = True
        if str(pub).strip().casefold() != "true":
            failing.append((name or "(unnamed)", pub))

    if not found_any:
        print("[BP Scripts Check] No processes have a 'published' value; treating aggregate as False.")
        return False

    if failing:
        print("[BP Scripts Check] Processes with non-True 'published':")
        for nm, val in failing:
            print(f"  - {nm}: published='{val}'")
        return False

    return True


def validate_and_write_dynamic(
    excel_path: str,
    sheet_arg,
    xml_proc: List[str],
    xml_bo: List[str],
    xml_wq: Dict[str, Dict[str, Optional[str]]],
    xml_env_prod: List[str],
    xml_local_data_names: List[str],
    proc_meta: Optional[Dict[str, Dict[str, Any]]] = None,
    all_procs_published_bool: Optional[bool] = None,
) -> None:
    sheet_arg_pd = int(sheet_arg) if (isinstance(sheet_arg, str) and str(sheet_arg).isdigit()) else sheet_arg
    df = pd.read_excel(excel_path, sheet_name=sheet_arg_pd, header=None, dtype=str, engine="openpyxl").fillna("")

    # 1) Detect markers
    markers = find_dynamic_markers(df)
    if not markers:
        print("⛔ No section letter markers found. Nothing to do.")
        return

    # 2) Build plans per marker
    plans: List[Dict[str, Any]] = []
    for marker_index, marker in enumerate(markers):
        marker_row = marker["row"]
        marker_type = marker["type"]
        next_marker_row: Optional[int] = markers[marker_index + 1]["row"] if marker_index + 1 < len(markers) else None

        if marker_type == "unknown":
            continue

        header_row = find_header_after_marker(df, marker_row, marker_type, lookahead=80)
        if header_row is None:
            print(f"[{marker_type}] ❌ Could not find a header row after marker at Excel row {_excel_row(marker_row)}. Skipping.")
            continue

        if next_marker_row is not None and next_marker_row - header_row > 20:
            print(f"[{marker_type}] Next marker is {next_marker_row - header_row} rows away (>20). Treating as last to EOF.")
            next_marker_row = None

        columns = pick_columns(df, header_row, marker_type)

        if marker_type in ("process", "business_object", "environment_variable_prod"):
            if columns["name_col"] is None:
                print(f"[{marker_type}] ❌ Could not resolve a 'Name' column at Excel row {_excel_row(header_row)}.")
                continue
            plan = build_plan_process_like(
                df, header_row, columns["name_col"], columns["no_col"], columns["check_col"],
                columns["validation_col"],
                xml_proc if marker_type == "process"
                else xml_bo if marker_type == "business_object"
                else xml_env_prod,
                next_marker_row=next_marker_row
            )
        elif marker_type == "bp_scripts_check":
            plan = build_plan_bp_scripts_check(df, header_row, columns, next_marker_row=next_marker_row)
        else:  # work_queue
            if columns["wq_name_col"] is None or columns["key_col"] is None:
                print(f"[work_queue] ❌ Need 'Work Queue Name' and 'Key Name' columns. Skipping.")
                continue
            plan = build_plan_workqueue(df, header_row, columns, xml_wq, next_marker_row=next_marker_row)

        plan["type"] = marker_type
        plan["header_row"] = header_row
        plans.append(plan)

    if not plans:
        print("⛔ No valid sections planned. Nothing to write.")
        return

    # 3) Open Excel; clone; set placement; write
    app = xw.App(visible=False, add_book=False)
    wb = None
    try:
        with excel_perf_mode(app):
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

            plans.sort(key=lambda p: p["header_row"])

            total_inserted_above = 0
            bp_plans = []

            for plan in plans:
                for key_name in ("start_row_excel", "insert_at_excel_row", "header_excel_row"):
                    if key_name in plan and isinstance(plan[key_name], int):
                        plan[key_name] += total_inserted_above

                if plan["type"] == "process":
                    # Insert the 3 cols physically, then add headers
                    _insert_process_extra_columns(ws, plan)
                    _write_process_extra_headers(ws, plan)

                # Write "Validation" header
                validation_col_1based = plan["validation_col"] + 1
                paste_formats_like_left(ws, plan["header_excel_row"], validation_col_1based)
                ws.range((plan["header_excel_row"], validation_col_1based)).value = "Validation"
                _apply_borders_like_left(ws, plan["header_excel_row"], validation_col_1based)
                try:
                    ws.api.Cells(plan["header_excel_row"], validation_col_1based).WrapText = True
                except Exception:
                    pass

                # Existing rows' validation (skip for BP Scripts Check)
                if plan["type"] != "bp_scripts_check":
                    write_existing_validation(ws, plan["start_row_excel"], plan["validation_col"], plan["validation_vals"])

                rows_added_here = 0
                if plan["type"] in ("process", "business_object", "environment_variable_prod"):
                    # Unmerge vertical merges for Name/Validation columns before filling
                    _unmerge_verticals_in_section(
                        ws,
                        plan["start_row_excel"],
                        plan["start_row_excel"] + plan["row_count"] - 1,
                        [plan["name_col"] + 1, plan["validation_col"] + 1]
                    )

                    consumed_count = fill_into_blanks_name_table(ws, plan, plan["missing"])
                    remaining_names = plan["missing"][consumed_count:]
                    rows_added_here = insert_new_rows_name_table(ws, plan, remaining_names)

                    if plan["type"] == "environment_variable_prod":
                        local_hits = adjust_env_local_validation(ws, plan, xml_local_data_names)
                        if local_hits:
                            print(f"[env_prod] {local_hits} name(s) are also defined locally (Data items).")

                    if plan["type"] == "process":
                        write_process_extras(ws, plan, proc_meta or {})

                elif plan["type"] == "bp_scripts_check":
                    bp_plans.append(plan)

                else:  # work_queue
                    consumed_count = fill_into_blanks_wq(ws, plan, plan["missing"], xml_wq)
                    remaining_names = plan["missing"][consumed_count:]
                    rows_added_here = insert_new_rows_wq(ws, plan, remaining_names, xml_wq)
                    mismatch_count = adjust_wq_key_validation(ws, plan, xml_wq)
                    print(f"[work_queue] filled={consumed_count}, inserted={rows_added_here}, key_mismatches={mismatch_count}")

                total_inserted_above += rows_added_here

                # Header/values wrapping + min width for our new columns (where applicable)
                hdr_map = _header_index_map(ws, plan["header_excel_row"])
                to_wrap = []
                for label in ("published status", "hard coded values", "exception types", "validation"):
                    idx = hdr_map.get(label)
                    if idx:
                        to_wrap.append(idx)
                if to_wrap:
                    section_start = plan["start_row_excel"]
                    section_rows  = plan["row_count"] + (plan.get("inserted_rows", 0) or 0)
                    section_end   = section_start + section_rows - 1
                    _wrap_columns_and_autofit_rows(ws, section_start, section_end, to_wrap)
                    _apply_alignment_like_header_for_rows(ws, plan["header_excel_row"], section_start, section_end, to_wrap)
                    _ensure_min_col_width(ws, to_wrap, min_width=18.0)

                # AutoFit validation at least
                try:
                    ws.api.Columns(validation_col_1based).AutoFit()
                except Exception:
                    pass

            for plan in bp_plans:
                write_bp_scripts_check_validation(ws, plan, bool(all_procs_published_bool))

            restore_placement(ws, placements)
            wb.save()
            print(f"✅ Completed. Wrote '{ws.name}'.")
    finally:
        try:
            app.api.EnableEvents = True
        except Exception:
            pass
        try:
            if wb is not None:
                wb.close()
        except Exception:
            pass
        app.quit()


# =============================== CLI ===============================

def main():
    parser = argparse.ArgumentParser(
        description="Validate sections (Process/Business Object/Work Queue/Environment Variables (PROD)/BP Scripts Check) against a Blue Prism release XML."
    )
    parser.add_argument("--xml", required=True, help="Path to Blue Prism .bprelease XML")
    parser.add_argument("--excel", required=True, help="Path to Excel file (.xlsx/.xlsm)")
    parser.add_argument("--sheet", default="0", help="Sheet name or 0-based index (default: 0)")
    parser.add_argument("--hardcoded-csv", default=None, help="Optional path to write a CSV report of detected hardcoded literals")
    args = parser.parse_args()

    print("🔍 Parsing XML…")
    xml_proc = extract_names_from_xml(args.xml, "process")
    xml_bo = extract_names_from_xml(args.xml, "object")
    xml_wq = {}
    try:
        from __main__ import extract_work_queues_from_xml as _maybe_wq
        xml_wq = _maybe_wq(args.xml)
    except Exception:
        xml_wq = {}

    xml_env_prod = extract_env_variables_from_xml(args.xml)
    xml_local_data_names = extract_local_data_item_names(args.xml)

    proc_meta = get_process_metadata(args.xml)
    all_published = _all_processes_published(proc_meta)

    print(f"✅ XML: processes={len(xml_proc)}, business_objects={len(xml_bo)}, "
          f"work_queues={len(xml_wq)}, env_vars={len(xml_env_prod)}, local_data_items={len(xml_local_data_names)}")
    print(f"ℹ️ BP Scripts Check aggregate — ALL processes published: {all_published}")

    if args.hardcoded_csv is not None:
        hits = find_hardcoded_literals(args.xml, min_len=3)
        print(f"🔎 Hardcoded literal candidates: {len(hits)}")
        with open(args.hardcoded_csv, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["where", "tag_or_attr", "value"])
            w.writeheader()
            for h in hits:
                w.writerow(h)
        print(f"📄 Wrote hardcoded-literals report to: {args.hardcoded_csv}")

    print("🧪 Validating & writing…")
    validate_and_write_dynamic(
        excel_path=args.excel,
        sheet_arg=args.sheet,
        xml_proc=xml_proc,
        xml_bo=xml_bo,
        xml_wq=xml_wq,
        xml_env_prod=xml_env_prod,
        xml_local_data_names=xml_local_data_names,
        all_procs_published_bool=all_published,
        proc_meta=proc_meta,
    )


if __name__ == "__main__":
    main()
