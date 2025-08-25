#!/usr/bin/env python3
"""
Validate Excel sections against a Blue Prism release XML.

Supports:
- A. Process
- B. Business Object
- S. Work Queues (Work Queue Name + Key Name validation)

Behaviors:
- Fill existing blank Name rows first before inserting.
- Continue numbering in 'No.' where needed.
- Trim sections to the last real row.
- Write Validation text + color (green/red/orange).
- Preserve borders/formatting; set embedded objects to "move with cells" during edits and restore later.
"""

import argparse
import re
import xml.etree.ElementTree as ET
from typing import List, Optional, Dict, Any, Tuple

import pandas as pd
import xlwings as xw


# =============== XML helpers ===============

def _local(tag: str) -> str:
    """Return local tag name by stripping namespace, e.g. '{ns}object' -> 'object'."""
    return tag.split('}', 1)[1] if tag.startswith('{') else tag


def extract_names_from_xml(xml_path: str, want: str) -> List[str]:
    """
    Extract an order-preserving unique list of names from the Blue Prism release XML.
    - want: 'process' or 'object'
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    names: List[str] = []
    for elem in root.iter():
        if _local(elem.tag) == want:
            name_attr = elem.attrib.get('name')
            if name_attr:
                names.append(name_attr.strip())
    return list(dict.fromkeys(names))  # order-preserving unique


def extract_work_queues_from_xml(xml_path: str) -> Dict[str, Dict[str, Optional[str]]]:
    """
    Extract work queues as a dict: { name: {'key': <key-field or ''>} }.
    Blue Prism <work-queue ... name="…" key-field="…"> (no environment split).
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    wq: Dict[str, Dict[str, Optional[str]]] = {}
    for elem in root.iter():
        if _local(elem.tag) == "work-queue":
            name = (elem.attrib.get("name") or "").strip()
            key = (elem.attrib.get("key-field") or "").strip()
            if name and name not in wq:
                wq[name] = {"key": key}
    return wq


# =============== DataFrame parsing helpers ===============

def _norm_cell(s) -> str:
    """Normalize a cell for case/whitespace tolerant matching."""
    if s is None:
        return ""
    return str(s).strip().casefold()


def _normalize_marker_token(s: str) -> str:
    """
    Lowercase + trim; if the token is a single letter with optional '.' or ':',
    strip that trailing punctuation so 's' == 's.' == 's:'.
    """
    v = _norm_cell(s)
    # Accept forms like 'a', 'a.', 'a:'
    if re.fullmatch(r"[a-zA-Z][\.:]?", v or ""):
        return v.rstrip(".:")
    return v


def _row_contains_all(row_series: pd.Series, keywords: List[str]) -> bool:
    """
    True if all keywords appear either spread across the row cells, or together in any one cell.
    """
    row = [_norm_cell(c) for c in row_series]
    keys = [_norm_cell(k) for k in keywords]
    if all(any(k in c for c in row) for k in keys):
        return True
    return any(all(k in c for k in keys) for c in row)


def _row_contains_all_markers(row_series: pd.Series, keywords: List[str]) -> bool:
    """
    Like _row_contains_all, but with dot/colon-insensitive matching for single-letter markers (A/A., S/S., etc.).
    """
    row = [_normalize_marker_token(c) for c in row_series]
    keys = [_normalize_marker_token(k) for k in keywords]
    if all(any(k in c for c in row) for k in keys):
        return True
    return any(all(k in c for k in keys) for c in row)


def _two_line_or_same_row_match(df: pd.DataFrame, i: int, group: List[str]) -> bool:
    """
    True if row i contains all tokens in 'group' (same row), OR
    row i contains the first token and row i+1 contains the remaining tokens.
    """
    if _row_contains_all(df.iloc[i], group):
        return True
    if len(group) >= 2 and i + 1 < len(df):
        if _row_contains_all(df.iloc[i], [group[0]]) and _row_contains_all(df.iloc[i + 1], group[1:]):
            return True
    return False


def _two_line_or_same_row_match_markers(df: pd.DataFrame, i: int, group: List[str]) -> bool:
    """
    Marker-aware version (A/A., S/S., etc.).
    """
    if _row_contains_all_markers(df.iloc[i], group):
        return True
    if len(group) >= 2 and i + 1 < len(df):
        if _row_contains_all_markers(df.iloc[i], [group[0]]) and _row_contains_all_markers(df.iloc[i + 1], group[1:]):
            return True
    return False


def _last_nonempty_col_index(row_series: pd.Series) -> int:
    """Return the last column index that contains a non-empty string; -1 if none."""
    last = -1
    for col_idx, val in row_series.items():
        if str(val).strip() != "":
            last = col_idx
    return last


def find_section(
    df: pd.DataFrame,
    section_keywords: List[str],
    header_keyword: str = "name",
    next_section_groups: Optional[List[List[str]]] = None,
) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[int], Optional[int]]:
    """
    Locate a section and its coarse boundaries (will be trimmed later):
      - marker row contains all 'section_keywords' (e.g., ['A.', 'Process'] or ['S.', 'work', 'queue'])
      - header row within ~12 rows contains 'header_keyword' ('name' for A/B, 'queue' for S)
      - content rows end at first row matching one of 'next_section_groups' or EOF
    Returns (content_start, content_end_exclusive, name_col_index, header_row_index, next_section_row_index)
    """
    # 1) section start marker row (marker-aware to allow 'S' or 'S.')
    section_start_idx = None
    for i in range(len(df) - 1):
        if _row_contains_all_markers(df.iloc[i], section_keywords):
            section_start_idx = i
            break
    if section_start_idx is None:
        print(f"❌ Could not find row containing all: {section_keywords}")
        return None, None, None, None, None

    # 2) header row (widened window: up to 12 rows)
    header_row_idx = None
    scan_limit = min(section_start_idx + 12, len(df))
    for j in range(section_start_idx + 1, scan_limit):
        if _row_contains_all(df.iloc[j], [header_keyword]):
            header_row_idx = j
            break
    if header_row_idx is None:
        print(f"❌ Could not find header row containing '{header_keyword}' after the section start.")
        return None, None, None, None, None

    content_start = header_row_idx + 1

    # 3) detect Name column from header row (generic; S will override below in plan_section)
    name_col_index = None
    for col_idx, val in df.iloc[header_row_idx].items():
        if header_keyword in _norm_cell(val):
            name_col_index = col_idx
            break
    if name_col_index is None:
        print(f"❌ Could not identify the '{header_keyword}' column in the header row.")
        return None, None, None, None, None

    # 4) coarse section end (next header or EOF)
    next_section_row_idx = None
    content_end = None
    if next_section_groups:
        for k in range(content_start, len(df)):
            if any(_two_line_or_same_row_match_markers(df, k, grp) for grp in next_section_groups):
                next_section_row_idx = k
                content_end = k
                break
    if content_end is None:
        content_end = len(df)

    print(
        f"🔎 Section found: header at row {header_row_idx}, "
        f"coarse content rows {content_start}..{content_end - 1}, "
        f"Name col = {name_col_index}"
    )
    return content_start, content_end, name_col_index, header_row_idx, next_section_row_idx


def find_no_col_in_header(df: pd.DataFrame, header_row_idx: int) -> Optional[int]:
    """Locate 'No.' column index (accepts 'No', 'No.' etc.)."""
    for col_idx, val in df.iloc[header_row_idx].items():
        v = _norm_cell(val)
        if v == "no." or v == "no" or v.startswith("no."):
            return col_idx
    return None


def find_check_col_in_header(df: pd.DataFrame, header_row_idx: int) -> Optional[int]:
    """Locate 'Check' column index (case-insensitive substring)."""
    for col_idx, val in df.iloc[header_row_idx].items():
        if "check" in _norm_cell(val):
            return col_idx
    return None


def find_keyname_col_in_header(df: pd.DataFrame, header_row_idx: int) -> Optional[int]:
    """Locate 'Key Name' column index (accepts 'Key Name' or cells containing both 'key' and 'name')."""
    for col_idx, val in df.iloc[header_row_idx].items():
        v = _norm_cell(val)
        if "key name" in v or ("key" in v and "name" in v):
            return col_idx
    return None


def _find_wq_name_col_in_header(df: pd.DataFrame, header_row_idx: int) -> Optional[int]:
    """
    Prefer a header containing both 'work' and 'queue' (e.g., 'Work Queue Name').
    Else a cell that has 'name' but NOT 'key'.
    Else the first cell that has 'name'.
    """
    candidates = []
    for col_idx, val in df.iloc[header_row_idx].items():
        v = _norm_cell(val)
        candidates.append((col_idx, v))

    for col_idx, v in candidates:
        if ("work" in v) and ("queue" in v):
            return col_idx

    for col_idx, v in candidates:
        if ("name" in v) and ("key" not in v):
            return col_idx

    for col_idx, v in candidates:
        if "name" in v:
            return col_idx

    return None


def _extract_int(s: str) -> Optional[int]:
    """Extract the first integer found in a string, or None."""
    m = re.search(r"\d+", str(s))
    return int(m.group(0)) if m else None


# =============== Shapes/objects placement snapshot ===============

def snapshot_and_set_placement(ws) -> List[tuple]:
    """
    Snapshot shapes/chart/OLE placement and set them to 'Move but don't size with cells' (2).
    Returns a list of (name, collection_type, original_placement) to restore later.
    """
    records: List[tuple] = []
    for getter, coll_name in (
        (lambda: ws.api.Shapes, "Shapes"),
        (lambda: ws.api.ChartObjects(), "ChartObjects"),
        (lambda: ws.api.OLEObjects(), "OLEObjects"),
    ):
        try:
            coll = getter()
            count = coll.Count
        except Exception:
            count = 0
        for i in range(1, count + 1):
            it = coll.Item(i)
            try:
                records.append((it.Name, coll_name, it.Placement))
                it.Placement = 2  # xlMove
            except Exception:
                pass
    return records


def restore_placement(ws, records: List[tuple]) -> None:
    """Restore placement for shapes, charts, and OLE objects from a snapshot."""
    by_type = {
        "Shapes": lambda: ws.api.Shapes,
        "ChartObjects": lambda: ws.api.ChartObjects(),
        "OLEObjects": lambda: ws.api.OLEObjects(),
    }
    for name, coll_type, placement in records:
        try:
            coll = by_type[coll_type]()
            obj = coll.Item(name)
            obj.Placement = placement
        except Exception:
            pass


# =============== Styling & borders ===============

COLOR_GREEN  = (198, 239, 206)  # Exists
COLOR_RED    = (255, 199, 206)  # Does not exist
COLOR_ORANGE = (255, 235, 156)  # Newly added / Key mismatch attention

def _clone_borders(from_cell_api, to_cell_api) -> None:
    """Clone all border edges and alignment from 'from_cell_api' to 'to_cell_api'."""
    for idx in (7, 8, 9, 10, 11, 12):  # edge/inside IDs
        try:
            bsrc = from_cell_api.Borders(idx)
            bdst = to_cell_api.Borders(idx)
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


def _apply_borders_like_left(ws, row: int, col_excel: int) -> None:
    """Copy borders from the immediate left cell in the same row."""
    if col_excel <= 1:
        return
    try:
        left = ws.api.Cells(row, col_excel - 1)
        dst  = ws.api.Cells(row, col_excel)
        _clone_borders(left, dst)
    except Exception:
        pass


def paste_formats_like_left(ws, row: int, col_excel: int) -> None:
    """Paste all formats from the immediate left cell into (row, col)."""
    if col_excel <= 1:
        return
    try:
        ws.api.Cells(row, col_excel - 1).Copy()
        ws.api.Cells(row, col_excel).PasteSpecial(Paste=-4122)  # xlPasteFormats
        ws.api.Application.CutCopyMode = False
    except Exception:
        pass


def clear_fill_preserve_borders(ws, row: int, col_excel: int) -> None:
    """
    Clear only the fill color in a cell, then restore borders from the cell above.
    Used to keep 'Check' column borders when we remove green fill on new rows.
    """
    try:
        dst = ws.api.Cells(row, col_excel)
        # clear fill only
        dst.Interior.Pattern = -4142  # xlPatternNone
        dst.Interior.TintAndShade = 0
        dst.Interior.PatternTintAndShade = 0
        # restore borders from the cell above (same column)
        if row > 1:
            _clone_borders(ws.api.Cells(row - 1, col_excel), dst)
    except Exception:
        pass


def insert_row_with_style(ws, row_num: int) -> None:
    """
    Insert a new row at 'row_num' with 'format from above', then:
      - paste formats from the above row,
      - match row height,
      - re-create same-row horizontal merges (1-row merge areas).
    """
    ws.api.Rows(row_num).Insert(Shift=-4121, CopyOrigin=0)  # down, format from above
    try:
        prev = row_num - 1
        ws.api.Rows(prev).Copy()
        ws.api.Rows(row_num).PasteSpecial(Paste=-4122)  # formats
        ws.api.Rows(row_num).RowHeight = ws.api.Rows(prev).RowHeight

        # recreate same-row merges (if any)
        last_col = ws.used_range.last_cell.column
        c = 1
        while c <= last_col:
            cell_above = ws.api.Cells(prev, c)
            try:
                if cell_above.MergeCells:
                    area = cell_above.MergeArea
                    if area.Column == c and area.Rows.Count == 1:
                        left_col = area.Column
                        width = area.Columns.Count
                        ws.api.Range(ws.api.Cells(row_num, left_col),
                                     ws.api.Cells(row_num, left_col + width - 1)).Merge()
                        c += width
                        continue
            except Exception:
                pass
            c += 1
        ws.api.Application.CutCopyMode = False
    except Exception:
        pass


# =============== Planning helpers ===============

def _row_has_any_nonblank(series: pd.Series) -> bool:
    """True if any cell in the row is non-blank (after strip)."""
    return any(str(v).strip() != "" for v in list(series.values))


def plan_section(
    df: pd.DataFrame,
    section_keys: List[str],
    xml_names: List[str],
    label: str,
    header_keyword: str = "name",
    next_section_groups: Optional[List[List[str]]] = None
) -> Optional[Dict[str, Any]]:
    """
    Build a plan describing how to write Validation and where to fill/insert rows.

    TRIMMING:
      - Start from header+1, end at first next header (coarse), then trim to the last real row:
          1) prefer last row with numeric 'No.'; else
          2) last row with ANY non-blank cell; else
          3) last row with a non-blank 'Name'.

    FILLABLE ROWS:
      - Rows inside the trimmed section where 'Name' is blank but the row is still "in the table":
          * 'No.' has a number, OR
          * any other cell is non-blank.
      We will place "Newly added" names into these rows before resorting to inserts.
    """
    found = find_section(df, section_keys, header_keyword=header_keyword, next_section_groups=next_section_groups)
    content_start, content_end_coarse, name_col, header_row_idx, _ = found
    if content_start is None:
        return None

    header_row = df.iloc[header_row_idx]
    # Validation column position (reuse if already present)
    validation_col = None
    for col_idx, val in header_row.items():
        if _norm_cell(val) == "validation":
            validation_col = col_idx
            break
    if validation_col is None:
        validation_col = _last_nonempty_col_index(header_row) + 1

    # Helper columns
    no_col    = find_no_col_in_header(df, header_row_idx)
    check_col = find_check_col_in_header(df, header_row_idx)
    key_col   = find_keyname_col_in_header(df, header_row_idx)  # may be None outside 'S'

    # If this is Work Queues, re-pick the Name column robustly
    if label.lower().startswith("work queue"):
        better = _find_wq_name_col_in_header(df, header_row_idx)
        if better is not None:
            name_col = better
        if name_col is None:
            print("❌ Could not identify a Work Queue Name column in the header row.")
            return None

    # Slice the coarse section for trimming
    coarse_df = df.iloc[content_start:content_end_coarse].copy().reset_index(drop=True)

    # (A) last row with a numeric "No."
    last_rel_by_no = None
    max_no = None
    if no_col is not None and no_col < coarse_df.shape[1]:
        for i in range(len(coarse_df)):
            val = _extract_int(coarse_df.iat[i, no_col])
            if val is not None:
                max_no = val if (max_no is None or val > max_no) else max_no
        for i in range(len(coarse_df) - 1, -1, -1):
            val = _extract_int(coarse_df.iat[i, no_col])
            if val is not None:
                last_rel_by_no = i
                break

    # (B) last row with any non-blank cell
    last_rel_by_any = None
    for i in range(len(coarse_df) - 1, -1, -1):
        if _row_has_any_nonblank(coarse_df.iloc[i]):
            last_rel_by_any = i
            break

    # (C) last row with a non-blank Name
    last_rel_by_name = None
    if name_col is not None and name_col < coarse_df.shape[1]:
        for i in range(len(coarse_df) - 1, -1, -1):
            if str(coarse_df.iat[i, name_col]).strip() != "":
                last_rel_by_name = i
                break

    # Decide true last row (relative)
    if last_rel_by_no is not None:
        last_rel = last_rel_by_no
    elif last_rel_by_any is not None:
        last_rel = last_rel_by_any
    else:
        last_rel = last_rel_by_name if last_rel_by_name is not None else -1

    true_end_rel_exclusive = last_rel + 1 if last_rel >= 0 else 0
    true_end_abs_exclusive = content_start + true_end_rel_exclusive  # 0-based exclusive

    # Build the trimmed section slice
    section_df = df.iloc[content_start:true_end_abs_exclusive].copy().reset_index(drop=True)

    # Names + validation for existing (trimmed) rows
    name_series = section_df[name_col].astype(str).fillna("").str.strip()
    excel_names = name_series.tolist()
    xml_set     = {x.strip() for x in xml_names}

    validation_vals = []
    for nm in excel_names:
        if not nm:
            validation_vals.append("")         # unlabelled for truly blank Name rows (will be filled later if missing)
        elif nm in xml_set:
            validation_vals.append("Exists")
        else:
            validation_vals.append("Does not exist")

    # Missing names to place
    excel_set = {n for n in excel_names if n}
    missing   = [n for n in xml_names if n not in excel_set]

    # Fillable rows (relative indices) inside the trimmed region
    fillable_rel = []
    for i in range(section_df.shape[0]):
        nm = str(section_df.iat[i, name_col]).strip()
        if nm != "":
            continue
        row = section_df.iloc[i]
        has_no_num = (_extract_int(row[no_col]) is not None) if (no_col is not None and no_col < len(row)) else False
        has_other  = any((str(row[j]).strip() != "" and j != name_col) for j in range(len(row)))
        if has_no_num or has_other:
            fillable_rel.append(i)

    # Insert anchor: AFTER the last real row we just computed
    last_abs = content_start + (true_end_rel_exclusive - 1) if true_end_rel_exclusive > 0 else header_row_idx
    insert_at_excel_row = last_abs + 2  # Excel 1-based & "+1 after"

    # Next value for No.
    # Prefer the maximum existing No. found in the coarse region; else count existing non-blank Names.
    if max_no is not None:
        next_no = max_no + 1
    else:
        nonblank_names = sum(1 for v in excel_names if v)
        next_no = nonblank_names + 1 if no_col is not None else None

    return {
        "label": label,
        "header_row_idx": header_row_idx,
        "content_start": content_start,
        "content_end": true_end_abs_exclusive,     # TRIMMED end (exclusive)
        "name_col": name_col,
        "validation_col": validation_col,
        "no_col": no_col,
        "check_col": check_col,
        "key_col": key_col,                         # may be None outside 'S'
        "header_excel_row": header_row_idx + 1,
        "section_first_excel_row": content_start + 1,
        "section_row_count": true_end_abs_exclusive - content_start,
        "insert_at_excel_row": insert_at_excel_row,
        "excel_names": excel_names,
        "validation_vals": validation_vals,
        "missing": missing,
        "next_no": next_no,
        "fillable_rel": fillable_rel,               # relative indices (0..section_row_count-1)
        "inserted_rows": 0,
    }


def apply_row_offset(plan: Optional[Dict[str, Any]], offset_rows: int) -> Optional[Dict[str, Any]]:
    """Shift key Excel row numbers in a plan by 'offset_rows' (used after inserts above)."""
    if offset_rows == 0 or plan is None:
        return plan
    newp = dict(plan)
    for key in ("header_excel_row", "section_first_excel_row", "insert_at_excel_row"):
        newp[key] = plan[key] + offset_rows
    return newp


# =============== Workbook helpers ===============

_ILLEGAL_SHEET_CHARS = r'[:\\/?*\[\]]'

def _sanitize_sheet_name(name: str) -> str:
    """Remove illegal chars and enforce Excel 31-char limit."""
    name = re.sub(_ILLEGAL_SHEET_CHARS, "_", name).strip()
    return (name[:31] or "Sheet")


def _unique_sheet_name(wb, base: str) -> str:
    """Return a unique sheet name based on 'base'."""
    base = _sanitize_sheet_name(base)
    names = [s.name for s in wb.sheets]
    name = base
    i = 2
    while name in names:
        suffix = f" ({i})"
        keep = 31 - len(suffix)
        name = _sanitize_sheet_name(base[:max(1, keep)] + suffix)
        i += 1
    return name


def _sheet_by_name_or_index(wb, sheet_arg):
    """Return an xlwings Sheet by name or 0-based index (string digits allowed)."""
    if isinstance(sheet_arg, int) or (isinstance(sheet_arg, str) and str(sheet_arg).isdigit()):
        return wb.sheets[int(sheet_arg)]
    return wb.sheets[str(sheet_arg)]


# =============== Main workflow ===============

def _write_existing_validation(ws, plan: Dict[str, Any]) -> None:
    """Write Validation TEXT + color for existing rows inside the trimmed section."""
    val_col_excel = plan["validation_col"] + 1
    start_row     = plan["section_first_excel_row"]
    for i, status in enumerate(plan["validation_vals"]):
        r = start_row + i
        paste_formats_like_left(ws, r, val_col_excel)
        ws.range((r, val_col_excel)).value = status if status else ""
        st = (status or "").lower()
        if st == "exists":
            ws.range((r, val_col_excel)).color = COLOR_GREEN
        elif st == "does not exist":
            ws.range((r, val_col_excel)).color = COLOR_RED
        else:
            ws.range((r, val_col_excel)).color = None
        _apply_borders_like_left(ws, r, val_col_excel)


def _fill_into_existing_blanks(ws, plan: Dict[str, Any], names_to_place: List[str],
                               xml_wq: Optional[Dict[str, Dict[str, Optional[str]]]] = None) -> int:
    """
    Place as many 'names_to_place' as possible into existing blank-Name rows (fillable rows).
    For S. Work Queues, also writes Key Name if column exists and XML has it.
    Returns the count of names consumed from names_to_place.
    """
    consumed = 0
    val_col_excel = plan["validation_col"] + 1
    name_col_excel = plan["name_col"] + 1
    no_col_excel   = plan["no_col"] + 1 if plan["no_col"] is not None else None
    check_col_excel= plan["check_col"] + 1 if plan["check_col"] is not None else None
    key_col_excel  = plan["key_col"] + 1 if plan.get("key_col") is not None else None

    next_no = plan["next_no"]

    for rel in plan["fillable_rel"]:
        if consumed >= len(names_to_place):
            break
        r = plan["section_first_excel_row"] + rel

        # Write Name
        nm = names_to_place[consumed]
        ws.range((r, name_col_excel)).value = nm

        # If this is S. Work Queues, also write Key Name
        if key_col_excel is not None and xml_wq is not None:
            key_val = (xml_wq.get(nm, {}) or {}).get("key") or ""
            if key_val:
                ws.range((r, key_col_excel)).value = key_val

        # Validation formatting/text
        paste_formats_like_left(ws, r, val_col_excel)
        ws.range((r, val_col_excel)).value = "Newly added"
        ws.range((r, val_col_excel)).color = COLOR_ORANGE
        _apply_borders_like_left(ws, r, val_col_excel)

        # No.: if blank, assign next_no; if already has a number, keep it and bump the cursor if needed
        if no_col_excel is not None:
            raw = str(ws.range((r, no_col_excel)).value or "").strip()
            cur = _extract_int(raw)
            if cur is None and next_no is not None:
                ws.range((r, no_col_excel)).value = next_no
                next_no += 1
            elif cur is not None and next_no is not None:
                next_no = max(next_no, cur + 1)

        # Check: clear fill only, keep borders
        if check_col_excel is not None:
            clear_fill_preserve_borders(ws, r, check_col_excel)

        consumed += 1

    # Update next_no back in plan for subsequent inserts
    plan["next_no"] = next_no
    return consumed


def _insert_new_rows(ws, plan: Dict[str, Any], remaining: List[str],
                     xml_wq: Optional[Dict[str, Dict[str, Optional[str]]]] = None) -> int:
    """
    Insert rows after the last real row and place remaining names there.
    For S. Work Queues, also writes Key Name if column exists and XML has it.
    Returns the number of rows inserted.
    """
    if not remaining:
        return 0

    val_col_excel   = plan["validation_col"] + 1
    name_col_excel  = plan["name_col"] + 1
    no_col_excel    = plan["no_col"] + 1 if plan["no_col"] is not None else None
    check_col_excel = plan["check_col"] + 1 if plan["check_col"] is not None else None
    key_col_excel   = plan["key_col"] + 1 if plan.get("key_col") is not None else None

    next_no = plan["next_no"]

    for i, nm in enumerate(remaining):
        r = plan["insert_at_excel_row"] + i
        insert_row_with_style(ws, r)
        ws.range((r, name_col_excel)).value = nm

        # If this is S. Work Queues, also write Key Name
        if key_col_excel is not None and xml_wq is not None:
            key_val = (xml_wq.get(nm, {}) or {}).get("key") or ""
            if key_val:
                ws.range((r, key_col_excel)).value = key_val

        paste_formats_like_left(ws, r, val_col_excel)
        ws.range((r, val_col_excel)).value = "Newly added"
        ws.range((r, val_col_excel)).color = COLOR_ORANGE
        _apply_borders_like_left(ws, r, val_col_excel)

        if no_col_excel is not None and next_no is not None:
            ws.range((r, no_col_excel)).value = next_no
            next_no += 1

        if check_col_excel is not None:
            clear_fill_preserve_borders(ws, r, check_col_excel)

    plan["next_no"] = next_no
    inserted = len(remaining)
    plan["inserted_rows"] = (plan.get("inserted_rows", 0) or 0) + inserted
    return inserted


def _adjust_wq_key_validation(ws, plan: Dict[str, Any], xml_wq: Dict[str, Dict[str, Optional[str]]]) -> None:
    """
    For S. Work Queues:
    - For rows with 'Exists' or 'Newly added', compare Excel 'Key Name' to XML 'key-field'.
    - If mismatch, annotate validation text and set ORANGE; if match, keep existing color/text.
    Includes newly inserted rows in this pass.
    """
    key_col = plan.get("key_col")
    if key_col is None:
        return

    name_col_excel = plan["name_col"] + 1
    key_col_excel  = key_col + 1
    val_col_excel  = plan["validation_col"] + 1

    total_rows_to_check = plan["section_row_count"] + (plan.get("inserted_rows", 0) or 0)

    for i in range(total_rows_to_check):
        r = plan["section_first_excel_row"] + i
        nm  = (ws.range((r, name_col_excel)).value or "").strip()
        if not nm or nm not in xml_wq:
            continue

        expected_key = (xml_wq[nm].get("key") or "").strip()
        if not expected_key:
            continue  # nothing to compare

        excel_key = (ws.range((r, key_col_excel)).value or "").strip()

        if excel_key.strip().casefold() != expected_key.strip().casefold():
            current = (ws.range((r, val_col_excel)).value or "").strip()
            if current.lower().startswith("exists"):
                ws.range((r, val_col_excel)).value = f'Exists (Key mismatch: expected "{expected_key}")'
            elif current.lower().startswith("newly added"):
                ws.range((r, val_col_excel)).value = f'Newly added (Key mismatch: expected "{expected_key}")'
            ws.range((r, val_col_excel)).color = COLOR_ORANGE


def validate_and_write_both(
    excel_path: str,
    sheet_arg,
    xml_process_names: List[str],
    xml_object_names: List[str],
    xml_work_queues: Dict[str, Dict[str, Optional[str]]],  # name -> {'key': ...}
) -> None:
    """
    End-to-end:
    - Read sheet into DataFrame to compute plans for A, B, S
    - Duplicate the sheet safely and rename to '<orig>_validated'
    - Write Validation cells (format like left, then color + TEXT)
    - Fill existing blank rows first; then insert new rows for any remaining names
    - For S: also fill 'Key Name' for newly added, and flag key mismatches (including inserted rows)
    - Preserve objects/placement and autofit Validation column
    """
    # Read once for planning
    sheet_arg_for_pd = int(sheet_arg) if (isinstance(sheet_arg, str) and str(sheet_arg).isdigit()) else sheet_arg
    df = pd.read_excel(excel_path, sheet_name=sheet_arg_for_pd, header=None, dtype=str, engine='openpyxl').fillna('')

    # Define section boundaries:
    # A ends at B; B ends at C/D/E; S likely near the end → until EOF
    next_for_A = [["B.", "Business Object"]]
    next_for_B = [
        ["C.", "Environment Variables"],
        ["C",  "Environment Variables"],
        ["C.", "environment", "variable"],
        ["D.", "environment", "variable"],
        ["D",  "environment", "variable"],
        ["E.", "Startup Parameters"],
        ["E",  "Startup Parameters"],
    ]
    next_for_S = None  # let it run to EOF (adjust if you have sections after S)

    plan_A = plan_section(df, ["A.", "Process"], xml_process_names, label="Process",
                          header_keyword="name", next_section_groups=next_for_A)
    plan_B = plan_section(df, ["B.", "Business Object"], xml_object_names, label="Business Object",
                          header_keyword="name", next_section_groups=next_for_B)
    # Work Queues — header_keyword='queue' (easier to find “Work Queue Name” rows)
    xml_wq_names = list(xml_work_queues.keys())
    plan_S = plan_section(df, ["S.", "work", "queue"], xml_wq_names, label="Work Queues",
                          header_keyword="queue", next_section_groups=next_for_S)

    if plan_A is None and plan_B is None and plan_S is None:
        print("⛔ Stopping: sections A, B, S were not detected.")
        return

    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(str(excel_path))
        src_sht = _sheet_by_name_or_index(wb, sheet_arg)

        # Safe copy/rename
        try:
            app.api.EnableEvents = False
        except Exception:
            pass
        before = [s.name for s in wb.sheets]
        src_sht.api.Copy(After=src_sht.api)
        after = [s.name for s in wb.sheets]
        added = [n for n in after if n not in before]
        ws = wb.sheets[added[0]] if len(added) == 1 else wb.sheets[-1]

        new_name = _unique_sheet_name(wb, f"{src_sht.name}_validated")
        try:
            ws.name = new_name
        except Exception as e:
            print(f"⚠️ Rename failed ({e}); keeping '{ws.name}'")

        placements = snapshot_and_set_placement(ws)

        rows_inserted_A = 0
        rows_inserted_B = 0

        # ----- Section A: Process -----
        if plan_A is not None:
            print("▶ Processing section A (Process)")
            val_col_excel = plan_A["validation_col"] + 1

            paste_formats_like_left(ws, plan_A["header_excel_row"], val_col_excel)
            ws.range((plan_A["header_excel_row"], val_col_excel)).value = "Validation"
            _apply_borders_like_left(ws, plan_A["header_excel_row"], val_col_excel)

            _write_existing_validation(ws, plan_A)

            consumed = _fill_into_existing_blanks(ws, plan_A, plan_A["missing"])
            remaining = plan_A["missing"][consumed:]
            rows_inserted_A = _insert_new_rows(ws, plan_A, remaining)

            try:
                ws.api.Columns(val_col_excel).AutoFit()
            except Exception:
                pass

        # ----- Section B: Business Object -----
        if plan_B is not None:
            print("▶ Processing section B (Business Object)")
            if plan_A is not None and plan_B["header_row_idx"] > plan_A["header_row_idx"]:
                plan_B = apply_row_offset(plan_B, rows_inserted_A)

            val_col_excel = plan_B["validation_col"] + 1

            paste_formats_like_left(ws, plan_B["header_excel_row"], val_col_excel)
            ws.range((plan_B["header_excel_row"], val_col_excel)).value = "Validation"
            _apply_borders_like_left(ws, plan_B["header_excel_row"], val_col_excel)

            _write_existing_validation(ws, plan_B)

            consumed = _fill_into_existing_blanks(ws, plan_B, plan_B["missing"])
            remaining = plan_B["missing"][consumed:]
            rows_inserted_B = _insert_new_rows(ws, plan_B, remaining)

            try:
                ws.api.Columns(val_col_excel).AutoFit()
            except Exception:
                pass

        # ----- Section S: Work Queues -----
        if plan_S is not None:
            print("▶ Processing section S (Work Queues)")
            # account for rows inserted above
            offset = 0
            if plan_A is not None and plan_S["header_row_idx"] > plan_A["header_row_idx"]:
                offset += rows_inserted_A
            if plan_B is not None and plan_S["header_row_idx"] > plan_B["header_row_idx"]:
                offset += rows_inserted_B
            if offset:
                plan_S = apply_row_offset(plan_S, offset)

            val_col_excel = plan_S["validation_col"] + 1
            paste_formats_like_left(ws, plan_S["header_excel_row"], val_col_excel)
            ws.range((plan_S["header_excel_row"], val_col_excel)).value = "Validation"
            _apply_borders_like_left(ws, plan_S["header_excel_row"], val_col_excel)

            # Existing rows' validation (name-based)
            _write_existing_validation(ws, plan_S)

            # Fill blanks first (write Name + Key if available), then insert remaining
            consumed = _fill_into_existing_blanks(ws, plan_S, plan_S["missing"], xml_work_queues)
            remaining = plan_S["missing"][consumed:]
            _insert_new_rows(ws, plan_S, remaining, xml_work_queues)

            # Post-pass: verify Key Name vs XML (includes inserted rows)
            _adjust_wq_key_validation(ws, plan_S, xml_work_queues)

            try:
                ws.api.Columns(val_col_excel).AutoFit()
            except Exception:
                pass

        restore_placement(ws, placements)
        wb.save()
        print(f"✅ Completed. Filled blank rows first, then inserted as needed. Wrote '{ws.name}'.")
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


# =============== CLI ===============

def main():
    parser = argparse.ArgumentParser(
        description="Validate 'A. Process', 'B. Business Object', and 'S. Work Queues' against a Blue Prism release XML. "
                    "Fills blank rows first; preserves styling/objects; adds colored 'Validation'."
    )
    parser.add_argument('--xml', required=True, help="Path to Blue Prism .bprelease XML")
    parser.add_argument('--excel', required=True, help="Path to Excel file (.xlsx/.xlsm)")
    parser.add_argument('--sheet', default="0", help="Sheet name or 0-based index (default: 0)")
    args = parser.parse_args()

    print("🔍 Parsing XML…")
    xml_process_names = extract_names_from_xml(args.xml, "process")
    xml_object_names  = extract_names_from_xml(args.xml, "object")
    xml_work_queues   = extract_work_queues_from_xml(args.xml)
    print(f"✅ Found {len(xml_process_names)} process names; {len(xml_object_names)} object names; "
          f"{len(xml_work_queues)} work queues.")

    print("🧪 Validating & writing…")
    validate_and_write_both(
        excel_path=args.excel,
        sheet_arg=args.sheet,
        xml_process_names=xml_process_names,
        xml_object_names=xml_object_names,
        xml_work_queues=xml_work_queues,
    )


if __name__ == "__main__":
    main()
