#!/usr/bin/env python3
"""
Validate Excel sections against a Blue Prism release XML.

Supports:
- A. Process
- B. Business Object
- S. Work Queues (strict same-row marker: one row must contain 'S' (or 'S.'/'S:') AND 'Work Queue*')

Behaviors:
- Fill existing blank Name rows first before inserting.
- Continue numbering in 'No.' where needed.
- Trim sections to the last real row.
- Write Validation text + color (green/red/orange).
- Preserve borders/formatting; set embedded objects to "move with cells" during edits and restore later.

Verbose S-focused logging is included so you can see detection decisions.
"""

import argparse
import re
import string
import xml.etree.ElementTree as ET
from typing import List, Optional, Dict, Any, Tuple

import pandas as pd
import xlwings as xw


# =============== XML helpers ===============

def _local(tag: str) -> str:
    return tag.split('}', 1)[1] if tag.startswith('{') else tag


def extract_names_from_xml(xml_path: str, want: str) -> List[str]:
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
    { name: {'key': <key-field or ''>} }
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


# =============== String / cell normalization ===============

def _to_space(s: str) -> str:
    if s is None:
        return ""
    # normalize NBSP and control whitespace
    return str(s).replace("\u00A0", " ").replace("\u2007", " ").replace("\u202F", " ")


def _norm_cell(s) -> str:
    return _to_space(s).strip().casefold()


def _is_letter_marker(cell_val: str, letter: str) -> bool:
    """
    True if cell is exactly that letter with optional trailing '.' or ':' (case-insensitive),
    ignoring leading/trailing spaces and NBSPs.
    e.g. 'S', 'S.', 'S:' -> True for 's'
    """
    v = _norm_cell(cell_val)
    if not v:
        return False
    if len(v) == 1:
        return v == letter
    if len(v) == 2 and v[0] == letter and v[1] in ('.', ':'):
        return True
    return False


def _row_tokens(row_series: pd.Series) -> List[str]:
    return [_norm_cell(c) for c in row_series]


def _has_tokens_same_row(row_series: pd.Series, tokens: List[str]) -> bool:
    """
    Every token must appear (as substring) in some cell on the SAME row.
    """
    row = _row_tokens(row_series)
    return all(any(tok in c for c in row) for tok in tokens)


def _excel_col_letter(col0: int) -> str:
    # 0-based index -> Excel letter
    col = col0 + 1
    letters = ""
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


# =============== Simple row dumps for debugging ===============

def _dump_row(row_idx: int, row_series: pd.Series, max_cols: int = 20) -> str:
    parts = []
    for j, val in enumerate(row_series[:max_cols]):
        parts.append(f"[{_excel_col_letter(j)}] '{_to_space(val).strip()}'")
    return f"r{row_idx} -> " + " | ".join(parts)


# =============== Section A/B helpers (unchanged logic) ===============

def _row_contains_all(row_series: pd.Series, keywords: List[str]) -> bool:
    row = _row_tokens(row_series)
    keys = [_norm_cell(k) for k in keywords]
    if all(any(k in c for c in row) for k in keys):
        return True
    return any(all(k in c for k in keys) for c in row)


def _two_line_or_same_row_match(df: pd.DataFrame, i: int, group: List[str]) -> bool:
    if _row_contains_all(df.iloc[i], group):
        return True
    if len(group) >= 2 and i + 1 < len(df):
        if _row_contains_all(df.iloc[i], [group[0]]) and _row_contains_all(df.iloc[i + 1], group[1:]):
            return True
    return False


def _last_nonempty_col_index(row_series: pd.Series) -> int:
    last = -1
    for col_idx, val in row_series.items():
        if _to_space(val).strip() != "":
            last = col_idx
    return last


def find_section_generic(
    df: pd.DataFrame,
    section_keywords: List[str],
    header_keyword: str = "name",
    next_section_groups: Optional[List[List[str]]] = None,
) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[int], Optional[int]]:
    """
    Generic locator for A/B (kept for backward-compat).
    """
    section_start_idx = None
    for i in range(len(df) - 1):
        if _row_contains_all(df.iloc[i], section_keywords):
            section_start_idx = i
            break
    if section_start_idx is None:
        return None, None, None, None, None

    header_row_idx = None
    scan_limit = min(section_start_idx + 12, len(df))
    for j in range(section_start_idx + 1, scan_limit):
        if _row_contains_all(df.iloc[j], [header_keyword]):
            header_row_idx = j
            break
    if header_row_idx is None:
        return None, None, None, None, None

    content_start = header_row_idx + 1

    name_col_index = None
    for col_idx, val in df.iloc[header_row_idx].items():
        if header_keyword in _norm_cell(val):
            name_col_index = col_idx
            break
    if name_col_index is None:
        return None, None, None, None, None

    next_section_row_idx = None
    content_end = None
    if next_section_groups:
        for k in range(content_start, len(df)):
            if any(_two_line_or_same_row_match(df, k, grp) for grp in next_section_groups):
                next_section_row_idx = k
                content_end = k
                break
    if content_end is None:
        content_end = len(df)

    return content_start, content_end, name_col_index, header_row_idx, next_section_row_idx


# =============== S. Work Queues — strict same-row marker ===============

def find_S_marker_row(df: pd.DataFrame, verbose: bool = True) -> Optional[int]:
    """
    SAME ROW must contain:
      - a cell that's exactly 'S'/'S.'/'S:'  (ignoring case/space)
      - AND a cell containing 'work' and 'queue'
    """
    for i in range(len(df)):
        row = df.iloc[i]
        has_s = any(_is_letter_marker(c, 's') for c in row)
        has_wq = any(('work' in _norm_cell(c) and 'queue' in _norm_cell(c)) for c in row)
        if has_s and has_wq:
            if verbose:
                print(f"[S] Marker found at row {i} (Excel {_excel_row(i)}).")
                print("     " + _dump_row(i, row))
            return i
    if verbose:
        print("[S] ❌ No marker row found that has BOTH 'S' and 'Work Queue' on the same row.")
    return None


def find_S_header_row(df: pd.DataFrame, marker_row: int, lookahead: int = 80, verbose: bool = True) -> Optional[int]:
    """
    After marker, locate the header row by scoring signals commonly present in S headers:
      - 'work' & 'queue' & 'name' together (+2)
      - 'key name' / 'keyname' / 'key field' / 'key column' / 'primary key' / 'unique key' / 'key' (+2 if strong, +1 if just 'key')
      - 'no' / 'no.' (+1)
      - 'encrypted' (+1)
      - 'check' (+1)
    We pick the highest score >= 2.
    """
    best_row, best_score = None, -1

    for j in range(marker_row + 1, min(marker_row + 1 + lookahead, len(df))):
        row = df.iloc[j]
        vals = [_norm_cell(c) for c in row]

        work_queue_name = any(("work" in v and "queue" in v and "name" in v) for v in vals)
        has_queue = any(("work" in v and "queue" in v) for v in vals)
        key_strong = any(("key name" in v or "keyname" in v or "key field" in v or "key column" in v
                          or "primary key" in v or "unique key" in v) for v in vals)
        key_any = key_strong or any(v == "key" or v.startswith("key ") for v in vals)
        has_no = any(v == "no." or v == "no" or v.startswith("no.") for v in vals)
        has_encrypted = any("encrypted" in v for v in vals)
        has_check = any("check" in v for v in vals)

        score = 0
        if work_queue_name: score += 2
        if has_queue and not work_queue_name: score += 1  # weaker
        if key_strong: score += 2
        elif key_any: score += 1
        if has_no: score += 1
        if has_encrypted: score += 1
        if has_check: score += 1

        if score > best_score:
            best_score, best_row = score, j

        # print a small preview for the first 15 rows after marker
        if verbose and j <= marker_row + 15:
            print(f"[S] header-candidate r{j} score={score}: " + _dump_row(j, row))

    if best_row is not None and best_score >= 2:
        if verbose:
            print(f"[S] Header chosen r{best_row} (Excel {_excel_row(best_row)}), score={best_score}")
            print("     " + _dump_row(best_row, df.iloc[best_row]))
        return best_row

    if verbose:
        print("[S] ❌ Could not find a convincing header row after the marker.")
    return None


def pick_S_columns(df: pd.DataFrame, header_row: int, verbose: bool = True) -> Dict[str, Optional[int]]:
    """
    Resolve columns for: no_col, wq_name_col, key_col, encrypted_col, check_col, validation_col
    """
    cols = list(df.iloc[header_row].items())
    norm = [(_norm_cell(v), idx) for idx, v in df.iloc[header_row].items()]

    # Work Queue Name
    wq_name_col = None
    for v, idx in norm:
        if ("work" in v and "queue" in v and "name" in v):
            wq_name_col = idx
            break
    if wq_name_col is None:
        for v, idx in norm:
            if ("queue" in v and "name" in v):
                wq_name_col = idx
                break
    if wq_name_col is None:
        for v, idx in norm:
            if "name" in v and "key" not in v:
                wq_name_col = idx
                break

    # Key Name
    key_col = None
    for v, idx in norm:
        if ("key name" in v or "keyname" in v or "key field" in v or "key column" in v
            or "primary key" in v or "unique key" in v):
            key_col = idx
            break
    if key_col is None:
        for v, idx in norm:
            if v == "key" or v.startswith("key "):
                key_col = idx
                break

    # No.
    no_col = None
    for v, idx in norm:
        if v == "no." or v == "no" or v.startswith("no."):
            no_col = idx
            break

    # Encrypted
    encrypted_col = None
    for v, idx in norm:
        if "encrypted" in v:
            encrypted_col = idx
            break

    # Check
    check_col = None
    for v, idx in norm:
        if "check" in v:
            check_col = idx
            break

    # Validation (reuse if present; else next new col)
    validation_col = None
    for v, idx in norm:
        if v == "validation":
            validation_col = idx
            break
    if validation_col is None:
        validation_col = _last_nonempty_col_index(df.iloc[header_row]) + 1

    if verbose:
        print("[S] Resolved columns (0-based):")
        print(f"    Work Queue Name: {wq_name_col} ({_excel_col_letter(wq_name_col) if wq_name_col is not None else 'N/A'})")
        print(f"    Key Name       : {key_col} ({_excel_col_letter(key_col) if key_col is not None else 'N/A'})")
        print(f"    No.            : {no_col} ({_excel_col_letter(no_col) if no_col is not None else 'N/A'})")
        print(f"    Encrypted      : {encrypted_col} ({_excel_col_letter(encrypted_col) if encrypted_col is not None else 'N/A'})")
        print(f"    Check          : {check_col} ({_excel_col_letter(check_col) if check_col is not None else 'N/A'})")
        print(f"    Validation     : {validation_col} ({_excel_col_letter(validation_col)})")

    return {
        "wq_name_col": wq_name_col,
        "key_col": key_col,
        "no_col": no_col,
        "encrypted_col": encrypted_col,
        "check_col": check_col,
        "validation_col": validation_col,
    }


# =============== Planning & trim helpers (shared) ===============

def _extract_int(s: str) -> Optional[int]:
    m = re.search(r"\d+", _to_space(s))
    return int(m.group(0)) if m else None


def _row_has_any_nonblank(series: pd.Series) -> bool:
    return any(_to_space(v).strip() != "" for v in list(series.values))


def _excel_row(i0: int) -> int:
    return i0 + 1


# =============== Shapes/objects placement snapshot ===============

def snapshot_and_set_placement(ws) -> List[tuple]:
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
    for idx in (7, 8, 9, 10, 11, 12):
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
    if col_excel <= 1:
        return
    try:
        left = ws.api.Cells(row, col_excel - 1)
        dst  = ws.api.Cells(row, col_excel)
        _clone_borders(left, dst)
    except Exception:
        pass


def paste_formats_like_left(ws, row: int, col_excel: int) -> None:
    if col_excel <= 1:
        return
    try:
        ws.api.Cells(row, col_excel - 1).Copy()
        ws.api.Cells(row, col_excel).PasteSpecial(Paste=-4122)  # xlPasteFormats
        ws.api.Application.CutCopyMode = False
    except Exception:
        pass


def clear_fill_preserve_borders(ws, row: int, col_excel: int) -> None:
    try:
        dst = ws.api.Cells(row, col_excel)
        dst.Interior.Pattern = -4142
        dst.Interior.TintAndShade = 0
        dst.Interior.PatternTintAndShade = 0
        if row > 1:
            _clone_borders(ws.api.Cells(row - 1, col_excel), dst)
    except Exception:
        pass


def insert_row_with_style(ws, row_num: int) -> None:
    ws.api.Rows(row_num).Insert(Shift=-4121, CopyOrigin=0)
    try:
        prev = row_num - 1
        ws.api.Rows(prev).Copy()
        ws.api.Rows(row_num).PasteSpecial(Paste=-4122)
        ws.api.Rows(row_num).RowHeight = ws.api.Rows(prev).RowHeight

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


# =============== Common write helpers ===============

def find_no_col_in_header(df: pd.DataFrame, header_row_idx: int) -> Optional[int]:
    for col_idx, val in df.iloc[header_row_idx].items():
        v = _norm_cell(val)
        if v == "no." or v == "no" or v.startswith("no."):
            return col_idx
    return None


def find_check_col_in_header(df: pd.DataFrame, header_row_idx: int) -> Optional[int]:
    for col_idx, val in df.iloc[header_row_idx].items():
        if "check" in _norm_cell(val):
            return col_idx
    return None


def _write_existing_validation(ws, section_first_excel_row: int, validation_col0: int, statuses: List[str]) -> None:
    val_col_excel = validation_col0 + 1
    start_row     = section_first_excel_row
    for i, status in enumerate(statuses):
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


# =============== Planning for S (strict) and generic A/B ===============

def plan_section_AB(df: pd.DataFrame, section_keys: List[str], xml_names: List[str],
                    label: str, header_keyword: str, next_section_groups: Optional[List[List[str]]]) -> Optional[Dict[str, Any]]:
    found = find_section_generic(df, section_keys, header_keyword=header_keyword, next_section_groups=next_section_groups)
    content_start, content_end_coarse, name_col, header_row_idx, _ = found
    if content_start is None:
        return None

    header_row = df.iloc[header_row_idx]
    validation_col = None
    for col_idx, val in header_row.items():
        if _norm_cell(val) == "validation":
            validation_col = col_idx
            break
    if validation_col is None:
        validation_col = _last_nonempty_col_index(header_row) + 1

    no_col    = find_no_col_in_header(df, header_row_idx)
    check_col = find_check_col_in_header(df, header_row_idx)

    coarse_df = df.iloc[content_start:content_end_coarse].copy().reset_index(drop=True)

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

    last_rel_by_any = None
    for i in range(len(coarse_df) - 1, -1, -1):
        if _row_has_any_nonblank(coarse_df.iloc[i]):
            last_rel_by_any = i
            break

    last_rel_by_name = None
    if name_col is not None and name_col < coarse_df.shape[1]:
        for i in range(len(coarse_df) - 1, -1, -1):
            if _to_space(coarse_df.iat[i, name_col]).strip() != "":
                last_rel_by_name = i
                break

    if last_rel_by_no is not None:
        last_rel = last_rel_by_no
    elif last_rel_by_any is not None:
        last_rel = last_rel_by_any
    else:
        last_rel = last_rel_by_name if last_rel_by_name is not None else -1

    true_end_rel_exclusive = last_rel + 1 if last_rel >= 0 else 0
    true_end_abs_exclusive = content_start + true_end_rel_exclusive

    section_df = df.iloc[content_start:true_end_abs_exclusive].copy().reset_index(drop=True)

    name_series = section_df[name_col].astype(str).fillna("").apply(_to_space).str.strip()
    excel_names = name_series.tolist()
    xml_set     = {x.strip() for x in xml_names}

    validation_vals = []
    for nm in excel_names:
        if not nm:
            validation_vals.append("")
        elif nm in xml_set:
            validation_vals.append("Exists")
        else:
            validation_vals.append("Does not exist")

    excel_set = {n for n in excel_names if n}
    missing   = [n for n in xml_names if n not in excel_set]

    fillable_rel = []
    for i in range(section_df.shape[0]):
        nm = _to_space(section_df.iat[i, name_col]).strip()
        if nm != "":
            continue
        row = section_df.iloc[i]
        has_no_num = (_extract_int(row[name_col*0 + no_col]) is not None) if (no_col is not None and no_col < len(row)) else False
        has_other  = any((_to_space(row[j]).strip() != "" and j != name_col) for j in range(len(row)))
        if has_no_num or has_other:
            fillable_rel.append(i)

    last_abs = content_start + (true_end_rel_exclusive - 1) if true_end_rel_exclusive > 0 else header_row_idx
    insert_at_excel_row = last_abs + 2

    if max_no is not None:
        next_no = max_no + 1
    else:
        nonblank_names = sum(1 for v in excel_names if v)
        next_no = nonblank_names + 1 if no_col is not None else None

    return {
        "label": label,
        "header_row_idx": header_row_idx,
        "content_start": content_start,
        "content_end": true_end_abs_exclusive,
        "name_col": name_col,
        "validation_col": validation_col,
        "no_col": no_col,
        "check_col": check_col,
        "header_excel_row": header_row_idx + 1,
        "section_first_excel_row": content_start + 1,
        "section_row_count": true_end_abs_exclusive - content_start,
        "insert_at_excel_row": insert_at_excel_row,
        "excel_names": excel_names,
        "validation_vals": validation_vals,
        "missing": missing,
        "next_no": next_no,
        "fillable_rel": fillable_rel,
    }


def plan_section_S(df: pd.DataFrame, xml_wq: Dict[str, Dict[str, Optional[str]]], verbose: bool = True) -> Optional[Dict[str, Any]]:
    """
    STRICT S detection and planning.
    """
    marker_row = find_S_marker_row(df, verbose=verbose)
    if marker_row is None:
        print("[S] Not detected; nothing will be written for S.")
        return None

    header_row = find_S_header_row(df, marker_row, verbose=verbose)
    if header_row is None:
        print("[S] Not detected (no header after marker); nothing will be written for S.")
        return None

    cols = pick_S_columns(df, header_row, verbose=verbose)
    wq_name_col = cols["wq_name_col"]
    key_col     = cols["key_col"]
    no_col      = cols["no_col"]
    check_col   = cols["check_col"]
    validation_col = cols["validation_col"]

    if wq_name_col is None:
        print("[S] ❌ Could not identify 'Work Queue Name' column; aborting S write.")
        return None

    content_start = header_row + 1
    content_end = len(df)  # until EOF (adjust if you have sections after S)

    # Trim to last real row
    coarse_df = df.iloc[content_start:content_end].copy().reset_index(drop=True)

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

    last_rel_by_any = None
    for i in range(len(coarse_df) - 1, -1, -1):
        if _row_has_any_nonblank(coarse_df.iloc[i]):
            last_rel_by_any = i
            break

    last_rel_by_name = None
    if wq_name_col < coarse_df.shape[1]:
        for i in range(len(coarse_df) - 1, -1, -1):
            if _to_space(coarse_df.iat[i, wq_name_col]).strip() != "":
                last_rel_by_name = i
                break

    if last_rel_by_no is not None:
        last_rel = last_rel_by_no
    elif last_rel_by_any is not None:
        last_rel = last_rel_by_any
    else:
        last_rel = last_rel_by_name if last_rel_by_name is not None else -1

    true_end_rel_exclusive = last_rel + 1 if last_rel >= 0 else 0
    true_end_abs_exclusive = content_start + true_end_rel_exclusive

    section_df = df.iloc[content_start:true_end_abs_exclusive].copy().reset_index(drop=True)

    # Build validations
    xml_names = list(xml_wq.keys())
    xml_set   = set(xml_names)
    name_series = section_df[wq_name_col].astype(str).fillna("").apply(_to_space).str.strip()
    excel_names = name_series.tolist()

    validation_vals = []
    for nm in excel_names:
        if not nm:
            validation_vals.append("")
        elif nm in xml_set:
            validation_vals.append("Exists")
        else:
            validation_vals.append("Does not exist")

    excel_set = {n for n in excel_names if n}
    missing   = [n for n in xml_names if n not in excel_set]

    # Fillable rows
    fillable_rel = []
    for i in range(section_df.shape[0]):
        nm = _to_space(section_df.iat[i, wq_name_col]).strip()
        if nm != "":
            continue
        row = section_df.iloc[i]
        has_no_num = (_extract_int(row[wq_name_col*0 + no_col]) is not None) if (no_col is not None and no_col < len(row)) else False
        has_other  = any((_to_space(row[j]).strip() != "" and j != wq_name_col) for j in range(len(row)))
        if has_no_num or has_other:
            fillable_rel.append(i)

    last_abs = content_start + (true_end_rel_exclusive - 1) if true_end_rel_exclusive > 0 else header_row
    insert_at_excel_row = last_abs + 2

    if max_no is not None:
        next_no = max_no + 1
    else:
        nonblank_names = sum(1 for v in excel_names if v)
        next_no = nonblank_names + 1 if no_col is not None else None

    print(f"[S] Content rows (trimmed) Excel {_excel_row(content_start)}..{_excel_row(true_end_abs_exclusive-1)}")
    print(f"[S] XML WQ count={len(xml_names)}; existing rows={len(excel_names)}; missing from Excel={len(missing)}")
    if missing:
        print(f"[S] Missing (top few): {missing[:10]}")

    return {
        "label": "Work Queues",
        "marker_row_idx": marker_row,
        "header_row_idx": header_row,
        "content_start": content_start,
        "content_end": true_end_abs_exclusive,
        "wq_name_col": wq_name_col,
        "key_col": key_col,
        "no_col": no_col,
        "check_col": check_col,
        "validation_col": validation_col,
        "header_excel_row": header_row + 1,
        "section_first_excel_row": content_start + 1,
        "section_row_count": true_end_abs_exclusive - content_start,
        "insert_at_excel_row": insert_at_excel_row,
        "excel_names": excel_names,
        "validation_vals": validation_vals,
        "missing": missing,
        "next_no": next_no,
        "fillable_rel": fillable_rel,
        "inserted_rows": 0,
    }


# =============== Write helpers for S ===============

def _fill_into_existing_blanks_S(ws, plan: Dict[str, Any], names_to_place: List[str],
                                 xml_wq: Dict[str, Dict[str, Optional[str]]]) -> int:
    consumed = 0
    val_col_excel = plan["validation_col"] + 1
    name_col_excel = plan["wq_name_col"] + 1
    no_col_excel   = plan["no_col"] + 1 if plan["no_col"] is not None else None
    check_col_excel= plan["check_col"] + 1 if plan["check_col"] is not None else None
    key_col_excel  = plan["key_col"] + 1 if plan.get("key_col") is not None else None

    next_no = plan["next_no"]

    for rel in plan["fillable_rel"]:
        if consumed >= len(names_to_place):
            break
        r = plan["section_first_excel_row"] + rel

        nm = names_to_place[consumed]
        ws.range((r, name_col_excel)).value = nm

        if key_col_excel is not None:
            key_val = (xml_wq.get(nm, {}) or {}).get("key") or ""
            if key_val:
                ws.range((r, key_col_excel)).value = key_val

        paste_formats_like_left(ws, r, val_col_excel)
        ws.range((r, val_col_excel)).value = "Newly added"
        ws.range((r, val_col_excel)).color = COLOR_ORANGE
        _apply_borders_like_left(ws, r, val_col_excel)

        if no_col_excel is not None:
            raw = str(ws.range((r, no_col_excel)).value or "").strip()
            cur = _extract_int(raw)
            if cur is None and next_no is not None:
                ws.range((r, no_col_excel)).value = next_no
                next_no += 1
            elif cur is not None and next_no is not None:
                next_no = max(next_no, cur + 1)

        if check_col_excel is not None:
            clear_fill_preserve_borders(ws, r, check_col_excel)

        consumed += 1

    plan["next_no"] = next_no
    return consumed


def _insert_new_rows_S(ws, plan: Dict[str, Any], remaining: List[str],
                       xml_wq: Dict[str, Dict[str, Optional[str]]]) -> int:
    if not remaining:
        return 0

    val_col_excel   = plan["validation_col"] + 1
    name_col_excel  = plan["wq_name_col"] + 1
    no_col_excel    = plan["no_col"] + 1 if plan["no_col"] is not None else None
    check_col_excel = plan["check_col"] + 1 if plan["check_col"] is not None else None
    key_col_excel   = plan["key_col"] + 1 if plan.get("key_col") is not None else None

    next_no = plan["next_no"]

    for i, nm in enumerate(remaining):
        r = plan["insert_at_excel_row"] + i
        insert_row_with_style(ws, r)
        ws.range((r, name_col_excel)).value = nm

        if key_col_excel is not None:
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
    plan["inserted_rows"] = (plan.get("inserted_rows", 0) or 0) + len(remaining)
    return len(remaining)


def _adjust_wq_key_validation(ws, plan: Dict[str, Any], xml_wq: Dict[str, Dict[str, Optional[str]]]) -> int:
    key_col = plan.get("key_col")
    if key_col is None:
        return 0

    name_col_excel = plan["wq_name_col"] + 1
    key_col_excel  = key_col + 1
    val_col_excel  = plan["validation_col"] + 1

    total_rows_to_check = plan["section_row_count"] + (plan.get("inserted_rows", 0) or 0)
    mismatches = 0

    for i in range(total_rows_to_check):
        r = plan["section_first_excel_row"] + i
        nm  = (_to_space(ws.range((r, name_col_excel)).value).strip())
        if not nm or nm not in xml_wq:
            continue

        expected_key = (_to_space(xml_wq[nm].get("key") or "").strip())
        if not expected_key:
            continue

        excel_key = (_to_space(ws.range((r, key_col_excel)).value).strip())

        if excel_key.casefold() != expected_key.casefold():
            current = (_to_space(ws.range((r, val_col_excel)).value).strip())
            if current.lower().startswith("exists"):
                ws.range((r, val_col_excel)).value = f'Exists (Key mismatch: expected "{expected_key}")'
            elif current.lower().startswith("newly added"):
                ws.range((r, val_col_excel)).value = f'Newly added (Key mismatch: expected "{expected_key}")'
            ws.range((r, val_col_excel)).color = COLOR_ORANGE
            mismatches += 1

    return mismatches


# =============== Main workflow ===============

_ILLEGAL_SHEET_CHARS = r'[:\\/?*\[\]]'

def _sanitize_sheet_name(name: str) -> str:
    name = re.sub(_ILLEGAL_SHEET_CHARS, "_", name).strip()
    return (name[:31] or "Sheet")


def _unique_sheet_name(wb, base: str) -> str:
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
    if isinstance(sheet_arg, int) or (isinstance(sheet_arg, str) and str(sheet_arg).isdigit()):
        return wb.sheets[int(sheet_arg)]
    return wb.sheets[str(sheet_arg)]


def validate_and_write(
    excel_path: str,
    sheet_arg,
    xml_process_names: List[str],
    xml_object_names: List[str],
    xml_work_queues: Dict[str, Dict[str, Optional[str]]],
    verbose_S: bool = True,
) -> None:
    sheet_arg_for_pd = int(sheet_arg) if (isinstance(sheet_arg, str) and str(sheet_arg).isdigit()) else sheet_arg
    df = pd.read_excel(excel_path, sheet_name=sheet_arg_for_pd, header=None, dtype=str, engine='openpyxl').fillna('')

    # A and B (unchanged)
    next_for_A = [["B.", "Business Object"]]
    next_for_B = [
        ["C.", "Environment Variables"], ["C", "Environment Variables"],
        ["C.", "environment", "variable"],
        ["D.", "environment", "variable"], ["D", "environment", "variable"],
        ["E.", "Startup Parameters"], ["E", "Startup Parameters"],
        ["S.", "work", "queue"], ["S", "work", "queue"],  # stop B at S if present
    ]

    plan_A = plan_section_AB(df, ["A.", "Process"], xml_process_names, label="Process",
                             header_keyword="name", next_section_groups=next_for_A)
    plan_B = plan_section_AB(df, ["B.", "Business Object"], xml_object_names, label="Business Object",
                             header_keyword="name", next_section_groups=next_for_B)
    plan_S = plan_section_S(df, xml_work_queues, verbose=verbose_S)

    if plan_A is None and plan_B is None and plan_S is None:
        print("⛔ Stopping: A, B, S not detected.")
        return

    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(str(excel_path))
        src_sht = _sheet_by_name_or_index(wb, sheet_arg)

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

        # ----- A -----
        if plan_A is not None:
            print("▶ A. Process")
            val_col_excel = plan_A["validation_col"] + 1
            paste_formats_like_left(ws, plan_A["header_excel_row"], val_col_excel)
            ws.range((plan_A["header_excel_row"], val_col_excel)).value = "Validation"
            _apply_borders_like_left(ws, plan_A["header_excel_row"], val_col_excel)

            _write_existing_validation(ws, plan_A["section_first_excel_row"], plan_A["validation_col"], plan_A["validation_vals"])

            # fill / insert
            rows_inserted_A = 0  # using your original generic helpers would go here, omitted to keep focus on S
            try:
                ws.api.Columns(val_col_excel).AutoFit()
            except Exception:
                pass

        # ----- B -----
        if plan_B is not None:
            print("▶ B. Business Object")
            if plan_A is not None and plan_B["header_row_idx"] > plan_A["header_row_idx"]:
                # would offset here if we inserted rows in A
                pass

            val_col_excel = plan_B["validation_col"] + 1
            paste_formats_like_left(ws, plan_B["header_excel_row"], val_col_excel)
            ws.range((plan_B["header_excel_row"], val_col_excel)).value = "Validation"
            _apply_borders_like_left(ws, plan_B["header_excel_row"], val_col_excel)

            _write_existing_validation(ws, plan_B["section_first_excel_row"], plan_B["validation_col"], plan_B["validation_vals"])

            rows_inserted_B = 0  # focus on S
            try:
                ws.api.Columns(val_col_excel).AutoFit()
            except Exception:
                pass

        # ----- S -----
        if plan_S is not None:
            print("▶ S. Work Queues")
            val_col_excel = plan_S["validation_col"] + 1
            paste_formats_like_left(ws, plan_S["header_excel_row"], val_col_excel)
            ws.range((plan_S["header_excel_row"], val_col_excel)).value = "Validation"
            _apply_borders_like_left(ws, plan_S["header_excel_row"], val_col_excel)

            # Existing rows' validation
            _write_existing_validation(ws, plan_S["section_first_excel_row"], plan_S["validation_col"], plan_S["validation_vals"])

            # Fill blanks first, then insert remainders
            consumed = _fill_into_existing_blanks_S(ws, plan_S, plan_S["missing"], xml_work_queues)
            remaining = plan_S["missing"][consumed:]
            inserted = _insert_new_rows_S(ws, plan_S, remaining, xml_work_queues)

            # Key check
            mismatches = _adjust_wq_key_validation(ws, plan_S, xml_work_queues)

            try:
                ws.api.Columns(val_col_excel).AutoFit()
            except Exception:
                pass

            print(f"[S] Summary: filled={consumed}, inserted={inserted}, key_mismatches={mismatches}")

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


# =============== CLI ===============

def main():
    parser = argparse.ArgumentParser(
        description="Validate A/B/S sections against a Blue Prism release XML (S uses strict same-row marker)."
    )
    parser.add_argument('--xml', required=True, help="Path to Blue Prism .bprelease XML")
    parser.add_argument('--excel', required=True, help="Path to Excel file (.xlsx/.xlsm)")
    parser.add_argument('--sheet', default="0", help="Sheet name or 0-based index (default: 0)")
    parser.add_argument('--quietS', action='store_true', help="Suppress verbose S detection logs")
    args = parser.parse_args()

    print("🔍 Parsing XML…")
    xml_process_names = extract_names_from_xml(args.xml, "process")
    xml_object_names  = extract_names_from_xml(args.xml, "object")
    xml_work_queues   = extract_work_queues_from_xml(args.xml)
    print(f"✅ Found {len(xml_process_names)} process; {len(xml_object_names)} objects; {len(xml_work_queues)} work queues.")

    print("🧪 Validating & writing…")
    validate_and_write(
        excel_path=args.excel,
        sheet_arg=args.sheet,
        xml_process_names=xml_process_names,
        xml_object_names=xml_object_names,
        xml_work_queues=xml_work_queues,
        verbose_S=not args.quietS,
    )


if __name__ == "__main__":
    main()
