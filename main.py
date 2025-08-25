#!/usr/bin/env python3
"""
Validate Excel sections against a Blue Prism release XML.

Supports:
- A. Process
- B. Business Object
- S. Work Queues (Work Queue Name + Key Name validation)

Behavior preserved from your original:
- Fill existing blank Name rows first, then insert remaining.
- Continue 'No.' numbering.
- Trim sections to the last real row.
- Write Validation text + color (green/red/orange).
- Preserve borders/formatting; set embedded objects to "move with cells" during edits and restore later.

Diagnostics for S:
- Tolerant marker (S / S. / S:)
- Header scanning window widened to 80 rows
- Multiple header heuristics + scored fallback
- Verbose prints: marker row, header row, chosen columns, content range, counts, key mismatches
"""

import argparse
import re
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


# =============== small utils ===============

def _norm_cell(s) -> str:
    if s is None:
        return ""
    return str(s).strip().casefold()


def _normalize_marker_token(s: str) -> str:
    v = _norm_cell(s)
    return v.rstrip(".:") if re.fullmatch(r"[a-zA-Z][\.:]?", v or "") else v


def _row_contains_all(row_series: pd.Series, keywords: List[str]) -> bool:
    row = [_norm_cell(c) for c in row_series]
    keys = [_norm_cell(k) for k in keywords]
    if all(any(k in c for c in row) for k in keys):
        return True
    return any(all(k in c for k in keys) for c in row)


def _row_contains_all_markers(row_series: pd.Series, keywords: List[str]) -> bool:
    row = [_normalize_marker_token(c) for c in row_series]
    keys = [_normalize_marker_token(k) for k in keywords]
    if all(any(k in c for c in row) for k in keys):
        return True
    return any(all(k in c for k in keys) for c in row)


def _two_line_or_same_row_match_markers(df: pd.DataFrame, i: int, group: List[str]) -> bool:
    # same row
    if _row_contains_all_markers(df.iloc[i], group):
        return True
    # split across two rows (e.g., "S" on row i, "Work Queue" on row i+1)
    if len(group) >= 2 and i + 1 < len(df):
        if _row_contains_all_markers(df.iloc[i], [group[0]]) and _row_contains_all_markers(df.iloc[i + 1], group[1:]):
            return True
    return False


def _last_nonempty_col_index(row_series: pd.Series) -> int:
    last = -1
    for col_idx, val in row_series.items():
        if str(val).strip() != "":
            last = col_idx
    return last


def _excel_col_letter(idx0: int) -> str:
    n = idx0 + 1
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _row_preview(row_series: pd.Series, max_cells: int = 40) -> str:
    vals = [str(x) for x in row_series.tolist()]
    trimmed = []
    for v in vals:
        vv = v.strip()
        if vv != "":
            trimmed.append(vv)
        if len(trimmed) >= max_cells:
            trimmed.append("…")
            break
    return " | ".join(trimmed) if trimmed else "<all blank>"


# =============== header / column finders ===============

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


def find_keyname_col_in_header(df: pd.DataFrame, header_row_idx: int) -> Optional[int]:
    for col_idx, val in df.iloc[header_row_idx].items():
        v = _norm_cell(val)
        if "key name" in v or ("key" in v and "name" in v):
            return col_idx
    return None


def _find_wq_name_col_in_header(df: pd.DataFrame, header_row_idx: int) -> Optional[int]:
    # prefer explicit "work queue name"
    candidates = [(col_idx, _norm_cell(val)) for col_idx, val in df.iloc[header_row_idx].items()]
    for col_idx, v in candidates:
        if "work" in v and "queue" in v and "name" in v:
            return col_idx
    # else "queue name"
    for col_idx, v in candidates:
        if "queue" in v and "name" in v:
            return col_idx
    # else any "name" not containing "key"
    for col_idx, v in candidates:
        if "name" in v and "key" not in v:
            return col_idx
    # else any "name"
    for col_idx, v in candidates:
        if "name" in v:
            return col_idx
    return None


# =============== WQ header scoring fallback ===============

def _score_row_for_wq_header(row_series: pd.Series) -> Tuple[int, Dict[str, bool]]:
    texts = [_norm_cell(x) for x in row_series.tolist()]
    has = lambda tok: any(tok in t for t in texts)
    both = lambda a, b: any(a in t for t in texts) and any(b in t for t in texts)

    signals = {
        "work_queue": both("work", "queue"),
        "queue_name": both("queue", "name"),
        "key_name":   both("key", "name"),
        "encrypted":  has("encrypted"),
        "check":      has("check"),
        "no":         any(t == "no." or t == "no" or t.startswith("no.") for t in texts),
    }
    score = (2 if signals["work_queue"] else 0) \
          + (2 if signals["key_name"] else 0) \
          + (1 if signals["queue_name"] else 0) \
          + (1 if signals["encrypted"] else 0) \
          + (1 if signals["check"] else 0) \
          + (1 if signals["no"] else 0)
    return score, signals


# =============== Section finder (robust for S) ===============

def find_section(
    df: pd.DataFrame,
    section_keywords: List[str],
    label_for_debug: str,
    header_any_of: Optional[List[List[str]]] = None,
    header_keyword: Optional[str] = "name",
    next_section_groups: Optional[List[List[str]]] = None,
) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[int], Optional[int]]:
    # 1) Section start marker row
    section_start_idx = None
    for i in range(len(df) - 1):
        if _two_line_or_same_row_match_markers(df, i, section_keywords):
            section_start_idx = i
            break
    if section_start_idx is None:
        print(f"❌ Could not find section marker {section_keywords} ({label_for_debug})")
        return None, None, None, None, None

    print(f"🔎 [{label_for_debug}] marker at row {section_start_idx+1}")
    print(f"   [{label_for_debug}] Next ~15 rows preview after marker:")
    for j in range(section_start_idx + 1, min(section_start_idx + 16, len(df))):
        print(f"     r{j+1}: {_row_preview(df.iloc[j])}")

    # 2) Header row (scan up to 80 rows)
    header_row_idx = None
    scan_limit = min(section_start_idx + 80, len(df))

    # explicit patterns (e.g., for S: 'work','queue','name' ; 'key','name' ; 'encrypted' ; 'check')
    if header_any_of:
        for j in range(section_start_idx + 1, scan_limit):
            if all(any(k in _norm_cell(c) for c in df.iloc[j].tolist()) for k in set(sum(header_any_of, []))):
                # If all tokens across all groups exist somewhere in the row, accept fast
                header_row_idx = j
                print(f"   [{label_for_debug}] Header matched by 'token set' presence at row {j+1}")
                break
        if header_row_idx is None:
            # Try group-by-group match (more strict)
            for j in range(section_start_idx + 1, scan_limit):
                if all(_row_contains_all(df.iloc[j], grp) for grp in header_any_of):
                    header_row_idx = j
                    print(f"   [{label_for_debug}] Header matched by 'all groups' at row {j+1}")
                    break

    # Fallback to single keyword
    if header_row_idx is None and header_keyword:
        for j in range(section_start_idx + 1, scan_limit):
            if _row_contains_all(df.iloc[j], [header_keyword]):
                header_row_idx = j
                print(f"   [{label_for_debug}] Header matched by fallback keyword '{header_keyword}' at row {j+1}")
                break

    # Scored fallback for S
    if header_row_idx is None and label_for_debug.lower().startswith("work queue"):
        best = (-1, None, None)  # (score, idx, signals)
        for j in range(section_start_idx + 1, scan_limit):
            score, sig = _score_row_for_wq_header(df.iloc[j])
            if score > best[0]:
                best = (score, j, sig)
        score, j, sig = best
        print(f"   [Work Queues] top-scored header candidate: row {j+1 if j is not None else '?'} "
              f"score={score} signals={sig}")
        if j is not None and score >= 2:
            header_row_idx = j
            print(f"   [Work Queues] Header chosen by scoring at row {j+1}")

    if header_row_idx is None:
        print(f"❌ Could not find header row for {label_for_debug} after marker at row {section_start_idx+1}")
        return None, None, None, None, None

    content_start = header_row_idx + 1

    # 3) Tentative name col from keyword (may be overridden for S)
    name_col_index = None
    if header_keyword:
        for col_idx, val in df.iloc[header_row_idx].items():
            if header_keyword in _norm_cell(val):
                name_col_index = col_idx
                break

    # 4) Coarse end
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

    print(f"   [{label_for_debug}] header row at {header_row_idx+1}: {_row_preview(df.iloc[header_row_idx])}")
    print(f"   [{label_for_debug}] tentative Name-col idx={name_col_index} "
          f"(Excel col {(_excel_col_letter(name_col_index) if name_col_index is not None else '?')}); "
          f"coarse content rows {content_start+1}..{content_end} (Excel rows)")
    return content_start, content_end, name_col_index, header_row_idx, next_section_row_idx


# =============== Styling & borders ===============

COLOR_GREEN  = (198, 239, 206)  # Exists
COLOR_RED    = (255, 199, 206)  # Does not exist
COLOR_ORANGE = (255, 235, 156)  # Newly added / Key mismatch

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
        ws.api.Cells(row, col_excel).PasteSpecial(Paste=-4122)  # formats
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
    ws.api.Rows(row_num).Insert(Shift=-4121, CopyOrigin=0)  # down, format from above
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


# =============== planning + writing ===============

def _extract_int(s: str) -> Optional[int]:
    m = re.search(r"\d+", str(s))
    return int(m.group(0)) if m else None


def _row_has_any_nonblank(series: pd.Series) -> bool:
    return any(str(v).strip() != "" for v in list(series.values))


def plan_section(
    df: pd.DataFrame,
    section_keys: List[str],
    xml_names: List[str],
    label: str,
    header_any_of: Optional[List[List[str]]] = None,
    header_keyword: str = "name",
    next_section_groups: Optional[List[List[str]]] = None
) -> Optional[Dict[str, Any]]:
    found = find_section(
        df,
        section_keys,
        label_for_debug=label,
        header_any_of=header_any_of,
        header_keyword=header_keyword,
        next_section_groups=next_section_groups
    )
    content_start, content_end_coarse, name_col, header_row_idx, _ = found
    if content_start is None:
        return None

    header_row = df.iloc[header_row_idx]

    # Existing "Validation" or the first new column to the right
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
    key_col   = find_keyname_col_in_header(df, header_row_idx)

    # S-specific: choose Work Queue Name column robustly
    if label.lower().startswith("work queue"):
        better = _find_wq_name_col_in_header(df, header_row_idx)
        if better is not None:
            name_col = better
        if name_col is None:
            print("❌ Could not identify a Work Queue Name column in the header row.")
            return None

        print(f"   [S] Chosen columns (0-based):")
        print(f"       Work Queue Name col = {name_col} ({_excel_col_letter(name_col)})")
        if key_col is not None:
            print(f"       Key Name col        = {key_col} ({_excel_col_letter(key_col)})")
        else:
            print(f"       Key Name col        = <not found in header>")
        if no_col is not None:
            print(f"       No. col             = {no_col} ({_excel_col_letter(no_col)})")
        if check_col is not None:
            print(f"       Check col           = {check_col} ({_excel_col_letter(check_col)})")
        print(f"       Validation col      = {validation_col} ({_excel_col_letter(validation_col)})")

    # Slice coarse section
    coarse_df = df.iloc[content_start:content_end_coarse].copy().reset_index(drop=True)

    # (A) last numeric No.
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

    # (B) last non-blank row
    last_rel_by_any = None
    for i in range(len(coarse_df) - 1, -1, -1):
        if _row_has_any_nonblank(coarse_df.iloc[i]):
            last_rel_by_any = i
            break

    # (C) last non-blank Name
    last_rel_by_name = None
    if name_col is not None and name_col < coarse_df.shape[1]:
        for i in range(len(coarse_df) - 1, -1, -1):
            if str(coarse_df.iat[i, name_col]).strip() != "":
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

    # Names + validation
    name_series = section_df[name_col].astype(str).fillna("").str.strip()
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

    # Missing → to place
    excel_set = {n for n in excel_names if n}
    missing   = [n for n in xml_names if n not in excel_set]

    # Fillable rows: blank Name but still in-table (No. numeric or any other non-blank cell)
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

    last_abs = content_start + (true_end_rel_exclusive - 1) if true_end_rel_exclusive > 0 else header_row_idx
    insert_at_excel_row = last_abs + 2

    if max_no is not None:
        next_no = max_no + 1
    else:
        nonblank_names = sum(1 for v in excel_names if v)
        next_no = nonblank_names + 1 if no_col is not None else None

    if label.lower().startswith("work queue"):
        print(f"   [S] Content rows (trimmed): Excel {content_start+1}..{true_end_abs_exclusive}")
        print(f"   [S] XML WQ count={len(xml_names)}; existing rows={len(excel_names)}; "
              f"missing from Excel={len(missing)}")
        if missing:
            print(f"   [S] First few missing: {missing[:10]}")

    return {
        "label": label,
        "header_row_idx": header_row_idx,
        "content_start": content_start,
        "content_end": true_end_abs_exclusive,
        "name_col": name_col,
        "validation_col": validation_col,
        "no_col": no_col,
        "check_col": check_col,
        "key_col": key_col,
        "header_excel_row": header_row_idx + 1,
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


def apply_row_offset(plan: Optional[Dict[str, Any]], offset_rows: int) -> Optional[Dict[str, Any]]:
    if offset_rows == 0 or plan is None:
        return plan
    newp = dict(plan)
    for key in ("header_excel_row", "section_first_excel_row", "insert_at_excel_row"):
        newp[key] = plan[key] + offset_rows
    return newp


# =============== write helpers ===============

def _write_existing_validation(ws, plan: Dict[str, Any]) -> None:
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

        nm = names_to_place[consumed]
        ws.range((r, name_col_excel)).value = nm

        if key_col_excel is not None and xml_wq is not None:
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


def _insert_new_rows(ws, plan: Dict[str, Any], remaining: List[str],
                     xml_wq: Optional[Dict[str, Dict[str, Optional[str]]]] = None) -> int:
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


def _adjust_wq_key_validation(ws, plan: Dict[str, Any], xml_wq: Dict[str, Dict[str, Optional[str]]]) -> int:
    key_col = plan.get("key_col")
    if key_col is None:
        return 0

    name_col_excel = plan["name_col"] + 1
    key_col_excel  = key_col + 1
    val_col_excel  = plan["validation_col"] + 1

    total_rows_to_check = plan["section_row_count"] + (plan.get("inserted_rows", 0) or 0)
    mismatches = 0

    for i in range(total_rows_to_check):
        r = plan["section_first_excel_row"] + i
        nm  = (ws.range((r, name_col_excel)).value or "").strip()
        if not nm or nm not in xml_wq:
            continue

        expected_key = (xml_wq[nm].get("key") or "").strip()
        if not expected_key:
            continue

        excel_key = (ws.range((r, key_col_excel)).value or "").strip()

        if excel_key.strip().casefold() != expected_key.strip().casefold():
            current = (ws.range((r, val_col_excel)).value or "").strip()
            if current.lower().startswith("exists"):
                ws.range((r, val_col_excel)).value = f'Exists (Key mismatch: expected "{expected_key}")'
            elif current.lower().startswith("newly added"):
                ws.range((r, val_col_excel)).value = f'Newly added (Key mismatch: expected "{expected_key}")'
            ws.range((r, val_col_excel)).color = COLOR_ORANGE
            mismatches += 1

    return mismatches


# =============== Orchestration ===============

def _sanitize_sheet_name(name: str) -> str:
    return re.sub(r'[:\\/?*\[\]]', "_", name).strip()[:31] or "Sheet"


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


def validate_and_write_both(
    excel_path: str,
    sheet_arg,
    xml_process_names: List[str],
    xml_object_names: List[str],
    xml_work_queues: Dict[str, Dict[str, Optional[str]]],
) -> None:
    # Read once for planning
    sheet_arg_for_pd = int(sheet_arg) if (isinstance(sheet_arg, str) and str(sheet_arg).isdigit()) else sheet_arg
    df = pd.read_excel(excel_path, sheet_name=sheet_arg_for_pd, header=None, dtype=str, engine='openpyxl').fillna('')

    # Section boundaries
    next_for_A = [["B.", "Business Object"], ["B", "Business Object"]]
    next_for_B = [
        ["C.", "Environment Variables"], ["C", "Environment Variables"],
        ["C.", "environment", "variable"], ["D.", "environment", "variable"], ["D", "environment", "variable"],
        ["E.", "Startup Parameters"], ["E", "Startup Parameters"],
        ["S.", "work", "queue"], ["S", "work", "queue"],
    ]
    next_for_S = None  # until EOF

    # A, B
    plan_A = plan_section(df, ["A.", "Process"], xml_process_names, label="Process",
                          header_any_of=[["name"]], header_keyword="name", next_section_groups=next_for_A)
    plan_B = plan_section(df, ["B.", "Business Object"], xml_object_names, label="Business Object",
                          header_any_of=[["name"]], header_keyword="name", next_section_groups=next_for_B)

    # S — tuned to your header: No. | Work Queue Name | Key Name | Encrypted | (empty) | Check
    xml_wq_names = list(xml_work_queues.keys())
    s_header_patterns = [
        ["work", "queue", "name"],
        ["key", "name"],
        ["encrypted"],
        ["check"],
    ]
    plan_S = plan_section(df, ["S.", "work", "queue"], xml_wq_names, label="Work Queues",
                          header_any_of=s_header_patterns, header_keyword="queue", next_section_groups=next_for_S)

    if plan_A is None and plan_B is None and plan_S is None:
        print("⛔ Stopping: sections A, B, S were not detected.")
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

        # ----- B -----
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

        # ----- S -----
        if plan_S is not None:
            print("▶ Processing section S (Work Queues)")
            offset = 0
            if plan_A is not None and plan_S["header_row_idx"] > plan_A["header_row_idx"]:
                offset += rows_inserted_A
            if plan_B is not None and plan_S["header_row_idx"] > plan_B["header_row_idx"]:
                offset += rows_inserted_B
            if offset:
                plan_S = apply_row_offset(plan_S, offset)
                print(f"   [S] Applied row offset: +{offset}")

            val_col_excel = plan_S["validation_col"] + 1
            paste_formats_like_left(ws, plan_S["header_excel_row"], val_col_excel)
            ws.range((plan_S["header_excel_row"], val_col_excel)).value = "Validation"
            _apply_borders_like_left(ws, plan_S["header_excel_row"], val_col_excel)
            print(f"   [S] Wrote 'Validation' header at Excel col {_excel_col_letter(plan_S['validation_col'])}, row {plan_S['header_excel_row']}")

            _write_existing_validation(ws, plan_S)
            print(f"   [S] Existing rows validated: {plan_S['section_row_count']}")

            consumed = _fill_into_existing_blanks(ws, plan_S, plan_S["missing"], xml_work_queues)
            remaining = plan_S["missing"][consumed:]
            inserted = _insert_new_rows(ws, plan_S, remaining, xml_work_queues)
            print(f"   [S] Filled blanks: {consumed}, inserted rows: {inserted}")

            mismatches = _adjust_wq_key_validation(ws, plan_S, xml_work_queues)
            print(f"   [S] Key Name mismatches found: {mismatches}")

            try:
                ws.api.Columns(val_col_excel).AutoFit()
            except Exception:
                pass
        else:
            print("⚠️ S (Work Queues) not detected; nothing written for S.")

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
        description="Validate 'A. Process', 'B. Business Object', and 'S. Work Queues' against a Blue Prism release XML."
    )
    parser.add_argument('--xml', required=True, help="Path to Blue Prism .bprelease XML")
    parser.add_argument('--excel', required=True, help="Path to Excel file (.xlsx/.xlsm)")
    parser.add_argument('--sheet', default="0", help="Sheet name or 0-based index (default: 0)")
    args = parser.parse_args()

    print("🔍 Parsing XML…")
    xml_process_names = extract_names_from_xml(args.xml, "process")
    xml_object_names  = extract_names_from_xml(args.xml, "object")
    xml_work_queues   = extract_work_queues_from_xml(args.xml)
    print(f"✅ XML: {len(xml_process_names)} processes; {len(xml_object_names)} objects; {len(xml_work_queues)} work queues.")

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
