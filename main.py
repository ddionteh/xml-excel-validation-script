#!/usr/bin/env python3
"""
Validate names in Excel 'A. Process' and 'B. Business Object' against a Blue Prism release XML.

Key points:
- Duplicates the target sheet (Excel-level copy; preserves shapes/objects/formatting)
- Adds a 'Validation' column at the end of the section header
- Writes text + color: Exists=green, Does not exist=red, Newly added=orange
- Inserts new rows for XML-only names (objects move with cells; borders kept)
- Continues 'No.' numbering for inserted rows
- Clears only the fill (not borders) in 'Check' for inserted rows
- Section B stops at C/D/E; section ends trimmed to the last numbered/real row
"""

import argparse
import re
import xml.etree.ElementTree as ET
from typing import List, Optional, Dict, Any

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


# =============== DataFrame parsing helpers ===============

def _norm_cell(s) -> str:
    """Normalize a cell for case/whitespace tolerant matching."""
    if s is None:
        return ""
    return str(s).strip().casefold()


def _row_contains_all(row_series: pd.Series, keywords: List[str]) -> bool:
    """
    True if all keywords appear either spread across the row cells, or together in any one cell.
    """
    row = [_norm_cell(c) for c in row_series]
    keys = [_norm_cell(k) for k in keywords]
    if all(any(k in c for c in row) for k in keys):
        return True
    return any(all(k in c for k in keys) for c in row)


def _two_line_or_same_row_match(df: pd.DataFrame, i: int, group: List[str]) -> bool:
    """
    True if row i contains all tokens in 'group' (same row), OR
    row i contains the first token and row i+1 contains the remaining tokens.
    Handles headers split across two lines (e.g., 'C.' then 'Environment Variables').
    """
    if _row_contains_all(df.iloc[i], group):
        return True
    if len(group) >= 2 and i + 1 < len(df):
        if _row_contains_all(df.iloc[i], [group[0]]) and _row_contains_all(df.iloc[i + 1], group[1:]):
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
):
    """
    Locate a section and its coarse boundaries (will be trimmed later):
      - marker row contains all 'section_keywords' (['A.', 'Process'] or ['B.', 'Business Object'])
      - header row within ~5 rows contains 'header_keyword' ('name')
      - content rows end at first row matching one of 'next_section_groups' or EOF
    Returns (content_start, content_end_exclusive, name_col_index, header_row_index, next_section_row_index)
    """
    # 1) section start marker row
    section_start_idx = None
    for i in range(len(df) - 1):
        if _row_contains_all(df.iloc[i], section_keywords):
            section_start_idx = i
            break
    if section_start_idx is None:
        print(f"❌ Could not find row containing all: {section_keywords}")
        return None, None, None, None, None

    # 2) header row (within next ~5 rows)
    header_row_idx = None
    scan_limit = min(section_start_idx + 6, len(df))
    for j in range(section_start_idx + 1, scan_limit):
        if _row_contains_all(df.iloc[j], [header_keyword]):
            header_row_idx = j
            break
    if header_row_idx is None:
        print(f"❌ Could not find header row containing '{header_keyword}' after the section start.")
        return None, None, None, None, None

    content_start = header_row_idx + 1

    # 3) detect Name column from header row
    name_col_index = None
    for col_idx, val in df.iloc[header_row_idx].items():
        if header_keyword in _norm_cell(val):
            name_col_index = col_idx
            break
    if name_col_index is None:
        print(f"❌ Could not identify the '{header_keyword}' column in the header row.")
        return None, None, None, None, None

    # 4) coarse section end (next header or EOF) — supports 1-row or 2-row headers
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
COLOR_ORANGE = (255, 235, 156)  # Newly added

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

def plan_section(df: pd.DataFrame, section_keys: List[str], xml_names: List[str], label: str,
                 header_keyword: str = "name", next_section_groups: Optional[List[List[str]]] = None) -> Optional[Dict[str, Any]]:
    """
    Build a plan describing how to write Validation and where to insert new rows.

    IMPORTANT TRIM: after coarse end is found, we trim the section to the **last real row**:
      1) Prefer the last row whose 'No.' column contains a number;
      2) Else, the last row with any non-blank cell in the row;
      3) Else, fall back to the last row with a non-blank Name.
    """
    found = find_section(df, section_keys, header_keyword=header_keyword, next_section_groups=next_section_groups)
    content_start, content_end_coarse, name_col, header_row_idx, _ = found
    if content_start is None:
        return None

    header_row = df.iloc[header_row_idx]
    # Determine where to place 'Validation' (reuse if already present)
    validation_col = None
    for col_idx, val in header_row.items():
        if _norm_cell(val) == "validation":
            validation_col = col_idx
            break
    if validation_col is None:
        validation_col = _last_nonempty_col_index(header_row) + 1

    # Identify helper columns
    no_col    = find_no_col_in_header(df, header_row_idx)
    check_col = find_check_col_in_header(df, header_row_idx)

    # Slice the coarse section for trimming
    coarse_df = df.iloc[content_start:content_end_coarse].copy().reset_index(drop=True)

    # (A) last row with a numeric "No."
    last_rel_by_no = None
    if no_col is not None and no_col < coarse_df.shape[1]:
        for i in range(len(coarse_df) - 1, -1, -1):
            val = _extract_int(coarse_df.iat[i, no_col])
            if val is not None:
                last_rel_by_no = i
                break

    # (B) last row with any non-blank cell
    last_rel_by_any = None
    for i in range(len(coarse_df) - 1, -1, -1):
        row_vals = [str(v).strip() for v in list(coarse_df.iloc[i].values)]
        if any(v != "" for v in row_vals):
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

    # Build section slice trimmed to the last real row
    section_df = df.iloc[content_start:true_end_abs_exclusive].copy().reset_index(drop=True)

    # Names and validation statuses for the trimmed region
    name_series = section_df[name_col].astype(str).fillna("").str.strip()
    excel_names = name_series.tolist()
    xml_set     = {x.strip() for x in xml_names}

    validation_vals = []
    for nm in excel_names:
        if not nm:
            validation_vals.append("")         # keep truly empty Name rows unlabelled
        elif nm in xml_set:
            validation_vals.append("Exists")
        else:
            validation_vals.append("Does not exist")

    # Missing names to append
    excel_set = {n for n in excel_names if n}
    missing   = [n for n in xml_names if n not in excel_set]

    # Insert position: AFTER the last real row we just computed
    last_abs = content_start + (true_end_rel_exclusive - 1) if true_end_rel_exclusive > 0 else header_row_idx
    insert_after_abs = last_abs
    insert_at_excel_row = insert_after_abs + 2  # Excel is 1-based; +1 to move after; +1 for 1-based

    # Next value for No.
    next_no = None
    if no_col is not None and true_end_rel_exclusive > 0:
        parsed = _extract_int(section_df.iat[true_end_rel_exclusive - 1, no_col])
        if parsed is not None:
            next_no = parsed + 1
        else:
            next_no = int(sum(1 for v in excel_names if v)) + 1
    elif no_col is not None:
        next_no = 1

    return {
        "label": label,
        "header_row_idx": header_row_idx,
        "content_start": content_start,
        "content_end": true_end_abs_exclusive,     # TRIMMED end (exclusive)
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
    if isinstance(sheet_arg, int) or (isinstance(sheet_arg, str) and sheet_arg.isdigit()):
        return wb.sheets[int(sheet_arg)]
    return wb.sheets[str(sheet_arg)]


# =============== Main workflow ===============

def validate_and_write_both(
    excel_path: str,
    sheet_arg,
    xml_process_names: List[str],
    xml_object_names: List[str],
) -> None:
    """
    End-to-end:
    - Read sheet into DataFrame to compute plans for A and B
    - Duplicate the sheet safely and rename to '<orig>_validated'
    - Write Validation cells (format like left, then color + TEXT)
    - Insert new rows and fill Name/No./Validation as needed
    - Preserve objects/placement and autofit Validation column
    """
    # Read once for planning
    sheet_arg_for_pd = int(sheet_arg) if (isinstance(sheet_arg, str) and str(sheet_arg).isdigit()) else sheet_arg
    df = pd.read_excel(excel_path, sheet_name=sheet_arg_for_pd, header=None, dtype=str, engine='openpyxl').fillna('')

    # Define section boundaries:
    # A ends at B; B ends at C/D/E (allow split headers)
    next_for_A = [["B.", "Business Object"]]
    next_for_B = [
        ["C.", "Environment Variables"],
        ["D.", "Environment Variables"],
        ["E.", "Startup Parameters"],
    ]

    plan_A = plan_section(df, ["A.", "Process"], xml_process_names, label="Process",
                          header_keyword="name", next_section_groups=next_for_A)
    plan_B = plan_section(df, ["B.", "Business Object"], xml_object_names, label="Business Object",
                          header_keyword="name", next_section_groups=next_for_B)
    if plan_A is None and plan_B is None:
        print("⛔ Stopping: neither section A nor B was detected.")
        return

    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(str(excel_path))
        src_sht = _sheet_by_name_or_index(wb, sheet_arg)

        # Safe copy/rename (avoid stray '(2)' sheet; disable events to dodge macros)
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

        # ----- Section A: Process -----
        if plan_A is not None:
            print("▶ Processing section A: Process")
            val_col_excel = plan_A["validation_col"] + 1

            # Header: paste formats like left, set caption, clone borders from left
            paste_formats_like_left(ws, plan_A["header_excel_row"], val_col_excel)
            ws.range((plan_A["header_excel_row"], val_col_excel)).value = "Validation"
            _apply_borders_like_left(ws, plan_A["header_excel_row"], val_col_excel)

            # Existing rows: format like left then color + TEXT
            if plan_A["section_row_count"] > 0:
                for i, status in enumerate(plan_A["validation_vals"]):
                    r = plan_A["section_first_excel_row"] + i
                    paste_formats_like_left(ws, r, val_col_excel)
                    if status:
                        ws.range((r, val_col_excel)).value = status  # <-- write the words
                    else:
                        ws.range((r, val_col_excel)).value = ""
                    # color
                    st = (status or "").lower()
                    if st == "exists":
                        ws.range((r, val_col_excel)).color = COLOR_GREEN
                    elif st == "does not exist":
                        ws.range((r, val_col_excel)).color = COLOR_RED
                    else:
                        ws.range((r, val_col_excel)).color = None
                    _apply_borders_like_left(ws, r, val_col_excel)

            # New rows: insert after last real row, keep styles/objects shifting
            if plan_A["missing"]:
                next_no = plan_A["next_no"]
                for i, name in enumerate(plan_A["missing"]):
                    r = plan_A["insert_at_excel_row"] + i
                    insert_row_with_style(ws, r)
                    ws.range((r, plan_A["name_col"] + 1)).value = name
                    paste_formats_like_left(ws, r, val_col_excel)
                    ws.range((r, val_col_excel)).value = "Newly added"
                    ws.range((r, val_col_excel)).color = COLOR_ORANGE
                    _apply_borders_like_left(ws, r, val_col_excel)
                    if plan_A["no_col"] is not None and next_no is not None:
                        ws.range((r, plan_A["no_col"] + 1)).value = next_no
                        next_no += 1
                    if plan_A["check_col"] is not None:
                        clear_fill_preserve_borders(ws, r, plan_A["check_col"] + 1)
                rows_inserted_A = len(plan_A["missing"])

            try:
                ws.api.Columns(val_col_excel).AutoFit()
            except Exception:
                pass

        # ----- Section B: Business Object -----
        if plan_B is not None:
            print("▶ Processing section B: Business Object")
            if plan_A is not None and plan_B["header_row_idx"] > plan_A["header_row_idx"]:
                plan_B = apply_row_offset(plan_B, rows_inserted_A)

            val_col_excel = plan_B["validation_col"] + 1

            paste_formats_like_left(ws, plan_B["header_excel_row"], val_col_excel)
            ws.range((plan_B["header_excel_row"], val_col_excel)).value = "Validation"
            _apply_borders_like_left(ws, plan_B["header_excel_row"], val_col_excel)

            if plan_B["section_row_count"] > 0:
                for i, status in enumerate(plan_B["validation_vals"]):
                    r = plan_B["section_first_excel_row"] + i
                    paste_formats_like_left(ws, r, val_col_excel)
                    if status:
                        ws.range((r, val_col_excel)).value = status  # <-- write the words
                    else:
                        ws.range((r, val_col_excel)).value = ""
                    st = (status or "").lower()
                    if st == "exists":
                        ws.range((r, val_col_excel)).color = COLOR_GREEN
                    elif st == "does not exist":
                        ws.range((r, val_col_excel)).color = COLOR_RED
                    else:
                        ws.range((r, val_col_excel)).color = None
                    _apply_borders_like_left(ws, r, val_col_excel)

            if plan_B["missing"]:
                next_no = plan_B["next_no"]
                for i, name in enumerate(plan_B["missing"]):
                    r = plan_B["insert_at_excel_row"] + i
                    insert_row_with_style(ws, r)
                    ws.range((r, plan_B["name_col"] + 1)).value = name
                    paste_formats_like_left(ws, r, val_col_excel)
                    ws.range((r, val_col_excel)).value = "Newly added"
                    ws.range((r, val_col_excel)).color = COLOR_ORANGE
                    _apply_borders_like_left(ws, r, val_col_excel)
                    if plan_B["no_col"] is not None and next_no is not None:
                        ws.range((r, plan_B["no_col"] + 1)).value = next_no
                        next_no += 1
                    if plan_B["check_col"] is not None:
                        clear_fill_preserve_borders(ws, r, plan_B["check_col"] + 1)

            try:
                ws.api.Columns(val_col_excel).AutoFit()
            except Exception:
                pass

        restore_placement(ws, placements)
        wb.save()
        print(f"✅ Completed. Validation texts restored; sections trimmed to last real row; wrote '{ws.name}'.")
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
        description="Validate 'A. Process' and 'B. Business Object' against <process name> / <object name> in a Blue Prism release XML. "
                    "Preserves styling/objects and adds a colored 'Validation' column."
    )
    parser.add_argument('--xml', required=True, help="Path to Blue Prism .bprelease XML")
    parser.add_argument('--excel', required=True, help="Path to Excel file (.xlsx/.xlsm)")
    parser.add_argument('--sheet', default="0", help="Sheet name or 0-based index (default: 0)")
    args = parser.parse_args()

    print("🔍 Parsing XML…")
    xml_process_names = extract_names_from_xml(args.xml, "process")
    xml_object_names  = extract_names_from_xml(args.xml, "object")
    print(f"✅ Found {len(xml_process_names)} process names; {len(xml_object_names)} object names.")

    print("🧪 Validating & writing…")
    validate_and_write_both(
        excel_path=args.excel,
        sheet_arg=args.sheet,
        xml_process_names=xml_process_names,
        xml_object_names=xml_object_names,
    )


if __name__ == "__main__":
    main()
