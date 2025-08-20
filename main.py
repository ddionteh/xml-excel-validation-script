#!/usr/bin/env python3
import argparse
import re
import xml.etree.ElementTree as ET

import pandas as pd
import xlwings as xw


# ---------------- XML helpers ----------------

def _local(tag: str) -> str:
    return tag.split('}', 1)[1] if tag.startswith('{') else tag


def extract_names_from_xml(xml_path: str, want: str) -> list[str]:
    tree = ET.parse(xml_path)
    root = tree.getroot()
    names: list[str] = []
    
    # Handle Blue Prism object references in resource elements
    if want == "object":
        # Look for <resource object="..." /> elements
        for elem in root.iter():
            if _local(elem.tag) == "resource" and elem.get('object'):
                object_name = elem.get('object')
                if object_name:
                    names.append(object_name.strip())
    else:
        # Handle other tag types (process, etc.)
        for elem in root.iter():
            if _local(elem.tag) == want:
                name_attr = elem.attrib.get('name')
                if name_attr:
                    names.append(name_attr.strip())
    
    return list(dict.fromkeys(names))


# ---------------- DataFrame parsing helpers ----------------

def _norm_cell(s) -> str:
    if s is None:
        return ""
    return str(s).strip().casefold()


def _row_contains_all(row_series, keywords: list[str]) -> bool:
    row = [_norm_cell(c) for c in row_series]
    keys = [_norm_cell(k) for k in keywords]
    if all(any(k in c for c in row) for k in keys):
        return True
    return any(all(k in c for k in keys) for c in row)


def _last_nonempty_col_index(row_series) -> int:
    last = -1
    for col_idx, val in row_series.items():
        if str(val).strip() != "":
            last = col_idx
    return last


def find_section(
    df: pd.DataFrame,
    section_keywords: list[str],
    header_keyword: str = "name",
    next_section_keywords: list[str] | None = None,
):
    section_start_idx = None
    for i in range(len(df) - 1):
        if _row_contains_all(df.iloc[i], section_keywords):
            section_start_idx = i
            break
    if section_start_idx is None:
        print(f"❌ Could not find row containing all: {section_keywords}")
        return None, None, None, None, None

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

    name_col_index = None
    for col_idx, val in df.iloc[header_row_idx].items():
        if header_keyword in _norm_cell(val):
            name_col_index = col_idx
            break
    if name_col_index is None:
        print(f"❌ Could not identify the '{header_keyword}' column in the header row.")
        return None, None, None, None, None

    next_section_row_idx = None
    content_end = None
    if next_section_keywords:
        for k in range(content_start, len(df)):
            if _row_contains_all(df.iloc[k], next_section_keywords):
                next_section_row_idx = k
                content_end = k
                break
    if content_end is None:
        content_end = len(df)

    print(
        f"🔎 Section found: header at row {header_row_idx}, "
        f"content rows {content_start}..{content_end - 1}, "
        f"Name col = {name_col_index}"
    )
    return content_start, content_end, name_col_index, header_row_idx, next_section_row_idx


def find_no_col_in_header(df: pd.DataFrame, header_row_idx: int) -> int | None:
    for col_idx, val in df.iloc[header_row_idx].items():
        v = _norm_cell(val)
        if v == "no." or v == "no" or v.startswith("no."):
            return col_idx
    return None


def find_check_col_in_header(df: pd.DataFrame, header_row_idx: int) -> int | None:
    for col_idx, val in df.iloc[header_row_idx].items():
        if "check" in _norm_cell(val):
            return col_idx
    return None


def _extract_int(s: str) -> int | None:
    m = re.search(r"\d+", str(s))
    return int(m.group(0)) if m else None


# ---------------- Shapes/objects placement snapshot ----------------
# Excel enum values: xlMoveAndSize=1, xlMove=2, xlFreeFloating=3

def snapshot_and_set_placement(ws) -> list[tuple[str, str, int]]:
    records: list[tuple[str, str, int]] = []
    try:
        shp = ws.api.Shapes
        for i in range(1, shp.Count + 1):
            it = shp.Item(i)
            try:
                records.append((it.Name, "Shapes", it.Placement))
                it.Placement = 2
            except Exception:
                pass
    except Exception:
        pass
    try:
        ch = ws.api.ChartObjects()
        for i in range(1, ch.Count + 1):
            it = ch.Item(i)
            try:
                records.append((it.Name, "ChartObjects", it.Placement))
                it.Placement = 2
            except Exception:
                pass
    except Exception:
        pass
    try:
        ole = ws.api.OLEObjects()
        for i in range(1, ole.Count + 1):
            it = ole.Item(i)
            try:
                records.append((it.Name, "OLEObjects", it.Placement))
                it.Placement = 2
            except Exception:
                pass
    except Exception:
        pass
    return records


def restore_placement(ws, records: list[tuple[str, str, int]]) -> None:
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


# ---------------- Styling & borders ----------------

COLOR_GREEN  = (198, 239, 206)  # Exists
COLOR_RED    = (255, 199, 206)  # Does not exist
COLOR_ORANGE = (255, 235, 156)  # Newly added
COLOR_HEADER = (217, 217, 217)  # Validation header

def style_validation_header(ws, header_row: int, val_col_excel: int):
    cell = ws.range((header_row, val_col_excel))
    cell.value = "Validation"
    cell.color = COLOR_HEADER
    try:
        cell.api.Font.Bold = True
    except Exception:
        pass
    # Match borders to the left neighbor (e.g., "Check")
    _apply_borders_like_left(ws, header_row, val_col_excel)

def color_validation_cells(ws, start_row: int, statuses: list[str], val_col_excel: int, end_row_inclusive: int):
    for i, status in enumerate(statuses):
        r = start_row + i
        if r > end_row_inclusive:
            break
        if not status:
            ws.range((r, val_col_excel)).color = None
        else:
            s = status.strip().lower()
            if s == "exists":
                ws.range((r, val_col_excel)).color = COLOR_GREEN
            elif s == "does not exist":
                ws.range((r, val_col_excel)).color = COLOR_RED
            elif s == "newly added":
                ws.range((r, val_col_excel)).color = COLOR_ORANGE
            else:
                ws.range((r, val_col_excel)).color = None
        # Match borders to the left neighbor for each row
        _apply_borders_like_left(ws, r, val_col_excel)

def _clone_borders(from_cell_api, to_cell_api):
    # Edge/inside ids: 7,8,9,10,11,12
    for idx in (7, 8, 9, 10, 11, 12):
        try:
            bsrc = from_cell_api.Borders(idx)
            bdst = to_cell_api.Borders(idx)
            bdst.LineStyle = bsrc.LineStyle
            bdst.Weight = bsrc.Weight
            bdst.Color = bsrc.Color
        except Exception:
            pass
    # Alignments to keep table look
    try:
        to_cell_api.HorizontalAlignment = from_cell_api.HorizontalAlignment
        to_cell_api.VerticalAlignment = from_cell_api.VerticalAlignment
    except Exception:
        pass

def _apply_borders_like_left(ws, row: int, col_excel: int):
    if col_excel <= 1:
        return
    try:
        left = ws.api.Cells(row, col_excel - 1)
        dst  = ws.api.Cells(row, col_excel)
        _clone_borders(left, dst)
    except Exception:
        pass

def insert_row_with_style(ws, row_num: int):
    ws.api.Rows(row_num).Insert(Shift=-4121, CopyOrigin=0)  # down, format from above
    try:
        prev = row_num - 1
        ws.api.Rows(prev).Copy()
        ws.api.Rows(row_num).PasteSpecial(Paste=-4122)  # formats
        ws.api.Rows(row_num).RowHeight = ws.api.Rows(prev).RowHeight
        # Recreate same-row horizontal merges
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


# ---------------- Planning helpers ----------------

def plan_section(df: pd.DataFrame, section_keys: list[str], xml_names: list[str], label: str,
                 header_keyword="name", next_keys=None):
    content_start, content_end, name_col, header_row_idx, _ = find_section(
        df, section_keys, header_keyword=header_keyword, next_section_keywords=next_keys
    )
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

    section_df  = df.iloc[content_start:content_end].copy().reset_index(drop=True)
    name_series = section_df[name_col].astype(str).fillna("").str.strip()
    excel_names = name_series.tolist()

    xml_set = {x.strip() for x in xml_names}
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

    last_rel = None
    for idx in range(len(name_series) - 1, -1, -1):
        if name_series.iloc[idx] != "":
            last_rel = idx
            break
    last_abs = header_row_idx if last_rel is None else content_start + last_rel

    next_no = None
    if no_col is not None and last_rel is not None:
        parsed = _extract_int(section_df.iloc[last_rel, no_col])
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
        "content_end": content_end,
        "name_col": name_col,
        "validation_col": validation_col,
        "no_col": no_col,
        "check_col": check_col,
        "header_excel_row": header_row_idx + 1,
        "section_first_excel_row": content_start + 1,
        "section_row_count": content_end - content_start,
        "insert_at_excel_row": last_abs + 2,
        "excel_names": excel_names,
        "validation_vals": validation_vals,
        "missing": missing,
        "next_no": next_no,
    }


def apply_row_offset(plan: dict, offset_rows: int) -> dict:
    if offset_rows == 0 or plan is None:
        return plan
    newp = dict(plan)
    for key in ("header_excel_row", "section_first_excel_row", "insert_at_excel_row"):
        newp[key] = plan[key] + offset_rows
    return newp


# ---------------- Main workflow ----------------

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
    if isinstance(sheet_arg, int) or (isinstance(sheet_arg, str) and sheet_arg.isdigit()):
        return wb.sheets[int(sheet_arg)]
    return wb.sheets[str(sheet_arg)]


def validate_and_write_both(
    excel_path: str,
    sheet_arg,
    xml_process_names: list[str],
    xml_object_names: list[str],
):
    # Read raw grid once to plan both sections
    sheet_arg_for_pd = int(sheet_arg) if (isinstance(sheet_arg, str) and sheet_arg.isdigit()) else sheet_arg
    df = pd.read_excel(excel_path, sheet_name=sheet_arg_for_pd, header=None, dtype=str, engine='openpyxl').fillna('')

    plan_A = plan_section(df, ["A.", "Process"], xml_process_names, label="Process",
                          header_keyword="name", next_keys=["B.", "Business Object"])
    plan_B = plan_section(df, ["B.", "Business Object"], xml_object_names, label="Business Object",
                          header_keyword="name", next_keys=None)
    if plan_A is None and plan_B is None:
        print("⛔ Stopping: neither section A nor B was detected.")
        return

    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(str(excel_path))
        src_sht = _sheet_by_name_or_index(wb, sheet_arg)

        src_sht.api.Copy(After=src_sht.api)
        ws = wb.sheets[-1]
        new_name = _unique_sheet_name(wb, f"{src_sht.name}_validated")
        try:
            ws.name = new_name
        except Exception as e:
            print(f"⚠️ Rename failed ({e}); keeping '{ws.name}'")

        placements = snapshot_and_set_placement(ws)

        rows_inserted_A = 0

        # ----- Section A -----
        if plan_A is not None:
            print("▶ Processing section A: Process")
            val_col_excel = plan_A["validation_col"] + 1

            # Header + borders like left
            style_validation_header(ws, plan_A["header_excel_row"], val_col_excel)

            # Existing rows: values + colors + borders like left
            if plan_A["section_row_count"] > 0:
                vals = [[v] for v in plan_A["validation_vals"]]
                ws.range((plan_A["section_first_excel_row"], val_col_excel)).options(
                    index=False, header=False
                ).value = vals
                color_validation_cells(
                    ws,
                    plan_A["section_first_excel_row"],
                    plan_A["validation_vals"],
                    val_col_excel,
                    plan_A["section_first_excel_row"] + plan_A["section_row_count"] - 1,
                )

            # Insert new rows contiguously
            if plan_A["missing"]:
                next_no = plan_A["next_no"]
                for i, name in enumerate(plan_A["missing"]):
                    r = plan_A["insert_at_excel_row"] + i
                    insert_row_with_style(ws, r)
                    ws.range((r, plan_A["name_col"] + 1)).value = name
                    ws.range((r, val_col_excel)).value = "Newly added"
                    ws.range((r, val_col_excel)).color = COLOR_ORANGE
                    _apply_borders_like_left(ws, r, val_col_excel)
                    if plan_A["no_col"] is not None and next_no is not None:
                        ws.range((r, plan_A["no_col"] + 1)).value = next_no
                        next_no += 1
                    if plan_A["check_col"] is not None:
                        ws.range((r, plan_A["check_col"] + 1)).color = None
                rows_inserted_A = len(plan_A["missing"])

            # Auto-fit the validation column
            try:
                ws.api.Columns(val_col_excel).AutoFit()
            except Exception:
                pass

        # ----- Section B -----
        if plan_B is not None:
            print("▶ Processing section B: Business Object")
            if plan_A is not None and plan_B["header_row_idx"] > plan_A["header_row_idx"]:
                plan_B = apply_row_offset(plan_B, rows_inserted_A)

            val_col_excel = plan_B["validation_col"] + 1

            style_validation_header(ws, plan_B["header_excel_row"], val_col_excel)

            if plan_B["section_row_count"] > 0:
                vals = [[v] for v in plan_B["validation_vals"]]
                ws.range((plan_B["section_first_excel_row"], val_col_excel)).options(
                    index=False, header=False
                ).value = vals
                color_validation_cells(
                    ws,
                    plan_B["section_first_excel_row"],
                    plan_B["validation_vals"],
                    val_col_excel,
                    plan_B["section_first_excel_row"] + plan_B["section_row_count"] - 1,
                )

            if plan_B["missing"]:
                next_no = plan_B["next_no"]
                for i, name in enumerate(plan_B["missing"]):
                    r = plan_B["insert_at_excel_row"] + i
                    insert_row_with_style(ws, r)
                    ws.range((r, plan_B["name_col"] + 1)).value = name
                    ws.range((r, val_col_excel)).value = "Newly added"
                    ws.range((r, val_col_excel)).color = COLOR_ORANGE
                    _apply_borders_like_left(ws, r, val_col_excel)
                    if plan_B["no_col"] is not None and next_no is not None:
                        ws.range((r, plan_B["no_col"] + 1)).value = next_no
                        next_no += 1
                    if plan_B["check_col"] is not None:
                        ws.range((r, plan_B["check_col"] + 1)).color = None

            # Auto-fit the validation column
            try:
                ws.api.Columns(val_col_excel).AutoFit()
            except Exception:
                pass

        restore_placement(ws, placements)

        wb.save()
        print(f"✅ Completed. Wrote Validation (colored, autofit) & appended new rows in '{ws.name}'.")
    finally:
        try:
            wb.close()
        except Exception:
            pass
        app.quit()


# ---------------- CLI ----------------

def main():
    parser = argparse.ArgumentParser(
        description="Validate A. Process (<process name>) and B. Business Object (<object name>) with colored & autofit Validation, preserving styling/objects."
    )
    parser.add_argument('--xml', required=True, help="Path to XML file")
    parser.add_argument('--excel', required=True, help="Path to Excel file (.xlsx/.xlsm)")
    parser.add_argument('--sheet', default="0", help="Sheet name or 0-based index (default: 0)")
    args = parser.parse_args()

    print("🔍 Parsing XML…")
    xml_process_names = extract_names_from_xml(args.xml, "process")
    xml_object_names  = extract_names_from_xml(args.xml, "object")
    print(f"✅ Found {len(xml_process_names)} process names; {len(xml_object_names)} object names.")

    print("🧪 Validating & writing (styling/objects preserved)…")
    validate_and_write_both(
        excel_path=args.excel,
        sheet_arg=args.sheet,
        xml_process_names=xml_process_names,
        xml_object_names=xml_object_names,
    )


if __name__ == "__main__":
    main()
