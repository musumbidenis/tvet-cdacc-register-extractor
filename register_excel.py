#!/usr/bin/env python3
"""
register_excel.py
=================
Turn the structured data from `extract_registers.extract()` into a formatted
multi-sheet Excel workbook.

Sheets produced:
  * Summary        – centre info + unit table grouped by course (if > 1 course)
  * All Candidates – flat table sorted by Reg No ASC, one row per
                     (unit, candidate); ready for filters
                     Columns: Centre Code | Course | Level | Report Type |
                              Reg No | Candidate Name | Unit Code | Unit Name
  * Roster         – unique candidates sorted by Reg No ASC, one row per person
                     Columns: Reg No | Candidate Name | Units Registered |
                              Unit Name(s) | Signature

Formatting:
  - Font         : Times New Roman 12 pt, Automatic (black) colour throughout
  - Row height   : 20 pts (title rows 24 pts)
  - Alignment    : every row vertically centred; column headers horizontally
                   centred; data cells left-aligned (except numeric/count
                   columns which are horizontally centred)
  - Column width : auto-fitted to the longest value in each column + 4 chars
                   padding; info-header and multi-col merged rows are excluded
                   from the measurement so they don't bloat data columns
  - Info header  : All Candidates and Roster each have a 4-row block
                   (title centred | Centre | Course(s) | spacer) above the
                   table; the title row is merged and centred
  - Multi-course : if the PDF covers more than one course, the Summary unit
                   table is grouped with a highlighted course-separator row;
                   the Course meta row shows all courses separated by "; "
"""

import re
from itertools import groupby

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from extract_registers import extract, build_roster


# ── Style constants ───────────────────────────────────────────────────────────

FONT  = "Times New Roman"
ROW_H = 20          # standard row height (Excel units ≈ points)
HDR_H = 24          # taller title row

HEADER_FILL     = PatternFill("solid", start_color="BDD7EE")   # soft blue
COURSE_GRP_FILL = PatternFill("solid", start_color="E2EFDA")   # soft green

# No explicit color= → openpyxl leaves it as Automatic (black)
TITLE_FONT  = Font(name=FONT, bold=True, size=14)
LABEL_FONT  = Font(name=FONT, bold=True, size=12)
HEADER_FONT = Font(name=FONT, bold=True, size=12)
BASE_FONT   = Font(name=FONT,            size=12)

THIN   = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MID_LEFT   = Alignment(horizontal="left",   vertical="center")
MID_CENTER = Alignment(horizontal="center", vertical="center")


# ── Helpers ───────────────────────────────────────────────────────────────────

def _safe_sheet_name(name, used):
    """Excel sheet names: ≤31 chars, no : \\ / ? * [ ], must be unique."""
    clean = re.sub(r"[:\\/?*\[\]]", "_", name)[:31].strip() or "Sheet"
    candidate, n = clean, 1
    while candidate.lower() in used:
        suffix = f"_{n}"
        candidate = clean[: 31 - len(suffix)] + suffix
        n += 1
    used.add(candidate.lower())
    return candidate


def _style_header_row(ws, row, ncols):
    """Apply header fill / font / centre-alignment / border and set row height."""
    ws.row_dimensions[row].height = ROW_H
    for c in range(1, ncols + 1):
        cell           = ws.cell(row=row, column=c)
        cell.fill      = HEADER_FILL
        cell.font      = HEADER_FONT
        cell.alignment = MID_CENTER
        cell.border    = BORDER


def _style_data_rows(ws, first_row, last_row, ncols, center_cols=()):
    """
    Apply base font / border / row-height to data rows.
    Columns in *center_cols* (1-based) receive MID_CENTER; all others MID_LEFT.
    """
    center_set = set(center_cols)
    for r in range(first_row, last_row + 1):
        ws.row_dimensions[r].height = ROW_H
        for c in range(1, ncols + 1):
            cell           = ws.cell(row=r, column=c)
            cell.font      = BASE_FONT
            cell.alignment = MID_CENTER if c in center_set else MID_LEFT
            cell.border    = BORDER


def _autofit(ws, start_row, padding=4):
    """
    Set each column's width = longest cell value (from *start_row* onward)
    plus *padding* character widths of white-space.

    Rows above *start_row* and cells that are the top-left of a multi-column
    merged range are excluded so that merged title / label rows do not inflate
    the data-column widths.  Formula strings are also skipped.
    """
    # Collect (row, col) of cells that are the anchor of a multi-column merge
    multi_col_anchors = set()
    for mc in ws.merged_cells.ranges:
        if mc.max_col > mc.min_col:
            multi_col_anchors.add((mc.min_row, mc.min_col))

    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = 0
        for cell in col_cells:
            if cell.row < start_row or cell.value is None:
                continue
            if (cell.row, cell.column) in multi_col_anchors:
                continue            # skip merged-spanning cell
            val = str(cell.value)
            if val.startswith("="):
                continue
            max_len = max(max_len, len(val))
        if max_len > 0:
            ws.column_dimensions[col_letter].width = max_len + padding


def _course_info(data):
    """
    Inspect every unit's course_name and return (label, display_value).
      Single course  → ("Course:", "COMPUTER SCIENCE")
      Multiple       → ("Courses:", "COMPUTER SCIENCE; ELECTRICAL ENGINEERING")
    """
    seen = list(dict.fromkeys(
        (u.get("course_name") or "").strip()
        for u in data["units"]
        if (u.get("course_name") or "").strip()
    ))
    if not seen:
        seen = [(data.get("course_name") or "—").strip()]
    label = "Courses:" if len(seen) > 1 else "Course:"
    return label, "; ".join(seen)


def _group_units_by_course(data):
    """
    Return a list of ((course_name, course_level), [unit, ...]) in document
    order, preserving the first-seen sequence for each course pair.
    """
    groups: dict = {}
    for u in data["units"]:
        key = (u.get("course_name", ""), u.get("course_level", ""))
        groups.setdefault(key, []).append(u)
    return list(groups.items())


def _info_header(ws, data, title, ncols):
    """
    Write a 4-row info block into rows 1-4:
        Row 1  – sheet title (merged across all columns, centred)
        Row 2  – Centre label / value
        Row 3  – Course(s) label / value  (all courses if > 1)
        Row 4  – blank spacer (height = 8)

    Value cells in rows 2-3 are merged from column 2 to *ncols* so that
    long strings don't inflate the data-column widths.

    Returns 5 — the row where the column-header row should be written.
    """
    last_col = get_column_letter(ncols)

    # Row 1 – title (merged + centred)
    ws.row_dimensions[1].height = HDR_H
    cell = ws.cell(row=1, column=1, value=title)
    cell.font      = TITLE_FONT
    cell.alignment = MID_CENTER          # centred as requested
    if ncols > 1:
        ws.merge_cells(f"A1:{last_col}1")

    # Rows 2-3 – metadata
    course_label, course_value = _course_info(data)
    meta_rows = [
        (2, "Centre:", f"{data['centre_name']}  ({data['centre_code']})"),
        (3, course_label, course_value),
    ]
    for row, label, value in meta_rows:
        ws.row_dimensions[row].height = ROW_H
        lc = ws.cell(row=row, column=1, value=label)
        lc.font      = LABEL_FONT
        lc.alignment = MID_LEFT
        vc = ws.cell(row=row, column=2, value=value)
        vc.font      = BASE_FONT
        vc.alignment = MID_LEFT
        if ncols > 2:
            ws.merge_cells(f"B{row}:{last_col}{row}")

    # Row 4 – blank spacer
    ws.row_dimensions[4].height = 8

    return 5   # column-header row index


# ── Workbook builder ──────────────────────────────────────────────────────────

def build_workbook(data, path):
    wb = Workbook()
    used_names = set()

    # ── All Candidates ────────────────────────────────────────────────────────
    ac       = wb.active
    ac.title = _safe_sheet_name("All Candidates", used_names)
    ac_cols  = ["Centre Code", "Course", "Level", "Report Type",
                "Reg No", "Candidate Name", "Unit Code", "Unit Name"]
    ncols    = len(ac_cols)

    hdr_row = _info_header(ac, data, "Assessment Register – All Candidates", ncols)

    for ci, h in enumerate(ac_cols, start=1):
        ac.cell(row=hdr_row, column=ci, value=h)
    _style_header_row(ac, hdr_row, ncols)

    # Collect all rows, sort by Reg No (index 4) ascending, then write
    all_rows = []
    for u in data["units"]:
        for c in u["candidates"]:
            all_rows.append([
                data["centre_code"], u["course_name"], u["course_level"],
                u["report_type"], c["reg_no"], c["name"],
                u["unit_code"], u["unit_name"],
            ])
    all_rows.sort(key=lambda x: x[4])   # Reg No

    r = hdr_row + 1
    first_data = r
    for row_vals in all_rows:
        for ci, val in enumerate(row_vals, start=1):
            ac.cell(row=r, column=ci, value=val)
        r += 1
    last_data = r - 1

    if last_data >= first_data:
        _style_data_rows(ac, first_data, last_data, ncols)
    ac.freeze_panes = f"A{first_data}"
    ac.auto_filter.ref = f"A{hdr_row}:{get_column_letter(ncols)}{last_data}"
    _autofit(ac, start_row=hdr_row)
    ac_name = ac.title

    # ── Roster ────────────────────────────────────────────────────────────────
    roster   = build_roster(data)
    rs       = wb.create_sheet(_safe_sheet_name("Roster", used_names))
    # Col 1 = Course (merged per course group for visual/print grouping)
    # Cols 2-6 = per-candidate data
    rs_cols  = ["Course", "Reg No", "Candidate Name", "Units Registered",
                "Unit Name(s)", "Signature"]
    rs_ncols = len(rs_cols)

    rs_hdr = _info_header(rs, data, "Assessment Register – Roster", rs_ncols)

    for ci, h in enumerate(rs_cols, start=1):
        rs.cell(row=rs_hdr, column=ci, value=h)
    _style_header_row(rs, rs_hdr, rs_ncols)

    # Sort by Reg No ASC, then group consecutive same-course runs so we can
    # merge column 1 across every candidate that belongs to the same course.
    sorted_entries = sorted(roster.items())   # (reg_no, info) sorted by reg_no
    r        = rs_hdr + 1
    rs_first = r

    for (cname, clevel), grp_iter in groupby(
            sorted_entries,
            key=lambda x: (x[1].get("course_name", ""),
                           x[1].get("course_level", ""))):

        entries     = list(grp_iter)
        group_start = r

        # ── data columns 2-6 for every candidate in this group ──────────────
        for reg, info in entries:
            rs.cell(row=r, column=2, value=reg)
            rs.cell(row=r, column=3, value=info["name"])
            rs.cell(row=r, column=4, value=len(info["units"]))
            rs.cell(row=r, column=5, value="; ".join(info["unit_names"]))
            rs.cell(row=r, column=6, value="")
            r += 1

        group_end = r - 1

        # ── column 1: merged cell spanning the whole group ───────────────────
        if group_end > group_start:
            rs.merge_cells(f"A{group_start}:A{group_end}")

        # Build label: "Computer Science  Level 6"
        # Guard against levels that already begin with the word "Level"
        if clevel:
            lvl_part = (clevel if clevel.strip().lower().startswith("level")
                        else f"Level {clevel}")
            course_label = f"{cname}  {lvl_part}".strip()
        else:
            course_label = cname

        cc           = rs.cell(row=group_start, column=1, value=course_label)
        cc.font      = LABEL_FONT
        # text_rotation=90 → reads bottom-to-top (standard vertical text in Excel)
        cc.alignment = Alignment(horizontal="center", vertical="center",
                                 text_rotation=90)
        # Border on every row of the merged range (outer edges render correctly)
        for ri in range(group_start, group_end + 1):
            rs.cell(row=ri, column=1).border = BORDER

    rs_last = r - 1

    # Style data columns 2-6; col 4 (Units Registered) is centred
    for ri in range(rs_first, rs_last + 1):
        rs.row_dimensions[ri].height = ROW_H
        for c in range(2, rs_ncols + 1):
            cell           = rs.cell(row=ri, column=c)
            cell.font      = BASE_FONT
            cell.alignment = MID_CENTER if c == 4 else MID_LEFT
            cell.border    = BORDER

    # Freeze rows (info + header) AND column A (course label stays visible
    # while scrolling right)
    rs.freeze_panes = f"B{rs_first}"
    # Filter on columns B-F only (col A has merged cells, not filterable)
    rs.auto_filter.ref = f"B{rs_hdr}:{get_column_letter(rs_ncols)}{rs_last}"
    _autofit(rs, start_row=rs_hdr)
    # Override: rotated text only needs enough width for the font height, not
    # the full string length that _autofit would calculate.
    rs.column_dimensions["A"].width = 8
    roster_name = rs.title

    # ── Summary ───────────────────────────────────────────────────────────────
    sm = wb.create_sheet(_safe_sheet_name("Summary", used_names))

    # Title row – merged + centred
    sm.row_dimensions[1].height = HDR_H
    tc = sm.cell(row=1, column=1, value="Assessment Registration Summary")
    tc.font      = TITLE_FONT
    tc.alignment = MID_CENTER
    sm.merge_cells("A1:D1")

    # Meta block
    course_label, course_value = _course_info(data)
    meta = [
        ("Centre",                   f"{data['centre_name']} ({data['centre_code']})"),
        (course_label.rstrip(":"),   course_value),
        ("Units",                    data["unit_count"]),
    ]
    r = 3
    for label, value in meta:
        sm.row_dimensions[r].height = ROW_H
        lc = sm.cell(row=r, column=1, value=label)
        lc.font = LABEL_FONT; lc.alignment = MID_LEFT
        vc = sm.cell(row=r, column=2, value=value)
        vc.font = BASE_FONT; vc.alignment = MID_LEFT
        r += 1

    # Unique-candidates count — anchored to Roster data rows so the
    # info-header block in that sheet is not included in the count.
    sm.row_dimensions[r].height = ROW_H
    sm.cell(row=r, column=1, value="Unique candidates").font      = LABEL_FONT
    sm.cell(row=r, column=1).alignment                            = MID_LEFT
    sm.cell(row=r, column=2,
            value=f"=COUNTA('{roster_name}'!B{rs_first}:B1048576)").font = BASE_FONT
    sm.cell(row=r, column=2).alignment                            = MID_LEFT
    r += 2

    # Unit table header
    sm_hdr = r
    for ci, h in enumerate(["Unit Code", "Unit Name", "Report Type", "Candidates"],
                            start=1):
        sm.cell(row=sm_hdr, column=ci, value=h)
    _style_header_row(sm, sm_hdr, 4)
    r += 1

    # Unit data rows, optionally grouped by course
    first_sm_data = r
    groups    = _group_units_by_course(data)
    is_multi  = len(groups) > 1

    for (cname, clevel), units in groups:
        if is_multi:
            # Course group separator row (merged A:D, highlighted)
            sm.row_dimensions[r].height = ROW_H
            lvl_part  = clevel if clevel.strip().lower().startswith("level") else f"Level {clevel}"
            grp_label = f"{cname}  –  {lvl_part}" if clevel else cname
            gc = sm.cell(row=r, column=1, value=grp_label)
            gc.font      = LABEL_FONT
            gc.alignment = MID_CENTER
            gc.fill      = COURSE_GRP_FILL
            sm.merge_cells(f"A{r}:D{r}")
            for c in range(1, 5):
                sm.cell(row=r, column=c).border = BORDER
            r += 1

        for u in units:
            sm.row_dimensions[r].height = ROW_H
            # Unit Code, Unit Name, Report Type – left-aligned
            for c, val in enumerate(
                    [u["unit_code"], u["unit_name"], u["report_type"]], start=1):
                cell           = sm.cell(row=r, column=c, value=val)
                cell.font      = BASE_FONT
                cell.alignment = MID_LEFT
                cell.border    = BORDER
            # Candidates count formula – centred
            cnt            = sm.cell(row=r, column=4,
                                     value=f"=COUNTIF('{ac_name}'!G:G,A{r})")
            cnt.font       = BASE_FONT
            cnt.alignment  = MID_CENTER
            cnt.border     = BORDER
            r += 1

    # Total row
    total_row = r
    sm.row_dimensions[total_row].height = ROW_H
    for c in range(1, 5):
        sm.cell(row=total_row, column=c).border = BORDER
    tc3            = sm.cell(row=total_row, column=3, value="TOTAL")
    tc3.font       = LABEL_FONT
    tc3.alignment  = MID_LEFT
    tc4            = sm.cell(row=total_row, column=4,
                             value=f"=SUM(D{first_sm_data}:D{total_row - 1})")
    tc4.font       = LABEL_FONT
    tc4.alignment  = MID_CENTER    # centred to match Candidates column above

    sm.freeze_panes = f"A{first_sm_data}"
    _autofit(sm, start_row=sm_hdr)

    # Put Summary first in tab order
    wb.move_sheet(sm, -(wb.sheetnames.index(sm.title)))

    wb.save(path)
    return path


def pdf_to_excel(pdf_path, xlsx_path):
    """Convenience one-shot: parse a register PDF straight into a workbook."""
    return build_workbook(extract(pdf_path), xlsx_path)


if __name__ == "__main__":
    import sys
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else "registers.xlsx"
    print("Wrote", pdf_to_excel(src, dst))
