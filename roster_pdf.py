#!/usr/bin/env python3
"""
roster_pdf.py — Professional PDF generator for the TVET CDACC register extractor.

Provides two public entry-points:

  build_roster_pdf(data, path, course_filter=None, report_type_filter=None)
      Landscape A4 candidate-roster PDF, with optional filtering by course
      and/or report type (Assessment / Re-Assessment).

  build_summary_pdf(data, path)
      Portrait A4 summary PDF: centre info, statistics, and per-unit table.

Requires: reportlab  (pip install reportlab)
"""

from itertools import groupby

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.platypus import (
    BaseDocTemplate, Frame, HRFlowable, KeepTogether,
    NextPageTemplate, PageBreak, PageTemplate,
    Paragraph, Spacer, Table, TableStyle,
)

from extract_registers import build_roster


# ═══════════════════════════════════════════════════════════════════════════════
#  Fonts  (PDF standard — no download needed)
# ═══════════════════════════════════════════════════════════════════════════════

_F  = "Times-Roman"
_FB = "Times-Bold"
_FI = "Times-Italic"


# ═══════════════════════════════════════════════════════════════════════════════
#  Colours
# ═══════════════════════════════════════════════════════════════════════════════

GREEN    = colors.Color(0.1608, 0.6588, 0.3804)   # #29A861
DKGREEN  = colors.Color(0.1137, 0.4902, 0.2784)   # darker accent
WHITE    = colors.HexColor("#FFFFFF")
DKGRAY   = colors.Color(0.3294, 0.3294, 0.3294)   # #545454
LTGRAY   = colors.HexColor("#F4F4F4")
MIDGRAY  = colors.HexColor("#E0E0E0")
BLACK    = colors.HexColor("#000000")
BORDER_C = colors.HexColor("#BBBBBB")
NAVY     = colors.Color(0.094, 0.322, 0.498)       # dark header for summary


# ═══════════════════════════════════════════════════════════════════════════════
#  Shared paragraph styles
# ═══════════════════════════════════════════════════════════════════════════════

_COL_HDR = ParagraphStyle(
    "col_hdr",
    fontName=_FB, fontSize=9.5, leading=13,
    textColor=WHITE, alignment=TA_CENTER,
)
_CELL = ParagraphStyle(
    "cell",
    fontName=_F, fontSize=9.5, leading=13,
    textColor=DKGRAY, alignment=TA_LEFT,
    leftIndent=2, rightIndent=2,
)
_CELL_C = ParagraphStyle(
    "cell_c", parent=_CELL,
    alignment=TA_CENTER, leftIndent=0, rightIndent=0,
)


# ═══════════════════════════════════════════════════════════════════════════════
#  Page-counter canvas  — two-pass, injects "PAGE X of Y" after build
# ═══════════════════════════════════════════════════════════════════════════════

class _PageCounter(pdfcanvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states: list = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self, page_size=None):
        total = len(self._saved_page_states)
        ps = page_size or A4
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.setFont(_FB, 9)
            self.setFillColor(DKGRAY)
            self.drawCentredString(
                ps[0] / 2,
                6 * mm,
                f"Page {self._pageNumber} of {total}",
            )
            pdfcanvas.Canvas.showPage(self)
        pdfcanvas.Canvas.save(self)


class _LandscapeCounter(_PageCounter):
    """Page counter configured for landscape A4."""
    def save(self):
        super().save(page_size=landscape(A4))


class _PortraitCounter(_PageCounter):
    """Page counter configured for portrait A4."""
    def save(self):
        super().save(page_size=A4)


# ═══════════════════════════════════════════════════════════════════════════════
#  Shared helpers
# ═══════════════════════════════════════════════════════════════════════════════

def _fit_string(canvas, text: str, font: str, start_size: float,
                max_width: float, min_size: float = 7.0) -> float:
    """Set canvas font to largest size ≤ start_size that fits max_width.
    Returns the font size used."""
    size = start_size
    while size >= min_size:
        if canvas.stringWidth(text, font, size) <= max_width:
            break
        size -= 0.5
    canvas.setFont(font, size)
    return size


def _fmt_course(cname: str, clevel: str) -> str:
    if not clevel:
        return cname
    lvl = clevel if clevel.strip().lower().startswith("level") else f"Level {clevel}"
    return f"{cname}  {lvl}"


# ═══════════════════════════════════════════════════════════════════════════════
#  ──────────────────────────  ROSTER PDF  ─────────────────────────────────────
# ═══════════════════════════════════════════════════════════════════════════════

# ── Landscape geometry ────────────────────────────────────────────────────────
_L_PAGE     = landscape(A4)          # 841.89 × 595.28 pts
_L_MARGIN   = 14 * mm
_L_TOP_PAD  = 8  * mm               # gap: page top edge → header rectangle
_L_HEADER_H = 22 * mm
_L_HDR_GAP  = 4  * mm
_L_FOOTER_H = 12 * mm
_L_FTR_PAD  = 3  * mm
_L_W = _L_PAGE[0] - 2 * _L_MARGIN   # ≈ 762 pts

_L_FRAME_X = _L_MARGIN
_L_FRAME_Y = _L_FOOTER_H + _L_FTR_PAD
_L_FRAME_W = _L_W
_L_FRAME_H = (_L_PAGE[1]
              - (_L_TOP_PAD + _L_HEADER_H + _L_HDR_GAP)
              - (_L_FOOTER_H + _L_FTR_PAD))

# ── Column widths  S/N | REG NO | CANDIDATE NAME | UNITS | UNIT NAME(S) | SIG ─
_L_CM   = [12, 57, 50, 20, 88, 42]   # mm;  Σ = 269 mm
_L_COLS = [c * mm for c in _L_CM]

_FILLER_ROWS = 5


def _roster_col_header_row() -> list:
    labels = ["S/N", "REG NO", "CANDIDATE NAME", "UNITS", "UNIT NAME(S)", "SIGNATURE"]
    return [Paragraph(lbl, _COL_HDR) for lbl in labels]


def _roster_data_table(entries: list) -> Table:
    rows    = [_roster_col_header_row()]
    row_hts = [None]

    for sn, (reg, info) in enumerate(entries, start=1):
        unit_str = "; ".join(info["unit_names"])
        rows.append([
            Paragraph(str(sn),                 _CELL_C),
            Paragraph(reg,                     _CELL),
            Paragraph(info["name"],            _CELL),
            Paragraph(str(len(info["units"])), _CELL_C),
            Paragraph(unit_str,                _CELL),
            "",
        ])
        row_hts.append(None)

    for _ in range(_FILLER_ROWS):
        rows.append(["", "", "", "", "", ""])
        row_hts.append(28)

    tbl = Table(rows, colWidths=_L_COLS, rowHeights=row_hts,
                repeatRows=1, hAlign="LEFT")
    tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1,  0),  GREEN),
        ("TOPPADDING",    (0, 0), (-1,  0),  8),
        ("BOTTOMPADDING", (0, 0), (-1,  0),  8),
        ("VALIGN",        (0, 0), (-1,  0),  "MIDDLE"),
        ("BACKGROUND",    (0, 1), (-1, -1),  WHITE),
        ("TOPPADDING",    (0, 1), (-1, -1),  7),
        ("BOTTOMPADDING", (0, 1), (-1, -1),  7),
        ("VALIGN",        (0, 1), (-1, -1),  "MIDDLE"),
        ("LEFTPADDING",   (0, 0), (-1, -1),  6),
        ("RIGHTPADDING",  (0, 0), (-1, -1),  6),
        ("ALIGN",         (0, 1), (0,  -1),  "CENTER"),
        ("ALIGN",         (3, 1), (3,  -1),  "CENTER"),
        ("BOX",           (0, 0), (-1, -1),  0.6, BORDER_C),
        ("INNERGRID",     (0, 0), (-1, -1),  0.3, BORDER_C),
    ]))
    return tbl


_INST_DEPT = "ICT DEPARTMENT"


def _make_roster_page_fn(course_label: str, centre_name: str, centre_code: str = ""):
    """Return the onPage callback that draws the green header on the first page
    of each course group only.

    Left  — top:    centre name
             bottom: Centre Code: {code}
    Right — top:    CDACC EXAM REGISTRATION(S)
             bottom: course name for this specific group
    """
    _pad  = 5 * mm
    _gap  = 6 * mm
    _half = _L_W / 2 - _pad - _gap / 2

    _code_line = f"CENTRE CODE: {centre_code}".upper() if centre_code else ""

    def draw(canvas, doc):
        canvas.saveState()

        rect_bottom = _L_PAGE[1] - _L_TOP_PAD - _L_HEADER_H
        canvas.setFillColor(GREEN)
        canvas.rect(_L_MARGIN, rect_bottom, _L_W, _L_HEADER_H, fill=1, stroke=0)

        y1 = _L_PAGE[1] - _L_TOP_PAD - 8  * mm
        y2 = _L_PAGE[1] - _L_TOP_PAD - 16 * mm
        canvas.setFillColor(WHITE)

        x_left  = _L_MARGIN + _pad
        x_right = _L_MARGIN + _L_W - _pad

        # LEFT: centre name (bold 12) / centre code (regular 11)
        _fit_string(canvas, centre_name, _FB, 12, _half)
        canvas.drawString(x_left, y1, centre_name)

        if _code_line:
            _fit_string(canvas, _code_line, _F, 11, _half)
            canvas.drawString(x_left, y2, _code_line)

        # RIGHT: CDACC label (bold 12) / course name (regular 11)
        reg_label = "CDACC EXAM REGISTRATION(S)"
        _fit_string(canvas, reg_label, _FB, 12, _half)
        canvas.drawRightString(x_right, y1, reg_label)

        course_up = course_label.upper()
        _fit_string(canvas, course_up, _F, 11, _half)
        canvas.drawRightString(x_right, y2, course_up)

        canvas.restoreState()

    return draw


def build_roster_pdf(data: dict, path: str,
                     course_filter: list | None = None,
                     report_type_filter: str | None = None) -> str:
    """
    Build a landscape A4 candidate-roster PDF.

    Parameters
    ----------
    data              : dict from extract_registers.extract()
    path              : output file path
    course_filter     : None → all courses; [(name, level), ...] → selected only
    report_type_filter: None → all; "Assessment Registrations" or
                        "Re-Assessment Registrations" → filter by report type
    """
    # Apply report-type filter on units before building roster
    working_data = data
    if report_type_filter:
        filtered_units = [
            u for u in data["units"]
            if (u.get("report_type") or "").strip().lower()
               == report_type_filter.strip().lower()
        ]
        working_data = {**data, "units": filtered_units}

    roster = build_roster(working_data)

    # Sort by (course_name, course_level, reg_no) so that all candidates for
    # the same course form one contiguous block (plain reg_no sort interleaves
    # courses whose reg numbers overlap, breaking groupby into many fragments).
    sorted_entries = sorted(
        roster.items(),
        key=lambda x: (
            x[1].get("course_name", ""),
            x[1].get("course_level", ""),
            x[0],   # reg_no within the course
        ),
    )
    groups = [
        ((cn, cl), list(grp))
        for (cn, cl), grp in groupby(
            sorted_entries,
            key=lambda x: (x[1].get("course_name", ""),
                           x[1].get("course_level", "")),
        )
    ]

    if course_filter:
        cf_set = {tuple(c) for c in course_filter}
        groups = [g for g in groups if g[0] in cf_set]

    # Course label for the right-side header
    if len(groups) == 1:
        course_label = _fmt_course(*groups[0][0])
    elif groups:
        course_label = "  |  ".join(_fmt_course(*k) for k, _ in groups)
    else:
        course_label = data.get("course_name") or "Assessment Register"

    centre_name = (data.get("centre_name") or "TVET CDACC CENTRE").upper()
    centre_code = data.get("centre_code") or ""

    # Continuation frame: taller than first-page frame (no header space needed)
    _L_FRAME_H_CONT = _L_PAGE[1] - _L_TOP_PAD - (_L_FOOTER_H + _L_FTR_PAD)

    # ── Two PageTemplates per course: "first" (with header) + "cont" (no header)
    page_templates: list = []
    for i, ((cname, clevel), _entries) in enumerate(groups):
        course_lbl = _fmt_course(cname, clevel)
        page_fn    = _make_roster_page_fn(course_lbl, centre_name, centre_code)

        frame_first = Frame(
            _L_FRAME_X, _L_FRAME_Y, _L_FRAME_W, _L_FRAME_H,
            leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
        )
        frame_cont = Frame(
            _L_FRAME_X, _L_FRAME_Y, _L_FRAME_W, _L_FRAME_H_CONT,
            leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
        )
        page_templates.append(
            PageTemplate(id=f"course_{i}", frames=frame_first,
                         onPage=page_fn, pagesize=_L_PAGE)
        )
        page_templates.append(
            PageTemplate(id=f"course_{i}_cont", frames=frame_cont,
                         pagesize=_L_PAGE)
        )

    # Fallback templates (used if groups is empty)
    if not page_templates:
        fb_fn = _make_roster_page_fn(
            data.get("course_name") or "Assessment Register", centre_name, centre_code)
        page_templates.append(PageTemplate(
            id="course_0", pagesize=_L_PAGE,
            frames=Frame(_L_FRAME_X, _L_FRAME_Y, _L_FRAME_W, _L_FRAME_H,
                         leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0),
            onPage=fb_fn,
        ))
        page_templates.append(PageTemplate(
            id="course_0_cont", pagesize=_L_PAGE,
            frames=Frame(_L_FRAME_X, _L_FRAME_Y, _L_FRAME_W, _L_FRAME_H_CONT,
                         leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0),
        ))

    # ── Story: each course on its own first-page template; cont pages use _cont
    story: list = []
    for i, ((cname, clevel), entries) in enumerate(groups):
        if i > 0:
            story.append(NextPageTemplate(f"course_{i}"))
            story.append(PageBreak())
        # After the first page of this course, switch to the cont template
        story.append(NextPageTemplate(f"course_{i}_cont"))
        # Sort candidates by registration number
        entries_sorted = sorted(entries, key=lambda x: x[0])
        tbl = _roster_data_table(entries_sorted)
        story.append(tbl)

    # ── Document ──────────────────────────────────────────────────────────────
    doc = BaseDocTemplate(
        str(path),
        pagesize=_L_PAGE,
        leftMargin=_L_MARGIN, rightMargin=_L_MARGIN,
        topMargin=_L_TOP_PAD + _L_HEADER_H + _L_HDR_GAP,
        bottomMargin=_L_FOOTER_H + _L_FTR_PAD,
        title="CDACC Assessment Register – Attendance Roster",
        author=centre_name,
    )
    doc.addPageTemplates(page_templates)
    doc.build(story, canvasmaker=_LandscapeCounter)
    return str(path)


# ═══════════════════════════════════════════════════════════════════════════════
#  ──────────────────────────  SUMMARY PDF  ────────────────────────────────────
# ═══════════════════════════════════════════════════════════════════════════════

# ── Portrait geometry ─────────────────────────────────────────────────────────
_P_PAGE     = A4                          # 595.28 × 841.89 pts
_P_MARGIN   = 12 * mm
_P_TOP_PAD  = 8  * mm               # gap: page top edge → header rectangle
_P_HEADER_H = 22 * mm
_P_HDR_GAP  = 6  * mm
_P_FOOTER_H = 12 * mm
_P_FTR_PAD  = 3  * mm
_P_W = _P_PAGE[0] - 2 * _P_MARGIN        # ≈ 555 pts

_P_FRAME_X = _P_MARGIN
_P_FRAME_Y = _P_FOOTER_H + _P_FTR_PAD
_P_FRAME_W = _P_W
_P_FRAME_H = (_P_PAGE[1]
              - (_P_TOP_PAD + _P_HEADER_H + _P_HDR_GAP)
              - (_P_FOOTER_H + _P_FTR_PAD))

# ── Summary paragraph styles ──────────────────────────────────────────────────
_SM_TITLE = ParagraphStyle(
    "sm_title",
    fontName=_FB, fontSize=13, leading=17,
    textColor=BLACK, alignment=TA_CENTER, spaceAfter=4,
)
_SM_SECTION = ParagraphStyle(
    "sm_section",
    fontName=_FB, fontSize=10.5, leading=14,
    textColor=BLACK, spaceBefore=8, spaceAfter=3,
)
_SM_LABEL = ParagraphStyle(
    "sm_label",
    fontName=_FB, fontSize=9.5, leading=13,
    textColor=BLACK,
)
_SM_VALUE = ParagraphStyle(
    "sm_value",
    fontName=_F, fontSize=9.5, leading=13,
    textColor=DKGRAY,
)
_SM_COL_H = ParagraphStyle(
    "sm_col_h",
    fontName=_FB, fontSize=7, leading=9,
    textColor=WHITE, alignment=TA_CENTER,
)
_SM_CELL = ParagraphStyle(
    "sm_cell",
    fontName=_F, fontSize=9, leading=12,
    textColor=DKGRAY, alignment=TA_LEFT,
    leftIndent=2, rightIndent=2,
)
_SM_CELL_C = ParagraphStyle(
    "sm_cell_c", parent=_SM_CELL,
    alignment=TA_CENTER, leftIndent=0, rightIndent=0,
)
_SM_CELL_TOTAL = ParagraphStyle(
    "sm_cell_total",
    fontName=_FB, fontSize=9, leading=12,
    textColor=BLACK, alignment=TA_RIGHT,
    leftIndent=0, rightIndent=4,
)
_SM_CELL_TOTAL_N = ParagraphStyle(
    "sm_cell_total_n", parent=_SM_CELL_TOTAL,
    alignment=TA_CENTER, rightIndent=0,
)


def _make_summary_page_fn(centre_name: str, centre_code: str, series: str = ""):
    """Return the onPage callback that draws the GREEN header on the first
    summary page only.

    Top row  — centre name, centred across full width, bold, dynamic font (max 13)
    Bottom row — Centre Code (left) | series (right)
    """
    _pad       = 5 * mm
    _gap       = 6 * mm
    _half      = _P_W / 2 - _pad - _gap / 2
    _full_name = _P_W - 2 * _pad   # full usable width for the centred name
    _x_centre  = _P_MARGIN + _P_W / 2

    _code_line   = f"CENTRE CODE: {centre_code}".upper() if centre_code else ""
    _series_line = series.upper() if series else ""

    def draw(canvas, doc):
        canvas.saveState()

        rect_bottom = _P_PAGE[1] - _P_TOP_PAD - _P_HEADER_H
        canvas.setFillColor(GREEN)
        canvas.rect(_P_MARGIN, rect_bottom, _P_W, _P_HEADER_H, fill=1, stroke=0)

        y1 = _P_PAGE[1] - _P_TOP_PAD - 8  * mm
        y2 = _P_PAGE[1] - _P_TOP_PAD - 16 * mm
        canvas.setFillColor(WHITE)

        x_left  = _P_MARGIN + _pad
        x_right = _P_MARGIN + _P_W - _pad

        # TOP: centre name centred, dynamic font max 13
        _fit_string(canvas, centre_name, _FB, 13, _full_name)
        canvas.drawCentredString(_x_centre, y1, centre_name)

        # BOTTOM: centre code (left) / series (right)
        if _code_line:
            _fit_string(canvas, _code_line, _F, 10, _half)
            canvas.drawString(x_left, y2, _code_line)

        if _series_line:
            _fit_string(canvas, _series_line, _F, 10, _half)
            canvas.drawRightString(x_right, y2, _series_line)

        canvas.restoreState()

    return draw


def _info_table(rows_data: list, col_widths: list) -> Table:
    """Two-column label/value table for the info block."""
    tbl = Table(rows_data, colWidths=col_widths, hAlign="LEFT")
    tbl.setStyle(TableStyle([
        ("FONTNAME",      (0, 0), (0, -1),  _FB),
        ("FONTNAME",      (1, 0), (1, -1),  _F),
        ("FONTSIZE",      (0, 0), (-1, -1), 9.5),
        ("LEADING",       (0, 0), (-1, -1), 14),
        ("TEXTCOLOR",     (0, 0), (0, -1),  BLACK),
        ("TEXTCOLOR",     (1, 0), (1, -1),  DKGRAY),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (0, -1),  8),
        ("LEFTPADDING",   (1, 0), (1, -1),  8),
        ("BACKGROUND",    (0, 0), (-1, -1), colors.transparent),
        ("BOX",           (0, 0), (-1, -1), 0.8, BORDER_C),
        ("INNERGRID",     (0, 0), (-1, -1), 0.5, BORDER_C),
    ]))
    return tbl


def build_summary_pdf(data: dict, path: str) -> str:
    """
    Build a portrait A4 summary PDF.

    Shows centre info, statistics (total / per-type breakdown), and a
    per-unit table with candidate counts.
    """
    roster     = build_roster(data)
    all_units  = data["units"]
    total_regs = sum(u["candidate_count"] for u in all_units)

    assess_units = [u for u in all_units
                    if (u.get("report_type") or "").strip().lower().startswith("assessment")]
    reassess_units = [u for u in all_units
                      if (u.get("report_type") or "").strip().lower().startswith("re")]
    assess_regs   = sum(u["candidate_count"] for u in assess_units)
    reassess_regs = sum(u["candidate_count"] for u in reassess_units)

    # Distinct courses (ordered by first appearance)
    seen_courses: list = list(dict.fromkeys(
        (u.get("course_name", ""), u.get("course_level", ""))
        for u in all_units
        if u.get("course_name", "").strip()
    ))

    centre_name = (data.get("centre_name") or "—").upper()
    centre_code = data.get("centre_code") or "—"
    series      = data.get("series") or ""

    # Sort units alphabetically (by course → level → unit name)
    all_units = sorted(
        all_units,
        key=lambda u: (
            u.get("course_name",  "").lower(),
            u.get("course_level", "").lower(),
            u.get("unit_name",    "").lower(),
        ),
    )

    # ── label / value column widths ───────────────────────────────────────────
    lw = 44 * mm
    vw = _P_W - lw

    # ── story ─────────────────────────────────────────────────────────────────
    # Continuation pages have no header so their frame is taller
    story: list = [NextPageTemplate("Cont")]

    # ── Section: Summary Statistics (Course(s) first, then numeric stats) ────────
    story.append(Paragraph("Summary Statistics", _SM_SECTION))

    _val_style = ParagraphStyle(
        "sm_val_wrap",
        fontName=_F, fontSize=9.5, leading=13, textColor=DKGRAY,
    )
    course_rows = []
    for i, (cn, cl) in enumerate(seen_courses or [("—", "")]):
        label = "Course(s):" if i == 0 else ""
        course_rows.append([label, Paragraph(_fmt_course(cn, cl), _val_style)])

    stats_rows = course_rows + [
        ["Total Units:",                   str(data.get("unit_count", len(all_units)))],
        ["Unique Candidates:",             str(len(roster))],
        ["Total Registrations:",           str(total_regs)],
        ["Assessment Registrations:",      str(assess_regs)],
        ["Re-Assessment Registrations:",   str(reassess_regs)],
    ]
    story.append(_info_table(stats_rows, [lw, vw]))
    story.append(Spacer(1, 6 * mm))

    # ── Section 3: Unit Details Table ─────────────────────────────────────────
    story.append(Paragraph("Unit Details", _SM_SECTION))

    # Column widths: Code | Unit Name | Report Type | Count
    # Total = _P_W
    uc_w  = 50 * mm
    rt_w  = 38 * mm
    cnt_w = 18 * mm
    un_w  = _P_W - uc_w - rt_w - cnt_w

    unit_rows = [[
        Paragraph("UNIT CODE",    _SM_COL_H),
        Paragraph("UNIT NAME",    _SM_COL_H),
        Paragraph("REPORT TYPE",  _SM_COL_H),
        Paragraph("CANDIDATES",   _SM_COL_H),
    ]]

    prev_course = None
    for u in all_units:
        course_key = (u.get("course_name", ""), u.get("course_level", ""))

        # Insert course-group separator row when course changes
        if len(seen_courses) > 1 and course_key != prev_course:
            cname, clevel = course_key
            grp_label = _fmt_course(cname, clevel) if cname else "Unknown Course"
            separator = Paragraph(f"<b>{grp_label}</b>", ParagraphStyle(
                "sep", fontName=_FB, fontSize=8.5, leading=12,
                textColor=BLACK, alignment=TA_CENTER, leftIndent=0,
            ))
            unit_rows.append([separator, "", "", ""])
            prev_course = course_key

        rt_raw = (u.get("report_type") or "").strip()
        if rt_raw.lower().startswith("re"):
            rt_short = "Re-Assessment"
        elif rt_raw.lower().startswith("assessment"):
            rt_short = "Assessment"
        else:
            rt_short = rt_raw[:14]

        unit_rows.append([
            Paragraph(u["unit_code"],   _SM_CELL),
            Paragraph(u["unit_name"],   _SM_CELL),
            Paragraph(rt_short,         _SM_CELL_C),
            Paragraph(str(u["candidate_count"]), _SM_CELL_C),
        ])

    # Totals row
    unit_rows.append([
        Paragraph("", _SM_CELL),
        Paragraph("TOTAL", _SM_CELL_TOTAL),
        Paragraph("", _SM_CELL),
        Paragraph(str(total_regs), _SM_CELL_TOTAL_N),
    ])

    n_data = len(unit_rows) - 1   # exclude header
    n_seps = sum(
        1 for row in unit_rows[1:]
        if isinstance(row[1], str) and row[1] == ""
        and isinstance(row[0], Paragraph) and row[0].text.startswith("<b>")
    )

    # Build row styles dynamically
    tbl_style = [
        ("BACKGROUND",    (0, 0), (-1,  0),  GREEN),
        ("TOPPADDING",    (0, 0), (-1,  0),  4),
        ("BOTTOMPADDING", (0, 0), (-1,  0),  4),
        ("VALIGN",        (0, 0), (-1,  0),  "MIDDLE"),
        ("BACKGROUND",    (0, 1), (-1, -2),  WHITE),
        ("TOPPADDING",    (0, 1), (-1, -1),  5),
        ("BOTTOMPADDING", (0, 1), (-1, -1),  5),
        ("VALIGN",        (0, 1), (-1, -1),  "MIDDLE"),
        ("LEFTPADDING",   (0, 0), (-1, -1),  5),
        ("RIGHTPADDING",  (0, 0), (-1, -1),  5),
        ("BOX",           (0, 0), (-1, -1),  0.6, BORDER_C),
        ("INNERGRID",     (0, 0), (-1, -1),  0.3, BORDER_C),
        # Total row
        ("BACKGROUND",    (0, -1), (-1, -1), LTGRAY),
        ("FONTNAME",      (0, -1), (-1, -1), _FB),
    ]

    # Shade alternating data rows (skip separator rows)
    data_row_idx = 0
    for ri, row in enumerate(unit_rows[1:], start=1):
        is_sep = (isinstance(row[1], str) and row[1] == ""
                  and not isinstance(row[0], str))
        if is_sep:
            tbl_style.append(("BACKGROUND", (0, ri), (-1, ri), colors.Color(0.93, 0.97, 0.94)))
            tbl_style.append(("SPAN",       (0, ri), (-1, ri)))
        else:
            if data_row_idx % 2 == 1 and ri < len(unit_rows) - 1:
                tbl_style.append(("BACKGROUND", (0, ri), (-1, ri), LTGRAY))
            data_row_idx += 1

    unit_tbl = Table(
        unit_rows,
        colWidths=[uc_w, un_w, rt_w, cnt_w],
        repeatRows=1, hAlign="LEFT",
    )
    unit_tbl.setStyle(TableStyle(tbl_style))
    story.append(unit_tbl)

    # ── Document ──────────────────────────────────────────────────────────────
    _P_FRAME_H_CONT = _P_PAGE[1] - _P_TOP_PAD - (_P_FOOTER_H + _P_FTR_PAD)

    page_fn = _make_summary_page_fn(centre_name, centre_code, series)

    frame_first = Frame(
        _P_FRAME_X, _P_FRAME_Y, _P_FRAME_W, _P_FRAME_H,
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
    )
    frame_cont = Frame(
        _P_FRAME_X, _P_FRAME_Y, _P_FRAME_W, _P_FRAME_H_CONT,
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
    )

    doc = BaseDocTemplate(
        str(path),
        pagesize=_P_PAGE,
        leftMargin=_P_MARGIN, rightMargin=_P_MARGIN,
        topMargin=_P_TOP_PAD + _P_HEADER_H + _P_HDR_GAP,
        bottomMargin=_P_FOOTER_H + _P_FTR_PAD,
        title="CDACC Assessment Register – Summary",
        author=centre_name,
    )
    doc.addPageTemplates([
        PageTemplate(id="First", frames=frame_first, onPage=page_fn, pagesize=_P_PAGE),
        PageTemplate(id="Cont",  frames=frame_cont,  pagesize=_P_PAGE),
    ])

    doc.build(story, canvasmaker=_PortraitCounter)
    return str(path)
