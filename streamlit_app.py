#!/usr/bin/env python3
"""
streamlit_app.py — Web UI for the TVET CDACC Register Extractor.

Run locally:
    streamlit run streamlit_app.py

Deploy free:
    Push to GitHub → connect repo on https://share.streamlit.io

Requires: streamlit, pdfplumber, reportlab, openpyxl
"""

import os
import tempfile
from datetime import datetime

import streamlit as st

from extract_registers import extract, build_roster
from register_excel import build_workbook

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TVET CDACC Register Extractor",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

_VERSION = "v2.0"

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Hide all Streamlit chrome ──────────────────────────────────────────── */
#MainMenu, footer,
[data-testid="stDeployButton"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
[data-testid="stAppCreatorBadge"],
.viewerBadge_container__r5tak,
.viewerBadge_link__qRIco,
iframe[title="st_app_creator_badge"] { display: none !important; }
a[href*="streamlit.io"]              { display: none !important; }

/* ── Layout ─────────────────────────────────────────────────────────────── */
.block-container {
    padding-top: 0 !important;
    padding-bottom: 2rem;
    max-width: 1150px;
}

/* ── App header ─────────────────────────────────────────────────────────── */
.app-header {
    background: linear-gradient(135deg, #1a252f 0%, #2c3e50 100%);
    padding: 14px 26px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    /* bleed to page edges */
    margin-left:  -5rem;
    margin-right: -5rem;
    margin-bottom: 1.4rem;
}
.app-header h1 {
    margin: 0; padding: 0;
    font-size: 1.3rem;
    font-weight: 700;
    color: #fff;
    letter-spacing: -0.2px;
}
.app-header .ver {
    font-size: 0.7rem;
    color: rgba(255,255,255,0.45);
    align-self: flex-start;
    margin-top: 3px;
    white-space: nowrap;
}

/* ── Section label ──────────────────────────────────────────────────────── */
.sec-lbl {
    font-size: 0.67rem;
    font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: #6B7280;
    margin: 0 0 7px;
    padding: 0;
}

/* ── Info grid ──────────────────────────────────────────────────────────── */
.info-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 10px;
    margin-bottom: 0.9rem;
}
.info-card {
    background: #F8FAFB;
    border: 1px solid rgba(0,0,0,0.07);
    border-radius: 8px;
    padding: 11px 14px;
}
.info-card .lbl {
    font-size: 0.65rem;
    font-weight: 700;
    letter-spacing: 0.8px;
    text-transform: uppercase;
    color: #6B7280;
    margin-bottom: 4px;
}
.info-card .val {
    font-size: 0.9rem;
    font-weight: 600;
    color: #111827;
    line-height: 1.35;
    word-break: break-word;
}

/* ── Metric cards ───────────────────────────────────────────────────────── */
.metrics-row {
    display: flex;
    gap: 10px;
    margin-bottom: 0.9rem;
}
.metric-card {
    flex: 1;
    background: #F8FAFB;
    border: 1px solid rgba(0,0,0,0.07);
    border-radius: 8px;
    padding: 13px 14px 11px;
    /* border-top colour set inline per card */
}
.metric-card .num {
    font-size: 1.9rem;
    font-weight: 800;
    line-height: 1.1;
    margin-bottom: 5px;
    /* colour inherited from parent inline style */
}
.metric-card .lbl {
    font-size: 0.66rem;
    font-weight: 600;
    letter-spacing: 0.5px;
    text-transform: uppercase;
    color: #6B7280;
    line-height: 1.3;
}

/* ── Activity panel ─────────────────────────────────────────────────────── */
.act-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 8px;
}
.act-title {
    font-size: 0.88rem;
    font-weight: 700;
    color: #111827;
}
.badge-done {
    background: #D1FAE5; color: #065F46;
    font-size: 0.67rem; font-weight: 700;
    padding: 2px 9px; border-radius: 12px; letter-spacing: 0.3px;
}
.badge-wait {
    background: #F3F4F6; color: #6B7280;
    font-size: 0.67rem; font-weight: 700;
    padding: 2px 9px; border-radius: 12px;
}
.log-box {
    font-family: Consolas, 'Courier New', monospace;
    font-size: 0.78rem;
    background: #FAFAFA;
    border: 1px solid rgba(0,0,0,0.07);
    border-radius: 6px;
    padding: 10px 13px;
    max-height: 230px;
    overflow-y: auto;
    line-height: 1.7;
    margin-top: 6px;
}

/* ── Export panel ───────────────────────────────────────────────────────── */
.export-title {
    font-size: 0.93rem;
    font-weight: 700;
    color: #111827;
    margin: 0 0 2px;
}
.export-caption {
    font-size: 0.75rem;
    color: #6B7280;
    margin: 0 0 10px;
    line-height: 1.4;
}

/* ── Download buttons ───────────────────────────────────────────────────── */
.stDownloadButton > button {
    width: 100% !important;
    font-weight: 600 !important;
    border-radius: 6px !important;
}
[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 10px !important;
}

/* ══════════════════════════════════════════════════════════════════════════
   DARK MODE  (follows OS/browser preference)
══════════════════════════════════════════════════════════════════════════ */
@media (prefers-color-scheme: dark) {
    .sec-lbl { color: #9CA3AF; }

    .info-card { background: #1E2530; border-color: rgba(255,255,255,0.08); }
    .info-card .lbl { color: #9CA3AF; }
    .info-card .val { color: #F3F4F6; }

    .metric-card { background: #1E2530; border-color: rgba(255,255,255,0.08); }
    .metric-card .lbl { color: #9CA3AF; }

    .act-title { color: #F3F4F6; }
    .badge-done { background: #064E3B; color: #6EE7B7; }
    .badge-wait { background: #374151; color: #9CA3AF; }

    .log-box {
        background: #161B22;
        border-color: rgba(255,255,255,0.08);
    }

    .export-title   { color: #F3F4F6; }
    .export-caption { color: #9CA3AF; }
}
</style>
""", unsafe_allow_html=True)


# ── Log colour maps (mirror desktop gui.py) ───────────────────────────────────
_LOG_COLORS = {
    "info":    "#999999",
    "step":    "#1565C0",
    "found":   "#2E7D32",
    "success": "#1B5E20",
    "warn":    "#E65100",
    "error":   "#B71C1C",
}
_LOG_ICONS = {
    "info":    "    ",
    "step":    " >> ",
    "found":   " +  ",
    "success": " ✓  ",
    "warn":    " !  ",
    "error":   " ✗  ",
}
_LOG_BOLD = {"step", "success", "error"}

# ttkbootstrap "flatly" accent colours — match the desktop metric cards exactly
_METRIC_COLORS = {
    "units":      "#2c3e50",   # PRIMARY  — dark navy
    "candidates": "#3498db",   # INFO     — blue
    "total":      "#18bc9c",   # SUCCESS  — teal
    "assess":     "#f39c12",   # WARNING  — amber
    "reassess":   "#e74c3c",   # DANGER   — red
}


# ── Helper functions ──────────────────────────────────────────────────────────

def _tmp_path(suffix: str) -> str:
    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    return path


def _gen_excel(data: dict) -> bytes:
    path = _tmp_path(".xlsx")
    try:
        build_workbook(data, path)
        with open(path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(path):
            os.unlink(path)


def _gen_roster_pdf(data: dict,
                    course_filter: list | None,
                    report_type_filter: str | None) -> bytes:
    from roster_pdf import build_roster_pdf
    path = _tmp_path(".pdf")
    try:
        build_roster_pdf(data, path,
                         course_filter=course_filter,
                         report_type_filter=report_type_filter)
        with open(path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(path):
            os.unlink(path)


def _gen_summary_pdf(data: dict) -> bytes:
    from roster_pdf import build_summary_pdf
    path = _tmp_path(".pdf")
    try:
        build_summary_pdf(data, path)
        with open(path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(path):
            os.unlink(path)


def _course_label(cn: str, cl: str) -> str:
    if not cl:
        return cn
    return f"{cn}  {cl}" if cl.lower().startswith("level") else f"{cn}  Level {cl}"


def _safe(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _render_info_grid(centre_name: str, centre_code: str,
                      courses_text: str, series: str) -> str:
    def card(lbl: str, val: str) -> str:
        safe_val = _safe(val or "—").replace("\n", "<br>")
        return (
            f'<div class="info-card">'
            f'<div class="lbl">{lbl}</div>'
            f'<div class="val">{safe_val}</div>'
            f'</div>'
        )
    return (
        '<div class="info-grid">'
        + card("Centre Name", centre_name)
        + card("Centre Code", centre_code)
        + card("Course(s)",   courses_text)
        + card("Exam Series", series)
        + '</div>'
    )


def _render_metrics(units: int, candidates: int,
                    total: int, assess: int, reassess: int) -> str:
    def card(num: int, lbl: str, key: str) -> str:
        c = _METRIC_COLORS[key]
        return (
            f'<div class="metric-card" '
            f'style="color:{c};border-top:3px solid {c};">'
            f'<div class="num">{num}</div>'
            f'<div class="lbl">{lbl}</div>'
            f'</div>'
        )
    return (
        '<div class="metrics-row">'
        + card(units,      "Units",                       "units")
        + card(candidates, "Unique<br>Candidates",        "candidates")
        + card(total,      "Total<br>Registrations",      "total")
        + card(assess,     "Assessment<br>Registrations", "assess")
        + card(reassess,   "Re-Assessment<br>Registrations", "reassess")
        + '</div>'
    )


def _render_log(entries: list) -> str:
    if not entries:
        return (
            '<div class="log-box">'
            '<span style="color:#999;">No activity yet.</span>'
            '</div>'
        )
    lines: list[str] = []
    for ts, level, msg in entries:
        color = _LOG_COLORS.get(level, "#999")
        icon  = _LOG_ICONS.get(level, "    ")
        bold  = "font-weight:700;" if level in _LOG_BOLD else ""
        lines.append(
            f'<span style="color:#BBBBBB;font-size:0.72rem;">{ts}</span>'
            f'<span style="color:{color};{bold}">{icon}{_safe(msg)}</span><br>'
        )
    return f'<div class="log-box">{"".join(lines)}</div>'


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="app-header">
  <h1>📋&nbsp;&nbsp;TVET CDACC &nbsp;Register Extractor</h1>
  <span class="ver">{_VERSION}</span>
</div>
""", unsafe_allow_html=True)

# ── File uploader ─────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Upload assessment register PDF",
    type=["pdf"],
    label_visibility="collapsed",
    help="Select any TVET CDACC assessment registration register PDF",
)

if not uploaded:
    st.info("Open a TVET CDACC assessment register PDF to begin.", icon="📂")
    st.stop()

# ── Extract (only when file changes) ─────────────────────────────────────────
if (st.session_state.get("pdf_name") != uploaded.name
        or "data" not in st.session_state):

    tmp_pdf     = _tmp_path(".pdf")
    log_entries: list = []

    try:
        with open(tmp_pdf, "wb") as f:
            f.write(uploaded.getvalue())

        with st.status("🔍  Reading PDF…", expanded=True) as status:
            prog = st.progress(0, text="Initialising…")

            def on_log(level: str, msg: str):
                ts = datetime.now().strftime("%H:%M:%S")
                log_entries.append((ts, level, msg))
                _icons = {
                    "step": "🔍", "found": "📋", "success": "✅",
                    "warn": "⚠️", "error": "❌", "info": "·",
                }
                st.write(f"{_icons.get(level, '·')}  {msg}")

            def on_progress(cur: int, total: int):
                prog.progress(cur / total, text=f"Page {cur} of {total}")

            data = extract(tmp_pdf, on_log=on_log, on_progress=on_progress)
            prog.progress(1.0, text="Complete ✓")
            status.update(
                label="✅  Extraction complete!",
                state="complete",
                expanded=False,
            )

        st.session_state.data            = data
        st.session_state.pdf_name        = uploaded.name
        st.session_state.log_entries     = log_entries
        st.session_state.extraction_done = True
        for key in list(st.session_state.keys()):
            if key.startswith("_cache_"):
                del st.session_state[key]

    except Exception as exc:
        st.error(f"❌  Extraction failed: {exc}")
        st.stop()
    finally:
        if os.path.exists(tmp_pdf):
            os.unlink(tmp_pdf)


data        = st.session_state.data
stem        = os.path.splitext(uploaded.name)[0]
log_entries = st.session_state.get("log_entries", [])
extr_done   = st.session_state.get("extraction_done", False)

# ── Derived values ────────────────────────────────────────────────────────────
roster   = build_roster(data)
total    = sum(u["candidate_count"] for u in data["units"])
assess   = sum(
    u["candidate_count"] for u in data["units"]
    if (u.get("report_type") or "").lower().startswith("assessment")
)
reassess = total - assess

seen_courses = list(dict.fromkeys(
    (u.get("course_name", ""), u.get("course_level", ""))
    for u in data["units"] if u.get("course_name", "")
))
courses_text = "\n".join(
    _course_label(cn, cl) for cn, cl in seen_courses
) or "—"

# ── Extracted Information ─────────────────────────────────────────────────────
st.markdown('<p class="sec-lbl">Extracted Information</p>', unsafe_allow_html=True)
st.markdown(
    _render_info_grid(
        data.get("centre_name") or "—",
        data.get("centre_code") or "—",
        courses_text,
        data.get("series") or "—",
    ),
    unsafe_allow_html=True,
)

# ── Metrics ───────────────────────────────────────────────────────────────────
st.markdown(
    _render_metrics(data["unit_count"], len(roster), total, assess, reassess),
    unsafe_allow_html=True,
)

# ── Activity ──────────────────────────────────────────────────────────────────
with st.container(border=True):
    badge = (
        '<span class="badge-done">Complete ✓</span>' if extr_done
        else '<span class="badge-wait">Waiting…</span>'
    )
    st.markdown(
        f'<div class="act-header">'
        f'<span class="act-title">⚡ Activity</span>{badge}</div>',
        unsafe_allow_html=True,
    )
    st.progress(1.0 if extr_done else 0.0)
    st.markdown(_render_log(log_entries), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Export ────────────────────────────────────────────────────────────────────
st.markdown('<p class="sec-lbl">Export</p>', unsafe_allow_html=True)

with st.container(border=True):
    ex_col, roster_col, summary_col = st.columns([1, 1.7, 1], gap="large")

    # ── Excel Workbook ────────────────────────────────────────────────────────
    with ex_col:
        st.markdown('<p class="export-title">📊 Excel Workbook</p>',
                    unsafe_allow_html=True)
        st.markdown(
            '<p class="export-caption">All candidates, units and roster '
            '— one workbook, three sheets.</p>',
            unsafe_allow_html=True,
        )
        _xk = "_cache_excel"
        if _xk not in st.session_state:
            with st.spinner("Building workbook…"):
                st.session_state[_xk] = _gen_excel(data)
        st.download_button(
            label="⬇  Download Excel",
            data=st.session_state[_xk],
            file_name=f"{stem}_extracted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

    # ── Roster PDF ────────────────────────────────────────────────────────────
    with roster_col:
        st.markdown('<p class="export-title">📄 Roster PDF</p>',
                    unsafe_allow_html=True)
        st.markdown(
            '<p class="export-caption">Landscape A4 attendance roster '
            'with candidate list and signature column.</p>',
            unsafe_allow_html=True,
        )

        has_assess   = any(
            (u.get("report_type") or "").lower().startswith("assessment")
            for u in data["units"]
        )
        has_reassess = any(
            (u.get("report_type") or "").lower().startswith("re")
            for u in data["units"]
        )

        rt_options: dict[str, str | None] = {"All report types": None}
        if has_assess:
            rt_options["Assessment only"]    = "Assessment Registrations"
        if has_reassess:
            rt_options["Re-Assessment only"] = "Re-Assessment Registrations"

        rt_label = st.radio(
            "Report type", list(rt_options.keys()),
            horizontal=True, key="roster_rt",
        )
        report_type_filter = rt_options[rt_label]

        all_courses = list(dict.fromkeys(
            (u.get("course_name", ""), u.get("course_level", ""))
            for u in data["units"]
        ))
        course_map  = {_course_label(cn, cl): (cn, cl) for cn, cl in all_courses}
        sel_labels  = st.multiselect(
            "Courses to include",
            options=list(course_map.keys()),
            default=list(course_map.keys()),
            key="roster_courses",
        )

        if not sel_labels:
            st.warning("Select at least one course.")
        else:
            course_filter = (
                None if len(sel_labels) == len(course_map)
                else [course_map[lbl] for lbl in sel_labels]
            )
            _rk = f"_cache_roster_{rt_label}_{'|'.join(sorted(sel_labels))}"
            if _rk not in st.session_state:
                with st.spinner("Building roster PDF…"):
                    st.session_state[_rk] = _gen_roster_pdf(
                        data, course_filter, report_type_filter
                    )
            _slug = (
                "all"          if not report_type_filter
                else "assessment" if "assessment" in report_type_filter.lower()
                else "reassessment"
            )
            st.download_button(
                label="⬇  Download Roster PDF",
                data=st.session_state[_rk],
                file_name=f"{stem}_{_slug}_roster.pdf",
                mime="application/pdf",
                use_container_width=True,
                type="primary",
            )

    # ── Summary PDF ───────────────────────────────────────────────────────────
    with summary_col:
        st.markdown('<p class="export-title">📋 Summary PDF</p>',
                    unsafe_allow_html=True)
        st.markdown(
            '<p class="export-caption">Portrait A4 — statistics '
            'and full unit breakdown.</p>',
            unsafe_allow_html=True,
        )
        _sk = "_cache_summary"
        if _sk not in st.session_state:
            with st.spinner("Building summary PDF…"):
                st.session_state[_sk] = _gen_summary_pdf(data)
        st.download_button(
            label="⬇  Download Summary PDF",
            data=st.session_state[_sk],
            file_name=f"{stem}_summary.pdf",
            mime="application/pdf",
            use_container_width=True,
            type="primary",
        )
