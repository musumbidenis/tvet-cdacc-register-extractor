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
from roster_pdf import build_roster_pdf, build_summary_pdf

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
    margin-left: -5rem;
    margin-right: -5rem;
    margin-bottom: 1.4rem;
}
.app-header h1 {
    margin: 0; padding: 0;
    font-size: 1.3rem; font-weight: 700;
    color: #fff; letter-spacing: -0.2px;
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
    font-size: 0.67rem; font-weight: 700;
    letter-spacing: 1px; text-transform: uppercase;
    color: #6B7280; margin: 0 0 7px; padding: 0;
}

/* ── Metric / info cards  (shared style) ────────────────────────────────── */
.metrics-row { display: flex; gap: 10px; margin-bottom: 0.9rem; }
.metric-card {
    flex: 1;
    background: #F8FAFB;
    border: 1px solid rgba(0,0,0,0.07);
    border-radius: 8px;
    padding: 13px 14px 11px;
    /* border-top colour and text colour set inline per card */
}
.metric-card .num {
    font-size: 1.9rem; font-weight: 800;
    line-height: 1.1; margin-bottom: 5px;
}
.metric-card .lbl {
    font-size: 0.66rem; font-weight: 600;
    letter-spacing: 0.5px; text-transform: uppercase;
    color: #6B7280; line-height: 1.3;
}

/* ── Activity panel ─────────────────────────────────────────────────────── */
.act-header {
    display: flex; align-items: center;
    justify-content: space-between; margin-bottom: 8px;
}
.act-title { font-size: 0.88rem; font-weight: 700; color: #111827; }
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
    font-size: 0.93rem; font-weight: 700;
    color: #111827; margin: 0 0 2px;
}
.export-caption {
    font-size: 0.75rem; color: #6B7280;
    margin: 0 0 10px; line-height: 1.4;
}

/* ── Filename input ─────────────────────────────────────────────────────── */
.fn-label {
    font-size: 0.67rem; font-weight: 600;
    letter-spacing: 0.5px; text-transform: uppercase;
    color: #6B7280; margin: 8px 0 3px;
}

/* ── Download buttons ───────────────────────────────────────────────────── */
.stDownloadButton > button {
    width: 100% !important;
    font-weight: 600 !important;
    border-radius: 6px !important;
}
[data-testid="stVerticalBlockBorderWrapper"] { border-radius: 10px !important; }

/* ─── DARK MODE ─────────────────────────────────────────────────────────── */
@media (prefers-color-scheme: dark) {
    .sec-lbl    { color: #9CA3AF; }

    .metric-card { background: #1E2530; border-color: rgba(255,255,255,0.08); }
    .metric-card .lbl { color: #9CA3AF; }

    .act-title  { color: #F3F4F6; }
    .badge-done { background: #064E3B; color: #6EE7B7; }
    .badge-wait { background: #374151; color: #9CA3AF; }

    .log-box { background: #161B22; border-color: rgba(255,255,255,0.08); }

    .export-title   { color: #F3F4F6; }
    .export-caption { color: #9CA3AF; }
    .fn-label       { color: #9CA3AF; }
}
</style>
""", unsafe_allow_html=True)


# ── ttkbootstrap "flatly" accent colours — same as desktop metric cards ───────
_C = {
    "primary":  "#2c3e50",
    "info":     "#3498db",
    "success":  "#18bc9c",
    "warning":  "#f39c12",
    "danger":   "#e74c3c",
}

# ── Log colour/icon maps (mirror desktop gui.py exactly) ─────────────────────
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


# ── Helpers ───────────────────────────────────────────────────────────────────

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
        if os.path.exists(path): os.unlink(path)


def _gen_roster_pdf(data: dict,
                    course_filter: list | None,
                    report_type_filter: str | None,
                    bw: bool = False) -> bytes:
    path = _tmp_path(".pdf")
    try:
        build_roster_pdf(data, path,
                         course_filter=course_filter,
                         report_type_filter=report_type_filter,
                         bw=bw)
        with open(path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(path): os.unlink(path)


def _gen_summary_pdf(data: dict, bw: bool = False) -> bytes:
    path = _tmp_path(".pdf")
    try:
        build_summary_pdf(data, path, bw=bw)
        with open(path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(path): os.unlink(path)


def _course_label(cn: str, cl: str) -> str:
    if not cl:
        return cn
    return f"{cn}  {cl}" if cl.lower().startswith("level") else f"{cn}  Level {cl}"


def _safe(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _merge_data(data_list: list) -> dict:
    """Merge extracted data from multiple PDFs into one."""
    if not data_list:
        return {}
    if len(data_list) == 1:
        return data_list[0]

    def _uniq(key: str) -> str:
        vals = list(dict.fromkeys(
            d.get(key, "").strip() for d in data_list
            if d.get(key, "").strip()
        ))
        if not vals:  return ""
        if len(vals) == 1: return vals[0]
        return " / ".join(vals[:3])

    all_units: list = []
    for d in data_list:
        all_units.extend(d.get("units", []))

    return {
        "centre_name":  _uniq("centre_name"),
        "centre_code":  _uniq("centre_code"),
        "series":       _uniq("series"),
        "course_name":  _uniq("course_name"),
        "course_level": _uniq("course_level"),
        "units":        all_units,
        "unit_count":   len(all_units),
    }


def _metric_card(val: str, lbl: str, color: str, big: bool = True) -> str:
    """Single card.  big=True → large number font; big=False → normal text.
    lbl may contain raw HTML (e.g. <br>) — it is NOT escaped."""
    num_style = "" if big else (
        "font-size:0.95rem;line-height:1.25;"
        "word-break:break-word;white-space:pre-line;"
    )
    return (
        f'<div class="metric-card" '
        f'style="color:{color};border-top:3px solid {color};">'
        f'<div class="num" style="{num_style}">{_safe(val)}</div>'
        f'<div class="lbl">{lbl}</div>'
        f'</div>'
    )


def _render_log(entries: list) -> str:
    if not entries:
        return (
            '<div class="log-box">'
            '<span style="color:#999;">No activity yet.</span></div>'
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

# ── File uploader  (multiple PDFs) ────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Upload assessment register PDF(s)",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    help="Select one or more TVET CDACC assessment register PDFs — "
         "you can keep adding more files while previous ones remain processed.",
)

if not uploaded_files:
    st.info(
        "Upload one or more TVET CDACC assessment register PDFs to begin.  "
        "You can keep adding files — each new one will be processed and merged.",
        icon="📂",
    )
    st.stop()

# ── Incremental extraction ────────────────────────────────────────────────────
current_map: dict = {f.name: f for f in uploaded_files}
processed:   dict = st.session_state.get("processed_files", {})   # {name: data}
all_logs:    list = st.session_state.get("log_entries", [])

removed = [name for name in list(processed.keys()) if name not in current_map]
added   = [f for name, f in current_map.items() if name not in processed]

for name in removed:
    del processed[name]

for uploaded in added:
    tmp = _tmp_path(".pdf")
    file_logs: list = []
    try:
        with open(tmp, "wb") as fh:
            fh.write(uploaded.getvalue())

        with st.status(f"🔍  Reading {uploaded.name}…", expanded=True) as status:
            prog = st.progress(0, text="Initialising…")

            def on_log(level: str, msg: str,
                       _logs: list = file_logs) -> None:
                ts = datetime.now().strftime("%H:%M:%S")
                _logs.append((ts, level, msg))
                _icons = {
                    "step": "🔍", "found": "📋", "success": "✅",
                    "warn": "⚠️", "error": "❌", "info": "·",
                }
                st.write(f"{_icons.get(level, '·')}  {msg}")

            def on_progress(cur: int, total: int,
                            _p=prog) -> None:
                _p.progress(cur / total, text=f"Page {cur} of {total}")

            data = extract(tmp, on_log=on_log, on_progress=on_progress)
            prog.progress(1.0, text="Complete ✓")
            status.update(
                label=f"✅  {uploaded.name} — extraction complete!",
                state="complete",
                expanded=False,
            )

        processed[uploaded.name] = data
        all_logs.extend(file_logs)

    except Exception as exc:
        st.error(f"❌  {uploaded.name}: {exc}")
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)

# Persist state; clear export caches whenever the file set changes
if removed or added:
    st.session_state.processed_files  = processed
    st.session_state.log_entries      = all_logs
    st.session_state.extraction_done  = bool(processed)
    st.session_state.data             = _merge_data(list(processed.values()))
    for key in list(st.session_state.keys()):
        if key.startswith("_cache_"):
            del st.session_state[key]

if not processed:
    st.warning("No files were processed successfully.")
    st.stop()

# ── Working values ────────────────────────────────────────────────────────────
data        = st.session_state.data
log_entries = st.session_state.get("log_entries", [])
extr_done   = st.session_state.get("extraction_done", False)

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

# Default stem for filenames
if len(uploaded_files) == 1:
    stem_default = os.path.splitext(uploaded_files[0].name)[0]
else:
    stem_default = "CDACC_registers"

# ── Info row  (same metric-card style, courses shows count) ──────────────────
st.markdown('<p class="sec-lbl">Extracted Information</p>', unsafe_allow_html=True)
st.markdown(
    '<div class="metrics-row">'
    + _metric_card(data.get("centre_name") or "—", "Centre Name", _C["primary"], big=False)
    + _metric_card(data.get("centre_code") or "—", "Centre Code", _C["info"],    big=False)
    + _metric_card(str(len(seen_courses)),          "Courses",     _C["success"], big=True)
    + _metric_card(data.get("series") or "—",       "Exam Series", _C["warning"], big=False)
    + '</div>',
    unsafe_allow_html=True,
)

# ── Metric row ────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="metrics-row">'
    + _metric_card(str(data["unit_count"]), "Units",                       _C["primary"])
    + _metric_card(str(len(roster)),        "Unique<br>Candidates",        _C["info"])
    + _metric_card(str(total),              "Total<br>Registrations",      _C["success"])
    + _metric_card(str(assess),             "Assessment<br>Registrations", _C["warning"])
    + _metric_card(str(reassess),           "Re-Assessment<br>Registrations", _C["danger"])
    + '</div>',
    unsafe_allow_html=True,
)

# ── Activity ──────────────────────────────────────────────────────────────────
with st.container(border=True):
    n_files = len(processed)
    badge = (
        f'<span class="badge-done">Complete ✓  '
        f'({n_files} file{"s" if n_files != 1 else ""})</span>'
        if extr_done else
        '<span class="badge-wait">Waiting…</span>'
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

        st.markdown('<p class="fn-label">Save as</p>', unsafe_allow_html=True)
        excel_fn = st.text_input(
            "Excel filename", value=f"{stem_default}_extracted",
            key="fn_excel", label_visibility="collapsed",
        )
        st.download_button(
            label="⬇  Download Excel",
            data=st.session_state[_xk],
            file_name=f"{excel_fn.strip() or stem_default}.xlsx",
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
        if has_assess:   rt_options["Assessment only"]    = "Assessment Registrations"
        if has_reassess: rt_options["Re-Assessment only"] = "Re-Assessment Registrations"

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

        roster_bw = st.radio(
            "Colour mode",
            ["🎨  Colour", "🖨️  B&W / Print"],
            horizontal=True, key="roster_bw",
        ) == "🖨️  B&W / Print"

        if not sel_labels:
            st.warning("Select at least one course.")
        else:
            course_filter = (
                None if len(sel_labels) == len(course_map)
                else [course_map[lbl] for lbl in sel_labels]
            )
            _bw_tag = "bw" if roster_bw else "col"
            _rk = f"_cache_roster_{rt_label}_{'|'.join(sorted(sel_labels))}_{_bw_tag}"
            if _rk not in st.session_state:
                with st.spinner("Building roster PDF…"):
                    st.session_state[_rk] = _gen_roster_pdf(
                        data, course_filter, report_type_filter, bw=roster_bw
                    )
            _slug = (
                "all"            if not report_type_filter
                else "assessment"    if "assessment" in report_type_filter.lower()
                else "reassessment"
            )
            _bw_suffix = "_bw" if roster_bw else ""
            st.markdown('<p class="fn-label">Save as</p>', unsafe_allow_html=True)
            roster_fn = st.text_input(
                "Roster filename",
                value=f"{stem_default}_{_slug}_roster{_bw_suffix}",
                key="fn_roster", label_visibility="collapsed",
            )
            st.download_button(
                label="⬇  Download Roster PDF",
                data=st.session_state[_rk],
                file_name=f"{roster_fn.strip() or stem_default}.pdf",
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

        summary_bw = st.radio(
            "Colour mode",
            ["🎨  Colour", "🖨️  B&W / Print"],
            horizontal=True, key="summary_bw",
        ) == "🖨️  B&W / Print"

        _bw_tag = "bw" if summary_bw else "col"
        _sk = f"_cache_summary_{_bw_tag}"
        if _sk not in st.session_state:
            with st.spinner("Building summary PDF…"):
                st.session_state[_sk] = _gen_summary_pdf(data, bw=summary_bw)

        _bw_suffix = "_bw" if summary_bw else ""
        st.markdown('<p class="fn-label">Save as</p>', unsafe_allow_html=True)
        summary_fn = st.text_input(
            "Summary filename", value=f"{stem_default}_summary{_bw_suffix}",
            key="fn_summary", label_visibility="collapsed",
        )
        st.download_button(
            label="⬇  Download Summary PDF",
            data=st.session_state[_sk],
            file_name=f"{summary_fn.strip() or stem_default}.pdf",
            mime="application/pdf",
            use_container_width=True,
            type="primary",
        )
