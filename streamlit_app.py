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

import streamlit as st

from extract_registers import extract, build_roster
from register_excel import build_workbook

# ── Page configuration ────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TVET CDACC Register Extractor",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Hide Streamlit chrome (menu, footer, deploy button) ───────────────────────
# Also wire up system dark/light theme for custom elements
st.markdown("""
<style>
/* ── Hide Streamlit chrome ─────────────────────────────── */
#MainMenu            { display: none !important; }
footer               { display: none !important; }
[data-testid="stDeployButton"]       { display: none !important; }
[data-testid="stToolbar"]            { display: none !important; }
[data-testid="stDecoration"]         { display: none !important; }

/* ── Layout ────────────────────────────────────────────── */
.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
    max-width: 1100px;
}

/* ── App header banner ──────────────────────────────────── */
.app-header {
    display: flex;
    align-items: center;
    gap: 14px;
    padding: 18px 24px;
    border-radius: 12px;
    margin-bottom: 1.5rem;
    background: linear-gradient(135deg, #1B7A48 0%, #29A861 100%);
    color: #ffffff;
    box-shadow: 0 4px 14px rgba(41,168,97,0.25);
}
.app-header .icon { font-size: 2.1rem; line-height: 1; }
.app-header h1 {
    margin: 0; padding: 0;
    font-size: 1.45rem; font-weight: 700; letter-spacing: -0.3px;
    color: #fff;
}
.app-header p {
    margin: 2px 0 0; padding: 0;
    font-size: 0.82rem; opacity: 0.85; color: #fff;
}

/* ── Info cards ──────────────────────────────────────────── */
.info-card {
    border-radius: 10px;
    padding: 14px 18px;
    min-height: 76px;
    border: 1px solid rgba(0,0,0,0.08);
    background: #F8FAFB;
}
.info-card .label {
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.6px;
    text-transform: uppercase;
    color: #6B7280;
    margin-bottom: 5px;
}
.info-card .value {
    font-size: 0.97rem;
    font-weight: 600;
    color: #111827;
    line-height: 1.35;
}

/* ── Metric cards ──────────────────────────────────────── */
[data-testid="stMetric"] {
    background: #F8FAFB;
    border: 1px solid rgba(0,0,0,0.08);
    border-radius: 10px;
    padding: 16px 18px 12px;
}
[data-testid="stMetricLabel"] {
    font-size: 0.72rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.5px;
    text-transform: uppercase;
    color: #6B7280 !important;
}
[data-testid="stMetricValue"] {
    font-size: 2rem !important;
    font-weight: 800 !important;
    color: #111827 !important;
}

/* ── Export cards ────────────────────────────────────────── */
.export-card {
    border-radius: 12px;
    padding: 20px 22px 18px;
    border: 1px solid rgba(0,0,0,0.09);
    background: #F8FAFB;
    height: 100%;
}
.export-card h4 {
    margin: 0 0 4px;
    font-size: 1rem;
    font-weight: 700;
    color: #111827;
}
.export-card p {
    margin: 0 0 16px;
    font-size: 0.79rem;
    color: #6B7280;
    line-height: 1.45;
}

/* ── Download buttons ────────────────────────────────────── */
.stDownloadButton > button {
    width: 100% !important;
    font-weight: 600 !important;
    border-radius: 7px !important;
    padding: 0.45rem 1rem !important;
}

/* ── Section divider label ───────────────────────────────── */
.section-label {
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 1.1px;
    text-transform: uppercase;
    color: #6B7280;
    margin: 0 0 10px;
}

/* ── Upload zone ─────────────────────────────────────────── */
[data-testid="stFileUploader"] {
    border-radius: 10px;
}

/* ── Status / progress tweaks ────────────────────────────── */
[data-testid="stStatusWidget"] {
    border-radius: 10px;
}

/* ════════════════════════════════════════════════════════════
   DARK MODE overrides  (system preference)
════════════════════════════════════════════════════════════ */
@media (prefers-color-scheme: dark) {
    .info-card {
        background: #1E2530;
        border-color: rgba(255,255,255,0.08);
    }
    .info-card .label  { color: #9CA3AF; }
    .info-card .value  { color: #F3F4F6; }

    .export-card {
        background: #1E2530;
        border-color: rgba(255,255,255,0.08);
    }
    .export-card h4    { color: #F3F4F6; }
    .export-card p     { color: #9CA3AF; }

    [data-testid="stMetric"] {
        background: #1E2530;
        border-color: rgba(255,255,255,0.08);
    }
    [data-testid="stMetricLabel"] { color: #9CA3AF !important; }
    [data-testid="stMetricValue"] { color: #F3F4F6 !important; }

    .section-label { color: #9CA3AF; }
}
</style>
""", unsafe_allow_html=True)


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


def _info_card(label: str, value: str) -> str:
    safe = value.replace("\n", "<br>") if value else "—"
    return (
        f'<div class="info-card">'
        f'<div class="label">{label}</div>'
        f'<div class="value">{safe}</div>'
        f'</div>'
    )


# ── App header ────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <div class="icon">📋</div>
  <div>
    <h1>TVET CDACC Register Extractor</h1>
    <p>Extract candidates, units &amp; rosters from assessment registration PDFs</p>
  </div>
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
    st.info("Upload a TVET CDACC assessment register PDF to get started.", icon="📂")
    st.stop()

# ── Extract (only when file changes) ─────────────────────────────────────────
if (st.session_state.get("pdf_name") != uploaded.name
        or "data" not in st.session_state):

    tmp_pdf = _tmp_path(".pdf")
    try:
        with open(tmp_pdf, "wb") as f:
            f.write(uploaded.getvalue())

        with st.status("🔍  Reading PDF…", expanded=True) as status:
            prog = st.progress(0, text="Initialising…")

            def on_log(level: str, msg: str):
                _icons = {
                    "step":    "🔍",
                    "found":   "📋",
                    "success": "✅",
                    "warn":    "⚠️",
                    "error":   "❌",
                    "info":    "·",
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

        st.session_state.data     = data
        st.session_state.pdf_name = uploaded.name
        for key in list(st.session_state.keys()):
            if key.startswith("_cache_"):
                del st.session_state[key]

    except Exception as exc:
        st.error(f"❌  Extraction failed: {exc}")
        st.stop()
    finally:
        if os.path.exists(tmp_pdf):
            os.unlink(tmp_pdf)


data = st.session_state.data
stem = os.path.splitext(uploaded.name)[0]

# ── Info panel ────────────────────────────────────────────────────────────────
st.markdown('<p class="section-label">Register Details</p>', unsafe_allow_html=True)

seen_courses = list(dict.fromkeys(
    (u.get("course_name", ""), u.get("course_level", ""))
    for u in data["units"]
    if u.get("course_name", "")
))
courses_text = "\n".join(
    _course_label(cn, cl) for cn, cl in seen_courses
) or "—"

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(_info_card("Centre Name", data.get("centre_name") or "—"),
                unsafe_allow_html=True)
with c2:
    st.markdown(_info_card("Centre Code", data.get("centre_code") or "—"),
                unsafe_allow_html=True)
with c3:
    st.markdown(_info_card("Course(s)", courses_text), unsafe_allow_html=True)
with c4:
    st.markdown(_info_card("Exam Series", data.get("series") or "—"),
                unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Metrics ───────────────────────────────────────────────────────────────────
st.markdown('<p class="section-label">Summary</p>', unsafe_allow_html=True)

roster   = build_roster(data)
total    = sum(u["candidate_count"] for u in data["units"])
assess   = sum(
    u["candidate_count"] for u in data["units"]
    if (u.get("report_type") or "").lower().startswith("assessment")
)
reassess = total - assess

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Units",               data["unit_count"])
m2.metric("Unique Candidates",   len(roster))
m3.metric("Total Registrations", total)
m4.metric("Assessment",          assess)
m5.metric("Re-Assessment",       reassess)

st.markdown("<br>", unsafe_allow_html=True)

# ── Export ────────────────────────────────────────────────────────────────────
st.markdown('<p class="section-label">Export</p>', unsafe_allow_html=True)

ex_col, roster_col, summary_col = st.columns([1, 1.7, 1], gap="large")

# ── Excel Workbook ────────────────────────────────────────────────────────────
with ex_col:
    st.markdown("""
    <div class="export-card">
      <h4>📊 Excel Workbook</h4>
      <p>All candidates, units, and roster — one workbook, three sheets.</p>
    </div>
    """, unsafe_allow_html=True)

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

# ── Roster PDF ────────────────────────────────────────────────────────────────
with roster_col:
    st.markdown("""
    <div class="export-card">
      <h4>📄 Roster PDF</h4>
      <p>Landscape A4 attendance roster with candidate list and signature column.</p>
    </div>
    """, unsafe_allow_html=True)

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
        "Report type",
        list(rt_options.keys()),
        horizontal=True,
        key="roster_rt",
    )
    report_type_filter = rt_options[rt_label]

    all_courses  = list(dict.fromkeys(
        (u.get("course_name", ""), u.get("course_level", ""))
        for u in data["units"]
    ))
    course_map   = {_course_label(cn, cl): (cn, cl) for cn, cl in all_courses}
    sel_labels   = st.multiselect(
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

        _slug = ("all"           if not report_type_filter
                 else "assessment" if "assessment" in report_type_filter.lower()
                 else "reassessment")
        st.download_button(
            label="⬇  Download Roster PDF",
            data=st.session_state[_rk],
            file_name=f"{stem}_{_slug}_roster.pdf",
            mime="application/pdf",
            use_container_width=True,
            type="primary",
        )

# ── Summary PDF ───────────────────────────────────────────────────────────────
with summary_col:
    st.markdown("""
    <div class="export-card">
      <h4>📋 Summary PDF</h4>
      <p>Portrait A4 — statistics and full unit breakdown.</p>
    </div>
    """, unsafe_allow_html=True)

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
