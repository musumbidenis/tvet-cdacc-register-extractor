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

# ── Custom styling ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Tighten top padding */
    .block-container { padding-top: 1.8rem; padding-bottom: 1rem; }

    /* Metric card tweaks */
    [data-testid="stMetric"] {
        background: #F8F9FA;
        border: 1px solid #E9ECEF;
        border-radius: 8px;
        padding: 14px 18px 10px;
    }
    [data-testid="stMetricLabel"] { font-size: 0.78rem; color: #6C757D; }
    [data-testid="stMetricValue"] { font-size: 1.9rem; font-weight: 700; }

    /* Download buttons */
    .stDownloadButton > button {
        width: 100%;
        font-weight: 600;
        border-radius: 6px;
    }

    /* Section headings */
    h4 { margin-bottom: 4px !important; }
</style>
""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _tmp_path(suffix: str) -> str:
    """Create an empty named temp file and return its path (already closed)."""
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


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("## 📋 TVET CDACC Register Extractor")
st.divider()

# ── File uploader ─────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Upload assessment register PDF",
    type=["pdf"],
    label_visibility="collapsed",
    help="Select any TVET CDACC assessment registration register PDF",
)

if not uploaded:
    st.info("👆  Upload a TVET CDACC assessment register PDF to begin.", icon="📂")
    st.stop()

# ── Extract (only when file changes) ─────────────────────────────────────────
if (st.session_state.get("pdf_name") != uploaded.name
        or "data" not in st.session_state):

    # Write uploaded bytes to a temp file (pdfplumber needs a path)
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
                prog.progress(cur / total, text=f"Page {cur} / {total}")

            data = extract(tmp_pdf, on_log=on_log, on_progress=on_progress)
            prog.progress(1.0, text="Complete ✓")
            status.update(
                label="✅  Extraction complete!",
                state="complete",
                expanded=False,
            )

        # Store in session; clear any previously cached export bytes
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
st.subheader("Extracted Information")

c1, c2, c3, c4 = st.columns(4)
c1.markdown(f"**Centre Name**  \n{data.get('centre_name') or '—'}")
c2.markdown(f"**Centre Code**  \n{data.get('centre_code') or '—'}")

seen_courses = list(dict.fromkeys(
    (u.get("course_name", ""), u.get("course_level", ""))
    for u in data["units"]
    if u.get("course_name", "")
))
courses_text = "  \n".join(
    _course_label(cn, cl) for cn, cl in seen_courses
) or "—"

c3.markdown(f"**Course(s)**  \n{courses_text}")
c4.markdown(f"**Exam Series**  \n{data.get('series') or '—'}")

st.divider()

# ── Metrics ───────────────────────────────────────────────────────────────────
roster   = build_roster(data)
total    = sum(u["candidate_count"] for u in data["units"])
assess   = sum(
    u["candidate_count"] for u in data["units"]
    if (u.get("report_type") or "").lower().startswith("assessment")
)
reassess = total - assess

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Units",                 data["unit_count"])
m2.metric("Unique Candidates",     len(roster))
m3.metric("Total Registrations",   total)
m4.metric("Assessment",            assess)
m5.metric("Re-Assessment",         reassess)

st.divider()

# ── Export ────────────────────────────────────────────────────────────────────
st.subheader("Export")

ex_col, roster_col, summary_col = st.columns([1, 1.8, 1], gap="large")

# ── Excel Workbook ────────────────────────────────────────────────────────────
with ex_col:
    st.markdown("#### 📊 Excel Workbook")
    st.caption("All candidates, units, and roster — one workbook, three sheets.")

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
    st.markdown("#### 📄 Roster PDF")
    st.caption("Landscape A4 attendance roster with signature column.")

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

        # Cache key encodes the filter state
        _rk = f"_cache_roster_{rt_label}_{'|'.join(sorted(sel_labels))}"
        if _rk not in st.session_state:
            with st.spinner("Building roster PDF…"):
                st.session_state[_rk] = _gen_roster_pdf(
                    data, course_filter, report_type_filter
                )

        _slug = ("all"          if not report_type_filter
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
    st.markdown("#### 📋 Summary PDF")
    st.caption("Portrait A4 — statistics and full unit breakdown.")

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
