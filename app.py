# =============================================================================
# app.py  —  Stage 3: Three feedback types, modular pipeline
# Run:  streamlit run app.py
#
# FLOW:
#   1. User selects feedback type (Course-Exit / Faculty End-Term / Faculty Mid-Term)
#   2. User uploads Excel file
#   3. App parses file using the correct parser
#   4. Auto-detected fields shown — user can edit any
#   5. Generate button → builds Analysis + ATR via correct builder
#   6. Download buttons (always) + Save to Drive button (if secrets configured)
# =============================================================================

import os
import sys
import tempfile

import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.parsers.course_exit_parser     import parse as parse_course_exit
from core.parsers.faculty_endterm_parser import parse as parse_faculty_endterm
from core.parsers.faculty_midterm_parser import parse as parse_faculty_midterm

from core.builders.course_exit_builder     import build_analysis as ce_analysis, build_atr as ce_atr
from core.builders.faculty_endterm_builder import build_analysis as fe_analysis, build_atr as fe_atr
from core.builders.faculty_midterm_builder import build_analysis as fm_analysis, build_atr as fm_atr

from core.shared.charts    import generate_all_charts
from core.shared.templates import (
    generate_all_course_exit_action_taken,
    generate_all_faculty_action_taken,
)
from core.shared.drive_saver import is_drive_configured, save_both_to_drive


# -----------------------------------------------------------------------
# PAGE CONFIG
# -----------------------------------------------------------------------
st.set_page_config(page_title="NBA Document Generator", page_icon="📄", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Serif+Display&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
#MainMenu, footer, header   { visibility: hidden; }
.stApp                      { background: #f7f6f2; }
.main .block-container { max-width: 720px; padding-top: 2.5rem; padding-bottom: 3rem; }
.app-title    { font-family:'DM Serif Display',serif; font-size:2rem; color:#1a1a2e; letter-spacing:-0.5px; }
.app-subtitle { font-size:0.9rem; color:#6b6b80; margin-bottom:2rem; }
.section-label { font-size:0.7rem; font-weight:600; letter-spacing:0.12em; text-transform:uppercase;
                 color:#9090a0; margin-bottom:0.5rem; margin-top:1.5rem; }
.field-row { display:flex; align-items:baseline; gap:0.75rem; padding:0.45rem 0; border-bottom:1px solid #ebebeb; }
.field-key { font-size:0.78rem; color:#888; min-width:130px; }
.field-val { font-size:0.9rem; color:#1a1a2e; font-weight:500; }
.badge-ok     { font-size:0.65rem; background:#e6f4ea; color:#2d7a3e; padding:2px 8px; border-radius:20px; font-weight:600; }
.badge-missing{ font-size:0.65rem; background:#fff3cd; color:#856404; padding:2px 8px; border-radius:20px; font-weight:600; }
.type-badge { display:inline-block; font-size:0.72rem; font-weight:600; padding:3px 10px; border-radius:20px; margin-bottom:0.5rem; }
.type-ce { background:#e8f0fe; color:#1a56d6; }
.type-fe { background:#fce8ff; color:#8b1ab5; }
.type-fm { background:#fff0e0; color:#b35a00; }
.co-table { width:100%; border-collapse:collapse; font-size:0.82rem; margin-top:0.5rem; }
.co-table th { text-align:left; padding:6px 10px; background:#1a1a2e; color:#fff; font-weight:500; }
.co-table td { padding:6px 10px; border-bottom:1px solid #ebebeb; color:#333; }
.co-table tr:last-child td { border-bottom:none; }
.stButton > button { background:#1a1a2e; color:#fff; border:none; border-radius:6px;
    font-family:'DM Sans',sans-serif; font-size:0.9rem; font-weight:500;
    padding:0.6rem 2rem; width:100%; cursor:pointer; transition:background 0.15s; }
.stButton > button:hover { background:#2e2e50; }
.stDownloadButton > button { background:#fff; color:#1a1a2e; border:1.5px solid #1a1a2e;
    border-radius:6px; font-family:'DM Sans',sans-serif; font-size:0.88rem; font-weight:500;
    width:100%; transition:all 0.15s; }
.stDownloadButton > button:hover { background:#1a1a2e; color:#fff; }
hr { border:none; border-top:1px solid #e0e0e0; margin:1.5rem 0; }
</style>
""", unsafe_allow_html=True)


# -----------------------------------------------------------------------
# CONSTANTS
# -----------------------------------------------------------------------
FEEDBACK_TYPES = {
    "Course-Exit Feedback":      "course_exit",
    "Faculty End-Term Feedback": "faculty_endterm",
    "Faculty Mid-Term Feedback": "faculty_midterm",
}
DEPT_OPTIONS = [
    "Department of Computer Science and Engineering (CSE)",
    "Department of Computer Science and Business Systems (CSBS)",
]
TYPE_BADGE = {
    "course_exit":     ("Course-Exit",        "type-ce"),
    "faculty_endterm": ("Faculty End-Term",    "type-fe"),
    "faculty_midterm": ("Faculty Mid-Term",    "type-fm"),
}
PARSERS = {
    "course_exit":     parse_course_exit,
    "faculty_endterm": parse_faculty_endterm,
    "faculty_midterm": parse_faculty_midterm,
}


# -----------------------------------------------------------------------
# HELPERS
# -----------------------------------------------------------------------
def field_row(label, value, detected):
    badge   = '<span class="badge-ok">Auto-detected</span>' if detected else '<span class="badge-missing">Not found</span>'
    display = value if value else "—"
    return f'<div class="field-row"><span class="field-key">{label}</span><span class="field-val">{display}</span>{badge}</div>'


@st.cache_data(show_spinner="Reading file…")
def cached_parse(file_bytes: bytes, file_name: str, feedback_type: str) -> dict:
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        return PARSERS[feedback_type](tmp_path)
    finally:
        os.unlink(tmp_path)


def build_docs(file_bytes: bytes, feedback_type: str, course_info: dict) -> tuple:
    with tempfile.TemporaryDirectory() as tmpdir:
        excel_path = os.path.join(tmpdir, "responses.xlsx")
        with open(excel_path, "wb") as f:
            f.write(file_bytes)

        data        = PARSERS[feedback_type](excel_path)
        charts_dir  = os.path.join(tmpdir, "charts")
        chart_paths = generate_all_charts(data["question_data"], charts_dir)

        if feedback_type == "course_exit":
            action_taken = generate_all_course_exit_action_taken(data["question_data"])
        else:
            action_taken = generate_all_faculty_action_taken(data["question_data"])

        analysis_path = os.path.join(tmpdir, "Analysis.docx")
        atr_path      = os.path.join(tmpdir, "ATR.docx")

        if feedback_type == "course_exit":
            ce_analysis(data["question_data"], chart_paths, course_info, analysis_path)
            ce_atr(data["question_data"], action_taken, course_info, data["date_range"], atr_path)
        elif feedback_type == "faculty_endterm":
            fe_analysis(data["question_data"], chart_paths, course_info, analysis_path)
            fe_atr(data["question_data"], action_taken, course_info, data["date_range"], atr_path)
        elif feedback_type == "faculty_midterm":
            fm_analysis(data["question_data"], chart_paths, course_info, analysis_path)
            fm_atr(data["question_data"], action_taken, course_info, data["date_range"], atr_path)

        with open(analysis_path, "rb") as f: analysis_bytes = f.read()
        with open(atr_path,      "rb") as f: atr_bytes      = f.read()

    return analysis_bytes, atr_bytes


# -----------------------------------------------------------------------
# SESSION STATE
# -----------------------------------------------------------------------
for key, default in [
    ("generated", False), ("analysis_bytes", None), ("atr_bytes", None),
    ("course_code", ""),  ("drive_saved", False),    ("drive_urls", {}),
]:
    if key not in st.session_state:
        st.session_state[key] = default


# -----------------------------------------------------------------------
# UI
# -----------------------------------------------------------------------
st.markdown('<div class="app-title">NBA Document Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="app-subtitle">Techno Main Salt Lake &nbsp;·&nbsp; Course-Exit &amp; Faculty Feedback</div>', unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

# STEP 1 — Feedback type
st.markdown('<div class="section-label">Step 1 — Select Feedback Type</div>', unsafe_allow_html=True)
feedback_label = st.selectbox("Feedback type", list(FEEDBACK_TYPES.keys()), label_visibility="collapsed")
feedback_type  = FEEDBACK_TYPES[feedback_label]
badge_text, badge_class = TYPE_BADGE[feedback_type]
st.markdown(f'<span class="type-badge {badge_class}">{badge_text}</span>', unsafe_allow_html=True)

# STEP 2 — File upload
st.markdown('<div class="section-label">Step 2 — Upload Excel File</div>', unsafe_allow_html=True)
uploaded = st.file_uploader("Google Form Responses (.xlsx)", type=["xlsx","xls"], label_visibility="collapsed")

if not uploaded:
    st.info("Upload your Google Form Excel responses file to continue.")
    st.stop()

# STEP 3 — Parse and show detected fields
file_bytes = uploaded.read()
data       = cached_parse(file_bytes, uploaded.name, feedback_type)
ci         = data["course_info"]

st.markdown("<hr>", unsafe_allow_html=True)
st.markdown('<div class="section-label">Step 3 — Auto-Detected Fields</div>', unsafe_allow_html=True)

rows = [
    field_row("Department",    ci.get("department", ""),    bool(ci.get("department"))),
    field_row("Course Code",   ci.get("course_code", ""),   bool(ci.get("course_code"))),
    field_row("Academic Year", ci.get("academic_year", ""), bool(ci.get("academic_year"))),
    field_row("Semester",      ci.get("semester", ""),      bool(ci.get("semester"))),
    field_row("Respondents",   str(data["respondent_count"]),   True),
    field_row("Date Range",    data["date_range"]["range_str"], True),
]
if feedback_type in ("faculty_endterm", "faculty_midterm"):
    rows.insert(2, field_row("Faculty Name", ci.get("faculty_name", ""), bool(ci.get("faculty_name"))))
st.markdown("".join(rows), unsafe_allow_html=True)

# Stats table
st.markdown('<div class="section-label" style="margin-top:1.2rem">Question Statistics</div>', unsafe_allow_html=True)
if feedback_type == "course_exit":
    q_rows = "".join(
        f"<tr><td>{q['code']}</td>"
        f"<td style='color:#2d7a3e;font-weight:600'>{q['high_pct']}%</td>"
        f"<td style='color:#856404'>{q['moderate_pct']}%</td>"
        f"<td style='color:#c0392b'>{q['low_pct']}%</td></tr>"
        for q in data["question_data"]
    )
    st.markdown(f'<table class="co-table"><tr><th>CO</th><th>High</th><th>Moderate</th><th>Low</th></tr>{q_rows}</table>', unsafe_allow_html=True)
else:
    q_rows = "".join(
        f"<tr><td>{q['code']}</td>"
        f"<td style='color:#2d7a3e;font-weight:600'>{q['strongly_agree_pct']}%</td>"
        f"<td style='color:#4caf82'>{q['agree_pct']}%</td>"
        f"<td style='color:#856404'>{q['neutral_pct']}%</td>"
        f"<td style='color:#c0392b'>{q['disagree_pct']}%</td></tr>"
        for q in data["question_data"]
    )
    st.markdown(f'<table class="co-table"><tr><th>Q</th><th>Strongly Agree</th><th>Agree</th><th>Neutral</th><th>Disagree</th></tr>{q_rows}</table>', unsafe_allow_html=True)

# STEP 4 — Edit fields
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown('<div class="section-label">Step 4 — Confirm or Edit Fields</div>', unsafe_allow_html=True)

dept = st.selectbox("Department", DEPT_OPTIONS, index=0 if "CSBS" not in ci.get("department","") else 1)
col1, col2 = st.columns(2)
with col1: course_code   = st.text_input("Course Code",   value=ci.get("course_code",""),   placeholder="e.g. PEC-CS702A")
with col2: academic_year = st.text_input("Academic Year", value=ci.get("academic_year",""), placeholder="e.g. 2025 - 2026")

course_name = st.text_input("Course Name", value=ci.get("course_name",""), placeholder="e.g. Deep Learning and Neural Network")
semester    = st.selectbox("Semester", ["Odd Semester","Even Semester"], index=0 if ci.get("semester")!="Even Semester" else 1)

faculty_name = ""
if feedback_type in ("faculty_endterm", "faculty_midterm"):
    faculty_name = st.text_input("Faculty Name", value=ci.get("faculty_name",""), placeholder="e.g. Dr. Biman Roy")

# STEP 5 — Generate
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown('<div class="section-label">Step 5 — Generate</div>', unsafe_allow_html=True)

ready = all([dept, course_code.strip(), academic_year.strip()])
if not ready:
    st.warning("Department, Course Code and Academic Year are required.")

if st.button("⚙  Generate Analysis + ATR Documents", disabled=not ready):
    course_info = {
        "department":       dept,
        "course_code":      course_code.strip(),
        "course_name":      course_name.strip(),
        "academic_year":    academic_year.strip(),
        "semester":         semester,
        "respondent_count": data["respondent_count"],
        "faculty_name":     faculty_name.strip(),
    }
    with st.spinner("Building documents… this takes a few seconds"):
        try:
            ab, tb = build_docs(file_bytes, feedback_type, course_info)
            st.session_state.generated      = True
            st.session_state.analysis_bytes = ab
            st.session_state.atr_bytes      = tb
            st.session_state.course_code    = course_code.strip()
            st.session_state.drive_saved    = False
            st.session_state.drive_urls     = {}
        except Exception as e:
            st.error(f"Something went wrong: {e}")
            st.session_state.generated = False

# STEP 6 — Download + Drive
if st.session_state.generated:
    code = st.session_state.course_code
    st.success("Documents ready!")

    st.markdown('<div class="section-label">Download</div>', unsafe_allow_html=True)
    col_a, col_b = st.columns(2)
    with col_a:
        st.download_button("⬇ Download Analysis.docx", st.session_state.analysis_bytes,
            f"Analysis_{code}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    with col_b:
        st.download_button("⬇ Download ATR.docx", st.session_state.atr_bytes,
            f"ATR_{code}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    if is_drive_configured():
        st.markdown('<div class="section-label" style="margin-top:1.2rem">Google Drive</div>', unsafe_allow_html=True)
        if not st.session_state.get("drive_saved"):
            if st.button("☁  Save to Google Drive", key="drive_btn"):
                with st.spinner("Saving to Google Drive…"):
                    urls = save_both_to_drive(st.session_state.analysis_bytes, st.session_state.atr_bytes, code)
                st.session_state.drive_urls  = urls
                st.session_state.drive_saved = True
                st.rerun()
        if st.session_state.get("drive_saved"):
            urls = st.session_state.get("drive_urls", {})
            if urls.get("analysis_url") or urls.get("atr_url"):
                st.markdown("**Saved to your NBA Documents folder:**")
                if urls.get("analysis_url"): st.markdown(f"📄 [Analysis_{code}.docx]({urls['analysis_url']})")
                if urls.get("atr_url"):      st.markdown(f"📄 [ATR_{code}.docx]({urls['atr_url']})")
            else:
                st.error("Drive save failed — check your secrets configuration.")
