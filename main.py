# =============================================================================
# main.py  —  Stage 1, Iteration 1
# Run:  python main.py
#
# Generates TWO documents from one Excel file:
#   1. Analysis_[code].docx  — CO% table + pie charts
#   2. ATR_[code].docx       — Feedback + Action Taken table
#
# Auto-detects from filename + Excel:
#   department, course code, course name, academic year, semester
# Only asks if something genuinely could not be detected.
# =============================================================================

import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.parser import parse_excel_file
from core.templates import generate_all_action_taken
from core.document_generator import (
    generate_all_charts,
    build_analysis_document,
    build_atr_document,
)

OUTPUT_DIR = "output"
CHARTS_DIR = "temp_charts"

DEPT_OPTIONS = {
    "1": "Department of Computer Science and Engineering (CSE)",
    "2": "Department of Computer Science and Business Systems (CSBS)",
}


def get_excel_path() -> str:
    """Ask for Excel path, keep asking until a valid file is found."""
    print("\nPaste the full path to your Google Form Excel responses file.")
    print("Tip: In Windows Explorer, Shift+Right-click the file -> 'Copy as path'\n")

    while True:
        path = input("  Excel file path: ").strip().strip('"').strip("'")
        if not path:
            print("  Please enter a file path.")
            continue
        if os.path.exists(path):
            return path
        if os.path.exists(path + ".xlsx"):
            return path + ".xlsx"
        print(f"\n  File not found: {path}")
        print("  Use Shift+Right-click -> 'Copy as path' to get the correct path.\n")


def ask_if_missing(field_name: str, current_value: str, prompt: str) -> str:
    """Only ask the user for a value if it wasn't auto-detected."""
    if current_value:
        return current_value
    print(f"\n  Could not auto-detect: {field_name}")
    while True:
        val = input(f"  Please enter {prompt}: ").strip()
        if val:
            return val
        print("  This field is required.")


def ask_department(current_value: str) -> str:
    """If department wasn't detected from filename, show a menu."""
    if current_value:
        return current_value
    print("\n  Could not detect department from filename.")
    print("  Please choose:")
    print("    1. Department of Computer Science and Engineering (CSE)")
    print("    2. Department of Computer Science and Business Systems (CSBS)")
    while True:
        choice = input("  Enter 1 or 2: ").strip()
        if choice in DEPT_OPTIONS:
            return DEPT_OPTIONS[choice]
        print("  Please enter 1 or 2.")


if __name__ == "__main__":

    print("\n" + "="*52)
    print("  NBA Accreditation Document Generator")
    print("="*52)

    # STEP 1 — Get Excel file path
    excel_path = get_excel_path()

    # STEP 2 — Parse Excel
    print("\n[1/5] Reading Excel file...")
    data = parse_excel_file(excel_path)
    ci   = data["course_info"]

    print(f"  ✓ Respondents  : {data['respondent_count']}")
    print(f"  ✓ COs found    : {len(data['co_data'])}")
    print(f"  ✓ Date range   : {data['date_range']['range_str']}")
    print(f"  ✓ Course code  : {ci['course_code']   or '(not detected)'}")
    print(f"  ✓ Course name  : {ci['course_name']   or '(not detected)'}")
    print(f"  ✓ Academic year: {ci['academic_year'] or '(not detected)'}")
    print(f"  ✓ Semester     : {ci['semester']      or '(not detected)'}")
    print(f"  ✓ Department   : {ci['department']    or '(not detected)'}")
    for co in data["co_data"]:
        print(f"  ✓ {co['code']}: High={co['high_pct']}%  Mod={co['moderate_pct']}%  Low={co['low_pct']}%")

    # STEP 3 — Ask only for fields that couldn't be detected
    print("\n[2/5] Confirming course details...")
    department    = ask_department(ci["department"])
    course_code   = ask_if_missing("Course code",   ci["course_code"],   "course code (e.g. PEC-CS702A)")
    course_name   = ask_if_missing("Course name",   ci["course_name"],   "course name (e.g. Deep Learning)")
    academic_year = ask_if_missing("Academic year", ci["academic_year"], "academic year (e.g. 2025 - 2026)")
    semester      = ask_if_missing("Semester",      ci["semester"],      "semester (Odd Semester / Even Semester)")

    course_info = {
        "department":       department,
        "course_code":      course_code,
        "course_name":      course_name,
        "academic_year":    academic_year,
        "semester":         semester,
        "respondent_count": data["respondent_count"],
    }

    # STEP 4 — Generate pie charts
    print("\n[3/5] Generating pie charts...")
    chart_paths = generate_all_charts(data["co_data"], CHARTS_DIR)

    # STEP 5 — Generate Action Taken sentences for ATR
    print("\n[4/5] Generating Action Taken sentences...")
    action_taken_texts = generate_all_action_taken(data["co_data"])
    for co, text in zip(data["co_data"], action_taken_texts):
        print(f"  ✓ {co['code']}: {text[:70]}...")

    # STEP 6 — Build both documents
    print("\n[5/5] Building documents...")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    analysis_path = os.path.join(OUTPUT_DIR, f"Analysis_{course_info['course_code']}.docx")
    build_analysis_document(
        co_data=data["co_data"],
        chart_paths=chart_paths,
        course_info=course_info,
        output_path=analysis_path,
    )

    atr_path = os.path.join(OUTPUT_DIR, f"ATR_{course_info['course_code']}.docx")
    build_atr_document(
        co_data=data["co_data"],
        action_taken_texts=action_taken_texts,
        course_info=course_info,
        date_range=data["date_range"],
        output_path=atr_path,
    )

    # STEP 7 — Done
    print("\n" + "="*52)
    print("Done! Your files are saved at:")
    print(f"  {os.path.abspath(analysis_path)}")
    print(f"  {os.path.abspath(atr_path)}")
    print("="*52 + "\n")
