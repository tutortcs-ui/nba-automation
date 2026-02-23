# =============================================================================
# core/parsers/course_exit_parser.py
# PURPOSE: Parse Course-Exit feedback Excel (Google Form responses).
#
# WHAT MAKES THIS FEEDBACK TYPE UNIQUE:
#   - Columns contain CO codes in the header: "PEC-702A.1 Students will..."
#   - Ratings are numeric: 3 = High, 2 = Moderate, 1 = Low
#   - One pie chart per CO with 3 slices
#
# OUTPUT FORMAT (returned by parse()):
# {
#   "feedback_type":    "course_exit",
#   "respondent_count": 26,
#   "date_range":       { "range_str": "20th – 21st November 2025", ... },
#   "course_info":      { "course_code": "PEC-702A", "course_name": ..., ... },
#   "question_data":    [
#       {
#           "code":          "PEC-702A.1",
#           "description":   "Students will understand...",
#           "high_pct":      89, "moderate_pct": 11, "low_pct": 0,
#           "high_count":    23, "moderate_count": 3, "low_count": 0,
#           "total":         26,
#           "slices": [{"label":"High","pct":89,"color":"#2864c8"}, ...]
#       }, ...
#   ]
# }
# =============================================================================

import os
import re
import pandas as pd
from core.shared.charts import make_course_exit_slices


# -----------------------------------------------------------------------
# INTERNAL HELPERS
# -----------------------------------------------------------------------

def _load_excel(file_path: str) -> pd.DataFrame:
    return pd.read_excel(file_path, sheet_name=0)


def _extract_co_columns(df: pd.DataFrame) -> list:
    """
    Find all CO columns — any column that is not Timestamp or Email.
    CO codes are extracted from the header text (e.g. "PEC-702A.1").
    """
    co_columns = []
    for i, col in enumerate(df.columns):
        stripped = str(col).strip()
        if stripped.lower() in ["timestamp", "email address", "e-mail"]:
            continue

        # Try to extract CO code from start of header
        match = re.match(r"^[\s\n]*([A-Z]+-?[A-Z0-9]*[.\-][0-9]+)", stripped)
        if match:
            code        = match.group(1).strip()
            description = stripped[len(code):].strip().lstrip(".")
        else:
            code        = f"CO{len(co_columns) + 1}"
            description = stripped

        co_columns.append({
            "column_name": col,
            "code":        code,
            "description": description.strip(),
            "index":       i,
        })
    return co_columns


def _calculate_stats(df: pd.DataFrame, co_columns: list) -> list:
    """
    Count High (3) / Moderate (2) / Low (1) ratings per CO.
    Percentages are rounded and corrected to sum to exactly 100.
    """
    results = []
    for co in co_columns:
        ratings = df[co["column_name"]].dropna()
        total   = len(ratings)

        if total == 0:
            h = m = l = h_pct = m_pct = l_pct = 0
        else:
            h = int((ratings == 3).sum())
            m = int((ratings == 2).sum())
            l = int((ratings == 1).sum())

            h_pct = round(h / total * 100)
            m_pct = round(m / total * 100)
            l_pct = round(l / total * 100)

            # Fix rounding so total is exactly 100
            diff = 100 - (h_pct + m_pct + l_pct)
            if diff != 0:
                if h_pct >= m_pct and h_pct >= l_pct:
                    h_pct += diff
                elif m_pct >= l_pct:
                    m_pct += diff
                else:
                    l_pct += diff

        results.append({
            **co,
            "total":          total,
            "high_count":     h,     "high_pct":     h_pct,
            "moderate_count": m,     "moderate_pct": m_pct,
            "low_count":      l,     "low_pct":      l_pct,
            "slices":         make_course_exit_slices(h_pct, m_pct, l_pct),
        })
    return results


def _extract_date_range(df: pd.DataFrame) -> dict:
    """Extract min/max submission dates and format as ordinal strings."""
    for col in df.columns:
        if str(col).strip().lower() == "timestamp":
            ts = pd.to_datetime(df[col], errors="coerce").dropna()
            if len(ts) == 0:
                break
            s, e = ts.min(), ts.max()

            def ord_str(dt):
                n = dt.day
                sfx = "th" if 11 <= n % 100 <= 13 else {1:"st",2:"nd",3:"rd"}.get(n%10,"th")
                return f"{n}{sfx} {dt.strftime('%B %Y')}"

            s_str = ord_str(s)
            e_str = ord_str(e)
            if s.date() == e.date():
                rng = s_str
            elif s.month == e.month and s.year == e.year:
                n = s.day
                sfx = "th" if 11 <= n % 100 <= 13 else {1:"st",2:"nd",3:"rd"}.get(n%10,"th")
                rng = f"{n}{sfx} – {ord_str(e)}"
            else:
                rng = f"{s_str} – {e_str}"

            return {"start": s, "end": e, "start_str": s_str, "end_str": e_str, "range_str": rng}

    return {"start": None, "end": None, "start_str": "N/A", "end_str": "N/A", "range_str": "N/A"}


def _detect_course_info(co_columns: list, df: pd.DataFrame, file_path: str) -> dict:
    """Auto-detect course code, name, year, semester, department from file."""
    info = {"course_code": "", "course_name": "", "academic_year": "", "semester": "", "department": ""}

    if co_columns:
        info["course_code"] = re.sub(r"\.[0-9]+$", "", co_columns[0]["code"])

    # Course name from CO description keywords
    all_desc = " ".join(co["description"] for co in co_columns).lower()
    domain_hints = [
        ("deep learning", "Deep Learning"), ("neural network", "Neural Network"),
        ("machine learning", "Machine Learning"), ("computer vision", "Computer Vision"),
        ("natural language", "Natural Language Processing"),
        ("data structure", "Data Structures"), ("algorithm", "Algorithms"),
        ("operating system", "Operating Systems"), ("database", "Database Management"),
        ("software engineering", "Software Engineering"),
        ("computer network", "Computer Networks"),
        ("artificial intelligence", "Artificial Intelligence"),
        ("cloud computing", "Cloud Computing"), ("cyber security", "Cyber Security"),
        ("data mining", "Data Mining"), ("compiler", "Compiler Design"),
        ("microprocessor", "Microprocessors"),
        ("fuzzy", "Fuzzy Systems"), ("reinforcement", "Reinforcement Learning"),
    ]
    matched = [disp for kw, disp in domain_hints if kw in all_desc]
    if matched:
        info["course_name"] = " and ".join(matched[:2])

    # Year and semester from timestamps
    for col in df.columns:
        if str(col).strip().lower() == "timestamp":
            ts = pd.to_datetime(df[col], errors="coerce").dropna()
            if len(ts) > 0:
                m, y = ts.max().month, ts.max().year
                if m >= 7:
                    info["academic_year"] = f"{y} - {y+1}"
                    info["semester"]      = "Odd Semester"
                else:
                    info["academic_year"] = f"{y-1} - {y}"
                    info["semester"]      = "Even Semester"
            break

    # Department from filename
    fname = os.path.basename(file_path).upper()
    if "CSBS" in fname:
        info["department"] = "Department of Computer Science and Business Systems (CSBS)"
    elif "CSE" in fname:
        info["department"] = "Department of Computer Science and Engineering (CSE)"

    return info


# -----------------------------------------------------------------------
# PUBLIC ENTRY POINT
# -----------------------------------------------------------------------

def parse(file_path: str) -> dict:
    """
    Parse a Course-Exit feedback Excel file.

    Returns the standard data dict used by course_exit_builder.py.
    """
    df          = _load_excel(file_path)
    co_columns  = _extract_co_columns(df)
    question_data = _calculate_stats(df, co_columns)
    date_range  = _extract_date_range(df)
    course_info = _detect_course_info(co_columns, df, file_path)

    print(f"  ✓ Respondents  : {len(df)}")
    print(f"  ✓ COs found    : {len(question_data)}")
    print(f"  ✓ Date range   : {date_range['range_str']}")
    print(f"  ✓ Course code  : {course_info['course_code']}")

    return {
        "feedback_type":    "course_exit",
        "respondent_count": len(df),
        "date_range":       date_range,
        "course_info":      course_info,
        "question_data":    question_data,
    }
