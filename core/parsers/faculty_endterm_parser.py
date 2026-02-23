# =============================================================================
# core/parsers/faculty_endterm_parser.py
# PURPOSE: Parse Faculty End-Term feedback Excel (Google Form responses).
#
# WHAT MAKES THIS FEEDBACK TYPE UNIQUE:
#   - Questions are plain teaching quality statements (no CO codes)
#   - Ratings are text: "Strongly agree", "Agree", "Neutral", "Disagree"
#   - One pie chart per question with 4 slices (SA / A / N / D)
#   - Document includes "Name of Faculty" line
#   - Typically collected Nov–Dec (End of Odd Semester)
#
# OUTPUT FORMAT (returned by parse()):
# {
#   "feedback_type":    "faculty_endterm",
#   "respondent_count": 22,
#   "date_range":       { "range_str": "17th – 24th November 2025", ... },
#   "course_info":      { "course_code": ..., "faculty_name": ..., ... },
#   "question_data":    [
#       {
#           "code":                 "Q1",
#           "description":         "COs for the course were disseminated well in advance.",
#           "strongly_agree_pct":  73, "agree_pct": 27,
#           "neutral_pct":         0,  "disagree_pct": 0,
#           "total":               22,
#           "slices": [...]
#       }, ...
#   ]
# }
# =============================================================================

import os
import re
import pandas as pd
from core.shared.charts import make_faculty_slices


# Rating values — case-insensitive matching
_SA_VALUES = {"strongly agree", "strongly agreed"}
_A_VALUES  = {"agree", "agreed"}
_N_VALUES  = {"neutral"}
_D_VALUES  = {"disagree", "disagreed", "strongly disagree", "strongly disagreed"}


def _load_excel(file_path: str) -> pd.DataFrame:
    return pd.read_excel(file_path, sheet_name=0)


def _extract_question_columns(df: pd.DataFrame) -> list:
    """
    All columns except Timestamp and Email are questions.
    Questions get Q1, Q2, Q3... as codes.
    """
    questions = []
    q_num = 1
    for col in df.columns:
        stripped = str(col).strip()
        if stripped.lower() in ["timestamp", "email address", "e-mail"]:
            continue
        questions.append({
            "column_name": col,
            "code":        f"Q{q_num}",
            "description": stripped,
        })
        q_num += 1
    return questions


def _normalise(val: str) -> str:
    return str(val).strip().lower()


def _calculate_stats(df: pd.DataFrame, questions: list) -> list:
    """
    Count SA / A / N / D responses per question.
    Percentages rounded and corrected to sum to 100.
    """
    results = []
    for i, q in enumerate(questions):
        ratings = df[q["column_name"]].dropna().apply(_normalise)
        total   = len(ratings)

        if total == 0:
            sa = a = n = d = 0
            sa_pct = a_pct = n_pct = d_pct = 0
        else:
            sa = int(ratings.isin(_SA_VALUES).sum())
            a  = int(ratings.isin(_A_VALUES).sum())
            n  = int(ratings.isin(_N_VALUES).sum())
            d  = int(ratings.isin(_D_VALUES).sum())

            sa_pct = round(sa / total * 100)
            a_pct  = round(a  / total * 100)
            n_pct  = round(n  / total * 100)
            d_pct  = round(d  / total * 100)

            # Fix rounding drift
            diff = 100 - (sa_pct + a_pct + n_pct + d_pct)
            if diff != 0:
                maxval = max(sa_pct, a_pct, n_pct, d_pct)
                if sa_pct == maxval:   sa_pct += diff
                elif a_pct == maxval:  a_pct  += diff
                elif n_pct == maxval:  n_pct  += diff
                else:                  d_pct  += diff

        results.append({
            **q,
            "total":               total,
            "strongly_agree_count":sa,    "strongly_agree_pct": sa_pct,
            "agree_count":         a,     "agree_pct":          a_pct,
            "neutral_count":       n,     "neutral_pct":        n_pct,
            "disagree_count":      d,     "disagree_pct":       d_pct,
            "slices": make_faculty_slices(sa_pct, a_pct, n_pct, d_pct),
        })
    return results


def _extract_date_range(df: pd.DataFrame) -> dict:
    for col in df.columns:
        if str(col).strip().lower() == "timestamp":
            ts = pd.to_datetime(df[col], errors="coerce").dropna()
            if len(ts) == 0:
                break
            s, e = ts.min(), ts.max()

            def ord_str(dt):
                n   = dt.day
                sfx = "th" if 11 <= n % 100 <= 13 else {1:"st",2:"nd",3:"rd"}.get(n%10,"th")
                return f"{n}{sfx} {dt.strftime('%B %Y')}"

            s_str, e_str = ord_str(s), ord_str(e)
            if s.date() == e.date():
                rng = s_str
            elif s.month == e.month and s.year == e.year:
                n   = s.day
                sfx = "th" if 11 <= n % 100 <= 13 else {1:"st",2:"nd",3:"rd"}.get(n%10,"th")
                rng = f"{n}{sfx} – {ord_str(e)}"
            else:
                rng = f"{s_str} – {e_str}"

            return {"start":s, "end":e, "start_str":s_str, "end_str":e_str, "range_str":rng}

    return {"start":None,"end":None,"start_str":"N/A","end_str":"N/A","range_str":"N/A"}


def _detect_course_info(df: pd.DataFrame, file_path: str) -> dict:
    """
    Detect course code, name, year, semester, department and faculty name
    from filename and timestamps.
    Faculty name defaults to empty string (shown in UI for user to fill).
    """
    info = {
        "course_code":   "",
        "course_name":   "",
        "academic_year": "",
        "semester":      "",
        "department":    "",
        "faculty_name":  "",   # unique to faculty feedback
    }

    # Course code from filename e.g. "PCC-CS702A" or "PEC-702A"
    fname = os.path.basename(file_path)
    code_match = re.search(r"([A-Z]+-?(?:CS)?[0-9]+[A-Z]?)", fname, re.IGNORECASE)
    if code_match:
        info["course_code"] = code_match.group(1).upper()

    # Year + semester from timestamps
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
    upper = fname.upper()
    if "CSBS" in upper:
        info["department"] = "Department of Computer Science and Business Systems (CSBS)"
    elif "CSE" in upper:
        info["department"] = "Department of Computer Science and Engineering (CSE)"

    return info


# -----------------------------------------------------------------------
# PUBLIC ENTRY POINT
# -----------------------------------------------------------------------

def parse(file_path: str) -> dict:
    """
    Parse a Faculty End-Term feedback Excel file.
    Returns the standard data dict used by faculty_endterm_builder.py.
    """
    df            = _load_excel(file_path)
    questions     = _extract_question_columns(df)
    question_data = _calculate_stats(df, questions)
    date_range    = _extract_date_range(df)
    course_info   = _detect_course_info(df, file_path)

    print(f"  ✓ Respondents : {len(df)}")
    print(f"  ✓ Questions   : {len(question_data)}")
    print(f"  ✓ Date range  : {date_range['range_str']}")
    print(f"  ✓ Course code : {course_info['course_code']}")

    return {
        "feedback_type":    "faculty_endterm",
        "respondent_count": len(df),
        "date_range":       date_range,
        "course_info":      course_info,
        "question_data":    question_data,
    }
