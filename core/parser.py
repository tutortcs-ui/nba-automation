# =============================================================================
# parser.py
# PURPOSE: Read the Excel file (Google Form responses) and extract everything
#          we need — CO names, ratings, respondent count, date range,
#          and optionally course details from the column headers.
# =============================================================================

import os
import pandas as pd
import re
from datetime import datetime


def load_excel(file_path: str) -> pd.DataFrame:
    """
    Load the Excel file into a pandas DataFrame.
    The first row is assumed to be headers (Timestamp, Email, CO1, CO2, ...).
    """
    df = pd.read_excel(file_path, sheet_name=0)
    return df


def extract_co_columns(df: pd.DataFrame) -> list[dict]:
    """
    Find all Course Outcome (CO) columns in the DataFrame.
    
    A CO column is any column that is NOT 'Timestamp' or 'Email Address'.
    Each CO column header contains the full CO description text.
    
    Returns a list of dicts like:
    [
        {
            "column_name": "PEC-702A.1 Students will understand...",
            "code": "PEC-702A.1",       # extracted from start of header
            "description": "Students will understand...",
            "index": 2                   # column position in DataFrame
        },
        ...
    ]
    """
    co_columns = []

    for i, col in enumerate(df.columns):
        col_stripped = str(col).strip()

        # Skip timestamp and email columns
        if col_stripped.lower() in ["timestamp", "email address"]:
            continue

        # Try to extract a CO code from the start of the header
        # Pattern: something like "PEC-702A.1" or "CO1" or "PCC-CS702A.1"
        code_match = re.match(r"^[\s\n]*([A-Z]+-?[A-Z0-9]*[\.\-][0-9]+)", col_stripped)

        if code_match:
            code = code_match.group(1).strip()
            # Description is everything after the code
            description = col_stripped[len(code):].strip().lstrip(".")
        else:
            # No recognizable code — use column position as fallback
            code = f"CO{len(co_columns) + 1}"
            description = col_stripped

        co_columns.append({
            "column_name": col,        # original column name (used to access DataFrame)
            "code": code,
            "description": description.strip(),
            "index": i
        })

    return co_columns


def calculate_statistics(df: pd.DataFrame, co_columns: list[dict]) -> list[dict]:
    """
    For each CO column, count how many students gave:
        3 = High, 2 = Moderate, 1 = Low
    Then convert to percentages.
    
    Returns co_columns list with added stats:
    {
        ...,
        "total": 26,
        "high_count": 23, "high_pct": 89,
        "moderate_count": 3, "moderate_pct": 11,
        "low_count": 0, "low_pct": 0
    }
    """
    results = []

    for co in co_columns:
        col = co["column_name"]

        # Get all non-null ratings for this CO
        ratings = df[col].dropna()
        total = len(ratings)

        if total == 0:
            high_count = moderate_count = low_count = 0
            high_pct = moderate_pct = low_pct = 0
        else:
            high_count = int((ratings == 3).sum())
            moderate_count = int((ratings == 2).sum())
            low_count = int((ratings == 1).sum())

            # Round percentages — ensure they sum to 100 using largest remainder method
            raw_high = (high_count / total) * 100
            raw_moderate = (moderate_count / total) * 100
            raw_low = (low_count / total) * 100

            high_pct = round(raw_high)
            moderate_pct = round(raw_moderate)
            low_pct = round(raw_low)

            # Correct rounding drift so percentages always sum to 100
            diff = 100 - (high_pct + moderate_pct + low_pct)
            if diff != 0:
                # Add the correction to the largest category
                maxval = max(high_pct, moderate_pct, low_pct)
                if high_pct == maxval:
                    high_pct += diff
                elif moderate_pct == maxval:
                    moderate_pct += diff
                else:
                    low_pct += diff

        results.append({
            **co,
            "total": total,
            "high_count": high_count,   "high_pct": high_pct,
            "moderate_count": moderate_count, "moderate_pct": moderate_pct,
            "low_count": low_count,     "low_pct": low_pct,
        })

    return results


def extract_date_range(df: pd.DataFrame) -> dict:
    """
    Extract the earliest and latest submission dates from the Timestamp column.
    
    Returns:
    {
        "start": datetime object,
        "end": datetime object,
        "start_str": "20th November 2025",
        "end_str": "21st November 2025",
        "range_str": "20th – 21st November 2025"   <-- used in ATR doc
    }
    """
    timestamp_col = None

    # Find the timestamp column (case-insensitive)
    for col in df.columns:
        if str(col).strip().lower() == "timestamp":
            timestamp_col = col
            break

    if timestamp_col is None:
        return {
            "start": None, "end": None,
            "start_str": "N/A", "end_str": "N/A",
            "range_str": "N/A"
        }

    timestamps = pd.to_datetime(df[timestamp_col], errors="coerce").dropna()

    if len(timestamps) == 0:
        return {
            "start": None, "end": None,
            "start_str": "N/A", "end_str": "N/A",
            "range_str": "N/A"
        }

    start_dt = timestamps.min()
    end_dt = timestamps.max()

    def ordinal(n: int) -> str:
        """Convert integer to ordinal string: 1 -> '1st', 2 -> '2nd', etc."""
        if 11 <= n % 100 <= 13:
            suffix = "th"
        else:
            suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
        return f"{n}{suffix}"

    start_str = f"{ordinal(start_dt.day)} {start_dt.strftime('%B %Y')}"
    end_str = f"{ordinal(end_dt.day)} {end_dt.strftime('%B %Y')}"

    # If same day, just show one date
    if start_dt.date() == end_dt.date():
        range_str = start_str
    else:
        # If same month/year, compress: "20th – 21st November 2025"
        if start_dt.month == end_dt.month and start_dt.year == end_dt.year:
            range_str = f"{ordinal(start_dt.day)} – {ordinal(end_dt.day)} {end_dt.strftime('%B %Y')}"
        else:
            range_str = f"{start_str} – {end_str}"

    return {
        "start": start_dt,
        "end": end_dt,
        "start_str": start_str,
        "end_str": end_str,
        "range_str": range_str
    }


def detect_department_from_filename(file_path: str) -> str:
    """
    Detect department from the Excel filename.
    Looks for 'CSBS' or 'CSE' in the filename.
    Returns the full department string, or "" if not detectable.
    """
    fname = os.path.basename(file_path).upper()
    if "CSBS" in fname:
        return "Department of Computer Science and Business Systems (CSBS)"
    if "CSE" in fname:
        return "Department of Computer Science and Engineering (CSE)"
    return ""


def extract_course_info_from_headers(co_columns: list[dict], df=None) -> dict:
    """
    Auto-detect as much course info as possible from the Excel file.

    Detects:
      - course_code: from CO header codes e.g. "PEC-702A.1" -> "PEC-702A"
      - course_name: inferred from CO description keywords
      - academic_year: from submission timestamps e.g. Nov 2025 -> "2025 - 2026"
      - semester: from submission month (Jul-Dec = Odd, Jan-Jun = Even)

    Returns dict (values may be empty string "" if undetectable):
    {
        "course_code":   "PEC-702A",
        "course_name":   "Deep Learning and Neural Network",
        "academic_year": "2025 - 2026",
        "semester":      "Odd Semester",
    }
    """
    result = {
        "course_code":   "",
        "course_name":   "",
        "academic_year": "",
        "semester":      "",
    }

    if not co_columns:
        return result

    # --- Course code: strip trailing .N from first CO code ---
    first_code = co_columns[0]["code"]
    result["course_code"] = re.sub(r"\.[0-9]+$", "", first_code)

    # --- Course name: collect unique meaningful keywords from all CO descriptions ---
    # Strategy: take the most distinctive noun phrases from descriptions
    # by finding words that appear across multiple COs (topic keywords)
    all_desc = " ".join(co["description"] for co in co_columns).lower()

    # Common domain keywords to look for
    domain_hints = [
        ("deep learning", "Deep Learning"),
        ("neural network", "Neural Network"),
        ("machine learning", "Machine Learning"),
        ("computer vision", "Computer Vision"),
        ("natural language", "Natural Language Processing"),
        ("data structure", "Data Structures"),
        ("algorithm", "Algorithms"),
        ("operating system", "Operating Systems"),
        ("database", "Database Management"),
        ("software engineering", "Software Engineering"),
        ("computer network", "Computer Networks"),
        ("artificial intelligence", "Artificial Intelligence"),
        ("cloud computing", "Cloud Computing"),
        ("cyber security", "Cyber Security"),
        ("data mining", "Data Mining"),
        ("compiler", "Compiler Design"),
        ("microprocessor", "Microprocessors"),
        ("digital signal", "Digital Signal Processing"),
        ("fuzzy", "Fuzzy Systems"),
        ("reinforcement", "Reinforcement Learning"),
    ]

    matched = []
    for keyword, display in domain_hints:
        if keyword in all_desc and display not in matched:
            matched.append(display)

    if matched:
        result["course_name"] = " and ".join(matched[:2])  # max 2 topics

    # --- Academic year and semester: from timestamps ---
    if df is not None:
        for col in df.columns:
            if str(col).strip().lower() == "timestamp":
                timestamps = pd.to_datetime(df[col], errors="coerce").dropna()
                if len(timestamps) > 0:
                    latest = timestamps.max()
                    month = latest.month
                    year  = latest.year

                    # Odd semester: July-December -> academic year starts that year
                    # Even semester: January-June -> academic year started previous year
                    if month >= 7:
                        result["academic_year"] = f"{year} - {year + 1}"
                        result["semester"]      = "Odd Semester"
                    else:
                        result["academic_year"] = f"{year - 1} - {year}"
                        result["semester"]      = "Even Semester"
                break

    return result


def parse_excel_file(file_path: str) -> dict:
    """
    MAIN ENTRY POINT for this module.

    Takes an Excel file path, returns a complete data dictionary
    with everything needed to generate both documents.

    Returns:
    {
        "df": DataFrame,
        "co_data": [list of CO dicts with stats],
        "respondent_count": 26,
        "date_range": { "range_str": "20th - 21st November 2025", ... },
        "course_info": {
            "course_code":   "PEC-702A",
            "course_name":   "Deep Learning and Neural Network",
            "academic_year": "2025 - 2026",
            "semester":      "Odd Semester",
            "department":    "Department of Computer Science and Engineering (CSE)",
        }
    }
    """
    df = load_excel(file_path)
    co_columns = extract_co_columns(df)
    co_data = calculate_statistics(df, co_columns)
    date_range = extract_date_range(df)
    course_info = extract_course_info_from_headers(co_columns, df=df)

    # Add department — detected from filename
    course_info["department"] = detect_department_from_filename(file_path)

    # Respondent count = number of rows (each row = one student response)
    respondent_count = len(df)

    return {
        "df": df,
        "co_data": co_data,
        "respondent_count": respondent_count,
        "date_range": date_range,
        "course_info": course_info,
        "course_code": course_info["course_code"],  # kept for convenience
    }
