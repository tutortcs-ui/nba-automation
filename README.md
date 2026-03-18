# NBA Accreditation Document Generator

Automates repetitive NBA (National Board of Accreditation) documentation for
course-exit feedback. Give it one Excel file → get two formatted Word documents.

---

## What it does

**Input:** Google Form Excel export (student course-exit feedback responses)

**Output — two Word documents, auto-named by course code:**
1. `Analysis_[code].docx` — CO% table + pie charts
2. `ATR_[code].docx`       — Feedback & Action Taken Report

Both documents include:
- TMSL header (top-right: college name + address + horizontal line)
- TMSL footer (© TMSL left | Page X of Y right + horizontal line)
- Correct fonts, borders, and formatting matching the original institutional style
- Signature block at the bottom

---

## Setup (one time only)

### 1. Install Python
Download Python 3.11 or newer from https://python.org and install it.
During install, check "Add Python to PATH".

### 2. Unzip the project
Extract `nba_automation.zip` anywhere on your computer, e.g. `D:\Projects\nba_automation`

### 3. Install dependencies
Open a terminal (Windows: press `Win + R`, type `cmd`, press Enter).
Navigate to the project folder:

```
cd D:\Projects\nba_automation
pip install -r requirements.txt
```

This installs: pandas, openpyxl, python-docx, matplotlib, lxml

### 4. Verify assets folder exists
The `assets/` folder must contain these 3 files (they come with the zip):
- `header1.xml`
- `footer1.xml`
- `image5.png`

These are the exact header/footer extracted from the original TMSL documents.
Do NOT delete or rename them.

---

## Running it

```
cd D:\Projects\nba_automation
python main.py
```

The script will ask for:

1. **Excel file path** — paste the full path to your Google Form responses file
   - Tip: In Windows Explorer, Shift+Right-click the file → "Copy as path", then paste

2. **Only if auto-detection fails** — it will ask for any field it couldn't figure out:
   - Department (shows a menu: 1 = CSE, 2 = CSBS)
   - Course code, course name, academic year, or semester

Everything else is detected automatically from the file.

### Example session

```
==================================================
  NBA Accreditation Document Generator
==================================================

Paste the full path to your Google Form Excel responses file.
Tip: In Windows Explorer, Shift+Right-click the file -> 'Copy as path'

  Excel file path: "C:\Users\you\Downloads\DL-CSE7-Course Exit... (Responses).xlsx"

[1/4] Reading Excel file...
  ✓ Respondents  : 26
  ✓ COs found    : 4
  ✓ Date range   : 20th – 21st November 2025
  ✓ Course code  : PEC-702A
  ✓ Course name  : Deep Learning and Neural Network
  ✓ Academic year: 2025 - 2026
  ✓ Semester     : Odd Semester
  ✓ Department   : Department of Computer Science and Engineering (CSE)

[2/4] Generating pie charts...
[3/4] Generating Action Taken sentences...
[4/4] Building documents...
  ✓ Analysis document saved: output\Analysis_PEC-702A.docx
  ✓ ATR document saved: output\ATR_PEC-702A.docx

==================================================
Done! Your files are saved at:
  D:\Projects\nba_automation\output\Analysis_PEC-702A.docx
  D:\Projects\nba_automation\output\ATR_PEC-702A.docx
==================================================
```

---

## Auto-detection logic

| Field | How it's detected |
|---|---|
| Department | `CSE` or `CSBS` found in the Excel filename |
| Course code | CO header codes e.g. `PEC-702A.1` → `PEC-702A` |
| Course name | Keywords in CO descriptions (Deep Learning, Neural Network, etc.) |
| Academic year | Submission timestamps — Nov 2025 → `2025 - 2026` |
| Semester | Submission month — Jul–Dec = Odd, Jan–Jun = Even |
| Respondent count | Number of rows in the Excel file |
| Date range | Min and max of Timestamp column |

---

## File structure

```
nba_automation/
│
├── main.py                  ← Run this. Asks for input, calls everything.
│
├── core/
│   ├── parser.py            ← Reads Excel, calculates stats, detects course info
│   ├── document_generator.py← Builds pie charts + both Word documents
│   └── templates.py         ← Generates Action Taken sentences for ATR
│
├── assets/                  ← DO NOT DELETE — header/footer XML from original docs
│   ├── header1.xml
│   ├── footer1.xml
│   └── image5.png
│
├── output/                  ← Generated .docx files appear here
├── temp_charts/             ← Pie chart PNGs (created during run, safe to delete)
├── requirements.txt
└── README.md
```

---

## Document specs

### Analysis document
- Header paragraphs: bold, 12.5pt, centered
  - Department name
  - "Analysis of Course-Exit Feedback (End-Term)"
  - "(CAY: 2025 - 2026)"
  - "No. of Respondents: 26"
- Title box: bold italic, "% of Course-Exit Feedback Analysis (code - name)"
- Data table: High / Moderate / Low % per CO, 10pt, black borders
- Pie charts: one per row, stacked vertically, 14cm wide
  - Blue = High, Amber = Moderate, Red = Low
- Signature bottom-right: dotted line + "(Submitted by – Dr. Biman Roy)"

### ATR document
- Same IQAC headers: bold, 12.5pt, centered
- Intro paragraph: 9.75pt justified, date range in bold italic
- Table: Sl.No (narrow ~1cm) | Feedback | Action Taken
  - Serial numbers 1, 2, 3... filled in automatically
  - Action Taken sentences generated from CO % data
- Signature bottom-right: dotted line + PAC ratification text

---

## Departments supported

| Code in filename | Full department name |
|---|---|
| `CSE` | Department of Computer Science and Engineering (CSE) |
| `CSBS` | Department of Computer Science and Business Systems (CSBS) |

If neither is found in the filename, the script asks you to choose.

---

## Adding more course name keywords

If the auto-detected course name is wrong or missing, edit `core/parser.py`
and find the `domain_hints` list inside `extract_course_info_from_headers()`.
Add a new entry:

```python
("your keyword", "Your Course Name"),
```

---

## Planned future stages

- **Stage 2:** Streamlit web UI — upload Excel in browser, download Word files
- **Stage 3:** Google Drive integration — auto-save outputs to your Drive
- **Stage 4:** Google OAuth login — for use by colleagues
.
