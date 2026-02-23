# =============================================================================
# core/shared/charts.py
# PURPOSE: Generate pie charts for any feedback type.
#
# Supports two modes:
#   - 3-slice: Course-Exit feedback (High / Moderate / Low)
#   - 4-slice: Faculty feedback (Strongly Agree / Agree / Neutral / Disagree)
#
# Both produce the same visual style — same colors, same layout,
# same file format — so the document builders treat them identically.
#
# USAGE:
#   from core.shared.charts import generate_pie_chart, generate_all_charts
#   paths = generate_all_charts(question_data, "temp_charts/")
# =============================================================================

import os
import matplotlib
matplotlib.use("Agg")   # non-interactive backend — required on servers
import matplotlib.pyplot as plt


# -----------------------------------------------------------------------
# Color palette — consistent across all feedback types
# -----------------------------------------------------------------------
# Course-Exit (3 slices)
COLOR_HIGH     = "#2864c8"   # blue
COLOR_MODERATE = "#f08c00"   # amber
COLOR_LOW      = "#dc2800"   # red

# Faculty (4 slices) — SA and A share the blue/green family; N and D warm
COLOR_STRONGLY_AGREE = "#2864c8"   # blue
COLOR_AGREE          = "#4caf82"   # green
COLOR_NEUTRAL        = "#f08c00"   # amber
COLOR_DISAGREE       = "#dc2800"   # red

COLOR_TEXT = "#1a1a2e"


def generate_pie_chart(question: dict, output_dir: str) -> str:
    """
    Generate one pie chart for a single question/CO.

    The question dict must have:
      - "code"        : label shown in chart title (e.g. "PEC-702A.1" or "Q1")
      - "description" : question text shown in chart title
      - "slices"      : list of {"label": str, "pct": int, "color": str}
                        only slices with pct > 0 are drawn

    Returns:
        Path to the saved PNG file.
    """
    os.makedirs(output_dir, exist_ok=True)

    # Filter out zero-value slices (empty wedges look bad)
    slices  = [s for s in question["slices"] if s["pct"] > 0]
    labels  = [s["label"] for s in slices]
    pcts    = [s["pct"]   for s in slices]
    colors  = [s["color"] for s in slices]

    fig, ax = plt.subplots(figsize=(6, 4.5))
    fig.patch.set_facecolor("white")

    wedges, texts, autotexts = ax.pie(
        pcts,
        labels=None,
        colors=colors,
        autopct="%1.0f%%",
        startangle=90,
        wedgeprops={"edgecolor": "white", "linewidth": 2},
        pctdistance=0.65,
    )

    # White bold % labels inside slices
    for autotext in autotexts:
        autotext.set_color("white")
        autotext.set_fontweight("bold")
        autotext.set_fontsize(12)

    # Legend below chart: "High – 89%", "Moderate – 11%", etc.
    ax.legend(
        wedges,
        [f"{l} – {p}%" for l, p in zip(labels, pcts)],
        loc="lower center",
        bbox_to_anchor=(0.5, -0.18),
        ncol=len(slices),
        fontsize=10,
        frameon=False,
        labelcolor=COLOR_TEXT,
    )

    # Title: code + truncated description
    short_desc = question["description"][:65] + ("..." if len(question["description"]) > 65 else "")
    ax.set_title(
        f"{question['code']}\n{short_desc}",
        fontsize=10,
        fontweight="bold",
        color=COLOR_TEXT,
        pad=12,
    )

    plt.tight_layout()

    # Safe filename: replace characters not allowed in filenames
    safe_code = question["code"].replace("/", "-").replace("\\", "-").replace(":", "-")
    filepath  = os.path.join(output_dir, f"chart_{safe_code}.png")
    plt.savefig(filepath, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    return filepath


def generate_all_charts(question_data: list, output_dir: str) -> list:
    """
    Generate one chart per question/CO.

    Args:
        question_data: list of question dicts (each must have "slices" key)
        output_dir:    folder to save PNGs into

    Returns:
        List of PNG file paths in same order as question_data.
    """
    paths = []
    for q in question_data:
        path = generate_pie_chart(q, output_dir)
        paths.append(path)
        print(f"  ✓ Chart: {path}")
    return paths


# -----------------------------------------------------------------------
# Slice builders — called by parsers to attach slice data to questions
# -----------------------------------------------------------------------

def make_course_exit_slices(high_pct: int, moderate_pct: int, low_pct: int) -> list:
    """
    Build slice list for Course-Exit feedback (3-slice: High/Moderate/Low).
    Used by course_exit_parser.py.
    """
    return [
        {"label": "High",     "pct": high_pct,     "color": COLOR_HIGH},
        {"label": "Moderate", "pct": moderate_pct, "color": COLOR_MODERATE},
        {"label": "Low",      "pct": low_pct,      "color": COLOR_LOW},
    ]


def make_faculty_slices(sa_pct: int, agree_pct: int, neutral_pct: int, disagree_pct: int) -> list:
    """
    Build slice list for Faculty feedback (4-slice: SA/A/N/D).
    Used by faculty_endterm_parser.py and faculty_midterm_parser.py.
    """
    return [
        {"label": "Strongly Agree", "pct": sa_pct,       "color": COLOR_STRONGLY_AGREE},
        {"label": "Agree",          "pct": agree_pct,    "color": COLOR_AGREE},
        {"label": "Neutral",        "pct": neutral_pct,  "color": COLOR_NEUTRAL},
        {"label": "Disagree",       "pct": disagree_pct, "color": COLOR_DISAGREE},
    ]
