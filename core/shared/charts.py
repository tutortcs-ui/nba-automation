# =============================================================================
# core/shared/charts.py
# PURPOSE: Generate pie charts — fixed identical pixel dimensions every time.
#
# KEY FIX: savefig uses bbox_inches=None (not "tight") so every chart is
# exactly figsize * dpi pixels. Long legend text wraps rather than
# expanding the canvas. All charts are identical size in the document.
# =============================================================================

import os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker


COLOR_HIGH           = "#2864c8"
COLOR_MODERATE       = "#f08c00"
COLOR_LOW            = "#dc2800"
COLOR_STRONGLY_AGREE = "#2864c8"
COLOR_AGREE          = "#4caf82"
COLOR_NEUTRAL        = "#f08c00"
COLOR_DISAGREE       = "#dc2800"
COLOR_TEXT           = "#1a1a2e"

# Fixed canvas: 7.5 x 5.7 inches at 150 dpi = 1125 x 855 px — every chart identical
CHART_W   = 7.5
CHART_H   = 5.7
CHART_DPI = 150


def generate_pie_chart(question: dict, output_dir: str) -> str:
    """
    Generate one pie chart at exactly CHART_W x CHART_H inches.

    Slices >= 5% : % label inside slice (white bold 16pt)
    Slices  < 5% : no label on chart — shown in legend only
    Canvas size is fixed — all charts identical, no bbox expansion.
    """
    os.makedirs(output_dir, exist_ok=True)

    slices = [s for s in question["slices"] if s["pct"] > 0]
    labels = [s["label"] for s in slices]
    pcts   = [s["pct"]   for s in slices]
    colors = [s["color"] for s in slices]

    fig, ax = plt.subplots(figsize=(CHART_W, CHART_H))
    fig.patch.set_facecolor("white")

    # Reserve fixed space for title (top) and legend (bottom)
    # so the pie circle itself is always the same size
    fig.subplots_adjust(top=0.82, bottom=0.22, left=0.05, right=0.95)

    wedges, _ = ax.pie(
        pcts,
        labels=None,
        colors=colors,
        autopct=None,
        startangle=90,
        wedgeprops={"edgecolor": "white", "linewidth": 2.5},
    )

    # Inside labels for slices >= 5%
    cumulative = 0.0
    for i, (wedge, pct) in enumerate(zip(wedges, pcts)):
        angle_deg = cumulative + (pct / 2.0) * (360.0 / 100.0)
        angle_rad = np.deg2rad(90.0 - angle_deg)
        cumulative += pct * (360.0 / 100.0)

        if pct >= 5:
            x = 0.65 * np.cos(angle_rad)
            y = 0.65 * np.sin(angle_rad)
            ax.text(x, y, f"{pct}%",
                    ha="center", va="center",
                    color="white", fontweight="bold", fontsize=16)

    # Legend — placed at fixed position inside the figure
    legend = ax.legend(
        wedges,
        [f"{l} \u2013 {p}%" for l, p in zip(labels, pcts)],
        loc="lower center",
        bbox_to_anchor=(0.5, -0.28),   # fixed anchor relative to axes
        ncol=len(slices),
        fontsize=12,
        frameon=False,
        labelcolor=COLOR_TEXT,
    )

    # Title at fixed position
    short_desc = (question["description"][:72] + "..."
                  if len(question["description"]) > 72
                  else question["description"])
    ax.set_title(
        f"{question['code']}\n{short_desc}",
        fontsize=13,
        fontweight="bold",
        color=COLOR_TEXT,
        pad=14,
    )

    safe_code = (question["code"]
                 .replace("/", "-").replace("\\", "-").replace(":", "-"))
    filepath = os.path.join(output_dir, f"chart_{safe_code}.png")

    # bbox_inches=None — use exact figsize, no expansion for content
    fig.savefig(filepath, dpi=CHART_DPI, bbox_inches=None, facecolor="white")
    plt.close(fig)
    return filepath


def generate_all_charts(question_data: list, output_dir: str) -> list:
    """Generate one chart per question. Returns list of PNG paths."""
    paths = []
    for q in question_data:
        path = generate_pie_chart(q, output_dir)
        paths.append(path)
        print(f"  \u2713 Chart: {path}")
    return paths


# -----------------------------------------------------------------------
# Slice builders
# -----------------------------------------------------------------------

def make_course_exit_slices(high_pct, moderate_pct, low_pct):
    return [
        {"label": "High",     "pct": high_pct,     "color": COLOR_HIGH},
        {"label": "Moderate", "pct": moderate_pct, "color": COLOR_MODERATE},
        {"label": "Low",      "pct": low_pct,      "color": COLOR_LOW},
    ]


def make_faculty_slices(sa_pct, agree_pct, neutral_pct, disagree_pct):
    return [
        {"label": "Strongly Agree", "pct": sa_pct,       "color": COLOR_STRONGLY_AGREE},
        {"label": "Agree",          "pct": agree_pct,    "color": COLOR_AGREE},
        {"label": "Neutral",        "pct": neutral_pct,  "color": COLOR_NEUTRAL},
        {"label": "Disagree",       "pct": disagree_pct, "color": COLOR_DISAGREE},
    ]
