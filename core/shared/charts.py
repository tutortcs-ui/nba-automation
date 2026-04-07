# =============================================================================
# core/shared/charts.py
# PURPOSE: Generate pie charts — white background, fixed identical pixel size.
#
# SIZE CALCULATION:
#   Canvas: 7.5 x 5.7 inches at 150 dpi = 1125 x 855 px
#   Inserted in Word at 11cm wide:
#     display height = (855/150)*2.54 * (11/19.05) = 8.35cm per chart
#     two charts per page = 16.7cm → fits A4 usable height (24.62cm) comfortably
# =============================================================================

import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


COLOR_HIGH           = "#2864c8"
COLOR_MODERATE       = "#f08c00"
COLOR_LOW            = "#dc2800"
COLOR_STRONGLY_AGREE = "#2864c8"
COLOR_AGREE          = "#4caf82"
COLOR_NEUTRAL        = "#f08c00"
COLOR_DISAGREE       = "#dc2800"
COLOR_TEXT           = "#1a1a2e"

CHART_W   = 7.5    # inches — 1125px at 150dpi
CHART_H   = 5.7    # inches —  855px at 150dpi
CHART_DPI = 150


def generate_pie_chart(question: dict, output_dir: str) -> str:
    """
    Fixed canvas 1125x855px, white background, autopct for correct labels.
    Insert in Word at 11cm wide → 8.35cm tall → two fit per A4 page.
    """
    os.makedirs(output_dir, exist_ok=True)

    slices = [s for s in question["slices"] if s["pct"] > 0]
    labels = [s["label"] for s in slices]
    pcts   = [s["pct"]   for s in slices]
    colors = [s["color"] for s in slices]

    fig, ax = plt.subplots(figsize=(CHART_W, CHART_H))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    # Fixed margins — title top 18%, legend bottom 22%, pie fills middle
    fig.subplots_adjust(top=0.82, bottom=0.22, left=0.05, right=0.95)

    def autopct_fn(pct):
        return f"{pct:.0f}%" if pct >= 5 else ""

    wedges, _, autotexts = ax.pie(
        pcts,
        labels=None,
        colors=colors,
        autopct=autopct_fn,
        pctdistance=0.65,
        startangle=90,
        wedgeprops={"edgecolor": "white", "linewidth": 2.5},
    )

    for autotext in autotexts:
        autotext.set_color("white")
        autotext.set_fontweight("bold")
        autotext.set_fontsize(16)

    ax.legend(
        wedges,
        [f"{l} \u2013 {p}%" for l, p in zip(labels, pcts)],
        loc="lower center",
        bbox_to_anchor=(0.5, -0.28),
        ncol=len(slices),
        fontsize=12,
        frameon=False,
        labelcolor=COLOR_TEXT,
    )

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

    fig.savefig(filepath, dpi=CHART_DPI, bbox_inches=None,
                facecolor="white", edgecolor="none")
    plt.close(fig)
    return filepath


def generate_all_charts(question_data: list, output_dir: str) -> list:
    paths = []
    for q in question_data:
        path = generate_pie_chart(q, output_dir)
        paths.append(path)
        print(f"  \u2713 Chart: {path}")
    return paths


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
