# =============================================================================
# charts.py
# PURPOSE: Generate one pie chart per CO showing High / Moderate / Low %
#          Save each chart as a PNG image file (to be embedded in Word doc)
# =============================================================================

import matplotlib.pyplot as plt
import matplotlib
import os

# Use non-interactive backend — no window popups, just saves to file
matplotlib.use("Agg")


# Colors for High / Moderate / Low — professional, NBA-report appropriate
COLORS = {
    "High":     "#2E75B6",   # Blue
    "Moderate": "#F4A100",   # Amber
    "Low":      "#C00000",   # Red
}


def generate_pie_chart(co: dict, output_dir: str) -> str:
    """
    Generate a pie chart for a single CO and save it as a PNG.
    
    Args:
        co: dict with keys code, description, high_pct, moderate_pct, low_pct
        output_dir: folder path where the PNG will be saved
    
    Returns:
        Full path to the saved PNG file
    """
    # Only include slices with value > 0 (avoid ugly zero-value slices)
    labels = []
    sizes = []
    colors = []

    for label in ["High", "Moderate", "Low"]:
        pct_key = f"{label.lower()}_pct"
        val = co[pct_key]
        if val > 0:
            labels.append(f"{label} ({val}%)")
            sizes.append(val)
            colors.append(COLORS[label])

    # Create figure
    fig, ax = plt.subplots(figsize=(5, 4))

    wedges, texts, autotexts = ax.pie(
        sizes,
        labels=labels,
        colors=colors,
        autopct="%1.0f%%",          # Show % inside each slice
        startangle=90,              # Start from top
        wedgeprops={"edgecolor": "white", "linewidth": 1.5},
        textprops={"fontsize": 10},
    )

    # Make percentage text inside slices bold and white
    for autotext in autotexts:
        autotext.set_color("white")
        autotext.set_fontweight("bold")
        autotext.set_fontsize(11)

    # Title = CO code only (description goes in the Word doc table)
    ax.set_title(co["code"], fontsize=12, fontweight="bold", pad=10)

    plt.tight_layout()

    # Save — filename uses the CO code (sanitize for filesystem)
    safe_code = co["code"].replace("/", "-").replace("\\", "-").replace(":", "-")
    filename = f"chart_{safe_code}.png"
    filepath = os.path.join(output_dir, filename)

    plt.savefig(filepath, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)   # Free memory — important when generating many charts

    return filepath


def generate_all_charts(co_data: list[dict], output_dir: str) -> list[str]:
    """
    Generate pie charts for all COs.
    
    Args:
        co_data: list of CO dicts (from parser.py)
        output_dir: folder to save PNG files
    
    Returns:
        List of file paths to generated PNG files, in same order as co_data
    """
    os.makedirs(output_dir, exist_ok=True)
    chart_paths = []

    for co in co_data:
        path = generate_pie_chart(co, output_dir)
        chart_paths.append(path)
        print(f"  ✓ Chart generated: {path}")

    return chart_paths
