import json
import shutil
from pathlib import Path


PROJECT_ROOT = Path(r"E:\UCL Final Essay")
NOTEBOOK_PATH = PROJECT_ROOT / "p8_uncertainty_framework" / "notebooks" / "P8_4_sector_linkage_deepening_local_reproducible.ipynb"
DOWNLOADS_PATH = Path(r"C:\Users\888\Downloads\P8_4_sector_linkage_deepening_local_reproducible.ipynb")


def md(source: str) -> dict:
    return {
        "cell_type": "markdown",
        "metadata": {},
        "source": source.strip().splitlines(keepends=True),
    }


def code(source: str) -> dict:
    return {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": source.strip().splitlines(keepends=True),
    }


cells = [
    md(
        r"""
# P8-4 Sector Linkage Deepening

This notebook deepens the P8 uncertainty framework by making the relationship between residual-emission sectors more explicit and more quantitative.

It responds to Neil's feedback that the dissertation should not only treat sectors independently. It should also ask whether some sectors move together, especially where clean electricity, electrification, hydrogen/CCUS, demand behaviour, land-use constraints, international technology costs, or removals accounting create shared constraints.

**Input data:** existing P8-2 outputs already generated from the DESNZ/CCC-aligned local workflow.

**Outputs produced by this notebook:**

- `p8_4_sector_linkage_deepening_metrics.csv`
- `p8_4_sector_dependency_similarity_matrix.csv`
- `p8_4_high_similarity_sector_pairs.csv`
- `p8_4_sector_typology_summary.csv`
- `p8_4_quality_checks.csv`
- `p8_4_results_summary.txt`
- `p8_4_linkage_score_vs_residual.jpg`
- `p8_4_sector_dependency_similarity_heatmap.jpg`
- `p8_4_linkage_weighted_residual_ranking.jpg`

**Interpretation note:** this is a diagnostic linkage framework, not a causal econometric model. The scores are designed to make the dissertation's sector-linkage discussion more transparent and defensible.
"""
    ),
    code(
        r"""
from pathlib import Path
import math
import textwrap
import warnings

import numpy as np
import pandas as pd
from PIL import Image as PILImage, ImageDraw, ImageFont

try:
    from IPython.display import display, Image as IPImage
except Exception:
    IPImage = None

    def display(obj):
        print(obj)


warnings.filterwarnings("ignore", category=UserWarning)


MANUAL_PROJECT_ROOT = Path(r"E:\UCL Final Essay")


def find_project_root():
    candidates = [
        MANUAL_PROJECT_ROOT,
        Path.cwd(),
        Path.home() / "Downloads",
        Path(r"E:\UCL Final Essay"),
    ]
    seen = set()
    required = Path("p8_uncertainty_framework") / "tables" / "p8_2_sector_linkage_driver_matrix.csv"
    for base in candidates:
        try:
            base = base.resolve()
        except Exception:
            continue
        for candidate in [base] + list(base.parents):
            if str(candidate).lower() in seen:
                continue
            seen.add(str(candidate).lower())
            if (candidate / required).exists():
                return candidate
    raise FileNotFoundError(
        "Could not find the project root containing p8_uncertainty_framework/tables. "
        "If needed, edit MANUAL_PROJECT_ROOT in the first code cell."
    )


PROJECT_ROOT = find_project_root()
P8_ROOT = PROJECT_ROOT / "p8_uncertainty_framework"
TABLE_DIR = P8_ROOT / "tables"
FIG_DIR = P8_ROOT / "figures"
TABLE_DIR.mkdir(parents=True, exist_ok=True)
FIG_DIR.mkdir(parents=True, exist_ok=True)

print("Project root:", PROJECT_ROOT)
print("P8 output root:", P8_ROOT)
"""
    ),
    md(
        r"""
## 1. Load Existing P8 Inputs

The analysis uses the local P8-2 sector tables:

- sector residual uncertainty in 2050;
- sector rate-anchor metrics;
- sector linkage driver scores.

No additional external dataset is introduced in this notebook.
"""
    ),
    code(
        r"""
linkage_path = TABLE_DIR / "p8_2_sector_linkage_driver_matrix.csv"
uncertainty_path = TABLE_DIR / "p8_2_sector_uncertainty_band_2050.csv"
rate_path = TABLE_DIR / "p8_2_sector_rate_anchor_metrics.csv"

for path in [linkage_path, uncertainty_path, rate_path]:
    assert path.exists(), f"Missing required input: {path}"

linkage = pd.read_csv(linkage_path)
uncertainty = pd.read_csv(uncertainty_path)
rate = pd.read_csv(rate_path)

driver_cols = [
    "clean_power_dependency",
    "electrification_infrastructure",
    "hydrogen_ccus_dependency",
    "demand_behaviour",
    "land_bio_nonco2",
    "international_fuel_technology",
    "removals_credits_accounting",
]

expected_cols = ["tes_sector", "rank_2050_residual"] + driver_cols
missing_cols = [c for c in expected_cols if c not in linkage.columns]
assert not missing_cols, f"Missing columns in linkage table: {missing_cols}"

for col in driver_cols:
    linkage[col] = pd.to_numeric(linkage[col], errors="coerce")

display(linkage)
display(uncertainty[["tes_sector", "desnz_central_2050_MtCO2e", "ccc7_2050_MtCO2e", "central_gap_vs_ccc7_2050_MtCO2e"]])
"""
    ),
    md(
        r"""
## 2. Build Sector Linkage Metrics

The aim is to distinguish three ideas:

1. **System linkage score:** how many shared transition drivers a sector depends on.
2. **Electrification coupling:** how strongly the sector depends on clean power and end-use electrification infrastructure.
3. **Linkage-weighted residual index:** whether a sector is both large in 2050 residual emissions and highly connected to other transition conditions.

This helps avoid treating each residual sector as an isolated accounting line.
"""
    ),
    code(
        r"""
metrics = linkage.copy()
metrics["system_linkage_score"] = metrics[driver_cols].sum(axis=1)
metrics["electrification_coupling_score"] = (
    metrics["clean_power_dependency"] + metrics["electrification_infrastructure"]
)
metrics["hard_to_abate_residual_score"] = (
    metrics["hydrogen_ccus_dependency"]
    + metrics["land_bio_nonco2"]
    + metrics["removals_credits_accounting"]
)
metrics["external_exposure_score"] = (
    metrics["international_fuel_technology"] + metrics["removals_credits_accounting"]
)

merge_cols_uncertainty = [
    "tes_sector",
    "desnz_central_2050_MtCO2e",
    "delayed_delivery_2050_MtCO2e",
    "accelerated_historical_anchor_2050_MtCO2e",
    "ccc7_2050_MtCO2e",
    "central_gap_vs_ccc7_2050_MtCO2e",
    "alignment_type",
]
merge_cols_rate = [
    "tes_sector",
    "2023",
    "2030",
    "2035",
    "2050",
    "projected_rate_2023_2050_MtCO2e_per_year",
    "best_positive_historical_rate_MtCO2e_per_year",
    "projected_vs_best_historical_rate_ratio",
    "rate_anchor_interpretation",
    "share_of_2050_inc_IAS_total_pct",
]

metrics = metrics.merge(uncertainty[merge_cols_uncertainty], on="tes_sector", how="left")
metrics = metrics.merge(rate[merge_cols_rate], on="tes_sector", how="left")

if metrics["share_of_2050_inc_IAS_total_pct"].isna().any():
    total_2050 = metrics["desnz_central_2050_MtCO2e"].sum()
    metrics["share_of_2050_inc_IAS_total_pct"] = (
        metrics["desnz_central_2050_MtCO2e"] / total_2050 * 100
    )

metrics["linkage_weighted_residual_index"] = (
    metrics["desnz_central_2050_MtCO2e"] * metrics["system_linkage_score"]
)
metrics["linkage_weighted_residual_share_pct"] = (
    metrics["linkage_weighted_residual_index"]
    / metrics["linkage_weighted_residual_index"].sum()
    * 100
)


def classify_typology(row):
    if row["electrification_coupling_score"] >= 3:
        return "Coupled electrification block"
    if row["hydrogen_ccus_dependency"] >= 2:
        return "Industrial hydrogen-CCUS block"
    if row["land_bio_nonco2"] >= 2:
        return "Land and non-CO2 residual block"
    if row["international_fuel_technology"] >= 2:
        return "International fuel and technology block"
    return "Lower-coupling residual block"


metrics["linkage_typology"] = metrics.apply(classify_typology, axis=1)
metrics["linkage_priority_rank"] = (
    metrics["linkage_weighted_residual_index"].rank(ascending=False, method="first").astype(int)
)

metrics = metrics.sort_values(["linkage_priority_rank", "rank_2050_residual"]).reset_index(drop=True)

output_cols = [
    "tes_sector",
    "rank_2050_residual",
    "linkage_priority_rank",
    "linkage_typology",
    "system_linkage_score",
    "electrification_coupling_score",
    "hard_to_abate_residual_score",
    "external_exposure_score",
    "desnz_central_2050_MtCO2e",
    "share_of_2050_inc_IAS_total_pct",
    "ccc7_2050_MtCO2e",
    "central_gap_vs_ccc7_2050_MtCO2e",
    "projected_rate_2023_2050_MtCO2e_per_year",
    "best_positive_historical_rate_MtCO2e_per_year",
    "projected_vs_best_historical_rate_ratio",
    "rate_anchor_interpretation",
    "linkage_weighted_residual_index",
    "linkage_weighted_residual_share_pct",
] + driver_cols

metrics_out = metrics[output_cols].copy()
metrics_path = TABLE_DIR / "p8_4_sector_linkage_deepening_metrics.csv"
metrics_out.to_csv(metrics_path, index=False)

print("Saved:", metrics_path)
display(metrics_out.round(3))
"""
    ),
    md(
        r"""
## 3. Quantify Whether Sectors Can Move Together

This step calculates a dependency-profile similarity matrix. It compares sectors by the seven driver scores above using cosine similarity.

High similarity does not prove statistical causality. It means two sectors depend on a similar mix of transition conditions, so they should be discussed together in the dissertation rather than treated as fully independent.
"""
    ),
    code(
        r"""
ordered = metrics.sort_values("rank_2050_residual").reset_index(drop=True)
vectors = ordered[driver_cols].to_numpy(dtype=float)
sectors = ordered["tes_sector"].tolist()


def cosine_similarity(a, b):
    denom = math.sqrt(float(np.dot(a, a))) * math.sqrt(float(np.dot(b, b)))
    if denom == 0:
        return 0.0
    return float(np.dot(a, b) / denom)


sim = np.zeros((len(sectors), len(sectors)))
for i in range(len(sectors)):
    for j in range(len(sectors)):
        sim[i, j] = cosine_similarity(vectors[i], vectors[j])

similarity = pd.DataFrame(sim, index=sectors, columns=sectors).round(3)
similarity_path = TABLE_DIR / "p8_4_sector_dependency_similarity_matrix.csv"
similarity.to_csv(similarity_path)

pairs = []
for i in range(len(sectors)):
    for j in range(i + 1, len(sectors)):
        pairs.append(
            {
                "sector_a": sectors[i],
                "sector_b": sectors[j],
                "dependency_profile_similarity": sim[i, j],
                "sector_a_2050_residual_MtCO2e": ordered.loc[i, "desnz_central_2050_MtCO2e"],
                "sector_b_2050_residual_MtCO2e": ordered.loc[j, "desnz_central_2050_MtCO2e"],
                "combined_2050_residual_MtCO2e": ordered.loc[i, "desnz_central_2050_MtCO2e"]
                + ordered.loc[j, "desnz_central_2050_MtCO2e"],
            }
        )

high_pairs = pd.DataFrame(pairs)
high_pairs = high_pairs.sort_values(
    ["dependency_profile_similarity", "combined_2050_residual_MtCO2e"],
    ascending=[False, False],
).reset_index(drop=True)
high_pairs_flagged = high_pairs[high_pairs["dependency_profile_similarity"] >= 0.85].copy()

high_pairs_path = TABLE_DIR / "p8_4_high_similarity_sector_pairs.csv"
high_pairs_flagged.to_csv(high_pairs_path, index=False)

print("Saved:", similarity_path)
print("Saved:", high_pairs_path)
display(similarity)
display(high_pairs_flagged.round(3))
"""
    ),
    md(
        r"""
## 4. Summarise Sector Typologies

This table converts the sector-level results into dissertation-ready claims. It shows which transition blocks explain most of the linkage-weighted residual risk.
"""
    ),
    code(
        r"""
typology = (
    metrics.groupby("linkage_typology", as_index=False)
    .agg(
        sectors=("tes_sector", lambda x: "; ".join(x)),
        sector_count=("tes_sector", "count"),
        central_2050_residual_MtCO2e=("desnz_central_2050_MtCO2e", "sum"),
        central_gap_vs_ccc7_2050_MtCO2e=("central_gap_vs_ccc7_2050_MtCO2e", "sum"),
        linkage_weighted_residual_index=("linkage_weighted_residual_index", "sum"),
    )
)
typology["share_of_linkage_weighted_residual_pct"] = (
    typology["linkage_weighted_residual_index"]
    / typology["linkage_weighted_residual_index"].sum()
    * 100
)
typology = typology.sort_values(
    "linkage_weighted_residual_index", ascending=False
).reset_index(drop=True)

typology_path = TABLE_DIR / "p8_4_sector_typology_summary.csv"
typology.to_csv(typology_path, index=False)

print("Saved:", typology_path)
display(typology.round(3))
"""
    ),
    md(
        r"""
## 5. Generate Figures

The figures are saved as JPG files to avoid PNG rendering issues on some local machines.
"""
    ),
    code(
        r"""
FIG_WIDE = 1700
FIG_TALL = 1200
BG = (255, 255, 255)
INK = (31, 31, 31)
MUTED = (92, 92, 92)
GRID = (218, 218, 218)

TYPOLOGY_COLORS = {
    "Coupled electrification block": "#2E7D32",
    "Industrial hydrogen-CCUS block": "#1F77B4",
    "Land and non-CO2 residual block": "#8E6C00",
    "International fuel and technology block": "#B23A48",
    "Lower-coupling residual block": "#666666",
}

SECTOR_SHORT = {
    "Buildings and product uses": "Buildings/products",
    "Domestic Transport": "Domestic transport",
    "Industry": "Industry",
    "Electricity supply": "Electricity",
    "Agriculture": "Agriculture",
    "IAS": "IAS",
    "Fuel supply": "Fuel supply",
    "Waste": "Waste",
    "LULUCF": "LULUCF",
}


def hex_to_rgb(value):
    value = value.lstrip("#")
    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))


def get_font(size, bold=False):
    candidates = []
    if bold:
        candidates += [
            Path(r"C:\Windows\Fonts\arialbd.ttf"),
            Path(r"C:\Windows\Fonts\calibrib.ttf"),
            Path(r"C:\Windows\Fonts\segoeuib.ttf"),
        ]
    candidates += [
        Path(r"C:\Windows\Fonts\arial.ttf"),
        Path(r"C:\Windows\Fonts\calibri.ttf"),
        Path(r"C:\Windows\Fonts\segoeui.ttf"),
    ]
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


FONT_TITLE = get_font(42, bold=True)
FONT_SUBTITLE = get_font(24)
FONT_LABEL = get_font(22)
FONT_SMALL = get_font(18)
FONT_TINY = get_font(15)
FONT_AXIS = get_font(19)
FONT_BOLD = get_font(22, bold=True)


def draw_text(draw, xy, text, font, fill=INK, anchor=None):
    draw.text(xy, str(text), font=font, fill=fill, anchor=anchor)


def text_size(draw, text, font):
    box = draw.textbbox((0, 0), str(text), font=font)
    return box[2] - box[0], box[3] - box[1]


def wrap_label(text, width=16):
    return "\n".join(textwrap.wrap(str(text), width=width, break_long_words=False))


def lighten(rgb, amount=0.85):
    return tuple(int(c + (255 - c) * amount) for c in rgb)


def save_jpg(image, path):
    image = image.convert("RGB")
    image.save(path, quality=96, optimize=True)
    return path


def display_jpg(path):
    if IPImage is not None:
        display(IPImage(filename=str(path)))
    else:
        print(path)
"""
    ),
    code(
        r"""
fig1_path = FIG_DIR / "p8_4_linkage_score_vs_residual.jpg"

plot_df = metrics.sort_values("rank_2050_residual").copy()
x_min = 0
x_max = max(10, int(plot_df["system_linkage_score"].max()) + 1)
y_min = 0
y_max = max(100, math.ceil(plot_df["desnz_central_2050_MtCO2e"].max() / 10) * 10 + 20)

img = PILImage.new("RGB", (FIG_WIDE, FIG_TALL), BG)
draw = ImageDraw.Draw(img)

left, right, top, bottom = 185, 90, 170, 210
plot_w = FIG_WIDE - left - right
plot_h = FIG_TALL - top - bottom

draw_text(draw, (70, 55), "Sector linkage score vs 2050 residual emissions", FONT_TITLE)
draw_text(
    draw,
    (70, 112),
    "Higher scores indicate stronger dependence on shared transition drivers; bubble size reflects each sector's 2050 residual share.",
    FONT_SUBTITLE,
    fill=MUTED,
)

for tick in range(0, x_max + 1, 2):
    x = left + (tick - x_min) / (x_max - x_min) * plot_w
    draw.line((x, top, x, top + plot_h), fill=GRID, width=1)
    draw_text(draw, (x, top + plot_h + 20), str(tick), FONT_AXIS, fill=MUTED, anchor="ma")

for tick in range(0, y_max + 1, 20):
    y = top + plot_h - (tick - y_min) / (y_max - y_min) * plot_h
    draw.line((left, y, left + plot_w, y), fill=GRID, width=1)
    draw_text(draw, (left - 18, y), str(tick), FONT_AXIS, fill=MUTED, anchor="rm")

draw.line((left, top, left, top + plot_h), fill=INK, width=2)
draw.line((left, top + plot_h, left + plot_w, top + plot_h), fill=INK, width=2)
draw_text(draw, (left + plot_w / 2, FIG_TALL - 105), "System linkage score", FONT_BOLD, anchor="ma")
draw_text(draw, (40, top + plot_h / 2), "DESNZ central 2050 residual emissions (MtCO2e)", FONT_BOLD, anchor="mm")

label_offsets = {
    "Buildings and product uses": (18, -46),
    "Domestic Transport": (20, -8),
    "Industry": (18, -38),
    "Electricity supply": (18, 20),
    "Agriculture": (18, -35),
    "IAS": (18, 20),
    "Fuel supply": (18, 20),
    "Waste": (18, -35),
    "LULUCF": (18, 16),
}

for _, row in plot_df.iterrows():
    x_val = row["system_linkage_score"]
    y_val = row["desnz_central_2050_MtCO2e"]
    x = left + (x_val - x_min) / (x_max - x_min) * plot_w
    y = top + plot_h - (y_val - y_min) / (y_max - y_min) * plot_h
    radius = 15 + max(3, row["share_of_2050_inc_IAS_total_pct"]) * 1.1
    color = hex_to_rgb(TYPOLOGY_COLORS[row["linkage_typology"]])
    draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=lighten(color, 0.35), outline=color, width=4)
    dx, dy = label_offsets.get(row["tes_sector"], (16, -28))
    draw_text(draw, (x + dx, y + dy), SECTOR_SHORT.get(row["tes_sector"], row["tes_sector"]), FONT_LABEL, fill=INK)

legend_x, legend_y = 70, FIG_TALL - 155
for i, (label, color_hex) in enumerate(TYPOLOGY_COLORS.items()):
    y = legend_y + i * 28
    color = hex_to_rgb(color_hex)
    draw.rectangle((legend_x, y - 10, legend_x + 20, y + 10), fill=color)
    draw_text(draw, (legend_x + 32, y), label, FONT_SMALL, fill=MUTED, anchor="lm")

draw_text(
    draw,
    (FIG_WIDE - 70, FIG_TALL - 42),
    "Source: Author calculations from P8-2 DESNZ/CCC-aligned sector tables. Diagnostic scores, not causal estimates.",
    FONT_TINY,
    fill=MUTED,
    anchor="ra",
)

save_jpg(img, fig1_path)
print("Saved:", fig1_path)
display_jpg(fig1_path)
"""
    ),
    code(
        r"""
fig2_path = FIG_DIR / "p8_4_sector_dependency_similarity_heatmap.jpg"

heat_df = similarity.copy()
n = len(heat_df)
cell = 112
left = 340
top = 285
right = 80
bottom = 120
width = left + n * cell + right
height = top + n * cell + bottom

img = PILImage.new("RGB", (width, height), BG)
draw = ImageDraw.Draw(img)
draw_text(draw, (70, 55), "Dependency-profile similarity between residual sectors", FONT_TITLE)
draw_text(
    draw,
    (70, 112),
    "Cosine similarity based on seven transition-driver scores. Values near 1 indicate sectors likely to move together under shared constraints.",
    FONT_SUBTITLE,
    fill=MUTED,
)

low = np.array([247, 251, 255])
high = np.array([8, 48, 107])

short_labels = [SECTOR_SHORT.get(s, s) for s in heat_df.index]

for i, label in enumerate(short_labels):
    y = top + i * cell + cell / 2
    draw_text(draw, (left - 18, y), label, FONT_AXIS, fill=INK, anchor="rm")

for j, label in enumerate(short_labels):
    x = left + j * cell + cell / 2
    wrapped = wrap_label(label, width=12)
    lines = wrapped.split("\n")
    for k, line in enumerate(lines):
        draw_text(draw, (x, top - 90 + k * 21), line, FONT_TINY, fill=INK, anchor="ma")

for i in range(n):
    for j in range(n):
        value = float(heat_df.iloc[i, j])
        rgb = tuple(np.round(low + (high - low) * value).astype(int))
        x0 = left + j * cell
        y0 = top + i * cell
        draw.rectangle((x0, y0, x0 + cell, y0 + cell), fill=rgb, outline=(245, 245, 245), width=2)
        txt_color = (255, 255, 255) if value >= 0.62 else INK
        draw_text(draw, (x0 + cell / 2, y0 + cell / 2), f"{value:.2f}", FONT_AXIS, fill=txt_color, anchor="mm")

draw.rectangle((left, top, left + n * cell, top + n * cell), outline=INK, width=2)
draw_text(
    draw,
    (70, height - 48),
    "Interpretation: this matrix identifies shared dependency structures, not observed causal correlations.",
    FONT_SMALL,
    fill=MUTED,
)

save_jpg(img, fig2_path)
print("Saved:", fig2_path)
display_jpg(fig2_path)
"""
    ),
    code(
        r"""
fig3_path = FIG_DIR / "p8_4_linkage_weighted_residual_ranking.jpg"

bar_df = metrics.sort_values("linkage_weighted_residual_index", ascending=True).reset_index(drop=True)
img = PILImage.new("RGB", (FIG_WIDE, FIG_TALL), BG)
draw = ImageDraw.Draw(img)

draw_text(draw, (70, 55), "Linkage-weighted residual risk by sector", FONT_TITLE)
draw_text(
    draw,
    (70, 112),
    "Index = DESNZ central 2050 residual emissions multiplied by the sector's system linkage score.",
    FONT_SUBTITLE,
    fill=MUTED,
)

left, right, top, bottom = 355, 160, 175, 170
plot_w = FIG_WIDE - left - right
plot_h = FIG_TALL - top - bottom
row_h = plot_h / len(bar_df)
max_index = float(bar_df["linkage_weighted_residual_index"].max())

for tick in np.linspace(0, max_index, 6):
    x = left + tick / max_index * plot_w
    draw.line((x, top, x, top + plot_h), fill=GRID, width=1)
    draw_text(draw, (x, top + plot_h + 20), f"{tick:.0f}", FONT_AXIS, fill=MUTED, anchor="ma")

for i, row in bar_df.iterrows():
    y = top + i * row_h + row_h * 0.5
    bar_len = row["linkage_weighted_residual_index"] / max_index * plot_w
    color = hex_to_rgb(TYPOLOGY_COLORS[row["linkage_typology"]])
    draw_text(draw, (left - 20, y), SECTOR_SHORT.get(row["tes_sector"], row["tes_sector"]), FONT_AXIS, fill=INK, anchor="rm")
    draw.rounded_rectangle(
        (left, y - row_h * 0.28, left + bar_len, y + row_h * 0.28),
        radius=8,
        fill=lighten(color, 0.25),
        outline=color,
        width=2,
    )
    annotation = (
        f"index {row['linkage_weighted_residual_index']:.0f}; "
        f"residual {row['desnz_central_2050_MtCO2e']:.1f}; "
        f"score {row['system_linkage_score']:.0f}"
    )
    draw_text(draw, (left + bar_len + 16, y), annotation, FONT_SMALL, fill=MUTED, anchor="lm")

draw.line((left, top + plot_h, left + plot_w, top + plot_h), fill=INK, width=2)
draw_text(draw, (left + plot_w / 2, FIG_TALL - 80), "Linkage-weighted residual index", FONT_BOLD, anchor="ma")

draw_text(
    draw,
    (FIG_WIDE - 70, FIG_TALL - 42),
    "Source: Author calculations from P8-2 sector residual and linkage-driver tables.",
    FONT_TINY,
    fill=MUTED,
    anchor="ra",
)

save_jpg(img, fig3_path)
print("Saved:", fig3_path)
display_jpg(fig3_path)
"""
    ),
    md(
        r"""
## 6. Results Summary and Quality Checks

The checks below are deliberately simple. They are meant to catch local path problems, missing sectors, missing values, and malformed similarity output before the results are used in the dissertation draft.
"""
    ),
    code(
        r"""
top_sector = metrics.sort_values("linkage_weighted_residual_index", ascending=False).iloc[0]
top_three = metrics.sort_values("linkage_weighted_residual_index", ascending=False).head(3)
top_pair = high_pairs.iloc[0]
electrification_block = typology[typology["linkage_typology"] == "Coupled electrification block"]

summary_lines = [
    "P8-4 Sector linkage deepening summary",
    "",
    "Main interpretation:",
    (
        f"1. The largest linkage-weighted residual priority is {top_sector['tes_sector']}, "
        f"with a 2050 residual of {top_sector['desnz_central_2050_MtCO2e']:.1f} MtCO2e, "
        f"a system linkage score of {top_sector['system_linkage_score']:.0f}, and "
        f"{top_sector['linkage_weighted_residual_share_pct']:.1f}% of the linkage-weighted residual index."
    ),
    (
        "2. The top three linkage-weighted priorities are "
        + ", ".join(top_three["tes_sector"].tolist())
        + ". This supports a dissertation discussion that gives more attention to cross-sector constraints, not only sector ranking."
    ),
    (
        f"3. The strongest dependency-profile pair is {top_pair['sector_a']} and {top_pair['sector_b']} "
        f"(similarity {top_pair['dependency_profile_similarity']:.2f}). This is useful evidence for Neil's question about whether some sectors move together."
    ),
]

if len(electrification_block) > 0:
    block = electrification_block.iloc[0]
    summary_lines.append(
        f"4. The coupled electrification block accounts for {block['central_2050_residual_MtCO2e']:.1f} MtCO2e "
        f"of central 2050 residual emissions and {block['share_of_linkage_weighted_residual_pct']:.1f}% "
        "of the linkage-weighted residual index."
    )

summary_lines += [
    "",
    "How to use in the dissertation:",
    "- Use the scatter plot to show that high residual emissions and high transition coupling are not the same concept.",
    "- Use the similarity heatmap to support the claim that buildings, domestic transport and electricity should be discussed as linked electrification systems.",
    "- Use the linkage-weighted residual ranking to prioritise which residual sectors need deeper policy discussion.",
    "",
    "Caveat:",
    "The analysis is based on expert-coded dependency scores from the existing P8-2 framework. It is a transparent diagnostic framework, not a causal estimate.",
]

summary_path = TABLE_DIR / "p8_4_results_summary.txt"
summary_path.write_text("\n".join(summary_lines), encoding="utf-8")

checks = []
checks.append({"check": "linkage_input_exists", "status": linkage_path.exists(), "detail": str(linkage_path)})
checks.append({"check": "uncertainty_input_exists", "status": uncertainty_path.exists(), "detail": str(uncertainty_path)})
checks.append({"check": "rate_input_exists", "status": rate_path.exists(), "detail": str(rate_path)})
checks.append({"check": "sector_count", "status": len(metrics) == len(linkage), "detail": f"{len(metrics)} rows"})
checks.append({"check": "unique_sectors", "status": metrics["tes_sector"].is_unique, "detail": f"{metrics['tes_sector'].nunique()} unique sectors"})
checks.append({"check": "driver_scores_between_0_and_2", "status": bool(((linkage[driver_cols] >= 0) & (linkage[driver_cols] <= 2)).all().all()), "detail": "0=low/not material; 1=moderate; 2=high"})
checks.append({"check": "no_missing_2050_residual", "status": not metrics["desnz_central_2050_MtCO2e"].isna().any(), "detail": "DESNZ central 2050 residual present"})
checks.append({"check": "similarity_matrix_symmetric", "status": bool(np.allclose(sim, sim.T, atol=1e-9)), "detail": "cosine similarity"})
checks.append({"check": "similarity_diagonal_one", "status": bool(np.allclose(np.diag(sim), 1.0, atol=1e-9)), "detail": "self-similarity"})
for fig_path in [fig1_path, fig2_path, fig3_path]:
    checks.append({"check": f"figure_created_{fig_path.name}", "status": fig_path.exists() and fig_path.stat().st_size > 10_000, "detail": str(fig_path)})

quality = pd.DataFrame(checks)
quality_path = TABLE_DIR / "p8_4_quality_checks.csv"
quality.to_csv(quality_path, index=False)

print("Saved:", summary_path)
print("Saved:", quality_path)
display(quality)
print("\n".join(summary_lines))

assert quality["status"].all(), "At least one quality check failed. Please inspect p8_4_quality_checks.csv."
"""
    ),
]


nb = {
    "cells": cells,
    "metadata": {
        "kernelspec": {
            "display_name": "Python 3",
            "language": "python",
            "name": "python3",
        },
        "language_info": {
            "codemirror_mode": {"name": "ipython", "version": 3},
            "file_extension": ".py",
            "mimetype": "text/x-python",
            "name": "python",
            "nbconvert_exporter": "python",
            "pygments_lexer": "ipython3",
            "version": "3",
        },
    },
    "nbformat": 4,
    "nbformat_minor": 5,
}

NOTEBOOK_PATH.parent.mkdir(parents=True, exist_ok=True)
NOTEBOOK_PATH.write_text(json.dumps(nb, ensure_ascii=False, indent=2), encoding="utf-8")
DOWNLOADS_PATH.parent.mkdir(parents=True, exist_ok=True)
shutil.copy2(NOTEBOOK_PATH, DOWNLOADS_PATH)

print(f"Notebook written to: {NOTEBOOK_PATH}")
print(f"Notebook copied to: {DOWNLOADS_PATH}")
