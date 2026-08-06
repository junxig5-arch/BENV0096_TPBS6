import json
from pathlib import Path


PROJECT_ROOT = Path(r"E:\UCL Final Essay")
NOTEBOOK_PATH = PROJECT_ROOT / "p8_uncertainty_framework" / "notebooks" / "P8_3_2x2_matrix_near_miss_mini_scenarios_local_reproducible.ipynb"
DOWNLOADS_PATH = Path(r"C:\Users\888\Downloads\P8_3_2x2_matrix_near_miss_mini_scenarios_local_reproducible.ipynb")


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
# P8-3 2x2 Matrix, Near-Miss Outcomes and Mini-Scenarios

This notebook turns the P8 uncertainty framework into a set of transparent mini-scenarios.

**Purpose:** respond to Neil's 17 July suggestion that the dissertation should use the DESNZ, CCC and NESO evidence to generate a richer set of possible "states of the world", rather than only comparing a small number of named pathways.

**This is not a probability model.** It is a stylised sensitivity framework that combines:

- domestic policy delivery strength;
- external technology / fuel / supply-chain conditions;
- sectoral linkages from P8-2;
- negative emissions or credits stress tests.

**Outputs produced by this notebook:**

- `p8_3_2x2_scenario_matrix.csv`
- `p8_3_mini_scenario_population.csv`
- `p8_3_negative_emissions_credit_stress_test.csv`
- `p8_3_near_miss_summary.csv`
- `p8_3_scenario_assumption_table.csv`
- `p8_3_quadrant_outcomes.jpg`
- `p8_3_mini_scenario_residual_gaps.jpg`
- `p8_3_negative_emissions_credit_sensitivity.jpg`
- `p8_3_results_summary_bilingual.txt`
"""
    ),
    code(
        r"""
from pathlib import Path
import warnings

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

warnings.filterwarnings("ignore", category=UserWarning)

plt.rcParams["figure.dpi"] = 140
plt.rcParams["savefig.dpi"] = 220
plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["axes.spines.top"] = False
plt.rcParams["axes.spines.right"] = False


def find_project_root(start=None):
    candidates = []
    if start is not None:
        candidates.append(Path(start).resolve())
    try:
        candidates.append(Path.cwd().resolve())
    except Exception:
        pass
    candidates += [Path(r"E:\UCL Final Essay"), Path.home() / "Downloads"]
    checked = []
    for base in candidates:
        for candidate in [base] + list(base.parents):
            if candidate in checked:
                continue
            checked.append(candidate)
            if (candidate / "p8_uncertainty_framework" / "tables").exists() and (candidate / "p7_neso_uncertainty" / "tables").exists():
                return candidate
    raise FileNotFoundError("Could not find project root. Please run this notebook from inside the dissertation project folder.")


PROJECT_ROOT = find_project_root()
P8_TABLE_DIR = PROJECT_ROOT / "p8_uncertainty_framework" / "tables"
P7_TABLE_DIR = PROJECT_ROOT / "p7_neso_uncertainty" / "tables"
FIG_DIR = PROJECT_ROOT / "p8_uncertainty_framework" / "figures"

P8_TABLE_DIR.mkdir(parents=True, exist_ok=True)
FIG_DIR.mkdir(parents=True, exist_ok=True)

print("Project root:", PROJECT_ROOT)
print("P8 output tables:", P8_TABLE_DIR)
"""
    ),
    md(
        r"""
## 1. Load P8-1, P8-2 and P7 Anchors

P8-3 uses the results already produced locally:

- P8-1: historical delivered-rate benchmark;
- P8-2: sectoral uncertainty bands and linkage driver matrix;
- P7: NESO target timing and external-condition evidence.
"""
    ),
    code(
        r"""
required_files = {
    "p8_1_rates": P8_TABLE_DIR / "p8_1_historical_rate_windows.csv",
    "p8_2_sector_band": P8_TABLE_DIR / "p8_2_sector_uncertainty_band_2050.csv",
    "p8_2_driver_matrix": P8_TABLE_DIR / "p8_2_sector_linkage_driver_matrix.csv",
    "p8_2_linkage_metadata": P8_TABLE_DIR / "p8_2_sector_linkage_metadata.csv",
    "p7_target_year": P7_TABLE_DIR / "p7_target_achievement_year_estimates.csv",
}

for name, path in required_files.items():
    assert path.exists(), f"Missing {name}: {path}. Please run P8-1/P8-2/P7 notebooks first."

historical_rates = pd.read_csv(required_files["p8_1_rates"])
sector_band = pd.read_csv(required_files["p8_2_sector_band"])
driver_matrix = pd.read_csv(required_files["p8_2_driver_matrix"])
linkage_metadata = pd.read_csv(required_files["p8_2_linkage_metadata"])
p7_target = pd.read_csv(required_files["p7_target_year"])

display(sector_band.head())
display(driver_matrix.head())
display(p7_target)
"""
    ),
    md(
        r"""
## 2. Define the 2x2 Scenario Matrix

The two axes follow the structure agreed after Neil's feedback:

- **Domestic policy delivery:** high vs low.
- **External conditions:** supportive vs constrained.

Each quadrant gets a short interpretation and a post-2050 continuation rate used only for mechanical delay estimates.
"""
    ),
    code(
        r"""
falling_behind_rate = abs(float(
    p7_target.loc[p7_target["pathway"].eq("Falling Behind"), "avg_annual_change_2040_2050_MtCO2e_per_year"].iloc[0]
))
ccc7_late_rate = abs(float(
    p7_target.loc[p7_target["pathway"].eq("CCC7 Balanced Pathway"), "avg_annual_change_2040_2050_MtCO2e_per_year"].iloc[0]
))

matrix_rows = [
    {
        "quadrant_id": "U1",
        "scenario_name": "Aligned transition",
        "uk_policy_delivery": "High",
        "external_conditions": "Supportive",
        "domestic_delivery_range": "0.75-0.95",
        "external_condition_adjustment": "+0.10 to high-external-dependency sectors",
        "default_removals_credit_available_MtCO2e": 50,
        "post_2050_reduction_rate_MtCO2e_per_year_for_stress_test": ccc7_late_rate,
        "interpretation": "Strong UK delivery plus supportive technology, fuel and supply-chain conditions. Closest to target-consistent outcomes.",
    },
    {
        "quadrant_id": "U2",
        "scenario_name": "Domestic delivery under external constraints",
        "uk_policy_delivery": "High",
        "external_conditions": "Constrained",
        "domestic_delivery_range": "0.65-0.85",
        "external_condition_adjustment": "-0.10 to high-external-dependency sectors",
        "default_removals_credit_available_MtCO2e": 30,
        "post_2050_reduction_rate_MtCO2e_per_year_for_stress_test": (ccc7_late_rate + falling_behind_rate) / 2,
        "interpretation": "Domestic policies are credible, but costs, supply chains, fuels or international technology conditions slow some sectors.",
    },
    {
        "quadrant_id": "U3",
        "scenario_name": "Weak delivery despite supportive conditions",
        "uk_policy_delivery": "Low",
        "external_conditions": "Supportive",
        "domestic_delivery_range": "0.30-0.50",
        "external_condition_adjustment": "+0.10 to high-external-dependency sectors",
        "default_removals_credit_available_MtCO2e": 30,
        "post_2050_reduction_rate_MtCO2e_per_year_for_stress_test": falling_behind_rate,
        "interpretation": "Global technology and cost conditions help, but weak UK implementation leaves a large domestic delivery gap.",
    },
    {
        "quadrant_id": "U4",
        "scenario_name": "Delayed transition risk",
        "uk_policy_delivery": "Low",
        "external_conditions": "Constrained",
        "domestic_delivery_range": "0.15-0.35",
        "external_condition_adjustment": "-0.10 to high-external-dependency sectors",
        "default_removals_credit_available_MtCO2e": 10,
        "post_2050_reduction_rate_MtCO2e_per_year_for_stress_test": max(5.0, falling_behind_rate * 0.85),
        "interpretation": "Weak domestic delivery and constrained external conditions. Closest to delayed-transition risk.",
    },
]

scenario_matrix = pd.DataFrame(matrix_rows)
scenario_matrix.to_csv(P8_TABLE_DIR / "p8_3_2x2_scenario_matrix.csv", index=False)
display(scenario_matrix)
"""
    ),
    md(
        r"""
## 3. Generate Mini-Scenarios

Each quadrant is represented by three mini-scenarios. These mini-scenarios vary the domestic delivery score within the quadrant range.

Sector emissions are interpolated between:

- **Delayed delivery value** from P8-2; and
- **Accelerated / historical-rate anchor value** from P8-2.

External conditions affect sectors that are more exposed to international technology, fuels, hydrogen/CCUS or removals/credits.
"""
    ),
    code(
        r"""
driver_cols = [
    "hydrogen_ccus_dependency",
    "international_fuel_technology",
    "removals_credits_accounting",
]
driver_matrix["external_dependency_score"] = driver_matrix[driver_cols].mean(axis=1) / 2.0
driver_matrix["clean_power_score"] = driver_matrix["clean_power_dependency"] / 2.0

sector_inputs = sector_band.merge(
    driver_matrix[["tes_sector", "external_dependency_score", "clean_power_score"]],
    on="tes_sector",
    how="left",
)

assert sector_inputs[["external_dependency_score", "clean_power_score"]].notna().all().all(), "Missing driver scores for some sectors."

mini_specs = []
for _, q in scenario_matrix.iterrows():
    if q["uk_policy_delivery"] == "High" and q["external_conditions"] == "Supportive":
        delivery_scores = [0.75, 0.85, 0.95]
        labels = ["credible", "strong", "front-loaded"]
        external_shift = 0.10
    elif q["uk_policy_delivery"] == "High" and q["external_conditions"] == "Constrained":
        delivery_scores = [0.65, 0.75, 0.85]
        labels = ["credible-but-costly", "strong-but-constrained", "front-loaded-domestic"]
        external_shift = -0.10
    elif q["uk_policy_delivery"] == "Low" and q["external_conditions"] == "Supportive":
        delivery_scores = [0.30, 0.40, 0.50]
        labels = ["partial-delivery", "uneven-delivery", "late-catch-up"]
        external_shift = 0.10
    else:
        delivery_scores = [0.15, 0.25, 0.35]
        labels = ["delayed", "fragmented", "late-and-costly"]
        external_shift = -0.10

    for idx, (score, label) in enumerate(zip(delivery_scores, labels), start=1):
        mini_specs.append({
            "mini_scenario_id": f"{q['quadrant_id']}-{idx}",
            "quadrant_id": q["quadrant_id"],
            "scenario_name": q["scenario_name"],
            "mini_scenario_label": label,
            "uk_policy_delivery": q["uk_policy_delivery"],
            "external_conditions": q["external_conditions"],
            "domestic_delivery_score": score,
            "external_shift": external_shift,
            "default_removals_credit_available_MtCO2e": q["default_removals_credit_available_MtCO2e"],
            "post_2050_reduction_rate_MtCO2e_per_year_for_stress_test": q["post_2050_reduction_rate_MtCO2e_per_year_for_stress_test"],
        })

mini_specs = pd.DataFrame(mini_specs)


def sector_value(row, domestic_score, external_shift):
    accelerated = row["accelerated_historical_anchor_2050_MtCO2e"]
    delayed = row["delayed_delivery_2050_MtCO2e"]
    external_dependency = row["external_dependency_score"]
    clean_power_score = row["clean_power_score"]

    # Supportive external conditions most affect externally exposed sectors,
    # and also slightly reinforce clean-power-dependent sectors.
    score = domestic_score + external_shift * (0.75 * external_dependency + 0.25 * clean_power_score)
    score = float(np.clip(score, 0, 1))
    return delayed - score * (delayed - accelerated)


scenario_records = []
sector_records = []

for _, spec in mini_specs.iterrows():
    values = []
    for _, sector in sector_inputs.iterrows():
        value = sector_value(sector, spec["domestic_delivery_score"], spec["external_shift"])
        values.append(value)
        sector_records.append({
            **spec.to_dict(),
            "tes_sector": sector["tes_sector"],
            "sector_2050_MtCO2e_before_offsets": value,
            "desnz_central_2050_MtCO2e": sector["desnz_central_2050_MtCO2e"],
            "accelerated_anchor_2050_MtCO2e": sector["accelerated_historical_anchor_2050_MtCO2e"],
            "delayed_delivery_2050_MtCO2e": sector["delayed_delivery_2050_MtCO2e"],
            "external_dependency_score": sector["external_dependency_score"],
            "clean_power_score": sector["clean_power_score"],
        })

    gross_2050 = float(np.sum(values))
    default_offsets = float(spec["default_removals_credit_available_MtCO2e"])
    net_after_default_offsets = gross_2050 - min(default_offsets, max(gross_2050, 0))
    post_rate = float(spec["post_2050_reduction_rate_MtCO2e_per_year_for_stress_test"])
    delay_years = max(0, net_after_default_offsets) / post_rate if post_rate > 0 else np.nan

    if net_after_default_offsets <= 0:
        status = "Meets net zero with default offsets"
    elif net_after_default_offsets <= 70:
        status = "Near miss: residual <= 70 MtCO2e"
    elif net_after_default_offsets <= 150:
        status = "Moderate miss: residual 70-150 MtCO2e"
    else:
        status = "Large miss: residual > 150 MtCO2e"

    scenario_records.append({
        **spec.to_dict(),
        "gross_2050_before_offsets_MtCO2e": gross_2050,
        "net_2050_after_default_offsets_MtCO2e": net_after_default_offsets,
        "indicative_delay_years_after_2050": delay_years,
        "indicative_target_year": 2050 + delay_years,
        "target_status": status,
        "method_note": "Stylised mini-scenario; delay is a mechanical stress-test, not a forecast.",
    })

mini_population = pd.DataFrame(scenario_records)
sector_detail = pd.DataFrame(sector_records)

mini_population.to_csv(P8_TABLE_DIR / "p8_3_mini_scenario_population.csv", index=False)
sector_detail.to_csv(P8_TABLE_DIR / "p8_3_mini_scenario_sector_detail.csv", index=False)

display(mini_population.round(2))
"""
    ),
    md(
        r"""
## 4. Negative Emissions / Credits Stress Test

Neil specifically suggested thinking about 10 or 50 MtCO2e of negative emissions or credit purchases. This stress test applies 0, 10 and 50 MtCO2e offsets to each mini-scenario.
"""
    ),
    code(
        r"""
offset_cases = [0, 10, 50]
credit_records = []

for _, scenario in mini_population.iterrows():
    for offsets in offset_cases:
        gross = scenario["gross_2050_before_offsets_MtCO2e"]
        net = gross - min(offsets, max(gross, 0))
        post_rate = scenario["post_2050_reduction_rate_MtCO2e_per_year_for_stress_test"]
        delay = max(0, net) / post_rate if post_rate > 0 else np.nan
        if net <= 0:
            status = "Meets net zero after offsets"
        elif net <= 70:
            status = "Near miss after offsets"
        elif net <= 150:
            status = "Moderate miss after offsets"
        else:
            status = "Large miss after offsets"

        credit_records.append({
            "mini_scenario_id": scenario["mini_scenario_id"],
            "quadrant_id": scenario["quadrant_id"],
            "scenario_name": scenario["scenario_name"],
            "mini_scenario_label": scenario["mini_scenario_label"],
            "uk_policy_delivery": scenario["uk_policy_delivery"],
            "external_conditions": scenario["external_conditions"],
            "gross_2050_before_offsets_MtCO2e": gross,
            "offsets_available_MtCO2e": offsets,
            "net_2050_after_offsets_MtCO2e": net,
            "indicative_delay_years_after_2050": delay,
            "indicative_target_year": 2050 + delay,
            "target_status_after_offsets": status,
        })

credit_stress = pd.DataFrame(credit_records)
credit_stress.to_csv(P8_TABLE_DIR / "p8_3_negative_emissions_credit_stress_test.csv", index=False)
display(credit_stress.round(2))
"""
    ),
    md(
        r"""
## 5. Near-Miss Summary and Assumption Table
"""
    ),
    code(
        r"""
near_miss_summary = (
    credit_stress
    .groupby(["quadrant_id", "scenario_name", "offsets_available_MtCO2e"], as_index=False)
    .agg(
        mini_scenarios=("mini_scenario_id", "count"),
        min_net_2050_MtCO2e=("net_2050_after_offsets_MtCO2e", "min"),
        median_net_2050_MtCO2e=("net_2050_after_offsets_MtCO2e", "median"),
        max_net_2050_MtCO2e=("net_2050_after_offsets_MtCO2e", "max"),
        median_delay_years=("indicative_delay_years_after_2050", "median"),
        near_miss_count=("target_status_after_offsets", lambda s: int(s.str.contains("Near miss|Meets net zero", regex=True).sum())),
    )
)
near_miss_summary["near_miss_share"] = near_miss_summary["near_miss_count"] / near_miss_summary["mini_scenarios"]
near_miss_summary.to_csv(P8_TABLE_DIR / "p8_3_near_miss_summary.csv", index=False)

scenario_assumptions = mini_specs.merge(
    scenario_matrix[["quadrant_id", "interpretation"]],
    on="quadrant_id",
    how="left",
)
scenario_assumptions.to_csv(P8_TABLE_DIR / "p8_3_scenario_assumption_table.csv", index=False)

display(near_miss_summary.round(2))
display(scenario_assumptions)
"""
    ),
    md(
        r"""
## 6. Visualisations

The figures are saved as `.jpg` files.
"""
    ),
    code(
        r"""
quadrant_summary = (
    mini_population.groupby(["quadrant_id", "scenario_name", "uk_policy_delivery", "external_conditions"], as_index=False)
    .agg(
        median_net_2050_after_default_offsets_MtCO2e=("net_2050_after_default_offsets_MtCO2e", "median"),
        min_net_2050_after_default_offsets_MtCO2e=("net_2050_after_default_offsets_MtCO2e", "min"),
        max_net_2050_after_default_offsets_MtCO2e=("net_2050_after_default_offsets_MtCO2e", "max"),
        median_delay_years=("indicative_delay_years_after_2050", "median"),
    )
)

x_map = {"Low": 0, "High": 1}
y_map = {"Constrained": 0, "Supportive": 1}

fig, ax = plt.subplots(figsize=(8, 7))
for _, row in quadrant_summary.iterrows():
    x = x_map[row["uk_policy_delivery"]]
    y = y_map[row["external_conditions"]]
    size = 320 + max(0, row["median_net_2050_after_default_offsets_MtCO2e"]) * 2.2
    color = row["median_net_2050_after_default_offsets_MtCO2e"]
    ax.scatter(x, y, s=size, c=[color], cmap="RdYlGn_r", vmin=0, vmax=max(200, quadrant_summary["median_net_2050_after_default_offsets_MtCO2e"].max()), edgecolor="black", linewidth=0.8)
    ax.text(x, y + 0.11, row["quadrant_id"], ha="center", va="center", fontsize=12, weight="bold")
    ax.text(x, y - 0.06, f"{row['median_net_2050_after_default_offsets_MtCO2e']:.0f} Mt", ha="center", va="center", fontsize=9)
    ax.text(x, y - 0.14, f"delay {row['median_delay_years']:.1f}y", ha="center", va="center", fontsize=8)

ax.set_xlim(-0.45, 1.45)
ax.set_ylim(-0.45, 1.45)
ax.set_xticks([0, 1])
ax.set_xticklabels(["Low UK delivery", "High UK delivery"])
ax.set_yticks([0, 1])
ax.set_yticklabels(["Constrained external conditions", "Supportive external conditions"])
ax.set_title("P8-3 2x2 Scenario Matrix: Median 2050 Residual After Default Offsets", fontsize=12, pad=12)
ax.grid(alpha=0.2)
fig.tight_layout()

fig_path_quadrant = FIG_DIR / "p8_3_quadrant_outcomes.jpg"
fig.savefig(fig_path_quadrant, format="jpg", bbox_inches="tight")
plt.show()
print("Saved:", fig_path_quadrant)
"""
    ),
    code(
        r"""
plot_df = mini_population.sort_values(["quadrant_id", "domestic_delivery_score"]).copy()
colors = {"U1": "#138a72", "U2": "#73a942", "U3": "#f2a65a", "U4": "#b35c00"}

fig, ax = plt.subplots(figsize=(11, 6.2))
bars = ax.bar(
    plot_df["mini_scenario_id"],
    plot_df["net_2050_after_default_offsets_MtCO2e"],
    color=[colors[q] for q in plot_df["quadrant_id"]],
)
ax.axhline(0, color="black", linewidth=0.8)
ax.axhspan(0, 70, color="#138a72", alpha=0.08, label="Near miss band: 0-70 MtCO2e")
ax.set_title("P8-3 Mini-Scenario Residual Gap After Default Offsets", fontsize=12, pad=12)
ax.set_ylabel("Net 2050 residual after default offsets (MtCO2e)")
ax.set_xlabel("Mini-scenario")
ax.grid(axis="y", alpha=0.25)
ax.legend(frameon=False, loc="upper left", fontsize=8.5)

for bar, value in zip(bars, plot_df["net_2050_after_default_offsets_MtCO2e"]):
    ax.text(bar.get_x() + bar.get_width()/2, value + 3, f"{value:.0f}", ha="center", va="bottom", fontsize=8)

fig.tight_layout()
fig_path_bars = FIG_DIR / "p8_3_mini_scenario_residual_gaps.jpg"
fig.savefig(fig_path_bars, format="jpg", bbox_inches="tight")
plt.show()
print("Saved:", fig_path_bars)
"""
    ),
    code(
        r"""
plot_df = credit_stress.copy()
fig, ax = plt.subplots(figsize=(10.5, 6.2))

for quadrant, group in plot_df.groupby("quadrant_id"):
    summary = (
        group.groupby("offsets_available_MtCO2e", as_index=False)
        .agg(median_net=("net_2050_after_offsets_MtCO2e", "median"))
        .sort_values("offsets_available_MtCO2e")
    )
    ax.plot(
        summary["offsets_available_MtCO2e"],
        summary["median_net"],
        marker="o",
        linewidth=2.2,
        label=quadrant,
        color=colors[quadrant],
    )

ax.axhline(0, color="black", linewidth=0.8)
ax.axhspan(0, 70, color="#138a72", alpha=0.08, label="Near miss band")
ax.set_title("P8-3 Negative Emissions / Credits Sensitivity", fontsize=12, pad=12)
ax.set_xlabel("Offsets available in 2050 (MtCO2e)")
ax.set_ylabel("Median net 2050 residual across mini-scenarios (MtCO2e)")
ax.grid(alpha=0.25)
ax.legend(frameon=False, fontsize=8.5, ncol=2)
fig.tight_layout()

fig_path_credits = FIG_DIR / "p8_3_negative_emissions_credit_sensitivity.jpg"
fig.savefig(fig_path_credits, format="jpg", bbox_inches="tight")
plt.show()
print("Saved:", fig_path_credits)
"""
    ),
    md(
        r"""
## 7. Bilingual Results Summary
"""
    ),
    code(
        r"""
u1_median = float(quadrant_summary.loc[quadrant_summary["quadrant_id"].eq("U1"), "median_net_2050_after_default_offsets_MtCO2e"].iloc[0])
u4_median = float(quadrant_summary.loc[quadrant_summary["quadrant_id"].eq("U4"), "median_net_2050_after_default_offsets_MtCO2e"].iloc[0])
near_with_50 = near_miss_summary.loc[near_miss_summary["offsets_available_MtCO2e"].eq(50), ["quadrant_id", "near_miss_share"]]
near_text = "; ".join([f"{r.quadrant_id}: {r.near_miss_share:.0%}" for r in near_with_50.itertuples()])

summary_en = f'''
P8-3 result summary

P8-3 converts the uncertainty framework into a transparent 2x2 scenario matrix and twelve mini-scenarios. The matrix combines UK policy delivery strength with external technology, fuel and supply-chain conditions. The resulting 2050 residuals are not forecasts; they are stylised stress tests built from the P8-2 sectoral uncertainty bands.

The aligned-transition quadrant (U1) has the lowest residual emissions after default offsets, with a median net 2050 residual of about {u1_median:.1f} MtCO2e. The delayed-transition quadrant (U4) has the highest residual, with a median net 2050 residual of about {u4_median:.1f} MtCO2e. This supports the dissertation argument that favourable global technology conditions are helpful but cannot substitute for domestic delivery; weak UK delivery still produces a large residual gap even when external conditions are supportive.

The negative-emissions / credits stress test shows why Neil's 10/50 MtCO2e question matters. With 50 MtCO2e of offsets, the share of mini-scenarios that become on-time or near-miss outcomes is: {near_text}. This should be presented cautiously: offsets can shrink residual gaps, but they do not remove the need for sectoral mitigation, especially in buildings, transport, industry, agriculture and IAS.
'''.strip()

summary_cn = f'''
P8-3 结果总结

P8-3 把 uncertainty framework 转化为一个清晰的 2x2 scenario matrix 和 12 个 mini-scenarios。矩阵结合了 UK policy delivery 强弱和 external technology/fuel/supply-chain conditions 好坏。这里的 2050 残余排放不是预测，而是基于 P8-2 部门不确定性区间构造的 stylised stress test。

Aligned-transition 象限 U1 的残余排放最低，默认 offsets 后 2050 年 median net residual 约为 {u1_median:.1f} MtCO2e。Delayed-transition 象限 U4 的残余排放最高，median net residual 约为 {u4_median:.1f} MtCO2e。这支持论文的核心判断：有利的全球技术和成本条件很有帮助，但不能替代英国国内政策交付；如果 UK delivery 弱，即使外部条件有利，也仍然会留下较大的 residual gap。

Negative emissions / credits stress test 说明 Neil 提到的 10/50 MtCO2e 问题很重要。在 50 MtCO2e offsets 情况下，各象限中变成 on-time 或 near-miss 的 mini-scenarios 比例为：{near_text}。这需要谨慎表述：offsets 可以缩小缺口，但不能替代部门减排，尤其是 buildings、transport、industry、agriculture 和 IAS。
'''.strip()

summary_text = summary_en + "\n\n" + summary_cn
summary_path = P8_TABLE_DIR / "p8_3_results_summary_bilingual.txt"
summary_path.write_text(summary_text, encoding="utf-8")

print(summary_text)
print("\nSaved:", summary_path)
"""
    ),
    md(
        r"""
## 8. Quality Checks
"""
    ),
    code(
        r"""
checks = []

checks.append({
    "check": "matrix_quadrants",
    "expected": "4 quadrants",
    "actual": str(scenario_matrix["quadrant_id"].nunique()),
    "status": "PASS" if scenario_matrix["quadrant_id"].nunique() == 4 else "CHECK",
})
checks.append({
    "check": "mini_scenario_count",
    "expected": "12 mini-scenarios",
    "actual": str(len(mini_population)),
    "status": "PASS" if len(mini_population) == 12 else "CHECK",
})
checks.append({
    "check": "credit_stress_cases",
    "expected": "36 rows = 12 scenarios x 3 offset cases",
    "actual": str(len(credit_stress)),
    "status": "PASS" if len(credit_stress) == 36 else "CHECK",
})
checks.append({
    "check": "no_missing_net_residuals",
    "expected": "no NA values",
    "actual": str(not mini_population["net_2050_after_default_offsets_MtCO2e"].isna().any()),
    "status": "PASS" if not mini_population["net_2050_after_default_offsets_MtCO2e"].isna().any() else "CHECK",
})
checks.append({
    "check": "figures_saved_as_jpg",
    "expected": "3 jpg figures",
    "actual": f"{fig_path_quadrant.exists()} / {fig_path_bars.exists()} / {fig_path_credits.exists()}",
    "status": "PASS" if fig_path_quadrant.exists() and fig_path_bars.exists() and fig_path_credits.exists() else "CHECK",
})

quality_checks = pd.DataFrame(checks)
quality_checks.to_csv(P8_TABLE_DIR / "p8_3_quality_checks.csv", index=False)
display(quality_checks)

if not all(quality_checks["status"].eq("PASS")):
    print("One or more checks need attention before using P8-3 outputs.")
else:
    print("All P8-3 quality checks passed.")
"""
    ),
]


notebook = {
    "cells": cells,
    "metadata": {
        "kernelspec": {
            "display_name": "Python 3",
            "language": "python",
            "name": "python3",
        },
        "language_info": {
            "name": "python",
            "pygments_lexer": "ipython3",
        },
    },
    "nbformat": 4,
    "nbformat_minor": 5,
}

NOTEBOOK_PATH.parent.mkdir(parents=True, exist_ok=True)
NOTEBOOK_PATH.write_text(json.dumps(notebook, ensure_ascii=False, indent=1), encoding="utf-8")
DOWNLOADS_PATH.write_text(json.dumps(notebook, ensure_ascii=False, indent=1), encoding="utf-8")

print(f"Notebook written to: {NOTEBOOK_PATH}")
print(f"Notebook copied to: {DOWNLOADS_PATH}")
