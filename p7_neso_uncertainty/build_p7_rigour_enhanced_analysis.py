from __future__ import annotations

import json
from pathlib import Path
from textwrap import dedent

import matplotlib.pyplot as plt
import pandas as pd


DEFAULT_PROJECT_ROOT = Path(r"E:\UCL Final Essay")


def find_project_root() -> Path:
    candidates = [Path.cwd(), *Path.cwd().parents, DEFAULT_PROJECT_ROOT]
    for candidate in candidates:
        if (candidate / "Reference" / "P7_external_conditions").exists():
            return candidate
    raise FileNotFoundError(
        "Could not find project root. Set DEFAULT_PROJECT_ROOT in this script "
        "to the folder containing Reference/P7_external_conditions."
    )


def clean_label(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).replace("\n", " ").strip()


def pct_change(new: float, old: float) -> float:
    return (new / old - 1.0) * 100.0


def extract_lcoe(econ_workbook: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    raw = pd.read_excel(econ_workbook, sheet_name="EA.03", header=None)
    rows = []
    for col in range(raw.shape[1]):
        technology = clean_label(raw.iat[7, col]) if col < raw.shape[1] else ""
        if not technology:
            continue
        value = raw.iat[8, col]
        if pd.isna(value):
            continue
        rows.append(
            {
                "technology": technology,
                "lcoe_2025_prices_gbp_per_MWh": float(value),
                "carbon_cost_component_gbp_per_MWh": float(raw.iat[9, col])
                if not pd.isna(raw.iat[9, col])
                else 0.0,
                "source_sheet": "EA.03",
            }
        )
    lcoe = pd.DataFrame(rows)

    pairs = {
        "Solar PV": ("2025 Solar PV", "2035 Solar PV"),
        "Fixed offshore wind": ("2025 Fixed Offshore Wind", "2035 Fixed Offshore Wind"),
        "Onshore wind": ("2025 Onshore Wind", "2035 Onshore Wind"),
        "Floating offshore wind": ("2025 Floating Offshore Wind", "2035 Floating Offshore Wind"),
    }
    learning_rows = []
    for label, (base_tech, future_tech) in pairs.items():
        base = float(lcoe.loc[lcoe["technology"].eq(base_tech), "lcoe_2025_prices_gbp_per_MWh"].iloc[0])
        future = float(lcoe.loc[lcoe["technology"].eq(future_tech), "lcoe_2025_prices_gbp_per_MWh"].iloc[0])
        learning_rows.append(
            {
                "technology": label,
                "lcoe_2025_gbp_per_MWh": base,
                "lcoe_2035_gbp_per_MWh": future,
                "absolute_change_gbp_per_MWh": future - base,
                "percentage_change_2025_to_2035": pct_change(future, base),
            }
        )
    learning = pd.DataFrame(learning_rows)
    return lcoe, learning


def extract_capital_intensity(econ_workbook: Path) -> pd.DataFrame:
    raw = pd.read_excel(econ_workbook, sheet_name="EA.04", header=None)
    rows = []
    current_unit = ""
    for row in range(raw.shape[0]):
        unit = clean_label(raw.iat[row, 13]) if raw.shape[1] > 13 else ""
        if unit and unit != "Unit":
            current_unit = unit
        technology = clean_label(raw.iat[row, 14]) if raw.shape[1] > 14 else ""
        if not technology or technology == "Technology":
            continue
        capex = raw.iat[row, 15]
        opex = raw.iat[row, 16]
        fuel = raw.iat[row, 17]
        if pd.isna(capex) or pd.isna(opex) or pd.isna(fuel):
            continue
        rows.append(
            {
                "unit": current_unit,
                "technology": technology,
                "capex": float(capex),
                "opex": float(opex),
                "fuel_cost": float(fuel),
                "total_cost_component": float(capex) + float(opex) + float(fuel),
                "source_sheet": "EA.04",
            }
        )
    return pd.DataFrame(rows)


def extract_cost_sensitivity(econ_workbook: Path) -> pd.DataFrame:
    raw = pd.read_excel(econ_workbook, sheet_name="EA.12", header=None)
    years = raw.iloc[7].to_dict()
    year_cols = {int(v): k for k, v in years.items() if pd.notna(v) and str(v).replace(".0", "").isdigit()}
    pathway = None
    rows = []
    for row in range(8, 23):
        maybe_pathway = clean_label(raw.iat[row, 13])
        if maybe_pathway:
            pathway = maybe_pathway
        case = clean_label(raw.iat[row, 14])
        if case not in {"Base case", "Low case", "High case"}:
            continue
        for year, col in year_cols.items():
            value = raw.iat[row, col]
            if pd.isna(value):
                continue
            rows.append(
                {
                    "pathway": pathway,
                    "case": case,
                    "year": year,
                    "total_energy_cost_gbp_bn_2025_prices": float(value),
                    "source_sheet": "EA.12",
                }
            )
    return pd.DataFrame(rows)


def extract_fuel_shock(econ_workbook: Path) -> pd.DataFrame:
    raw = pd.read_excel(econ_workbook, sheet_name="EA.16", header=None)
    rows = []
    for row in range(9, 13):
        case = clean_label(raw.iat[row, 13])
        if not case:
            continue
        rows.extend(
            [
                {
                    "case": case,
                    "year_or_pathway": "2022 historical shock",
                    "change_in_energy_cost_pct_gdp": float(raw.iat[row, 14]),
                    "source_sheet": "EA.16",
                },
                {
                    "case": case,
                    "year_or_pathway": "Falling Behind 2050",
                    "change_in_energy_cost_pct_gdp": float(raw.iat[row, 15]),
                    "source_sheet": "EA.16",
                },
                {
                    "case": case,
                    "year_or_pathway": "Holistic Transition 2050",
                    "change_in_energy_cost_pct_gdp": float(raw.iat[row, 16]),
                    "source_sheet": "EA.16",
                },
            ]
        )
    return pd.DataFrame(rows)


def extract_gas_imports(econ_workbook: Path) -> pd.DataFrame:
    raw = pd.read_excel(econ_workbook, sheet_name="EA.17", header=None)
    rows = []
    for row in range(8, 12):
        pathway = clean_label(raw.iat[row, 13])
        for col in range(14, 18):
            year = int(raw.iat[7, col])
            rows.append(
                {
                    "pathway": pathway,
                    "year": year,
                    "gb_gas_imports_by_volume": float(raw.iat[row, col]),
                    "source_sheet": "EA.17",
                }
            )
    return pd.DataFrame(rows)


def extract_benchmark_gap(project_root: Path) -> pd.DataFrame:
    p5_file = project_root / "p4_p5_local_reproduction" / "data_processed" / "p5_cleaned_desnz_ccc_annual_comparison.csv"
    p5 = pd.read_csv(p5_file)
    selected = p5[p5["year"].isin([2030, 2035, 2040, 2045, 2050])].copy()
    cols = [
        "year",
        "DESNZ_EEP_2024_inc_IAS_MtCO2e",
        "DESNZ_EEP_2024_excl_IAS_MtCO2e",
        "CCC7_Balanced_Pathway_MtCO2e",
        "gap_inc_IAS_DESNZ_minus_CCC7_MtCO2e",
        "gap_excl_IAS_DESNZ_minus_CCC7_MtCO2e",
    ]
    return selected[cols].reset_index(drop=True)


def build_metric_summary(
    target_years: pd.DataFrame,
    learning: pd.DataFrame,
    lcoe: pd.DataFrame,
    cost_sensitivity: pd.DataFrame,
    fuel_shock: pd.DataFrame,
    gas_imports: pd.DataFrame,
) -> pd.DataFrame:
    rows = []

    def add(metric, value, unit, interpretation, source):
        rows.append(
            {
                "metric": metric,
                "value": value,
                "unit": unit,
                "interpretation": interpretation,
                "source": source,
            }
        )

    for _, row in learning.iterrows():
        add(
            f"{row['technology']} LCOE change, 2025 to 2035",
            row["percentage_change_2025_to_2035"],
            "%",
            f"{row['technology']} becomes cheaper in NESO's 2035 cost assumptions.",
            "NESO FES 2025 Economics Tables and Graphs Data Workbook, EA.03",
        )

    solar_2035 = float(lcoe.loc[lcoe["technology"].eq("2035 Solar PV"), "lcoe_2025_prices_gbp_per_MWh"].iloc[0])
    gas_baseload = float(lcoe.loc[lcoe["technology"].eq("Gas CCGT"), "lcoe_2025_prices_gbp_per_MWh"].iloc[-1])
    add(
        "2035 Solar PV LCOE relative to baseload Gas CCGT",
        pct_change(solar_2035, gas_baseload),
        "%",
        "A negative value means solar PV is cheaper on this plant-level LCOE measure.",
        "NESO FES 2025 Economics Tables and Graphs Data Workbook, EA.03",
    )

    shock = fuel_shock[fuel_shock["case"].eq("2022 price shock")]
    shock_2022 = float(shock.loc[shock["year_or_pathway"].eq("2022 historical shock"), "change_in_energy_cost_pct_gdp"].iloc[0])
    shock_fb = float(shock.loc[shock["year_or_pathway"].eq("Falling Behind 2050"), "change_in_energy_cost_pct_gdp"].iloc[0])
    shock_ht = float(shock.loc[shock["year_or_pathway"].eq("Holistic Transition 2050"), "change_in_energy_cost_pct_gdp"].iloc[0])
    add(
        "Holistic Transition 2050 fossil price shock exposure vs 2022 shock",
        (1.0 - shock_ht / shock_2022) * 100.0,
        "% reduction",
        "Holistic Transition leaves much lower GDP-scale exposure to a 2022-style fossil fuel price shock.",
        "NESO FES 2025 Economics Tables and Graphs Data Workbook, EA.16",
    )
    add(
        "Falling Behind 2050 fossil price shock exposure vs Holistic Transition 2050",
        shock_fb / shock_ht,
        "ratio",
        "Falling Behind remains materially more exposed to fossil price volatility than Holistic Transition.",
        "NESO FES 2025 Economics Tables and Graphs Data Workbook, EA.16",
    )

    gas_2024 = gas_imports[gas_imports["year"].eq(2024)].set_index("pathway")["gb_gas_imports_by_volume"]
    gas_2050 = gas_imports[gas_imports["year"].eq(2050)].set_index("pathway")["gb_gas_imports_by_volume"]
    for pathway in ["Holistic Transition", "Electric Engagement", "Hydrogen Evolution", "Falling Behind"]:
        add(
            f"{pathway} gas import change, 2024 to 2050",
            pct_change(gas_2050[pathway], gas_2024[pathway]),
            "%",
            "This is a proxy for changing exposure to international gas markets.",
            "NESO FES 2025 Economics Tables and Graphs Data Workbook, EA.17",
        )

    add(
        "Falling Behind gas imports relative to Holistic Transition in 2050",
        gas_2050["Falling Behind"] / gas_2050["Holistic Transition"],
        "ratio",
        "In 2050, Falling Behind uses a much higher imported-gas volume than Holistic Transition.",
        "NESO FES 2025 Economics Tables and Graphs Data Workbook, EA.17",
    )

    cost_2050 = cost_sensitivity[cost_sensitivity["year"].eq(2050)]
    pivot = cost_2050.pivot(index="pathway", columns="case", values="total_energy_cost_gbp_bn_2025_prices")
    for pathway in pivot.index:
        range_value = pivot.loc[pathway, "High case"] - pivot.loc[pathway, "Low case"]
        add(
            f"{pathway} 2050 total energy cost sensitivity range",
            range_value,
            "GBP bn, 2025 prices",
            "High-minus-low sensitivity range from NESO's capex/fuel-cost sensitivity cases.",
            "NESO FES 2025 Economics Tables and Graphs Data Workbook, EA.12",
        )
    add(
        "Base-case 2050 total energy cost: Falling Behind minus Holistic Transition",
        pivot.loc["Falling Behind", "Base case"] - pivot.loc["Holistic Transition", "Base case"],
        "GBP bn, 2025 prices",
        "Central-cost comparison: Falling Behind is more costly than Holistic Transition in 2050 in this NESO table.",
        "NESO FES 2025 Economics Tables and Graphs Data Workbook, EA.12",
    )

    missed = target_years[target_years["target_status"].str.contains("Does not meet", na=False)]
    for _, row in missed.iterrows():
        add(
            f"{row['pathway']} estimated delay after 2050",
            row["delay_years_after_2050"],
            "years",
            row["method_note"],
            "Local DESNZ/CCC/NESO target-year analysis",
        )

    return pd.DataFrame(rows)


def plot_outputs(
    fig_dir: Path,
    target_years: pd.DataFrame,
    learning: pd.DataFrame,
    cost_sensitivity: pd.DataFrame,
    fuel_shock: pd.DataFrame,
    gas_imports: pd.DataFrame,
) -> dict[str, Path]:
    fig_dir.mkdir(parents=True, exist_ok=True)
    out = {}
    plt.style.use("seaborn-v0_8-whitegrid")

    delay = target_years.copy()
    delay["delay_years_after_2050"] = delay["delay_years_after_2050"].fillna(0.0)
    delay = delay.sort_values("delay_years_after_2050", ascending=True)
    fig, ax = plt.subplots(figsize=(8.5, 4.8))
    colors = ["#2E7D32" if v == 0 else "#B23A48" for v in delay["delay_years_after_2050"]]
    ax.barh(delay["pathway"], delay["delay_years_after_2050"], color=colors)
    ax.set_xlabel("Estimated delay after 2050 (years)")
    ax.set_title("Target-achievement timing across DESNZ, CCC and NESO pathways")
    ax.axvline(0, color="#333333", linewidth=0.8)
    for i, v in enumerate(delay["delay_years_after_2050"]):
        label = "on time" if v == 0 else f"{v:.1f} years"
        ax.text(max(v, 0) + 1.0, i, label, va="center", fontsize=9)
    fig.tight_layout()
    path = fig_dir / "p7_target_delay_years.jpg"
    fig.savefig(path, dpi=220, format="jpg")
    plt.close(fig)
    out["target_delay"] = path

    fig, ax = plt.subplots(figsize=(8.5, 4.8))
    x = range(len(learning))
    ax.bar([i - 0.18 for i in x], learning["lcoe_2025_gbp_per_MWh"], width=0.36, label="2025", color="#7A869A")
    ax.bar([i + 0.18 for i in x], learning["lcoe_2035_gbp_per_MWh"], width=0.36, label="2035", color="#2E74B5")
    ax.set_xticks(list(x))
    ax.set_xticklabels(learning["technology"], rotation=20, ha="right")
    ax.set_ylabel("LCOE (GBP/MWh, 2025 prices)")
    ax.set_title("NESO clean-technology cost assumptions: 2025 vs 2035")
    ax.legend(frameon=False)
    for i, row in learning.iterrows():
        ax.text(i + 0.18, row["lcoe_2035_gbp_per_MWh"] + 3, f"{row['percentage_change_2025_to_2035']:.0f}%", ha="center", fontsize=9)
    fig.tight_layout()
    path = fig_dir / "p7_lcoe_learning_2025_2035.jpg"
    fig.savefig(path, dpi=220, format="jpg")
    plt.close(fig)
    out["lcoe_learning"] = path

    shock = fuel_shock[fuel_shock["case"].eq("2022 price shock")].copy()
    fig, ax = plt.subplots(figsize=(7.6, 4.5))
    labels = shock["year_or_pathway"].replace(
        {
            "2022 historical shock": "2022 historical shock",
            "Falling Behind 2050": "Falling Behind 2050",
            "Holistic Transition 2050": "Holistic Transition 2050",
        }
    )
    ax.bar(labels, shock["change_in_energy_cost_pct_gdp"] * 100, color=["#7A869A", "#B23A48", "#2E7D32"])
    ax.set_ylabel("Change in energy costs (% of GDP)")
    ax.set_title("Fossil-fuel price-shock exposure")
    ax.tick_params(axis="x", rotation=15)
    fig.tight_layout()
    path = fig_dir / "p7_fossil_price_shock_exposure.jpg"
    fig.savefig(path, dpi=220, format="jpg")
    plt.close(fig)
    out["fossil_shock"] = path

    gas_wide = gas_imports.pivot(index="pathway", columns="year", values="gb_gas_imports_by_volume").loc[
        ["Holistic Transition", "Electric Engagement", "Hydrogen Evolution", "Falling Behind"]
    ]
    fig, ax = plt.subplots(figsize=(8.3, 4.8))
    x = range(len(gas_wide))
    ax.bar([i - 0.18 for i in x], gas_wide[2024], width=0.36, label="2024", color="#7A869A")
    ax.bar([i + 0.18 for i in x], gas_wide[2050], width=0.36, label="2050", color="#1F4D78")
    ax.set_xticks(list(x))
    ax.set_xticklabels(gas_wide.index, rotation=18, ha="right")
    ax.set_ylabel("GB gas imports by volume")
    ax.set_title("Gas import exposure by NESO pathway")
    ax.legend(frameon=False)
    fig.tight_layout()
    path = fig_dir / "p7_gas_import_exposure.jpg"
    fig.savefig(path, dpi=220, format="jpg")
    plt.close(fig)
    out["gas_imports"] = path

    cost_2050 = cost_sensitivity[cost_sensitivity["year"].eq(2050)].pivot(
        index="pathway", columns="case", values="total_energy_cost_gbp_bn_2025_prices"
    )
    cost_2050 = cost_2050.loc[["Holistic Transition", "Electric Engagement", "Hydrogen Evolution", "Falling Behind"]]
    fig, ax = plt.subplots(figsize=(8.4, 4.8))
    centers = cost_2050["Base case"]
    lower = centers - cost_2050["Low case"]
    upper = cost_2050["High case"] - centers
    ax.errorbar(
        centers,
        range(len(cost_2050)),
        xerr=[lower, upper],
        fmt="o",
        color="#0B2545",
        ecolor="#8392A5",
        elinewidth=4,
        capsize=6,
    )
    ax.set_yticks(list(range(len(cost_2050))))
    ax.set_yticklabels(cost_2050.index)
    ax.set_xlabel("2050 total energy cost (GBP bn, 2025 prices)")
    ax.set_title("2050 total energy cost sensitivity range")
    fig.tight_layout()
    path = fig_dir / "p7_2050_cost_sensitivity_range.jpg"
    fig.savefig(path, dpi=220, format="jpg")
    plt.close(fig)
    out["cost_sensitivity"] = path

    return out


def write_notebook(project_root: Path, notebook_path: Path) -> None:
    notebook_path.parent.mkdir(parents=True, exist_ok=True)
    cells = [
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": "# P7 rigour-enhanced analysis: target timing, costs and external conditions\n\nThis notebook reproduces the additional P7 analysis from local DESNZ/CCC/NESO outputs and the downloaded NESO FES 2025 economics workbooks.",
        },
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": "## 1. Setup\nRun the next cell first. If your project is not at `E:/UCL Final Essay`, edit `DEFAULT_PROJECT_ROOT`.",
        },
        {
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "outputs": [],
            "source": dedent(
                r"""
                from pathlib import Path
                import runpy
                import pandas as pd

                DEFAULT_PROJECT_ROOT = Path(r"E:\UCL Final Essay")

                def find_project_root():
                    candidates = [Path.cwd(), *Path.cwd().parents, DEFAULT_PROJECT_ROOT]
                    for candidate in candidates:
                        if (candidate / "Reference" / "P7_external_conditions").exists():
                            return candidate
                    raise FileNotFoundError(
                        "Could not find the project root. Update DEFAULT_PROJECT_ROOT to the folder containing Reference/P7_external_conditions."
                    )

                project_root = find_project_root()
                script_path = project_root / "p7_neso_uncertainty" / "build_p7_rigour_enhanced_analysis.py"
                print("Project root:", project_root)
                print("Analysis script:", script_path)
                """
            ).strip(),
        },
        {
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "outputs": [],
            "source": "runpy.run_path(str(script_path), run_name='__main__')",
        },
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": "## 2. Read the main result tables",
        },
        {
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "outputs": [],
            "source": "project_root = find_project_root()\nout_dir = project_root / 'p7_neso_uncertainty' / 'tables'\nmetric_summary = pd.read_csv(out_dir / 'p7_rigour_enhanced_external_metrics.csv')\ntarget_years = pd.read_csv(out_dir / 'p7_target_achievement_year_estimates.csv')\nbenchmark_gap = pd.read_csv(out_dir / 'p7_benchmark_year_gap_summary.csv')\ndisplay(metric_summary)\ndisplay(target_years)\ndisplay(benchmark_gap)",
        },
    ]
    nb = {
        "cells": cells,
        "metadata": {
            "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
            "language_info": {"name": "python", "pygments_lexer": "ipython3"},
        },
        "nbformat": 4,
        "nbformat_minor": 5,
    }
    notebook_path.write_text(json.dumps(nb, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> dict[str, str]:
    project_root = find_project_root()
    external_dir = project_root / "Reference" / "P7_external_conditions"
    table_dir = project_root / "p7_neso_uncertainty" / "tables"
    fig_dir = project_root / "p7_neso_uncertainty" / "figures"
    notebook_dir = project_root / "p7_neso_uncertainty" / "notebooks"
    table_dir.mkdir(parents=True, exist_ok=True)
    fig_dir.mkdir(parents=True, exist_ok=True)
    notebook_dir.mkdir(parents=True, exist_ok=True)

    econ_workbook = external_dir / "FES 2025 Economics Tables and Graphs Data Workbook v1.xlsx"
    if not econ_workbook.exists():
        raise FileNotFoundError(f"Missing required workbook: {econ_workbook}")

    lcoe, learning = extract_lcoe(econ_workbook)
    capital = extract_capital_intensity(econ_workbook)
    cost_sensitivity = extract_cost_sensitivity(econ_workbook)
    fuel_shock = extract_fuel_shock(econ_workbook)
    gas_imports = extract_gas_imports(econ_workbook)
    benchmark_gap = extract_benchmark_gap(project_root)
    target_years = pd.read_csv(table_dir / "p7_target_achievement_year_estimates.csv")
    metrics = build_metric_summary(target_years, learning, lcoe, cost_sensitivity, fuel_shock, gas_imports)

    lcoe.to_csv(table_dir / "p7_lcoe_2035_selected.csv", index=False)
    learning.to_csv(table_dir / "p7_lcoe_renewable_learning_2025_2035.csv", index=False)
    capital.to_csv(table_dir / "p7_capital_intensity_2035.csv", index=False)
    cost_sensitivity.to_csv(table_dir / "p7_total_energy_cost_sensitivity_long.csv", index=False)
    cost_sensitivity[cost_sensitivity["year"].eq(2050)].to_csv(table_dir / "p7_total_energy_cost_sensitivity_2050.csv", index=False)
    fuel_shock.to_csv(table_dir / "p7_fossil_price_shock_exposure.csv", index=False)
    gas_imports.to_csv(table_dir / "p7_gas_imports_by_pathway.csv", index=False)
    benchmark_gap.to_csv(table_dir / "p7_benchmark_year_gap_summary.csv", index=False)
    metrics.to_csv(table_dir / "p7_rigour_enhanced_external_metrics.csv", index=False)

    figures = plot_outputs(fig_dir, target_years, learning, cost_sensitivity, fuel_shock, gas_imports)
    notebook_path = notebook_dir / "P7_rigour_enhanced_external_conditions_local_reproducible.ipynb"
    write_notebook(project_root, notebook_path)

    outputs = {
        "metrics": str(table_dir / "p7_rigour_enhanced_external_metrics.csv"),
        "benchmark_gap": str(table_dir / "p7_benchmark_year_gap_summary.csv"),
        "notebook": str(notebook_path),
    }
    outputs.update({name: str(path) for name, path in figures.items()})
    print(json.dumps(outputs, indent=2))
    return outputs


if __name__ == "__main__":
    main()
