# -*- coding: utf-8 -*-
from pathlib import Path
import json


ROOT = Path(r"E:\UCL Final Essay")
OUT_DIR = ROOT / "p7_neso_uncertainty" / "notebooks"
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT_PATH = OUT_DIR / "P7_NESO_FES_compact_indicator_extraction_local_reproducible.ipynb"


def md(source):
    return {
        "cell_type": "markdown",
        "metadata": {},
        "source": source.strip().splitlines(True),
    }


def code(source):
    return {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": source.strip().splitlines(True),
    }


cells = [
    md(
        r"""
# P7 NESO FES 2025 compact indicator extraction

This notebook supports P7 of the dissertation timeline: NESO Future Energy Scenarios, focused literature draft and uncertainty setup.

It does **not** treat NESO as a third main benchmark equal to CCC7. The purpose is narrower: extract a compact set of NESO indicators that can be used as supporting external modelling context for pathway feasibility, electrification, power-sector decarbonisation, transport, heat and flexibility assumptions.

Outputs are saved under `p7_neso_uncertainty/tables/`.
"""
    ),
    code(
        r"""
from pathlib import Path
import warnings

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning)
pd.set_option("display.max_columns", 80)
pd.set_option("display.width", 160)
"""
    ),
    md(
        r"""
## 1. Locate project root and NESO workbook

Run this notebook from anywhere inside the dissertation project folder. If automatic root detection fails, set `PROJECT_ROOT` manually.
"""
    ),
    code(
        r"""
def find_project_root(start=None):
    start = Path.cwd() if start is None else Path(start)
    for candidate in [start, *start.parents]:
        if (candidate / "Data_raw").exists():
            return candidate
    fallback = Path(r"E:\UCL Final Essay")
    if (fallback / "Data_raw").exists():
        return fallback
    raise FileNotFoundError("Could not find project root containing Data_raw.")


PROJECT_ROOT = find_project_root()
DATA_RAW = PROJECT_ROOT / "Data_raw"
P7_ROOT = PROJECT_ROOT / "p7_neso_uncertainty"
TABLE_DIR = P7_ROOT / "tables"
DATA_PROCESSED = P7_ROOT / "data_processed"
TABLE_DIR.mkdir(parents=True, exist_ok=True)
DATA_PROCESSED.mkdir(parents=True, exist_ok=True)

workbook_candidates = list((DATA_RAW / "NESO Future Energy Scenarios 2025 Data Workbook").glob("*.xlsx"))
if not workbook_candidates:
    raise FileNotFoundError("Could not find NESO FES 2025 workbook under Data_raw.")

NESO_WORKBOOK = workbook_candidates[0]
print("Project root:", PROJECT_ROOT)
print("NESO workbook:", NESO_WORKBOOK)
"""
    ),
    md(
        r"""
## 2. Helper functions

The NESO workbook has several metadata rows before the actual data table. These functions detect the header row, trim blank leading columns and convert the wide year columns into tidy long format.
"""
    ),
    code(
        r"""
SELECTED_YEARS = [2030, 2035, 2040, 2050]
CORE_PATHWAYS = ["Holistic Transition", "Electric Engagement", "Hydrogen Evolution", "Falling Behind"]
OPTIONAL_PATHWAYS = ["Ten Year Forecast"]
ALL_PATHWAYS = CORE_PATHWAYS + OPTIONAL_PATHWAYS


def read_neso_data_table(workbook_path, sheet_name):
    raw = pd.read_excel(workbook_path, sheet_name=sheet_name, header=None, engine="openpyxl")
    header_idx = None
    for i, row in raw.iterrows():
        values = row.tolist()
        non_null_positions = [j for j, value in enumerate(values) if pd.notna(value)]
        has_year = any(
            isinstance(values[j], (int, float, np.integer, np.floating)) and 1900 <= float(values[j]) <= 2100
            for j in non_null_positions
        )
        if len(non_null_positions) >= 3 and has_year:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError(f"Could not detect header row for sheet {sheet_name}.")

    header_row = raw.iloc[header_idx]
    non_null = [j for j, value in enumerate(header_row) if pd.notna(value)]
    start_col, end_col = min(non_null), max(non_null) + 1
    header = list(header_row.iloc[start_col:end_col])

    df = raw.iloc[header_idx + 1 :, start_col:end_col].copy()
    df.columns = header
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def year_columns(df, start=2023, end=2050):
    years = []
    for column in df.columns:
        if isinstance(column, (int, float, np.integer, np.floating)) and start <= int(column) <= end:
            years.append(column)
    return years


def sum_year_columns(df, by, years):
    # min_count=1 prevents all-missing years from becoming false zeros.
    return df.groupby(by, dropna=False)[years].sum(min_count=1).reset_index()


def to_indicator_long(wide_df, indicator_id, theme, source_sheet, unit, method_note, value_scale=1.0):
    years = year_columns(wide_df)
    id_cols = [col for col in wide_df.columns if col not in years]
    long = wide_df.melt(id_vars=id_cols, value_vars=years, var_name="year", value_name="value")
    long = long.rename(columns={"Pathway": "pathway"})
    if "pathway" not in long.columns:
        raise ValueError(f"{indicator_id} does not include a pathway column.")
    long["year"] = long["year"].astype(int)
    long["value"] = pd.to_numeric(long["value"], errors="coerce") * value_scale
    long["indicator_id"] = indicator_id
    long["theme"] = theme
    long["source_sheet"] = source_sheet
    long["unit"] = unit
    long["method_note"] = method_note
    return long[["indicator_id", "theme", "source_sheet", "pathway", "year", "value", "unit", "method_note"]]


def show_selected(indicators, indicator_id):
    view = indicators[(indicators["indicator_id"] == indicator_id) & (indicators["year"].isin(SELECTED_YEARS))]
    return view.pivot_table(index="pathway", columns="year", values="value", aggfunc="first").round(1)
"""
    ),
    md(
        r"""
## 3. Load only the P7-relevant NESO data tables

The workbook contains 137 sheets, but P7 only uses a deliberately small subset.
"""
    ),
    code(
        r"""
SHEETS_TO_LOAD = ["WS2", "ED1", "ES1", "ED5", "ED7", "FLX1"]
neso = {sheet: read_neso_data_table(NESO_WORKBOOK, sheet) for sheet in SHEETS_TO_LOAD}

for sheet, df in neso.items():
    years = year_columns(df)
    print(f"{sheet}: {df.shape[0]} rows x {df.shape[1]} columns; year range {int(min(years))}-{int(max(years))}")
"""
    ),
    md(
        r"""
## 4. Extract compact indicators

The first P7 extraction uses seven indicators:

1. Whole-system emissions pathway, from `WS2`.
2. Total electricity demand, from `ED1`.
3. Power-generation carbon intensity excluding BECCS, from `ES1`.
4. Electric vehicle stock, from `ED5`.
5. Heat-pump electricity demand, from `ED7`.
6. Electricity storage connection capacity, from `FLX1`.
7. Hydrogen dispatchable electricity capacity, from `FLX1`.
"""
    ),
    code(
        r"""
indicator_frames = []

# 4.1 Whole-system emissions: sum WS2 sector rows by pathway.
ws2 = neso["WS2"].copy()
ws2_years = year_columns(ws2, start=2021, end=2050)
ws2_sum = sum_year_columns(ws2, by=["Pathway"], years=ws2_years)
indicator_frames.append(
    to_indicator_long(
        ws2_sum,
        indicator_id="neso_total_emissions",
        theme="Emissions",
        source_sheet="WS2",
        unit="MtCO2e",
        method_note="Sum of WS2 sector rows by pathway, including removals/offset categories where present. Used as supporting external pathway context, not as the main CCC benchmark.",
    )
)

# 4.2 Total electricity demand: ED1 GBFES System Demand Total, converted from GWh to TWh.
ed1 = neso["ED1"].copy()
ed1_total = ed1[
    (ed1["Data item"] == "GBFES System Demand: Total")
    & (ed1["Unit"] == "GWh")
    & (ed1["Peak/ Annual/ Minimum"] == "Annual [Fiscal]")
].copy()
indicator_frames.append(
    to_indicator_long(
        ed1_total[["Pathway", *year_columns(ed1_total)]],
        indicator_id="neso_total_electricity_demand",
        theme="Electricity demand",
        source_sheet="ED1",
        unit="TWh",
        method_note="GBFES System Demand: Total, annual fiscal demand; converted from GWh to TWh.",
        value_scale=1 / 1000,
    )
)

# 4.3 Power generation carbon intensity, excluding BECCS to avoid negative-intensity interpretation dominating the comparison.
es1 = neso["ES1"].copy()
es1_intensity = es1[
    es1["Variable"].eq("CO2 Intensity of Generation excluding BECCS (gCO2/kWh)")
].copy()
indicator_frames.append(
    to_indicator_long(
        es1_intensity[["Pathway", *year_columns(es1_intensity, start=2024)]],
        indicator_id="neso_power_intensity_excluding_beccs",
        theme="Power-sector decarbonisation",
        source_sheet="ES1",
        unit="gCO2/kWh",
        method_note="CO2 intensity of generation excluding BECCS. This is easier to interpret than net intensity when BECCS produces negative values.",
    )
)

# 4.4 Electric vehicle stock, converted to million vehicles.
ed5 = neso["ED5"].copy()
ed5_ev = ed5[ed5["Data item"].eq("# Vehicles - Electric Vehicles")].copy()
indicator_frames.append(
    to_indicator_long(
        ed5_ev[["Pathway", *year_columns(ed5_ev)]],
        indicator_id="neso_electric_vehicle_stock",
        theme="Road transport electrification",
        source_sheet="ED5",
        unit="million vehicles",
        method_note="Total electric vehicle stock across vehicle classes; converted from number of vehicles to million vehicles.",
        value_scale=1 / 1_000_000,
    )
)

# 4.5 Heat-pump electricity demand: ASHP + GSHP, summed across residential, commercial and industrial sectors.
ed7 = neso["ED7"].copy()
ed7_hp = ed7[
    (ed7["Data Type"] == "Annual")
    & (ed7["Units"] == "TWh")
    & (ed7["Technology"].isin(["ASHP", "GSHP"]))
].copy()
ed7_hp_sum = sum_year_columns(ed7_hp, by=["Pathway"], years=year_columns(ed7_hp))
indicator_frames.append(
    to_indicator_long(
        ed7_hp_sum,
        indicator_id="neso_heat_pump_electricity_demand",
        theme="Buildings and heat electrification",
        source_sheet="ED7",
        unit="TWh",
        method_note="Annual ASHP + GSHP electricity demand summed across residential, commercial and industrial sectors.",
    )
)

# 4.6 Electricity storage connection capacity.
flx1 = neso["FLX1"].copy()
flx_storage = flx1[
    (flx1["Data item"] == "Electricity storage connection capacity")
    & (flx1["Detail"] == "Total (excluding vehicle-to-grid)")
    & (flx1["Unit"] == "GW")
].copy()
indicator_frames.append(
    to_indicator_long(
        flx_storage[["Pathway", *year_columns(flx_storage, start=2024)]],
        indicator_id="neso_storage_connection_capacity",
        theme="System flexibility",
        source_sheet="FLX1",
        unit="GW",
        method_note="Electricity storage connection capacity, total excluding vehicle-to-grid.",
    )
)

# 4.7 Hydrogen dispatchable electricity capacity.
flx_h2 = flx1[
    (flx1["Data item"] == "Dispatchable electricity supply capacity")
    & (flx1["Detail"] == "Hydrogen generation")
    & (flx1["Unit"] == "GW")
].copy()
indicator_frames.append(
    to_indicator_long(
        flx_h2[["Pathway", *year_columns(flx_h2, start=2024)]],
        indicator_id="neso_hydrogen_dispatchable_capacity",
        theme="Hydrogen and dispatchable power",
        source_sheet="FLX1",
        unit="GW",
        method_note="Dispatchable electricity supply capacity for hydrogen generation.",
    )
)

indicator_long = pd.concat(indicator_frames, ignore_index=True)
indicator_long = indicator_long[indicator_long["pathway"].isin(ALL_PATHWAYS)].copy()
indicator_long["value_rounded"] = indicator_long["value"].round(2)

display(indicator_long.head(12))
print("Extracted rows:", len(indicator_long))
print("Indicators:", indicator_long["indicator_id"].nunique())
"""
    ),
    md(
        r"""
## 5. Create selected-year tables

These are the tables most likely to be useful for the dissertation.
"""
    ),
    code(
        r"""
selected_years = indicator_long[indicator_long["year"].isin(SELECTED_YEARS)].copy()
selected_years = selected_years.sort_values(["indicator_id", "pathway", "year"]).reset_index(drop=True)

snapshot_2050 = selected_years[selected_years["year"] == 2050].copy()
snapshot_2050 = snapshot_2050.sort_values(["theme", "indicator_id", "pathway"]).reset_index(drop=True)

wide_selected = (
    selected_years
    .pivot_table(index=["indicator_id", "theme", "unit", "pathway"], columns="year", values="value", aggfunc="first")
    .reset_index()
)
wide_selected.columns.name = None

display(wide_selected.round(2))
"""
    ),
    md(
        r"""
## 6. Quality checks

These checks are intended to catch missing pathways, accidental duplicate rows and false zeros from missing data.
"""
    ),
    code(
        r"""
checks = []

def add_check(name, status, details):
    checks.append({"check": name, "status": status, "details": details})

add_check("NESO workbook found", "PASS" if NESO_WORKBOOK.exists() else "FAIL", str(NESO_WORKBOOK))

for sheet in SHEETS_TO_LOAD:
    add_check(
        f"{sheet} loaded",
        "PASS" if sheet in neso and not neso[sheet].empty else "FAIL",
        f"{len(neso[sheet]) if sheet in neso else 0} data rows",
    )

available_core = sorted(set(indicator_long["pathway"]).intersection(CORE_PATHWAYS))
add_check(
    "Core FES pathways present",
    "PASS" if set(CORE_PATHWAYS).issubset(set(indicator_long["pathway"])) else "FAIL",
    ", ".join(available_core),
)

expected_indicators = {
    "neso_total_emissions",
    "neso_total_electricity_demand",
    "neso_power_intensity_excluding_beccs",
    "neso_electric_vehicle_stock",
    "neso_heat_pump_electricity_demand",
    "neso_storage_connection_capacity",
    "neso_hydrogen_dispatchable_capacity",
}
available_indicators = set(indicator_long["indicator_id"].unique())
missing_indicators = sorted(expected_indicators - available_indicators)
add_check(
    "Expected indicators extracted",
    "PASS" if not missing_indicators else "FAIL",
    "missing: " + ", ".join(missing_indicators) if missing_indicators else f"{len(available_indicators)} indicators extracted",
)

dupes = indicator_long.duplicated(["indicator_id", "pathway", "year"]).sum()
add_check("No duplicate indicator-pathway-year rows", "PASS" if dupes == 0 else "FAIL", f"duplicates={dupes}")

core_2050 = indicator_long[
    (indicator_long["pathway"].isin(CORE_PATHWAYS))
    & (indicator_long["year"] == 2050)
]
missing_core_2050 = core_2050["value"].isna().sum()
add_check(
    "Core pathway 2050 values available",
    "PASS" if missing_core_2050 == 0 else "WARN",
    f"missing core 2050 values={missing_core_2050}",
)

optional_2050 = indicator_long[
    (indicator_long["pathway"].isin(OPTIONAL_PATHWAYS))
    & (indicator_long["year"] == 2050)
]
optional_missing = optional_2050["value"].isna().sum()
add_check(
    "Ten Year Forecast treated as optional",
    "PASS",
    f"Ten Year Forecast has {optional_missing} missing 2050 values; this is acceptable because it is not a full 2050 pathway.",
)

quality_checks = pd.DataFrame(checks)
display(quality_checks)
"""
    ),
    md(
        r"""
## 7. Save outputs
"""
    ),
    code(
        r"""
indicator_long.to_csv(TABLE_DIR / "p7_neso_compact_indicator_long.csv", index=False)
selected_years.to_csv(TABLE_DIR / "p7_neso_selected_year_indicators.csv", index=False)
snapshot_2050.to_csv(TABLE_DIR / "p7_neso_2050_indicator_snapshot.csv", index=False)
wide_selected.to_csv(TABLE_DIR / "p7_neso_selected_year_indicators_wide.csv", index=False)
quality_checks.to_csv(TABLE_DIR / "p7_neso_extraction_quality_checks.csv", index=False)

scope_notes = pd.DataFrame(
    [
        {
            "decision": "NESO role",
            "note": "NESO FES 2025 is used as supporting external modelling context, not as the main target-consistent benchmark.",
        },
        {
            "decision": "Main benchmark remains CCC7",
            "note": "CCC7 remains the dissertation's principal target-consistent benchmark; NESO supports pathway-feasibility discussion.",
        },
        {
            "decision": "Ten Year Forecast handling",
            "note": "Ten Year Forecast is retained where available but treated as optional because it does not provide full 2050 values for all indicators.",
        },
        {
            "decision": "First P7 cut-off",
            "note": "Only seven clean indicators are extracted in this first notebook. Additional NESO indicators should be added only if they directly support the dissertation argument.",
        },
    ]
)
scope_notes.to_csv(TABLE_DIR / "p7_neso_indicator_scope_notes.csv", index=False)

print("Saved outputs to:", TABLE_DIR)
for path in sorted(TABLE_DIR.glob("p7_neso_*.csv")):
    print("-", path.name)
"""
    ),
    md(
        r"""
## 8. Quick interpretation snapshot

Use this table to decide whether a simple NESO comparison table is enough, or whether an additional figure is worth making in the next step.
"""
    ),
    code(
        r"""
for indicator in [
    "neso_total_emissions",
    "neso_total_electricity_demand",
    "neso_power_intensity_excluding_beccs",
    "neso_electric_vehicle_stock",
    "neso_heat_pump_electricity_demand",
    "neso_storage_connection_capacity",
    "neso_hydrogen_dispatchable_capacity",
]:
    print("\n" + "=" * 90)
    print(indicator)
    display(show_selected(indicator_long, indicator))
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

OUT_PATH.write_text(json.dumps(notebook, indent=1, ensure_ascii=False), encoding="utf-8")
print(OUT_PATH)
