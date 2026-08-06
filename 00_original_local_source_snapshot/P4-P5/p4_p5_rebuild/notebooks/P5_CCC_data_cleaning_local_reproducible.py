# %% [markdown]
# # P5 CCC Data Cleaning - Local Reproducible Notebook
#
# This notebook cleans the CCC pathway datasets needed for P5:
# - CCC Seventh Carbon Budget Balanced Pathway, economy-wide annual emissions.
# - CCC Seventh Carbon Budget sector-level annual direct emissions.
# - CCC Seventh Carbon Budget sector classification metadata.
# - CCC Sixth Carbon Budget Balanced Net Zero Pathway as a sensitivity benchmark.
#
# Outputs are written to `p4_p5_local_reproduction/data_processed` and
# `p4_p5_local_reproduction/tables`.

# %%
from pathlib import Path
import re
import sys

import numpy as np
import pandas as pd

try:
    display
except NameError:
    def safe_print(text) -> None:
        encoding = sys.stdout.encoding or "utf-8"
        print(str(text).encode(encoding, errors="replace").decode(encoding, errors="replace"))


    def display(obj):
        if isinstance(obj, pd.DataFrame):
            safe_print(obj.to_string(index=False))
        elif isinstance(obj, pd.Series):
            safe_print(obj.to_string())
        else:
            safe_print(obj)


# %% [markdown]
# ## 1. Locate project folders and source files

# %%
# If Jupyter is launched from another folder, keep this path pointing to the
# dissertation project folder that contains Data_raw.
PROJECT_ROOT_OVERRIDE = r"E:\UCL Final Essay"


def find_project_root(start: Path | None = None) -> Path:
    """Find the project folder containing Data_raw."""
    candidates = []

    if PROJECT_ROOT_OVERRIDE:
        candidates.append(Path(PROJECT_ROOT_OVERRIDE).expanduser())

    start = Path.cwd() if start is None else Path(start)
    candidates.extend([start, *start.parents])

    if "__file__" in globals():
        script_path = Path(__file__).resolve()
        candidates.extend([script_path.parent, *script_path.parents])

    for candidate in candidates:
        if (candidate / "Data_raw").exists():
            return candidate

    raise FileNotFoundError(
        "Could not find a project root containing Data_raw. "
        "Set PROJECT_ROOT_OVERRIDE to the folder that contains Data_raw, "
        "for example r'E:\\UCL Final Essay'."
    )


PROJECT_ROOT = find_project_root()
DATA_RAW = PROJECT_ROOT / "Data_raw"
OUTPUT_ROOT = PROJECT_ROOT / "p4_p5_local_reproduction"
DATA_PROCESSED = OUTPUT_ROOT / "data_processed"
TABLES_DIR = OUTPUT_ROOT / "tables"

DATA_PROCESSED.mkdir(parents=True, exist_ok=True)
TABLES_DIR.mkdir(parents=True, exist_ok=True)

print("Project root:", PROJECT_ROOT)
print("Data raw:", DATA_RAW)
print("Output data:", DATA_PROCESSED)


def find_data_file(filename: str, preferred_contains: str | None = None) -> Path:
    """Find a file in Data_raw, optionally preferring a path containing a phrase."""
    matches = sorted(DATA_RAW.rglob(filename))
    if not matches:
        raise FileNotFoundError(f"Could not find {filename} under {DATA_RAW}")
    if preferred_contains:
        preferred = [
            path for path in matches
            if preferred_contains.lower() in str(path).lower()
        ]
        if preferred:
            return preferred[0]
    return matches[0]


CCC7_DATA = find_data_file(
    "The-Seventh-Carbon-Budget-full-dataset.xlsx",
    preferred_contains="CCC Seventh Carbon Budget",
)
CCC6_DATA = find_data_file(
    "The-Sixth-Carbon-Budget-Dataset_v2.xlsx",
    preferred_contains="CCC Sixth Carbon Budget",
)

source_files = pd.DataFrame(
    [
        {
            "dataset": "CCC7 Seventh Carbon Budget full dataset",
            "path": str(CCC7_DATA),
            "role": "Main P5 target-consistent benchmark",
        },
        {
            "dataset": "CCC6 Sixth Carbon Budget dataset",
            "path": str(CCC6_DATA),
            "role": "Sensitivity benchmark only",
        },
    ]
)
display(source_files)


# %% [markdown]
# ## 2. Helper functions

# %%
DIRECT_EMISSIONS_TOTAL = "Emissions: direct emissions total"
UNIT_MTCO2E = "MtCO2e"


def clean_text(value) -> str:
    """Normalize spreadsheet labels without changing their meaning."""
    if pd.isna(value):
        return ""
    text = str(value).replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def unique_column_names(columns) -> list[str]:
    """Return non-empty, unique column names while preserving the original labels."""
    seen = {}
    out = []
    for index, col in enumerate(columns):
        base = clean_text(col) or f"unnamed_{index}"
        if base not in seen:
            seen[base] = 0
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base}_{seen[base]}")
    return out


def require_columns(df: pd.DataFrame, required: set[str], label: str) -> None:
    missing = required.difference(df.columns)
    if missing:
        raise ValueError(f"{label} is missing expected columns: {sorted(missing)}")


def clean_ccc7_long_table(
    df: pd.DataFrame,
    label: str,
    extra_dimension_cols: list[str] | None = None,
) -> pd.DataFrame:
    """Clean one CCC7 long-format table and keep direct total emissions."""
    extra_dimension_cols = extra_dimension_cols or []
    required = {"scenario", "country", "variable", "variable_unit", "year", "value"}
    require_columns(df, required.union(extra_dimension_cols), label)

    cleaned = df.copy()
    for col in ["scenario", "country", "variable", "variable_unit", *extra_dimension_cols]:
        cleaned[col] = cleaned[col].map(clean_text)
    cleaned["year"] = pd.to_numeric(cleaned["year"], errors="coerce").astype("Int64")
    cleaned["emissions_MtCO2e"] = pd.to_numeric(cleaned["value"], errors="coerce")

    cleaned = cleaned[
        (cleaned["country"] == "United Kingdom")
        & (cleaned["variable"] == DIRECT_EMISSIONS_TOTAL)
        & (cleaned["variable_unit"] == UNIT_MTCO2E)
        & cleaned["year"].notna()
        & cleaned["emissions_MtCO2e"].notna()
    ].copy()

    cleaned["year"] = cleaned["year"].astype(int)
    cleaned = cleaned.drop(columns=["value"])

    cols = [
        "scenario",
        "country",
        *extra_dimension_cols,
        "variable",
        "variable_unit",
        "year",
        "emissions_MtCO2e",
    ]
    cleaned = cleaned[cols].sort_values(["scenario", *extra_dimension_cols, "year"])
    cleaned = cleaned.reset_index(drop=True)
    return cleaned


def expected_years(start: int, end: int) -> set[int]:
    return set(range(start, end + 1))


def check_year_coverage(df: pd.DataFrame, scenario: str, start: int, end: int) -> tuple[bool, str]:
    years = set(df.loc[df["scenario"] == scenario, "year"].astype(int))
    missing = sorted(expected_years(start, end).difference(years))
    extra = sorted(years.difference(expected_years(start, end)))
    ok = not missing and not extra
    details = f"missing={missing}; extra={extra}"
    return ok, details


# %% [markdown]
# ## 3. Clean CCC7 economy-wide direct emissions
#
# The main P5 benchmark is:
# - workbook: CCC Seventh Carbon Budget full dataset;
# - sheet: `Economy-wide data`;
# - scenario: `Balanced Pathway`;
# - country: `United Kingdom`;
# - variable: `Emissions: direct emissions total`;
# - unit: `MtCO2e`.

# %%
ccc7_economy_raw = pd.read_excel(CCC7_DATA, sheet_name="Economy-wide data")
ccc7_economy_clean = clean_ccc7_long_table(
    ccc7_economy_raw,
    label="CCC7 Economy-wide data",
)

ccc7_economy_balanced = ccc7_economy_clean[
    ccc7_economy_clean["scenario"] == "Balanced Pathway"
].copy()

ccc7_economy_all_path = DATA_PROCESSED / "ccc7_clean_economy_wide_direct_emissions_all_scenarios.csv"
ccc7_economy_balanced_path = DATA_PROCESSED / "ccc7_clean_economy_wide_balanced_pathway.csv"
ccc7_economy_clean.to_csv(ccc7_economy_all_path, index=False, encoding="utf-8-sig")
ccc7_economy_balanced.to_csv(ccc7_economy_balanced_path, index=False, encoding="utf-8-sig")

display(ccc7_economy_balanced.head())
display(ccc7_economy_balanced.tail())


# %% [markdown]
# ## 4. Clean CCC7 sector-level direct emissions
#
# This is used for sector-alignment checks and broad sector-level context. It should
# not be treated as a strict one-to-one DESNZ sector decomposition without caveats.

# %%
ccc7_sector_raw = pd.read_excel(CCC7_DATA, sheet_name="Sector-level data")
ccc7_sector_clean = clean_ccc7_long_table(
    ccc7_sector_raw,
    label="CCC7 Sector-level data",
    extra_dimension_cols=["sector"],
)

ccc7_sector_balanced = ccc7_sector_clean[
    ccc7_sector_clean["scenario"] == "Balanced Pathway"
].copy()

ccc7_sector_all_path = DATA_PROCESSED / "ccc7_clean_sector_direct_emissions_all_scenarios.csv"
ccc7_sector_balanced_path = DATA_PROCESSED / "ccc7_clean_sector_direct_emissions_balanced_pathway.csv"
ccc7_sector_clean.to_csv(ccc7_sector_all_path, index=False, encoding="utf-8-sig")
ccc7_sector_balanced.to_csv(ccc7_sector_balanced_path, index=False, encoding="utf-8-sig")

display(ccc7_sector_balanced.head())
display(
    ccc7_sector_balanced.loc[ccc7_sector_balanced["year"] == 2050]
    .sort_values("emissions_MtCO2e", ascending=False)
)


# %% [markdown]
# ## 5. Clean CCC7 sector classification and variable definitions
#
# These metadata outputs make the cleaning auditable and help justify sector
# alignment decisions.

# %%
ccc7_sector_classification = pd.read_excel(CCC7_DATA, sheet_name="Sector classification")
ccc7_sector_classification.columns = unique_column_names(ccc7_sector_classification.columns)
for col in ccc7_sector_classification.columns:
    if ccc7_sector_classification[col].dtype == "object":
        ccc7_sector_classification[col] = ccc7_sector_classification[col].map(clean_text)

sector_class_path = DATA_PROCESSED / "ccc7_clean_sector_classification.csv"
ccc7_sector_classification.to_csv(sector_class_path, index=False, encoding="utf-8-sig")

ccc7_variable_definitions = pd.read_excel(CCC7_DATA, sheet_name="Variable definitions")
ccc7_variable_definitions.columns = unique_column_names(ccc7_variable_definitions.columns)
for col in ccc7_variable_definitions.columns:
    if ccc7_variable_definitions[col].dtype == "object":
        ccc7_variable_definitions[col] = ccc7_variable_definitions[col].map(clean_text)

emissions_variable_definitions = ccc7_variable_definitions[
    ccc7_variable_definitions["variable_name"].str.contains("Emissions:", regex=False, na=False)
].copy()
variable_def_path = DATA_PROCESSED / "ccc7_clean_emissions_variable_definitions.csv"
emissions_variable_definitions.to_csv(variable_def_path, index=False, encoding="utf-8-sig")

display(ccc7_sector_classification.head())
display(emissions_variable_definitions.head())


# %% [markdown]
# ## 6. Validate CCC7 sector sums against economy-wide totals

# %%
sector_sums = (
    ccc7_sector_clean
    .groupby(["scenario", "country", "variable", "variable_unit", "year"], as_index=False)
    ["emissions_MtCO2e"].sum()
    .rename(columns={"emissions_MtCO2e": "sector_sum_MtCO2e"})
)

economy_for_validation = ccc7_economy_clean.rename(
    columns={"emissions_MtCO2e": "economy_wide_MtCO2e"}
)

sector_sum_validation = sector_sums.merge(
    economy_for_validation[
        ["scenario", "country", "variable", "variable_unit", "year", "economy_wide_MtCO2e"]
    ],
    on=["scenario", "country", "variable", "variable_unit", "year"],
    how="left",
)
sector_sum_validation["sector_minus_economy_MtCO2e"] = (
    sector_sum_validation["sector_sum_MtCO2e"]
    - sector_sum_validation["economy_wide_MtCO2e"]
)

sector_sum_validation_path = TABLES_DIR / "ccc7_sector_sum_validation.csv"
sector_sum_validation.to_csv(sector_sum_validation_path, index=False, encoding="utf-8-sig")

display(sector_sum_validation.head())
display(sector_sum_validation.tail())


# %% [markdown]
# ## 7. Clean CCC6 Balanced Net Zero Pathway sensitivity benchmark
#
# CCC6 is retained only as a sensitivity check. The main benchmark remains CCC7.

# %%
ccc6_raw = pd.read_excel(CCC6_DATA, sheet_name="Scenario key metrics", header=None)


def extract_ccc6_uk_balanced_net_zero_pathway(df: pd.DataFrame) -> pd.DataFrame:
    """Extract UK total emissions for the CCC6 Balanced Net Zero Pathway."""
    header_idx = None
    for idx, row in df.iterrows():
        numeric_values = pd.to_numeric(row, errors="coerce").dropna().astype(int).tolist()
        if 2020 in numeric_values and 2050 in numeric_values:
            header_idx = idx
            break
    if header_idx is None:
        raise ValueError("Could not find CCC6 year header row.")

    header = df.loc[header_idx].map(clean_text)

    def find_header_col(label: str) -> int:
        matches = [col for col, value in header.items() if value == label]
        if not matches:
            raise ValueError(f"Could not find CCC6 header column {label!r}.")
        return matches[0]

    scenario_col = find_header_col("Scenario")
    category_col = find_header_col("Category")
    element_col = find_header_col("Element")

    start_col = next(
        col for col, value in df.loc[header_idx].items()
        if pd.to_numeric(value, errors="coerce") == 2020
    )
    year_cols = []
    for col in range(start_col, len(df.columns)):
        year = pd.to_numeric(df.loc[header_idx, col], errors="coerce")
        if pd.notna(year) and 2020 <= int(year) <= 2050:
            year_cols.append((col, int(year)))
        elif year_cols:
            break

    scenario = df[scenario_col].map(clean_text)
    category = df[category_col].map(clean_text)
    element = df[element_col].map(clean_text)
    mask = (
        scenario.eq("Balanced Net Zero Pathway")
        & category.str.contains("Scenario emissions", regex=False, na=False)
        & element.str.contains("Total emissions", regex=False, na=False)
    )
    if mask.sum() < 1:
        preview = df.loc[:, [scenario_col, category_col, element_col]].dropna(how="all").head(20)
        raise ValueError(
            "Could not find CCC6 Balanced Net Zero Pathway total emissions row. "
            f"Preview:\n{preview.to_string(index=False)}"
        )

    row = df.loc[mask].iloc[0]
    records = []
    for col, year in year_cols:
        records.append(
            {
                "scenario": "Balanced Net Zero Pathway",
                "country": "United Kingdom",
                "variable": "Scenario emissions: Total emissions",
                "variable_unit": "MtCO2e",
                "year": year,
                "emissions_MtCO2e": pd.to_numeric(row[col], errors="coerce"),
            }
        )
    out = pd.DataFrame(records)
    out = out.dropna(subset=["emissions_MtCO2e"]).sort_values("year").reset_index(drop=True)
    return out


ccc6_balanced_clean = extract_ccc6_uk_balanced_net_zero_pathway(ccc6_raw)
ccc6_balanced_path = DATA_PROCESSED / "ccc6_clean_balanced_net_zero_pathway.csv"
ccc6_balanced_clean.to_csv(ccc6_balanced_path, index=False, encoding="utf-8-sig")

display(ccc6_balanced_clean.head())
display(ccc6_balanced_clean.tail())


# %% [markdown]
# ## 8. Summary tables for P5 use

# %%
benchmark_years = [2025, 2030, 2035, 2040, 2045, 2050]
ccc7_benchmark_years = ccc7_economy_balanced[
    ccc7_economy_balanced["year"].isin(benchmark_years)
].copy()
ccc7_benchmark_years = ccc7_benchmark_years[
    ["scenario", "country", "variable", "variable_unit", "year", "emissions_MtCO2e"]
]
ccc7_benchmark_years_path = TABLES_DIR / "ccc7_clean_benchmark_years.csv"
ccc7_benchmark_years.to_csv(ccc7_benchmark_years_path, index=False, encoding="utf-8-sig")

ccc7_sector_2050 = (
    ccc7_sector_balanced.loc[ccc7_sector_balanced["year"] == 2050]
    .sort_values("emissions_MtCO2e", ascending=False)
    .reset_index(drop=True)
)
ccc7_sector_2050_path = TABLES_DIR / "ccc7_clean_sector_2050_summary.csv"
ccc7_sector_2050.to_csv(ccc7_sector_2050_path, index=False, encoding="utf-8-sig")

display(ccc7_benchmark_years)
display(ccc7_sector_2050)


# %% [markdown]
# ## 9. Quality checks

# %%
checks = []

ok, details = check_year_coverage(ccc7_economy_clean, "Balanced Pathway", 2025, 2050)
checks.append(
    {
        "check": "CCC7 economy-wide Balanced Pathway has complete 2025-2050 coverage",
        "status": "PASS" if ok else "FAIL",
        "details": details,
    }
)

balanced_duplicates = ccc7_economy_balanced.duplicated(subset=["scenario", "year"]).sum()
checks.append(
    {
        "check": "CCC7 economy-wide Balanced Pathway has no duplicate years",
        "status": "PASS" if balanced_duplicates == 0 else "FAIL",
        "details": f"duplicates={balanced_duplicates}",
    }
)

sector_counts = ccc7_sector_balanced.groupby("year")["sector"].nunique()
sector_count_ok = sector_counts.min() == sector_counts.max() == 13
checks.append(
    {
        "check": "CCC7 sector-level Balanced Pathway has 13 sectors in every year",
        "status": "PASS" if sector_count_ok else "FAIL",
        "details": f"min_sectors={sector_counts.min()}; max_sectors={sector_counts.max()}",
    }
)

max_abs_sector_diff = sector_sum_validation["sector_minus_economy_MtCO2e"].abs().max()
checks.append(
    {
        "check": "CCC7 sector sums match economy-wide totals",
        "status": "PASS" if max_abs_sector_diff < 1e-8 else "WARN",
        "details": f"max_abs_difference_MtCO2e={max_abs_sector_diff:.12f}",
    }
)

ccc6_years = set(ccc6_balanced_clean["year"].astype(int))
ccc6_missing_p5_years = sorted(expected_years(2025, 2050).difference(ccc6_years))
checks.append(
    {
        "check": "CCC6 sensitivity series covers 2025-2050",
        "status": "PASS" if not ccc6_missing_p5_years else "FAIL",
        "details": f"missing={ccc6_missing_p5_years}",
    }
)

quality_checks = pd.DataFrame(checks)
quality_checks_path = TABLES_DIR / "p5_ccc_data_cleaning_quality_checks.csv"
quality_checks.to_csv(quality_checks_path, index=False, encoding="utf-8-sig")

display(quality_checks)

if (quality_checks["status"] == "FAIL").any():
    raise ValueError("One or more CCC data cleaning quality checks failed.")


# %% [markdown]
# ## 10. Cleaning note

# %%
note = f"""P5 CCC data cleaning note

Purpose
This note documents the cleaned CCC datasets used for P5 benchmark comparison.

Source files
1. CCC7 main source:
{CCC7_DATA}

2. CCC6 sensitivity source:
{CCC6_DATA}

CCC7 economy-wide cleaning
- Sheet: Economy-wide data.
- Rows retained: country = United Kingdom; variable = Emissions: direct emissions total; unit = MtCO2e.
- Main scenario retained for P5: Balanced Pathway.
- Years retained: 2025-2050.
- Output files:
  - {ccc7_economy_all_path}
  - {ccc7_economy_balanced_path}

CCC7 sector-level cleaning
- Sheet: Sector-level data.
- Rows retained: country = United Kingdom; variable = Emissions: direct emissions total; unit = MtCO2e.
- Main scenario retained for P5: Balanced Pathway.
- Sectors include Agriculture, Aviation, Electricity supply, Engineered removals, F-gases, Fuel supply, Industry, Land use, Non-residential buildings, Residential buildings, Shipping, Surface transport, and Waste.
- Output files:
  - {ccc7_sector_all_path}
  - {ccc7_sector_balanced_path}
  - {ccc7_sector_2050_path}

CCC7 metadata cleaning
- Sector classification output: {sector_class_path}
- Emissions variable definitions output: {variable_def_path}

CCC6 sensitivity cleaning
- Sheet: Scenario key metrics.
- Rows retained: Balanced Net Zero Pathway; Scenario emissions; Total emissions.
- Role: sensitivity check only, not the main P5 benchmark.
- Output file: {ccc6_balanced_path}

Quality checks
- CCC7 Balanced Pathway economy-wide series covers 2025-2050.
- CCC7 Balanced Pathway has no duplicate economy-wide years.
- CCC7 sector-level Balanced Pathway has 13 sectors in every year.
- CCC7 sector sums match economy-wide totals within rounding tolerance.
- CCC6 sensitivity series covers 2025-2050.
- Quality check file: {quality_checks_path}

Recommended P5 use
Use CCC7 Balanced Pathway economy-wide direct emissions as the main target-consistent benchmark. Use CCC6 only as a sensitivity check. Use CCC7 sector-level data for broad sector context and sector-alignment discussion, not for strict one-to-one attribution against DESNZ TES sectors.
"""

note_path = TABLES_DIR / "p5_ccc_data_cleaning_note.txt"
note_path.write_text(note, encoding="utf-8")

print(note)


# %% [markdown]
# ## 11. Output manifest

# %%
output_manifest = pd.DataFrame(
    [
        {"output": str(ccc7_economy_all_path), "role": "Clean CCC7 economy-wide direct emissions, all scenarios"},
        {"output": str(ccc7_economy_balanced_path), "role": "Clean CCC7 Balanced Pathway economy-wide direct emissions"},
        {"output": str(ccc7_sector_all_path), "role": "Clean CCC7 sector-level direct emissions, all scenarios"},
        {"output": str(ccc7_sector_balanced_path), "role": "Clean CCC7 Balanced Pathway sector-level direct emissions"},
        {"output": str(sector_class_path), "role": "Clean CCC7 sector classification metadata"},
        {"output": str(variable_def_path), "role": "Clean CCC7 emissions variable definitions"},
        {"output": str(sector_sum_validation_path), "role": "Validation of CCC7 sector sums against economy-wide totals"},
        {"output": str(ccc6_balanced_path), "role": "Clean CCC6 Balanced Net Zero sensitivity pathway"},
        {"output": str(ccc7_benchmark_years_path), "role": "CCC7 benchmark-year values for P5"},
        {"output": str(ccc7_sector_2050_path), "role": "CCC7 2050 sector summary for P5"},
        {"output": str(quality_checks_path), "role": "P5 CCC data cleaning quality checks"},
        {"output": str(note_path), "role": "P5 CCC data cleaning note"},
    ]
)
manifest_path = TABLES_DIR / "p5_ccc_data_cleaning_output_manifest.csv"
output_manifest.to_csv(manifest_path, index=False, encoding="utf-8-sig")
display(output_manifest)

print("CCC data cleaning complete.")
