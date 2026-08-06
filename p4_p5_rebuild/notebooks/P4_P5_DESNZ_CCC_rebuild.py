# %% [markdown]
# # P4-P5 Rebuild: DESNZ Baseline and CCC Benchmark Comparison
# 
# This notebook reproduces the P4-P5 workflow from raw data files. It is designed to be auditable: every headline number is calculated from the source workbooks, not copied from an AI summary.
# 
# Sections:
# 
# 1. Setup and file discovery
# 2. P4 DESNZ annual baseline, excluding and including IAS
# 3. P4 official carbon-budget-period gap metrics
# 4. P5 CCC Seventh and Sixth Carbon Budget benchmark extraction
# 5. DESNZ vs CCC annual gap calculation
# 6. Tables, figures, and consistency checks
# 
# Before running, make sure these packages are installed in the Jupyter environment: `pandas`, `openpyxl`, `odfpy`, `matplotlib`, `numpy`.

# %%
# Optional, run only if your Jupyter environment is missing dependencies:
# %pip install pandas openpyxl odfpy matplotlib numpy

# %%
from pathlib import Path
import re
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

plt.rcParams.update({
    "figure.dpi": 130,
    "savefig.dpi": 300,
    "font.family": "Arial",
    "axes.spines.top": False,
    "axes.spines.right": False,
})

PROJECT_ROOT = Path(r"E:\UCL Final Essay")
RAW = PROJECT_ROOT / "Data_raw"
OUTPUT = PROJECT_ROOT / "p4_p5_rebuild_python"

if not RAW.exists():
    raise FileNotFoundError(
        f"Cannot find Data_raw at {RAW}. Update PROJECT_ROOT to your dissertation folder."
    )

for folder in [
    OUTPUT,
    OUTPUT / "data_inventory",
    OUTPUT / "data_processed",
    OUTPUT / "tables",
    OUTPUT / "figures",
    OUTPUT / "notes",
]:
    folder.mkdir(parents=True, exist_ok=True)

print("Project root:", PROJECT_ROOT)
print("Raw data folder:", RAW)
print("Output folder:", OUTPUT)

# %% [markdown]
# ## Helper Functions
# 
# These functions keep the notebook readable. They do not contain any hidden assumptions: they only clean text labels, find files, extract year columns, and save tables.

# %%
def clean_text(value):
    """Normalise labels from Excel/ODS cells."""
    if pd.isna(value):
        return ""
    text = str(value)
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.replace("<text:s/>", " ")
    return re.sub(r"\s+", " ", text).strip()


def find_data_file(filename, preferred_contains=None):
    """Find a raw data file, preferring a folder name when duplicate copies exist."""
    matches = sorted(RAW.rglob(filename), key=lambda p: len(str(p)))
    if preferred_contains:
        preferred = [p for p in matches if preferred_contains.lower() in str(p).lower()]
        if preferred:
            return preferred[0]
    if not matches:
        raise FileNotFoundError(f"Could not find {filename} under {RAW}")
    return matches[0]


def read_ods(path, sheet_name, header=None):
    """Read an ODS sheet using pandas' odf engine."""
    try:
        return pd.read_excel(path, sheet_name=sheet_name, header=header, engine="odf")
    except ImportError as exc:
        raise ImportError("Reading .ods files requires odfpy. Run: %pip install odfpy") from exc


def year_columns(df, first_year=1990, last_year=2050):
    """Return a mapping {year: column_name} for year-labelled columns."""
    out = {}
    for col in df.columns:
        try:
            year = int(float(col))
        except (TypeError, ValueError):
            continue
        if first_year <= year <= last_year:
            out[year] = col
    return out


def extract_series_by_label_and_coverage(df, label, coverage):
    """Extract a wide annual series from DESNZ Annex A style tables."""
    label_col, coverage_col = df.columns[0], df.columns[1]
    mask = (
        df[label_col].map(clean_text).eq(label)
        & df[coverage_col].map(clean_text).eq(coverage)
    )
    if mask.sum() != 1:
        matches = df.loc[mask, [label_col, coverage_col]].to_string(index=False)
        raise ValueError(f"Expected one row for {label!r} / {coverage!r}, found {mask.sum()}
{matches}")
    row = df.loc[mask].iloc[0]
    cols = year_columns(df)
    series = pd.Series({year: pd.to_numeric(row[col], errors="coerce") for year, col in cols.items()})
    series.index.name = "year"
    series.name = coverage
    return series.sort_index()


def save_csv(df, path):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False)
    print(f"Saved {path}")


def get_value(series, year):
    return float(series.loc[year])


def period_sum(series, years):
    return float(series.loc[list(years)].sum())

# %% [markdown]
# ## 1. Locate Source Files
# 
# The file names are searched under `Data_raw`. This avoids hard-coding every folder name, while still preferring the official DESNZ/CCC folders when duplicate copies exist.

# %%
DESNZ_ANNEX_TES = find_data_file(
    "Annex_A_GHG_by_TES_sector.ods",
    preferred_contains="DESNZ Energy and Emissions Projections",
)
DESNZ_WEB_FIGURES = find_data_file(
    "Web_figures_EEP_2024_2050.ods",
    preferred_contains="DESNZ Energy and Emissions Projections",
)
DESNZ_WEB_TABLES = find_data_file(
    "Web_tables_EEP_2024_2050.ods",
    preferred_contains="DESNZ Energy and Emissions Projections",
)
CCC7_DATA = find_data_file(
    "The-Seventh-Carbon-Budget-full-dataset.xlsx",
    preferred_contains="CCC Seventh Carbon Budget",
)
CCC6_DATA = find_data_file(
    "The-Sixth-Carbon-Budget-Dataset_v2.xlsx",
    preferred_contains="CCC Sixth Carbon Budget",
)

source_files = pd.DataFrame([
    ["DESNZ Annex A TES sectors", DESNZ_ANNEX_TES],
    ["DESNZ web figures", DESNZ_WEB_FIGURES],
    ["DESNZ web tables", DESNZ_WEB_TABLES],
    ["CCC Seventh Carbon Budget dataset", CCC7_DATA],
    ["CCC Sixth Carbon Budget dataset", CCC6_DATA],
], columns=["role", "path"])
source_files

# %% [markdown]
# ## 2. P4 DESNZ Annual Baseline
# 
# DESNZ Annex A contains both total emissions excluding IAS and including IAS. We keep both:
# 
# - **excluding IAS**: useful for a consistent territorial historical/projection series.
# - **including IAS**: more appropriate for comparing with CCC economy-wide total pathway data.

# %%
tes = read_ods(DESNZ_ANNEX_TES, sheet_name="Reference", header=2)

# Show the relevant top rows so the source labels are visible.
tes[[tes.columns[0], tes.columns[1], tes.columns[2]]].head(8)

# %%
desnz_excl_ias = extract_series_by_label_and_coverage(
    tes,
    label="GHG (All)",
    coverage="Total emissions (exc. IAS)",
)
desnz_inc_ias = extract_series_by_label_and_coverage(
    tes,
    label="GHG (All)",
    coverage="Total emissions (inc. IAS)",
)

desnz_annual = pd.DataFrame({
    "year": desnz_excl_ias.index,
    "DESNZ_EEP_2024_excl_IAS_MtCO2e": desnz_excl_ias.values,
    "DESNZ_EEP_2024_inc_IAS_MtCO2e": desnz_inc_ias.reindex(desnz_excl_ias.index).values,
})

desnz_annual.tail()

# %% [markdown]
# ### Sanity Check Against DESNZ Web Figures
# 
# DESNZ web figures provide the annual territorial emissions pathway used in the published figure. It should match the Annex A **excluding IAS** total. This check helps catch accidental row or accounting-scope errors.

# %%
web_fig = read_ods(DESNZ_WEB_FIGURES, sheet_name="Fig__i_and_2_1", header=1)
label_col = web_fig.columns[0]
key_col = web_fig.columns[1]

mask = (
    web_fig[label_col].map(clean_text).eq("EEP 2024-2050")
    & web_fig[key_col].map(clean_text).eq("Territorial emissions")
)
if mask.sum() != 1:
    raise ValueError(f"Expected one DESNZ web-figure row, found {mask.sum()}")

web_row = web_fig.loc[mask].iloc[0]
web_year_cols = year_columns(web_fig)
web_series = pd.Series({year: pd.to_numeric(web_row[col], errors="coerce") for year, col in web_year_cols.items()}).sort_index()

common_years = sorted(set(web_series.index).intersection(desnz_excl_ias.index))
max_abs_diff = (web_series.loc[common_years] - desnz_excl_ias.loc[common_years]).abs().max()
print("Max absolute difference between Web figures and Annex A excluding-IAS total:", max_abs_diff)
assert max_abs_diff < 1e-6

# %% [markdown]
# ## 3. P4 Official Carbon-Budget-Period Gap Metrics
# 
# DESNZ Web Table 2.1 gives the official carbon-budget-period comparison. The key accounting rule is:
# 
# - CB4 and CB5 official comparisons exclude IAS.
# - CB6 official comparison is on the Net Carbon Account / including-IAS basis.
# 
# So the notebook records both the raw excluding/including IAS totals and the official comparison basis.

# %%
table21 = read_ods(DESNZ_WEB_TABLES, sheet_name="Table_2_1", header=None)
table21.head(12)

# %%
def find_first_row_contains(df, text, column=0):
    labels = df[column].map(clean_text)
    mask = labels.str.contains(text, regex=False, na=False)
    if mask.sum() < 1:
        raise ValueError(f"Could not find row containing {text!r}")
    return df.loc[mask].iloc[0]

header_row_idx = table21.index[
    table21.apply(lambda row: row.map(clean_text).str.contains("CB4", regex=False).any(), axis=1)
][0]
header_row = table21.loc[header_row_idx]
period_cols = {
    clean_text(value): col
    for col, value in header_row.items()
    if clean_text(value).startswith("CB")
}

carbon_budget_target = find_first_row_contains(table21, "Carbon budget target")
territorial_excl = find_first_row_contains(table21, "Territorial emissions exc. IAS")
territorial_inc = find_first_row_contains(table21, "Territorial emissions inc. IAS")
official_gap = find_first_row_contains(table21, "Projected performance vs target")

period_years = {
    "CB4 (2023-27)": list(range(2023, 2028)),
    "CB5 (2028-32)": list(range(2028, 2033)),
    "CB6 (2033-37)": list(range(2033, 2038)),
}
official_basis = {
    "CB4 (2023-27)": "NCA / excluding IAS",
    "CB5 (2028-32)": "NCA / excluding IAS",
    "CB6 (2033-37)": "NCA / including IAS",
}

cb_rows = []
for period, col in period_cols.items():
    target = float(carbon_budget_target[col])
    excl = float(territorial_excl[col])
    inc = float(territorial_inc[col])
    gap = float(official_gap[col])
    cb_rows.append({
        "period": period.split()[0],
        "years": "-".join(map(str, period_years[period])),
        "carbon_budget_target_MtCO2e": target,
        "DESNZ_excl_IAS_MtCO2e": excl,
        "DESNZ_inc_IAS_MtCO2e": inc,
        "official_comparison_basis": official_basis[period],
        "DESNZ_official_comparison_emissions_MtCO2e": target + gap,
        "DESNZ_official_gap_MtCO2e": gap,
    })

carbon_budget_metrics = pd.DataFrame(cb_rows)
carbon_budget_metrics

# %% [markdown]
# ## 4. P5 CCC Benchmark Extraction
# 
# The main benchmark is CCC Seventh Carbon Budget **Balanced Pathway**, economy-wide data, variable `Emissions: direct emissions total`.
# 
# The CCC Sixth Carbon Budget **Balanced Net Zero Pathway** is extracted as a secondary comparison.

# %%
ccc7 = pd.read_excel(CCC7_DATA, sheet_name="Economy-wide data")
ccc7.head()

# %%
ccc7_balanced = (
    ccc7[
        (ccc7["scenario"].map(clean_text) == "Balanced Pathway")
        & (ccc7["country"].map(clean_text) == "United Kingdom")
        & (ccc7["variable"].map(clean_text) == "Emissions: direct emissions total")
    ]
    .assign(year=lambda d: pd.to_numeric(d["year"], errors="coerce"),
            value=lambda d: pd.to_numeric(d["value"], errors="coerce"))
    .dropna(subset=["year", "value"])
    .sort_values("year")
)

ccc7_series = ccc7_balanced.set_index("year")["value"]
ccc7_series.name = "CCC7_Balanced_Pathway_MtCO2e"
ccc7_series.loc[[2025, 2030, 2035, 2040, 2045, 2050]]

# %%
ccc6_raw = pd.read_excel(CCC6_DATA, sheet_name="Scenario key metrics", header=None)
ccc6_raw.head(10)

# %%
def extract_ccc6_uk_balanced_net_zero_pathway(df):
    """Extract the UK 2020-2050 total-emissions row from CCC6 Scenario key metrics."""
    header_idx = None
    for idx, row in df.iterrows():
        numeric_values = pd.to_numeric(row, errors="coerce").dropna().astype(int).tolist()
        if 2020 in numeric_values and 2050 in numeric_values:
            header_idx = idx
            break
    if header_idx is None:
        raise ValueError("Could not find CCC6 year header row")

    header = df.loc[header_idx]
    start_col = next(col for col, value in header.items() if pd.to_numeric(value, errors="coerce") == 2020)
    year_cols = []
    for col in range(start_col, len(header)):
        year = pd.to_numeric(header[col], errors="coerce")
        if pd.notna(year) and 2020 <= int(year) <= 2050:
            year_cols.append((col, int(year)))
        elif year_cols:
            break

    scenario = df.iloc[:, 0].map(clean_text)
    category = df.iloc[:, 2].map(clean_text)
    element = df.iloc[:, 3].map(clean_text)
    mask = (
        scenario.eq("Balanced Net Zero Pathway")
        & category.str.contains("Scenario emissions", regex=False, na=False)
        & element.str.contains("Total emissions", regex=False, na=False)
    )
    if mask.sum() < 1:
        raise ValueError("Could not find CCC6 Balanced Net Zero Pathway total emissions row")

    row = df.loc[mask].iloc[0]
    values = {year: pd.to_numeric(row[col], errors="coerce") for col, year in year_cols}
    series = pd.Series(values).sort_index()
    series.index.name = "year"
    series.name = "CCC6_Balanced_Net_Zero_Pathway_MtCO2e"
    return series

ccc6_series = extract_ccc6_uk_balanced_net_zero_pathway(ccc6_raw)
ccc6_series.loc[[2025, 2030, 2035, 2040, 2045, 2050]]

# %% [markdown]
# ## 5. DESNZ vs CCC Annual Gap Calculation
# 
# The main P5 comparison uses DESNZ **including IAS** minus CCC7 Balanced Pathway.
# 
# The excluding-IAS gap is retained as an accounting-scope sensitivity, not as the primary benchmark gap.

# %%
years = pd.Index(range(2025, 2051), name="year")
annual_comparison = pd.DataFrame({
    "year": years,
    "DESNZ_EEP_2024_excl_IAS_MtCO2e": desnz_excl_ias.reindex(years).values,
    "DESNZ_EEP_2024_inc_IAS_MtCO2e": desnz_inc_ias.reindex(years).values,
    "CCC7_Balanced_Pathway_MtCO2e": ccc7_series.reindex(years).values,
    "CCC6_Balanced_Net_Zero_Pathway_MtCO2e": ccc6_series.reindex(years).values,
})

annual_comparison["gap_inc_IAS_DESNZ_minus_CCC7_MtCO2e"] = (
    annual_comparison["DESNZ_EEP_2024_inc_IAS_MtCO2e"]
    - annual_comparison["CCC7_Balanced_Pathway_MtCO2e"]
)
annual_comparison["gap_inc_IAS_DESNZ_minus_CCC6_MtCO2e"] = (
    annual_comparison["DESNZ_EEP_2024_inc_IAS_MtCO2e"]
    - annual_comparison["CCC6_Balanced_Net_Zero_Pathway_MtCO2e"]
)
annual_comparison["gap_excl_IAS_DESNZ_minus_CCC7_MtCO2e"] = (
    annual_comparison["DESNZ_EEP_2024_excl_IAS_MtCO2e"]
    - annual_comparison["CCC7_Balanced_Pathway_MtCO2e"]
)

annual_comparison.head()

# %%
benchmark_years = [2030, 2035, 2040, 2045, 2050]
key_benchmark_years = annual_comparison[annual_comparison["year"].isin(benchmark_years)].copy()
key_benchmark_years

# %% [markdown]
# ## 6. Carbon-Budget-Period Benchmark Additions
# 
# CCC7 starts in 2025, so it cannot provide a full CB4 (2023-2027) period sum. CB5 and CB6 can be summed from CCC7.

# %%
def safe_sum_for_years(series, years):
    years = list(years)
    if not set(years).issubset(set(series.index.astype(int))):
        return np.nan
    return float(series.loc[years].sum())

period_lookup = {
    "CB4": list(range(2023, 2028)),
    "CB5": list(range(2028, 2033)),
    "CB6": list(range(2033, 2038)),
}

carbon_budget_metrics["CCC7_balanced_sum_MtCO2e"] = carbon_budget_metrics["period"].map(
    lambda p: safe_sum_for_years(ccc7_series, period_lookup[p])
)
carbon_budget_metrics["CCC7_minus_target_MtCO2e"] = (
    carbon_budget_metrics["CCC7_balanced_sum_MtCO2e"]
    - carbon_budget_metrics["carbon_budget_target_MtCO2e"]
)
carbon_budget_metrics["CCC6_balanced_sum_MtCO2e"] = carbon_budget_metrics["period"].map(
    lambda p: safe_sum_for_years(ccc6_series, period_lookup[p])
)
carbon_budget_metrics["comparability_note"] = np.where(
    carbon_budget_metrics["CCC7_balanced_sum_MtCO2e"].isna(),
    "CCC7 starts in 2025; full CB4 comparison not available from annual CCC7 series.",
    "CCC7 covers full period; compare cautiously because pathway/accounting scope may differ from statutory budget basis.",
)

carbon_budget_metrics

# %% [markdown]
# ## 7. Save Clean Tables

# %%
dataset_inventory = pd.DataFrame([
    {
        "dataset_role": "DESNZ annual pathway",
        "source_file": str(DESNZ_ANNEX_TES.relative_to(PROJECT_ROOT)),
        "sheet": "Reference",
        "rows_used": "GHG (All); Total emissions (exc. IAS) and Total emissions (inc. IAS)",
        "accounting_basis": "Territorial emissions excluding and including IAS",
        "notes": "Primary annual DESNZ baseline, 1990-2050; inc-IAS version used for CCC annual comparison.",
    },
    {
        "dataset_role": "DESNZ official carbon budget gap",
        "source_file": str(DESNZ_WEB_TABLES.relative_to(PROJECT_ROOT)),
        "sheet": "Table_2_1",
        "rows_used": "Carbon budget target; projected territorial and NCA emissions; projected performance vs target",
        "accounting_basis": "Official NCA basis; CB4-CB5 exclude IAS, CB6 includes IAS",
        "notes": "Used for official CB4-CB6 policy gap metrics.",
    },
    {
        "dataset_role": "CCC Seventh Carbon Budget benchmark",
        "source_file": str(CCC7_DATA.relative_to(PROJECT_ROOT)),
        "sheet": "Economy-wide data",
        "rows_used": "Balanced Pathway; United Kingdom; Emissions: direct emissions total",
        "accounting_basis": "CCC economy-wide total pathway emissions",
        "notes": "Primary target-consistent benchmark for annual comparison, 2025-2050.",
    },
    {
        "dataset_role": "CCC Sixth Carbon Budget benchmark",
        "source_file": str(CCC6_DATA.relative_to(PROJECT_ROOT)),
        "sheet": "Scenario key metrics",
        "rows_used": "Balanced Net Zero Pathway; Scenario emissions; Total emissions; UK columns",
        "accounting_basis": "CCC scenario total emissions",
        "notes": "Older benchmark / sensitivity comparison, 2020-2050.",
    },
])

save_csv(dataset_inventory, OUTPUT / "data_inventory" / "p4_p5_dataset_inventory.csv")
save_csv(desnz_annual, OUTPUT / "data_processed" / "desnz_annual_territorial_pathway.csv")
save_csv(annual_comparison, OUTPUT / "data_processed" / "annual_desnz_ccc_comparison.csv")
save_csv(carbon_budget_metrics, OUTPUT / "tables" / "carbon_budget_gap_metrics.csv")
save_csv(key_benchmark_years, OUTPUT / "tables" / "key_benchmark_year_gap_metrics.csv")

# %% [markdown]
# ## 8. Generate Figures

# %%
fig, ax = plt.subplots(figsize=(9.5, 5.6))

ax.plot(
    annual_comparison["year"],
    annual_comparison["DESNZ_EEP_2024_inc_IAS_MtCO2e"],
    label="DESNZ EEP 2024 including IAS",
    linewidth=2.6,
    color="#005f73",
)
ax.plot(
    annual_comparison["year"],
    annual_comparison["DESNZ_EEP_2024_excl_IAS_MtCO2e"],
    label="DESNZ EEP 2024 excluding IAS",
    linewidth=2.0,
    linestyle="--",
    color="#64748b",
)
ax.plot(
    annual_comparison["year"],
    annual_comparison["CCC7_Balanced_Pathway_MtCO2e"],
    label="CCC7 Balanced Pathway",
    linewidth=2.6,
    color="#ca6702",
)
ax.plot(
    annual_comparison["year"],
    annual_comparison["CCC6_Balanced_Net_Zero_Pathway_MtCO2e"],
    label="CCC6 Balanced Net Zero Pathway",
    linewidth=2.2,
    color="#6d597a",
)

ax.axhline(0, color="#111827", linewidth=0.8)
ax.set_title("DESNZ current-policy baseline vs CCC target-consistent pathways")
ax.set_ylabel("MtCO2e")
ax.set_xlabel("Year")
ax.set_xlim(2025, 2050)
ax.grid(axis="y", alpha=0.25)
ax.legend(frameon=False, loc="upper right")
ax.text(
    2025, -55,
    "Source: DESNZ EEP Annex A TES sectors; CCC Seventh Carbon Budget full dataset; CCC Sixth Carbon Budget dataset.",
    fontsize=8,
    color="#475569",
)
fig.tight_layout()

fig.savefig(OUTPUT / "figures" / "desnz_vs_ccc_annual_pathways.png")
fig.savefig(OUTPUT / "figures" / "desnz_vs_ccc_annual_pathways.svg")
plt.show()

# %%
plot_df = carbon_budget_metrics.copy()
x = np.arange(len(plot_df))
width = 0.35

fig, ax = plt.subplots(figsize=(8.5, 5.2))
ax.bar(
    x - width / 2,
    plot_df["carbon_budget_target_MtCO2e"],
    width,
    label="Carbon budget target",
    color="#94a3b8",
)
ax.bar(
    x + width / 2,
    plot_df["DESNZ_official_comparison_emissions_MtCO2e"],
    width,
    label="DESNZ official comparison emissions",
    color="#005f73",
)

for i, row in plot_df.iterrows():
    gap = row["DESNZ_official_gap_MtCO2e"]
    label = f"gap {gap:+.0f}"
    ax.text(i, max(row["carbon_budget_target_MtCO2e"], row["DESNZ_official_comparison_emissions_MtCO2e"]) + 45, label,
            ha="center", va="bottom", fontsize=9, fontweight="bold",
            color="#b42318" if gap > 0 else "#067647")

ax.set_xticks(x)
ax.set_xticklabels(plot_df["period"])
ax.set_ylabel("MtCO2e over budget period")
ax.set_title("Official carbon budget performance under DESNZ EEP 2024")
ax.grid(axis="y", alpha=0.25)
ax.legend(frameon=False)
ax.text(
    -0.45, -230,
    "Note: CB4-CB5 official comparison excludes IAS; CB6 uses NCA / including-IAS basis. Source: DESNZ EEP Web Table 2.1.",
    fontsize=8,
    color="#475569",
)
fig.tight_layout()

fig.savefig(OUTPUT / "figures" / "carbon_budget_period_comparison.png")
fig.savefig(OUTPUT / "figures" / "carbon_budget_period_comparison.svg")
plt.show()

# %% [markdown]
# ## 9. Consistency Checks
# 
# These checks are deliberately explicit. If a future data update changes the source files, these assertions will flag that the headline values have changed.

# %%
headline_2050_gap = float(
    key_benchmark_years.loc[
        key_benchmark_years["year"] == 2050,
        "gap_inc_IAS_DESNZ_minus_CCC7_MtCO2e",
    ].iloc[0]
)
cb6_official_gap = float(
    carbon_budget_metrics.loc[
        carbon_budget_metrics["period"] == "CB6",
        "DESNZ_official_gap_MtCO2e",
    ].iloc[0]
)

print(f"2050 DESNZ including-IAS vs CCC7 gap: {headline_2050_gap:.1f} MtCO2e")
print(f"CB6 official DESNZ projected gap: {cb6_official_gap:.1f} MtCO2e")

assert abs(headline_2050_gap - 325.398436382106) < 1e-6
assert abs(cb6_official_gap - 737.469408760892) < 1e-6

print("All checks passed.")

# %% [markdown]
# ## 10. Headline Results for Writing
# 
# Use these printed numbers when drafting P4-P5, but keep the generated CSVs as the auditable source.

# %%
summary = {
    "DESNZ 2050 excluding IAS (MtCO2e)": get_value(desnz_excl_ias, 2050),
    "DESNZ 2050 including IAS (MtCO2e)": get_value(desnz_inc_ias, 2050),
    "CCC7 2050 Balanced Pathway (MtCO2e)": float(ccc7_series.loc[2050]),
    "2050 gap, DESNZ inc IAS minus CCC7 (MtCO2e)": headline_2050_gap,
    "CB6 official DESNZ projected gap (MtCO2e)": cb6_official_gap,
    "CB6 DESNZ excluding-IAS gap (MtCO2e)": float(
        carbon_budget_metrics.loc[carbon_budget_metrics["period"] == "CB6", "DESNZ_excl_IAS_MtCO2e"].iloc[0]
        - carbon_budget_metrics.loc[carbon_budget_metrics["period"] == "CB6", "carbon_budget_target_MtCO2e"].iloc[0]
    ),
}

pd.Series(summary).round(1)