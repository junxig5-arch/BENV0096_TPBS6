# P4-P5 Source Code Index

This index maps each dissertation workflow part to the generated Jupyter notebook sections.

## Main files

- `P4_P5_DESNZ_CCC_rebuild.ipynb`: Jupyter notebook to run in order.
- `P4_P5_DESNZ_CCC_rebuild.py`: the same code in `# %%` cell format for review or VS Code/Jupyter import.

## Section mapping

| Workflow part | Notebook section | Main raw source | Main outputs |
|---|---|---|---|
| Setup and file discovery | 1. Locate Source Files | `Data_raw` recursive file search | Printed source file table |
| P4 DESNZ annual baseline | 2. P4 DESNZ Annual Baseline | DESNZ `Annex_A_GHG_by_TES_sector.ods`, sheet `Reference` | `data_processed/desnz_annual_territorial_pathway.csv` |
| DESNZ sanity check | 2. Sanity Check Against DESNZ Web Figures | DESNZ `Web_figures_EEP_2024_2050.ods`, sheet `Fig__i_and_2_1` | Assertion that Annex A excluding-IAS matches web figure |
| P4 official carbon budget gap | 3. P4 Official Carbon-Budget-Period Gap Metrics | DESNZ `Web_tables_EEP_2024_2050.ods`, sheet `Table_2_1` | `tables/carbon_budget_gap_metrics.csv` |
| P5 CCC7 benchmark | 4. P5 CCC Benchmark Extraction | CCC7 `The-Seventh-Carbon-Budget-full-dataset.xlsx`, sheet `Economy-wide data` | CCC7 Balanced Pathway series |
| P5 CCC6 benchmark | 4. P5 CCC Benchmark Extraction | CCC6 `The-Sixth-Carbon-Budget-Dataset_v2.xlsx`, sheet `Scenario key metrics` | CCC6 Balanced Net Zero Pathway series |
| Annual gap calculation | 5. DESNZ vs CCC Annual Gap Calculation | DESNZ inc/exc IAS + CCC6/CCC7 series | `data_processed/annual_desnz_ccc_comparison.csv`, `tables/key_benchmark_year_gap_metrics.csv` |
| Carbon-budget-period CCC additions | 6. Carbon-Budget-Period Benchmark Additions | CCC6/CCC7 annual series + DESNZ Table 2.1 | Updated `carbon_budget_gap_metrics.csv` |
| Figure generation | 8. Generate Figures | Cleaned output tables | `figures/desnz_vs_ccc_annual_pathways.png/svg`, `figures/carbon_budget_period_comparison.png/svg` |
| QA checks | 9. Consistency Checks | Generated tables | Assertions for 2050 gap and CB6 official gap |

## Headline assertions

The notebook checks these exact values from the current local `Data_raw` files:

- 2050 DESNZ including-IAS minus CCC7 gap: `325.398436382106 MtCO2e`.
- CB6 official DESNZ projected gap: `737.469408760892 MtCO2e`.

If these assertions fail, either the raw files have changed, the wrong row/sheet was read, or the accounting basis needs to be revisited.
