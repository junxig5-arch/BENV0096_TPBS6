# BENV0096 Dissertation Code and Output Package

Generated: 2026-08-06T00:13:54

This folder organises the code and analysis outputs currently supporting the dissertation draft:

**Thesis used for comparison:** `C:\Users\888\Desktop\Final Draft\TBPS6_Dissertation_Final_Draft_Cleaned.docx`

**Original local code folder supplied by user:** `C:\Users\888\BENV0096_Disertation`

The package is structured so that the dissertation's Appendix A relative paths can be found directly from this folder.

## Folder Structure

- `p4_p5_rebuild/notebooks/` - P4-P5 local notebooks and Python exports from the supplied local folder.
- `p4_p5_local_reproduction/` - final P4-P5 processed outputs, tables and figures used in the thesis.
- `p6_sector_analysis/` - P6 sector analysis notebooks, supplement notebook, figures and tables.
- `p7_neso_uncertainty/` - P7 NESO/FES notebooks, builder scripts, external-condition figures and tables.
- `p8_uncertainty_framework/` - P8 historical-rate, uncertainty, scenario and linkage notebooks plus final figures/tables.
- `document_build_audit_scripts/` - scripts used for final figure cleanup, page cache repair and meeting-material generation.
- `00_original_local_source_snapshot/` - snapshot of the original supplied folder, excluding notebook checkpoints and cache folders.
- `MANIFESTS/` - crosswalks between thesis figures/tables, code files and output files.

## Recommended Reproduction Order

1. Run P4-P5 notebooks:
   - `p4_p5_rebuild/notebooks/P5_CCC_data_cleaning_local_reproducible.ipynb`
   - `p4_p5_rebuild/notebooks/P5_benchmark_comparison_from_cleaned_CCC_local_reproducible.ipynb`
   - `p4_p5_rebuild/notebooks/P4_P5_local_reproducible_with_historical_check.ipynb`

2. Run P6 notebooks:
   - `p6_sector_analysis/notebooks/P6_sectoral_drivers_local_reproducible.ipynb`
   - `p6_sector_analysis/notebooks/P6_supplement_historical_sector_bridge_local_reproducible.ipynb`

3. Run P7 notebooks/scripts:
   - `p7_neso_uncertainty/notebooks/P7_NESO_FES_compact_indicator_extraction_local_reproducible.ipynb`
   - `p7_neso_uncertainty/notebooks/P7_rigour_enhanced_external_conditions_local_reproducible.ipynb`

4. Run P8 notebooks:
   - `p8_uncertainty_framework/notebooks/P8_1_historical_delivered_rate_benchmark_local_reproducible.ipynb`
   - `p8_uncertainty_framework/notebooks/P8_2_sectoral_uncertainty_and_linkages_local_reproducible.ipynb`
   - `p8_uncertainty_framework/notebooks/P8_3_2x2_matrix_near_miss_mini_scenarios_local_reproducible.ipynb`
   - `p8_uncertainty_framework/notebooks/P8_4_sector_linkage_deepening_local_reproducible.ipynb`

## Important Note About Raw Data

The package includes processed outputs and some local P8 raw files, but it does not copy all large official raw workbooks/PDFs. See:

- `MANIFESTS/data_required_manifest.csv`
- `MANIFESTS/code_and_output_manifest.csv`

The raw official sources are documented there so that they can be downloaded or placed in the expected local folders if full rerun from raw data is required.

## Key Manifest Files

- `MANIFESTS/thesis_figure_table_code_crosswalk.csv` maps thesis figures/tables to the relevant code and output files.
- `MANIFESTS/current_thesis_captions_extracted.csv` lists the figure/table captions extracted from the current thesis.
- `MANIFESTS/appendix_a_output_locations_extracted.csv` lists Appendix A output-location rows extracted from the thesis.
- `MANIFESTS/missing_from_supplied_local_folder.md` explains which current-thesis code/output items were missing from `C:/Users/888/BENV0096_Disertation` and were filled from the dissertation workspace.
- `MANIFESTS/copied_file_inventory.csv` lists every copied file with source, destination, size and SHA256.
