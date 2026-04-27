2025 SS045 Trend Plate Codex Source Bundle v01

Purpose
This source bundle contains the files needed to rebuild the 2025 SS045 Appendix D trend plate workbook in Codex as a polished macro-enabled Excel workbook. The intended output file is:

04_Trend_Plates\2025_SS045_Trend_Plate_Master_v03_CodexBuild_Working.xlsm

Use this bundle with the Codex_Build_Prompt.txt file in this folder.

Important controls
- Do not use 2025_SS045_Trend_Plate_Master_v02_Working.xlsx as a production source. It was a concept-only prototype with chart source issues.
- Use the user's prototype workbook in 01_Prototype_Template as the visual/design reference, especially the Plate - Well #15 Rev concept.
- Use the AMR working tables workbook as the controlling summary source for 2025 exceedance, monitoring-scope, MNA, and recommendation logic.
- Use the historical data workbooks as the controlling source for trend plate time-series plots.
- Use Attachment 1A and the GIS input table for well metadata, coordinates, groundwater zone, and 2025 inspection status.
- The full 2025 analytical package is not included because it is too large and is not needed for workbook engineering. If a value discrepancy is identified, request the analytical package separately.
- The DUVR/DUSR package is included as optional validation support for qualifier treatment.
- Do not add password protection until the template is finalized and reviewed.

Recommended output review sequence
1. Confirm Dynamic Plate dropdowns for Plate Number and Well ID stay synchronized.
2. Test D-1 BH01MW02, D-9 Well #15, D-12 A39MW094, and D-16 A39L036DPR2.
3. Confirm VOC charts plot correctly from fixed helper ranges.
4. Confirm all axis titles are outside the plot area.
5. Confirm 11x17 landscape print areas are set and export cleanly to PDF.
6. Run workbook QA/QC and produce a QA/QC log.

Folder contents
00_README_AND_PROMPT: README, source-file manifest, and Codex build prompt.
01_Prototype_Template: user-provided macro-enabled prototype workbook.
02_Control_Documents: plate register, Appendix D cover memo, client rules, and AMR table/figure/appendix register.
03_AMR_Working_Tables: current AMR working tables through Table 5-1.
04_Historical_Data: historical data workbooks and precipitation file.
05_Metadata_and_GIS: monitoring well inspection table and GIS input/request workbook.
06_Data_Validation_Optional: DUVR/DUSR package for optional qualifier reference.
