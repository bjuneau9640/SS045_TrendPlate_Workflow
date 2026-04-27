# SS045_TrendPlate_Workflow

## Plate Navigator synchronization macro

This repo now includes VBA source to keep Plate Navigator controls in sync:

- `vba/PlateNavigatorSync.bas`
  - If Plate Number (`B4`) changes, it updates Well ID (`B5`).
  - If Well ID (`B5`) changes, it updates Plate Number (`B4`).
  - After either change, it recalculates helper data and refreshes charts.
- `vba/PlateNavigator_WorksheetChange_Snippet.bas`
  - Worksheet event hook that calls the sync routine.

### How to wire it into the workbook

1. Open the `.xlsm` workbook in Excel.
2. Press `Alt+F11` to open the VBA editor.
3. Import `vba/PlateNavigatorSync.bas` into the project.
4. Open the `Plate Navigator` worksheet code module.
5. Paste in the contents of `vba/PlateNavigator_WorksheetChange_Snippet.bas`.
6. Save the workbook as `.xlsm`.
