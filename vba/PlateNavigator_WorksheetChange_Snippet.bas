Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    PlateNavigatorSync.HandlePlateNavigatorChange Me, Target
End Sub
