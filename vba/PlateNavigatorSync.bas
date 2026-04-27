Attribute VB_Name = "PlateNavigatorSync"
Option Explicit

'---------------------------------------------------------------------------------------
' Purpose:
'   Keep Plate Number and Well ID selections synchronized on the Plate Navigator sheet.
'   - If user changes Plate Number (B4), update Well ID (B5).
'   - If user changes Well ID (B5), update Plate Number (B4).
'   Then force helper recalculation and chart refresh.
'
' Hook:
'   Call HandlePlateNavigatorChange from the Plate Navigator sheet's Worksheet_Change event.
'---------------------------------------------------------------------------------------
Public Sub HandlePlateNavigatorChange(ByVal ws As Worksheet, ByVal Target As Range)
    On Error GoTo SafeExit

    If ws Is Nothing Then Exit Sub
    If Intersect(Target, ws.Range("B4:B5")) Is Nothing Then Exit Sub

    Application.EnableEvents = False

    Dim changedPlate As Boolean
    Dim changedWell As Boolean

    changedPlate = Not Intersect(Target, ws.Range("B4")) Is Nothing
    changedWell = Not Intersect(Target, ws.Range("B5")) Is Nothing

    If changedPlate Then
        ws.Range("B5").Value = LookupMappedValue( _
            lookupValue:=ws.Range("B4").Value2, _
            lookupRange:=ws.Range("I2:I3"), _
            returnRange:=ws.Range("L2:L3") _
        )
    ElseIf changedWell Then
        ws.Range("B4").Value = LookupMappedValue( _
            lookupValue:=ws.Range("B5").Value2, _
            lookupRange:=ws.Range("L2:L27"), _
            returnRange:=ws.Range("I2:I27") _
        )
    End If

    RecalculateHelpersAndRefreshCharts ws.Parent

SafeExit:
    Application.EnableEvents = True
End Sub

Private Function LookupMappedValue( _
    ByVal lookupValue As Variant, _
    ByVal lookupRange As Range, _
    ByVal returnRange As Range _
) As Variant

    Dim idx As Variant

    If Len(Trim$(CStr(lookupValue))) = 0 Then
        LookupMappedValue = vbNullString
        Exit Function
    End If

    idx = Application.Match(lookupValue, lookupRange, 0)

    If IsError(idx) Then
        LookupMappedValue = vbNullString
    Else
        LookupMappedValue = returnRange.Cells(CLng(idx), 1).Value
    End If
End Function

Private Sub RecalculateHelpersAndRefreshCharts(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim objChart As ChartObject

    Application.CalculateFullRebuild

    On Error Resume Next
    wb.RefreshAll
    On Error GoTo 0

    For Each ws In wb.Worksheets
        For Each objChart In ws.ChartObjects
            objChart.Chart.Refresh
        Next objChart
    Next ws
End Sub
