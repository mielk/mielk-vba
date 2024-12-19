Attribute VB_Name = "modVersioning"
Option Explicit

Private Const CLASS_NAME As String = "modVersioning"
'----------------------------------------------------------------------------------------------------------

Public Function getFileVersion(wkb As Excel.Workbook) As String
    getFileVersion = f.ExcelNames.getValue(wkb, props.version.getName)
End Function

Public Sub setFileVersion(wkb As Excel.Workbook, value As String, Optional description As String)
    Call f.ExcelNames.assignValue(wkb, props.version.getName, value)
    Call addLogToVersionsHistory(wkb, value, description)
End Sub

Private Sub addLogToVersionsHistory(wkb As Excel.Workbook, version As String, description As String)
    Const SHEET_NAME_PATTERN As String = "versions?.?(?:history)?"
    Const DEFAULT_SHEET_NAME As String = "versions_history"
    '------------------------------------------------------------------------------------------------------
    Const COL_VERSION As Long = 1
    Const COL_TIMESTAMP As Long = 2
    Const COL_DESCRIPTION As Long = 3
    '------------------------------------------------------------------------------------------------------
    Dim wks As Excel.Worksheet
    Dim lastRow As Long
    '------------------------------------------------------------------------------------------------------
    
    Set wks = f.sheets.getSheet(wkb, SHEET_NAME_PATTERN, True)
    If f.sheets.IsValid(wks) Then
        wks.name = DEFAULT_SHEET_NAME
    Else
        Set wks = wkb.Worksheets.Add
        With wks
            .name = DEFAULT_SHEET_NAME
            .cells(1, COL_VERSION).value = "version"
            .cells(2, COL_TIMESTAMP).value = "timestamp"
            .cells(3, COL_DESCRIPTION).value = "description"
        End With
    End If

    lastRow = f.ranges.getLastNonEmptyRow(wks)
    With wks
        .cells(lastRow + 1, COL_VERSION).value = version
        .cells(lastRow + 1, COL_TIMESTAMP).value = VBA.Now
        .cells(lastRow + 1, COL_DESCRIPTION).value = description
    End With

End Sub

