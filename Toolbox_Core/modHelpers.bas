Attribute VB_Name = "modHelpers"
Option Explicit

Private Const CLASS_NAME As String = "modHelpers"
'----------------------------------------------------------------------------------------------------------



Public Function getToolboxWorkbook() As Excel.Workbook
    Dim wkb As Excel.Workbook
    '------------------------------------------------------------------------------------------------------
    
    For Each wkb In Excel.Workbooks
        If F.Strings.compareStrings(wkb.name, VIEW_WORKBOOK_NAME) Then
            Set getToolboxWorkbook = wkb
            Exit For
        End If
    Next wkb
    
End Function
