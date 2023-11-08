Attribute VB_Name = "modHelpers"
Option Explicit

Private Const CLASS_NAME As String = "modHelpers"
'----------------------------------------------------------------------------------------------------------



Public Function getYourProjectCodeNameWorkbook() As Excel.Workbook
    Dim wkb As Excel.Workbook
    '------------------------------------------------------------------------------------------------------
    
    For Each wkb In Excel.Workbooks
        If F.regex.checkIfMatch(wkb.name, VIEW_WORKBOOK_NAME) Then
            Set getYourProjectCodeNameWorkbook = wkb
            Exit For
        End If
    Next wkb
    
End Function

Public Function getYourProjectCodeNameWorkbookPath() As String
    Dim wkb As Excel.Workbook
    '------------------------------------------------------------------------------------------------------
    Set wkb = getYourProjectCodeNameWorkbook
    If Not wkb Is Nothing Then
        getYourProjectCodeNameWorkbookPath = wkb.FullName
    End If
End Function


Public Sub setupServices()
    Call setParentApp
    Call ActionLogger(inject:=TextfileActionLogger)
    Call F.Excel.adjustExcelSettings(Excel.Application)
    Call RibbonManager.setUpdateDisabled(True)
End Sub

Public Sub setParentApp()
    With App
        Call .setName(APPLICATION_NAME)
        Call .setVersion(APPLICATION_VERSION)
        Call .setPath(getYourProjectCodeNameWorkbookPath)
    End With
End Sub

