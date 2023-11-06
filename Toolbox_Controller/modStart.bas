Attribute VB_Name = "modStart"
Option Explicit

Private Const CLASS_NAME As String = "modStart"
'----------------------------------------------------------------------------------------------------------




Public Sub auto_open()
    Call ActionLogger.addLog("Start", , True)
    Call Session
End Sub

Public Sub setupServices()
    Call setParentApp
    Call ActionLogger(inject:=TextfileActionLogger)
    Call f.Excel.adjustExcelSettings(Excel.Application)
    Call RibbonManager.setUpdateDisabled(True)
End Sub


Sub test()
    Const FILE_PATH As String = "D:\vba-tests\apps\Testowa aplikacja\test-controller.xlam"
    Const LIB_PATH As String = "D:\Dropbox\tm\mielk\mielk-vba\code\mielk-db.xlam"
    Dim xls As Excel.Application
    Dim wkb As Excel.Workbook
    Dim project As vbide.VBProject
    
    Set xls = New Excel.Application
    With xls
        .Visible = True
        '.AutomationSecurity = msoAutomationSecurityForceDisable
    End With
        
    Set wkb = f.Books.open_(FILE_PATH, , xls)
    Set project = f.Developer.getVbProject(wkb)
    
    On Error Resume Next
    Call Err.Clear
    project.references.addFromFile (LIB_PATH)
    Debug.Print Err.Number & " | " & Err.Description
    On Error GoTo 0
    
    
    Stop
    
End Sub
