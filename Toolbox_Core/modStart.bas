Attribute VB_Name = "modStart"
Option Explicit

Private Const CLASS_NAME As String = "modStart"
'----------------------------------------------------------------------------------------------------------

Public Sub setupServices()
    Call ActionLogger(inject:=TextfileActionLogger)
    Call ActionLogger.addLog("start", , True)
    Call setParentApp
    Call F.Excel.adjustExcelSettings(Excel.Application)
    'Call RibbonManager.setUpdateDisabled(True)
End Sub
