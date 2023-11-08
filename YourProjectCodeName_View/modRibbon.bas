Attribute VB_Name = "modRibbon"
Option Explicit

Private Const CLASS_NAME As String = "modRibbon"
'----------------------------------------------------------------------------------------------------------





'[Initializer]
Private Sub ribbon_YourProjectCodeName_afterLoaded(ByVal ribbon As IRibbonUI)
    Call ErrorManager.Clear
    Call Msg
    Call RibbonManager(F.Create.RibbonManager. _
                        setRibbon(ribbon). _
                        setWorkbook(Excel.ThisWorkbook). _
                        setJsonFilePath(Paths.RibbonConfigFilePath))
    
    Call ribbon.ActivateTab("tab.YourProjectCodeName")
    Call RibbonManager.setUpdateDisabled(False)
End Sub

Private Sub ribbon_showNotifier_afterLoaded(ByVal ribbon As IRibbonUI)
    Const MESSAGE_TAG As String = "Warning.RibbonXmlNotAdded"
    '------------------------------------------------------------------------------------------------------
    Call setParentApp
    Call wksMissingRibbonXml.Activate
    Call VBA.MsgBox(Msg.getText(MESSAGE_TAG), vbExclamation, App.getNameVersion)
End Sub





'[Callback functions]
Private Sub getLabel_YourProjectCodeName(ByVal control As IRibbonControl, ByRef returnedVal)
    Call RibbonManager.assignProperty(control.ID, "label", returnedVal)
End Sub

Private Sub getEnabled_YourProjectCodeName(ByVal control As IRibbonControl, ByRef returnedVal)
    Call RibbonManager.assignProperty(control.ID, "enabled", returnedVal)
End Sub

Private Sub getVisible_YourProjectCodeName(ByVal control As IRibbonControl, ByRef returnedVal)
    Call RibbonManager.assignProperty(control.ID, "visible", returnedVal)
End Sub

Private Sub getImage_YourProjectCodeName(ByVal control As IRibbonControl, ByRef image)
    Call RibbonManager.assignControlImage(control.ID, image)
End Sub

Private Sub getScreentip_YourProjectCodeName(ByVal control As IRibbonControl, ByRef returnedVal)
    Call RibbonManager.assignProperty(control.ID, "screentip", returnedVal)
End Sub
