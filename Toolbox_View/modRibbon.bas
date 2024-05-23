Attribute VB_Name = "modRibbon"
Option Explicit

Private Const CLASS_NAME As String = "modRibbon"
'----------------------------------------------------------------------------------------------------------



'[Initializer]
Private Sub ribbon_toolbox_afterLoaded(ByVal ribbon As IRibbonUI)
    Call ErrorManager.Clear
    Call Msg
    Call RibbonManager(F.Create.RibbonManager. _
                        setRibbon(ribbon). _
                        setWorkbook(Excel.ThisWorkbook). _
                        setJsonFilePath(Paths.RibbonConfigFilePath))
    
    Call ribbon.ActivateTab("tab.toolbox")
    Call RibbonManager.setUpdateDisabled(False)
    
End Sub



'[Callback functions]
Private Sub getLabel_toolbox(ByVal control As IRibbonControl, ByRef returnedVal)
    Call RibbonManager.assignProperty(control.ID, "label", returnedVal)
End Sub

Private Sub getEnabled_toolbox(ByVal control As IRibbonControl, ByRef returnedVal)
    Call RibbonManager.assignProperty(control.ID, "enabled", returnedVal)
End Sub

Private Sub getVisible_toolbox(ByVal control As IRibbonControl, ByRef returnedVal)
    Call RibbonManager.assignProperty(control.ID, "visible", returnedVal)
End Sub

Private Sub getImage_toolbox(ByVal control As IRibbonControl, ByRef image)
    Call RibbonManager.assignControlImage(control.ID, image)
End Sub

Private Sub getScreentip_toolbox(ByVal control As IRibbonControl, ByRef returnedVal)
    Call RibbonManager.assignProperty(control.ID, "screentip", returnedVal)
End Sub
