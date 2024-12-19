Attribute VB_Name = "modSetup"
Option Explicit

Private Const CLASS_NAME As String = "modSetup"
'----------------------------------------------------------------------------------------------------------
Private Const MENU_CAPTION_ADD_ERROR_HANDLING As String = "VBE.Actions.AddErrorHandling.label"
Private Const MENU_CAPTION_CREATE_FRAMED_SECTION As String = "VBE.Actions.CreateFramedSection.label"
Private Const MENU_CAPTION_ADD_SEPARATOR_LINE As String = "VBE.Actions.AddSeparatorLine.label"
Private Const MENU_CAPTION_ADD_CLASS As String = "VBE.Actions.AddClass.label"
Private Const MENU_CAPTION_ADD_SETTERS_AND_GETTERS As String = "VBE.Actions.AddSettersAndGetters.label"
Private Const MENU_CAPTION_ADD_PROP As String = "VBE.Actions.AddProps.label"
Private Const MENU_CAPTION_ADD_REPOSITORY As String = "VBE.Actions.AddRepo.label"
'----------------------------------------------------------------------------------------------------------

Public Sub test()
    Call addErrorHandlingToCurrentMethod
End Sub

Public Sub adjustContextMenu()
    With ContextManager
        Call .removeCustomMenuItems
        Call .addItem(CUSTOM_MENU_CAPTION, msg.getText(MENU_CAPTION_ADD_SEPARATOR_LINE), "addSeparatorLine", 130)
        Call .addItem(CUSTOM_MENU_CAPTION, msg.getText(MENU_CAPTION_CREATE_FRAMED_SECTION), "createFramedSection", 131)
        'Call .addItem(CUSTOM_MENU_CAPTION, Msg.getText(MENU_CAPTION_ADD_ERROR_HANDLING), "addErrorHandlingToCurrentMethod", 348)
        Call .addItem(CUSTOM_MENU_CAPTION, msg.getText(MENU_CAPTION_ADD_CLASS), "addClass", 137)
        Call .addItem(CUSTOM_MENU_CAPTION, msg.getText(MENU_CAPTION_ADD_SETTERS_AND_GETTERS), "addSettersAndGetters", 2104)
        Call .addItem(CUSTOM_MENU_CAPTION, msg.getText(MENU_CAPTION_ADD_PROP), "addProp", 2145)
        Call .addItem(CUSTOM_MENU_CAPTION, msg.getText(MENU_CAPTION_ADD_REPOSITORY), "addRepository", 123)
    End With
End Sub
