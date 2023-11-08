VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAppSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "SAppSettings"
'[Captions] -----------------------------------------------------------------------------------------------
Private Const LANGUAGE_CHANGED As String = "Languages.ChangeLanguage.SuccessInfo"
'----------------------------------------------------------------------------------------------------------
Private Const REG_ITEM_LANGUAGE As String = "language"
'[Default values] -----------------------------------------------------------------------------------------
Private Const DEFAULT_ASK_BEFORE_CLOSE As Boolean = False
'[Basic settings] -----------------------------------------------------------------------------------------
Private pAskBeforeClose As Boolean
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    Call setInitialValues
End Sub

Private Sub setInitialValues()
    pAskBeforeClose = DEFAULT_ASK_BEFORE_CLOSE
End Sub




'[Basic settings]
Public Function AskBeforeClose() As Boolean
    AskBeforeClose = pAskBeforeClose
End Function


'[Languages]
Public Sub changeLanguage()
    Dim selectedLanguage As Language
    Dim message As String
    '------------------------------------------------------------------------------------------------------
    Call Msg
    Set selectedLanguage = MApp.selectLanguage
    If selectedLanguage Is Nothing Then
        'User cancelled
    ElseIf selectedLanguage Is Msg.getCurrentLanguage Then
        'Nothing changed
    Else
        Call setCurrentLanguage(selectedLanguage)
        message = F.Strings.Format(Msg.getText(LANGUAGE_CHANGED), selectedLanguage.getNameForCurrentLanguage)
        Call VBA.MsgBox(message, vbInformation + vbOKOnly, App.getNameVersion)
    End If
End Sub

Friend Sub loadLanguageFromRegistry()
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    value = F.Registry.getValue(HKEY_CURRENT_USER, REG_KEY_NAME, REG_ITEM_LANGUAGE)
    If Not VBA.IsNull(value) Then
        Call Msg.setCurrentLanguageByKey(VBA.CStr(value))
    Else
        Call setCurrentLanguage(Msg.getCurrentLanguage)
    End If
End Sub

Private Sub setCurrentLanguage(objLanguage As Language)
    Dim value As Variant
    '----------------------------------------------------------------------------------------------------------
    If Not objLanguage Is Nothing Then
        value = objLanguage.getKey
        Call F.Registry.createValue(HKEY_CURRENT_USER, REG_KEY_NAME, REG_ITEM_LANGUAGE, value, True)
    End If
End Sub
