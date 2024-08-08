VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "UI"
'[Functions] ----------------------------------------------------------------------------------------------
Private pCast As FCast
Private pControls As FControls
Private pCreate As FConstructors
Private pDialogs As FDialogs
Private pForms As FForms
Private pIcons As ufImages
Private pNormalize As FNormalize
'[Enums] --------------------------------------------------------------------------------------------------
Private pProps As UIPropsEnum
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    Set pCast = New FCast
    Set pControls = New FControls
    Set pCreate = New FConstructors
    Set pDialogs = New FDialogs
    Set pForms = New FForms
    Set pIcons = ufImages
    Set pNormalize = New FNormalize
    '[Enums]
    Set pProps = uiProps
End Sub



Public Function Cast() As FCast
    Set Cast = pCast
End Function

Public Function controls() As FControls
    Set controls = pControls
End Function

Public Function Create() As FConstructors
    Set Create = pCreate
End Function

Public Function Dialogs() As FDialogs
    Set Dialogs = pDialogs
End Function

Public Function Forms() As FForms
    Set Forms = pForms
End Function

Public Function Icons() As ufImages
    Set Icons = pIcons
End Function

Public Function Normalize() As FNormalize
    Set Normalize = pNormalize
End Function



'[Enums]
Public Function Events() As EventsEnum
    Set Events = MielkCore.Events
End Function

Public Function props() As UIPropsEnum
    Set props = pProps
End Function
