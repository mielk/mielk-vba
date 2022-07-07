VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StandarizingAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "StandarizingAction"
'----------------------------------------------------------------------------------------------------------
Private pAlias As String
Private pBefore As Variant
Private pAfter As Variant
'----------------------------------------------------------------------------------------------------------


'[Setters]
Public Function setAlias(value As String) As StandarizingAction
    Set setAlias = Me
    pAlias = value
End Function

Public Function setBefore(value As Variant) As StandarizingAction
    Set setBefore = Me
    Call F.Variables.assign(pBefore, value)
End Function

Public Function setAfter(value As Variant) As StandarizingAction
    Set setAfter = Me
    Call F.Variables.assign(pBefore, value)
End Function



'[Getters]
Public Function getAlias() As String
    getAlias = pAlias
End Function

Public Function getBefore() As Variant
    Call F.Variables.assign(getBefore, pBefore)
End Function

Public Function getAfter() As Variant
    Call F.Variables.assign(getAfter, pAfter)
End Function
