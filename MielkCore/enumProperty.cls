VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "enumProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "enumProperty"
'----------------------------------------------------------------------------------------------------------
Private Const TRANSLATION_PREFIX As String = "props."
'----------------------------------------------------------------------------------------------------------
Private pName As String
Private pDbField As String
'----------------------------------------------------------------------------------------------------------


'[SETTERS]
Public Function setName(value As String) As enumProperty
    Set setName = Me
    pName = value
End Function

Public Function setDbField(value As String) As enumProperty
    Set setDbField = Me
    pDbField = value
End Function



'[GETTERS]
Public Function getName() As String
    getName = pName
End Function

Public Function getLowerCaseName() As String
    getLowerCaseName = VBA.LCase$(pName)
End Function

Public Function getDbField() As String
Attribute getDbField.VB_UserMemId = 0
    getDbField = pDbField
End Function

Public Function getTag() As String
    getTag = "[" & pName & "]"
End Function

Public Function getDisplayedValue() As String
    getDisplayedValue = MsgService.getText(TRANSLATION_PREFIX + pName)
End Function

Public Function toString() As String
    toString = pName
End Function
