VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Exception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "Exception"
'----------------------------------------------------------------------------------------------------------
Private pNumber As Long
Private pName As String
'----------------------------------------------------------------------------------------------------------



'[SETTERS]

Public Function setNumber(value As Long) As Exception
    Set setNumber = Me
    pNumber = value
End Function

Public Function setName(value As String) As Exception
    Set setName = Me
    pName = value
End Function


'[GETTERS]

Public Function getNumber() As Long
    getNumber = pNumber
End Function

Public Function getName() As String
    getName = pName
End Function


Public Function getDescription(ParamArray params() As Variant) As String
    getDescription = F.strings.format(MsgService.getText(pName), params)
End Function

Public Sub raise(ParamArray params() As Variant)
    Dim description As String
    '----------------------------------------------------------------------------------------------------------
    
    description = F.strings.format(MsgService.getText(pName), params)
    Call VBA.err.raise(number:=pNumber, description:=description)
    
End Sub

Public Sub raiseWithCustomMessage(message As String)
    Call VBA.err.raise(number:=pNumber, description:=message)
End Sub
