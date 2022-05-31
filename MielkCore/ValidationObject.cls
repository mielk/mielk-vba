VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "Validation"
'----------------------------------------------------------------------------------------------------------
Private pStatus As Boolean
Private pMessage As String
Private pCode As Variant
'----------------------------------------------------------------------------------------------------------



'[SETTERS]
Public Function setStatus(value As Boolean) As ValidationObject
    Set setStatus = Me
    pStatus = value
End Function

Public Function setMessage(value As String) As ValidationObject
    Set setMessage = Me
    pMessage = value
End Function

Public Function setCode(value As String) As ValidationObject
    Set setCode = Me
    pCode = value
End Function



'[GETTERS]
Public Function getStatus() As Boolean
    getStatus = pStatus
End Function

Public Function getMessage() As String
    getMessage = pMessage
End Function

Public Function getCode() As Variant
    getCode = pCode
End Function



'[API]
Public Function clone() As ValidationObject
    Set clone = New ValidationObject
    With clone
        Call .setStatus(pStatus)
        Call .setMessage(pMessage)
    End With
End Function

Public Function appendMessage(text As String) As ValidationObject
    Dim appendNewLine As Boolean
    Set appendMessage = Me
    If VBA.Len(text) > 0 And VBA.Len(pMessage) > 0 Then appendNewLine = True
    pMessage = pMessage & VBA.IIf(appendNewLine, VBA.vbCrLf, vbNullString) & text
End Function

Public Function appendValidation(validation As ValidationObject) As ValidationObject
    Set appendValidation = Me
    If validation Is Nothing Then
        
    ElseIf validation.getStatus = False Then
        pStatus = False
        Call appendMessage(validation.getMessage)
    End If
End Function

Public Function compare(validation As ValidationObject) As Boolean
    If validation Is Nothing Then
        compare = False
    ElseIf pStatus <> validation.getStatus Then
        compare = False
    ElseIf Not F.strings.compareStrings(pMessage, validation.getMessage) Then
        compare = False
    Else
        compare = True
    End If
End Function
