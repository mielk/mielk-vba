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
Private pStatus As Variant
Private pMessage As String
Private pErrors As Scripting.Dictionary
'Private pCode As Variant
'----------------------------------------------------------------------------------------------------------



'[SETTERS]
Public Function setStatus(value As Variant) As ValidationObject
    Set setStatus = Me
    pStatus = value
End Function

Public Function setMessage(value As String) As ValidationObject
    Set setMessage = Me
    pMessage = value
End Function

Public Function addError(code As Variant, message As Variant) As ValidationObject
    Set addError = Me
    If pErrors Is Nothing Then Set pErrors = f.dictionaries.Create(False)
    Call f.dictionaries.addItem(pErrors, code, message)
End Function

'Public Function setCode(value As String) As ValidationObject
'    Set setCode = Me
'    pCode = value
'End Function



'[GETTERS]
Public Function getStatus() As Variant
    getStatus = pStatus
End Function

Public Function getMessage() As String
    getMessage = pMessage
End Function

Public Function getErrors() As Scripting.Dictionary
    If pErrors Is Nothing Then
        Set getErrors = f.dictionaries.Create(False)
    Else
        Set getErrors = pErrors
    End If
End Function

'Public Function getCode() As Variant
'    getCode = pCode
'End Function



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

Public Function appendValidation(Validation As ValidationObject) As ValidationObject
    Set appendValidation = Me
    If Validation Is Nothing Then
        
    ElseIf Validation.getStatus = False Then
        pStatus = False
        Call appendMessage(Validation.getMessage)
    End If
End Function

Public Function compare(Validation As ValidationObject) As Boolean
    If Validation Is Nothing Then
        compare = False
    ElseIf pStatus <> Validation.getStatus Then
        compare = False
    ElseIf Not f.strings.compareStrings(pMessage, Validation.getMessage) Then
        compare = False
    Else
        compare = True
    End If
End Function

Public Sub merge(Validation As ValidationObject)
    If Not Validation Is Nothing Then
        If Validation.getStatus <> True Then pStatus = Validation.getStatus
        
        If VBA.Len(Validation.getMessage) = 0 Then
        ElseIf VBA.Len(pMessage) > 0 Then
            pMessage = pMessage & VBA.vbCrLf & Validation.getMessage
        Else
            pMessage = Validation.getMessage
        End If
        
        If Not pErrors Is Nothing Then
            Call f.dictionaries.appendDictionary(pErrors, Validation.getErrors)
        End If
        
    End If
End Sub
