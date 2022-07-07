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
    ElseIf Not f.strings.compareStrings(pMessage, validation.getMessage) Then
        compare = False
    Else
        compare = True
    End If
End Function

Public Sub merge(validation As ValidationObject)
    If Not validation Is Nothing Then
        If validation.getStatus <> True Then pStatus = validation.getStatus
        
        If VBA.Len(validation.getMessage) = 0 Then
        ElseIf VBA.Len(pMessage) > 0 Then
            pMessage = pMessage & VBA.vbCrLf & validation.getMessage
        Else
            pMessage = validation.getMessage
        End If
        
        If Not pErrors Is Nothing Then
            Call f.dictionaries.appendDictionary(pErrors, validation.getErrors)
        End If
        
    End If
End Sub
