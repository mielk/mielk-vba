VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FCast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FCast"
'----------------------------------------------------------------------------------------------------------

Public Function toObject(value As Variant) As Object
    If VBA.IsObject(value) Then
        Set toObject = value
    End If
End Function

Public Function toDictionary(value As Variant) As Scripting.Dictionary
    If F.Dictionaries.isDictionary(value) Then
        Set toDictionary = value
    End If
End Function

Public Function toCollection(value As Variant) As VBA.Collection
    If F.Collections.isCollection(value) Then
        Set toCollection = value
    End If
End Function

Public Function toValidationObject(value As Variant) As ValidationObject
    On Error Resume Next
    Set toValidationObject = value
End Function
