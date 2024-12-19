VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    '----------------------------------------------------------------------------------------------------------
Private Const CLASS_NAME As String = "FObjects"
    '----------------------------------------------------------------------------------------------------------


Public Function coalesce(ParamArray objects() As Variant) As Object
    Dim varObject As Variant
    '------------------------------------------------------------------------------------------------------
    For Each varObject In objects
        If VBA.IsObject(varObject) Then
            If Not varObject Is Nothing Then
                Set coalesce = varObject
                Exit For
            End If
        End If
    Next varObject
End Function

Public Function cast(value As Variant) As Object
    If f.Variables.isUndefined(value, includeZero:=False, includeUndefinedArray:=False) Then
        Set cast = Nothing
    ElseIf VBA.IsObject(value) Then
        Set cast = value
    Else
        Call VBA.err.raise(1, description:="Illegal data type. Object was expected")
    End If
End Function

#If VBA7 Then
    Public Function getByPointer(pointer As LongPtr) As Object
#Else
    Public Function getByPointer(pointer As Long) As Object
#End If
    Call CopyMemory(getByPointer, pointer, 4)
End Function

Public Function isTheSameObject(obj1 As Variant, obj2 As Variant) As Boolean
    If VBA.IsObject(obj1) And VBA.IsObject(obj2) Then
        If obj1 Is Nothing Then
            isTheSameObject = False
        ElseIf obj2 Is Nothing Then
            isTheSameObject = False
        ElseIf obj1 Is obj2 Then
            isTheSameObject = True
        End If
    End If
End Function
