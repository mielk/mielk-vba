VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EWhere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "EWhere"
'----------------------------------------------------------------------------------------------------------
Private pFieldName As String
Private pComparisonMode As ComparisonModeEnum
Private pComparedValue As Variant
'----------------------------------------------------------------------------------------------------------



'[SETTERS]
Public Function setFieldName(value As String) As EWhere
    Set setFieldName = Me
    pFieldName = value
End Function

Public Function setComparisonMode(value As ComparisonModeEnum) As EWhere
    Set setComparisonMode = Me
    pComparisonMode = value
End Function

Public Function setComparedValue(value As Variant) As EWhere
    Set setComparedValue = Me
    Call F.Variables.assign(pComparedValue, value)
End Function




'[GETTERS]
Public Function getFieldName() As String
    getFieldName = pFieldName
End Function

Public Function getComparisonMode() As ComparisonModeEnum
    getComparisonMode = pComparisonMode
End Function

Public Function getComparedValue() As Variant
    getComparedValue = pComparedValue
End Function





'[API]
Public Function toString() As String
    Dim col As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    If pComparisonMode = ComparisonMode_In Then
        If VBA.IsArray(pComparedValue) Then
            Set col = F.Collections.fromArray(pComparedValue)
        Else
            Set col = pComparedValue
        End If
        toString = pFieldName & operatorAsString & " (" & F.Collections.toString(col, StringifyMode_Db) & ")"
    Else
        toString = pFieldName & operatorAsString & F.Strings.toDbValue(pComparedValue)
    End If
End Function

Private Function operatorAsString() As String
    Select Case pComparisonMode
        Case ComparisonMode_Equal: operatorAsString = " = "
        Case ComparisonMode_GreaterThan: operatorAsString = " > "
        Case ComparisonMode_LessThan: operatorAsString = " < "
        Case ComparisonMode_NotEqual: operatorAsString = " <> "
        Case ComparisonMode_In: operatorAsString = " IN "
    End Select
End Function
