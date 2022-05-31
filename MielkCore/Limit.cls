VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Limit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "Limit"
'----------------------------------------------------------------------------------------------------------
Private pName As String
Private pLowValue As Variant
Private pLowValueExcluded As Boolean
Private pHighValue As Variant
Private pHighValueEXcluded As Boolean
'----------------------------------------------------------------------------------------------------------


'[SETTERS]

Public Function setName(value As String) As Limit
    Set setName = Me
    pName = value
End Function

Public Function setLowValue(value As Variant, Optional excluded As Boolean = False) As Limit
    Set setLowValue = Me
    pLowValue = value
    pLowValueExcluded = excluded
End Function

Public Function setHighValue(value As Variant, Optional excluded As Boolean = False) As Limit
    Set setHighValue = Me
    pHighValue = value
    pHighValueEXcluded = excluded
End Function



'[GETTERS]
Public Function getLow() As Variant
    getLow = pLowValue
End Function

Public Function getLowExcluded() As Boolean
    getLowExcluded = pLowValueExcluded
End Function

Public Function getHigh() As Variant
    getHigh = pHighValue
End Function

Public Function getHighExcluded() As Boolean
    getHighExcluded = pHighValueEXcluded
End Function

Public Function getRange() As Variant
    getRange = pHighValue - pLowValue
End Function

Public Sub assignIfLower(value As Variant)
    If VBA.IsEmpty(pLowValue) Then
        pLowValue = value
    ElseIf value < pLowValue Then
        pLowValue = value
    End If
End Sub

Public Sub assignIfHigher(value As Variant)
    If VBA.IsEmpty(pHighValue) Then
        pHighValue = value
    ElseIf value > pHighValue Then
        pHighValue = value
    End If
End Sub
