VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Interface for objects that can be used as a filter.

Private Const CLASS_NAME As String = "IFilter"
'----------------------------------------------------------------------------------------------------------



'[Setters & initialization]
Public Function setParent(value As IFilterable) As IFilter
End Function

Public Function setKey(value As String) As IFilter
End Function

Public Function setColumnName(value As String) As IFilter
End Function

Public Function setIsActive(value As Boolean) As IFilter
End Function

Public Function setSelection(value As Variant) As IFilter
End Function



'[Getters]
Public Function getSelectionString() As String
End Function

Public Function getSelection() As Variant
End Function



'[Actions]
Public Function match(value As Variant) As Boolean
End Function

Public Sub edit()
End Sub

Public Sub clear()
End Sub



'[Properties]
Public Function isActive() As Boolean
End Function

Public Function getKey() As String
End Function
