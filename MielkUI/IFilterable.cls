VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFilterable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Interface for objects that can have filter applied (i.e. XGridManager)

Private Const CLASS_NAME As String = "IFilterable"
'----------------------------------------------------------------------------------------------------------

Public Sub applyFilter(filter As IFilter)
End Sub

Public Sub removeFilter(filter As IFilter)
End Sub

Public Function getAvailableValues(key As String) As Scripting.Dictionary
End Function

Public Function getFilterRegex(key As String) As String
End Function
