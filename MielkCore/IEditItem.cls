VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEditItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Interface grouping all entities that can be edited.

Private Const CLASS_NAME As String = "IEditItem"
'----------------------------------------------------------------------------------------------------------

Public Function clone() As IEditItem
End Function

Public Function getValue(valueKey As String) As Variant
End Function

Public Function getValidation(validationFunction As String, Optional validationParams As Scripting.Dictionary) As ValidationObject
End Function

Public Function getBasicPropertiesDictionary(Optional includeId As Boolean = False) As Scripting.Dictionary
End Function

Public Sub updateValue(valueKey As String, value As Variant)
End Sub
