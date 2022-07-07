VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGridItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Interface for objects that can be listed on custom list view.

Private Const CLASS_NAME As String = "IGridItem"
'----------------------------------------------------------------------------------------------------------

Public Function getKey() As Variant
End Function

Public Function getValue(propName As String) As Variant
End Function

Public Function getDeleteInfoString() As String
End Function

Public Function getActionName(actionTag As String) As String
End Function

Public Function isActionAllowed(actionName As String) As Boolean
End Function

Public Sub run(actionName As String)
End Sub


