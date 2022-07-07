VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISubitemsParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "IAliasable"

Public Function getId() As Long
End Function

Public Function getLabel() As String
End Function

Public Function countAliases() As Long
End Function

Public Function getSubitems() As VBA.Collection
End Function

Public Function getSubitemsForAllObjects() As Scripting.Dictionary
End Function

Public Function getSubitemCurrentParent(alias As String, Optional displayValue As Boolean = True) As Variant
End Function

Public Function getObjectsAffectedByReassignment() As VBA.Collection
End Function


Public Sub addSubitem(subitem As String)
End Sub

Public Sub removeSubitem(subitem As String)
End Sub


