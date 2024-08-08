VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IItemSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "IItemSelector"
'----------------------------------------------------------------------------------------------------------

Public Sub run()
End Sub

Public Sub setCurrentSelection(value As Variant)
End Sub

Public Function getValue() As Variant
End Function

Public Function getDisplayValueForId(id As Long) As String
End Function

Public Function getDetailsForId(id As Long) As Variant
End Function

Public Function isCancelledByUser() As Boolean
End Function
