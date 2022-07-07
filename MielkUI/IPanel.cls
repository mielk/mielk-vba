VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "IPanel"
'----------------------------------------------------------------------------------------------------------


'[Parent]
Public Function setParentFrame(value As MSForms.frame) As IPanel
End Function

Public Function setWindow(value As WTemplate) As IPanel
End Function

Public Function setListener(value As IListener) As IPanel
End Function



'[Index]
Public Function setName(value As String) As IPanel
End Function

Public Function getName() As String
End Function

Public Function setIndex(value As Long) As IPanel
End Function

Public Function getIndex() As Long
End Function




'[UI properties]
Public Function getUiProperty(prop As UIProp) As Variant
End Function

Public Function setUiProperty(prop As UIProp, value As Variant) As IPanel
End Function

Public Function isVisible() As Boolean
End Function



'[Rendering]
Public Sub render()
End Sub

Public Sub insertBorder(index As XlBordersIndex, color As Long)
End Sub

Public Sub addValidationLabel(value As UIValidationLabel)
End Sub



'[Results]
Public Function getValue() As Variant
End Function

Public Function getValidation() As ValidationObject
End Function
