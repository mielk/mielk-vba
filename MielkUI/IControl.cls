VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Interface describing all classes that can be used as a control

Private Const CLASS_NAME As String = "IControl"
'----------------------------------------------------------------------------------------------------------

Public Sub setVisibility(value As Boolean)
End Sub

Public Sub setEnability(value As Boolean)
End Sub

Public Sub setWidth(value As Single)
End Sub


Public Function getHeight() As Single
End Function

Public Function getWidth() As Single
End Function

Public Function getLeft() As Single
End Function

Public Function getTop() As Single
End Function

Public Function getRectangle() As RECT
End Function


Public Function Move(Optional x As Variant, Optional y As Variant) As IControl
End Function

Public Function offset(Optional x As Single, Optional y As Single) As IControl
End Function

Public Sub alignVertically(baseCtrl As MSForms.control)
End Sub

Public Sub resize(Optional width As Variant, Optional height As Variant)
End Sub





Public Sub triggerEvent(eventId As Long)
End Sub
