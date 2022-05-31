VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "WBackground"
'[State] --------------------------------------------------------------------------------------------------
Private pWindow As WTemplate
'[UI components] ------------------------------------------------------------------------------------------
Private pForm As ufBackground
'----------------------------------------------------------------------------------------------------------




Private Sub Class_Initialize()
    Set pForm = New ufBackground
End Sub



'[SETTERS]
Public Function setWindow(value As WTemplate) As WBackground
    Set setWindow = Me
    Set pWindow = value
    Call pForm.setWindow(value)
End Function

Public Function setWidth(value As Single) As WBackground
    Set setWidth = Me
    Call pForm.setWidth(value)
End Function

Public Function setHeight(value As Single) As WBackground
    Set setHeight = Me
    Call pForm.setHeight(value)
End Function

Public Function setLeft(value As Single) As WBackground
    Set setLeft = Me
    Call pForm.setLeft(value)
End Function

Public Function setTop(value As Single) As WBackground
    Set setTop = Me
    Call pForm.setTop(value)
End Function

Public Sub show()
    Call pForm.show(vbModal)
End Sub

Public Sub hide()
    Call pForm.hide
End Sub

Public Sub makeOpaque()
    Call pForm.makeOpaque
End Sub

Public Sub makeTransparent()
    Call pForm.makeTransparent
End Sub
