VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "MApp"
'----------------------------------------------------------------------------------------------------------
Private pConstructors As FConstructors
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    Set pConstructors = New FConstructors
End Sub


Public Function Create() As FConstructors
    Set Create = pConstructors
End Function
