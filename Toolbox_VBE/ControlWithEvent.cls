VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlWithEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "ControlWithEvent"
'----------------------------------------------------------------------------------------------------------
Public WithEvents pControl As VBIDE.CommandBarEvents
Attribute pControl.VB_VarHelpID = -1
'----------------------------------------------------------------------------------------------------------

Private Sub pControl_Click(ByVal control As Object, handled As Boolean, CancelDefault As Boolean)
    
    On Error Resume Next
    
    'codeSnippet_range = CommandBarControl.Parameter
    
    Application.run control.OnAction
    
    handled = True
    CancelDefault = True
    
End Sub

