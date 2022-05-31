VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FDialogs"
'----------------------------------------------------------------------------------------------------------


Public Function getUserConfirmation(question As String, Optional ByVal title As String = "Confirm") As Boolean
    Dim answer As VBA.VbMsgBoxResult
    '----------------------------------------------------------------------------------------------------------
    
    answer = VBA.MsgBox(question, vbYesNo + vbQuestion, title)
    getUserConfirmation = (answer = VBA.VbMsgBoxResult.vbYes)
    
End Function

