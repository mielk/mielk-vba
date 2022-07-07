Attribute VB_Name = "modHandlers"
Option Explicit

Private Const CLASS_NAME As String = "modHandlers"
'----------------------------------------------------------------------------------------------------------





Public Sub catchClick()
    Call ActionListener.click(Application.caller)
End Sub
