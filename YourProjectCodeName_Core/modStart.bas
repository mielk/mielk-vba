Attribute VB_Name = "modStart"
Option Explicit

Private Const CLASS_NAME As String = "modStart"
'----------------------------------------------------------------------------------------------------------


Public Sub auto_open()
    Call ActionLogger.addLog("Start", , True)
    Call Session
End Sub

