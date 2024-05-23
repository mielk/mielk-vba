VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProgressIndicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function setCaption(value As String) As IProgressIndicator
End Function

Public Function setCaptionLocked(value As Boolean) As IProgressIndicator
End Function

Public Function setProgress(value As Double) As IProgressIndicator
End Function

Public Function setProgressAndMessage(progress As Double, caption As String) As IProgressIndicator
End Function

Public Sub openSubProcess(ByVal initialProgress As Double, ByVal finalProgress As Double)
End Sub

Public Sub closeCurrentSubProcess()
End Sub
