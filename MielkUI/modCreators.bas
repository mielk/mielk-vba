Attribute VB_Name = "modCreators"
Option Explicit

Private Const CLASS_NAME As String = "modCreators"
'----------------------------------------------------------------------------------------------------------



'[SINGLETONS]
Public Function UI() As UI
    Static instance As UI
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then
        Set instance = New UI
    End If
    Set UI = instance
End Function

Public Function Icons() As ufImages
    Set Icons = ufImages
End Function

Public Function ProgressBar() As WProgressBar
    Static instance As WProgressBar
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then
        Set instance = New WProgressBar
    End If
    Set ProgressBar = instance
End Function

