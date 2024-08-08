Attribute VB_Name = "modSingletons"
Option Explicit

Private Const CLASS_NAME As String = "modSingletons"
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

Public Function WindowsCache() As WindowsCache
    Static instance As WindowsCache
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then Set instance = New WindowsCache
    Set WindowsCache = instance
End Function

Public Function Notifier() As WNotifier
    Static instance As WNotifier
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then Set instance = New WNotifier
    'Call instance.checkUserForm
    Set Notifier = instance
End Function


