Attribute VB_Name = "modConstructors"
Option Explicit

Private Const CLASS_NAME As String = "modConstructors"
'----------------------------------------------------------------------------------------------------------


'[Generic]
Public Function Session() As ESession
    Static instance As ESession
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then
        Call setupServices
        Set instance = New ESession
        Call instance.setup
    End If
    Set Session = instance
End Function

Public Function RibbonManager(Optional inject As RibbonManager) As RibbonManager
    Static instance As RibbonManager
    '------------------------------------------------------------------------------------------------------
    If Not inject Is Nothing Then Set instance = inject
    Set RibbonManager = instance
End Function

Public Function RibbonControlTypes(Optional inject As CRibbonControlTypes) As CRibbonControlTypes
    Static instance As CRibbonControlTypes
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then
        Set instance = New CRibbonControlTypes
    End If
    Set RibbonControlTypes = instance
End Function


'[Project-specific]
Public Function Toolbox() As Toolbox
    Static instance As Toolbox
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then Set instance = New Toolbox
    Set Toolbox = instance
End Function

Public Function CodeCompactor() As CodeCompactor
    Static instance As CodeCompactor
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then Set instance = New CodeCompactor
    Set CodeCompactor = instance
End Function


