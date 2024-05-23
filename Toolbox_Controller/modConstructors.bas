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

Public Function CodeComparisonManager() As CodeComparisonManager
    Static instance As CodeComparisonManager
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then Set instance = New CodeComparisonManager
    Set CodeComparisonManager = instance
End Function

Public Function CodeComparisonPrinter() As CodeComparisonPrinter
    Static instance As CodeComparisonPrinter
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then Set instance = New CodeComparisonPrinter
    Set CodeComparisonPrinter = instance
End Function



Public Function newProjectsComparison() As EProjectsComparison
    Set newProjectsComparison = New EProjectsComparison
End Function

Public Function newModulesComparison() As EModulesComparison
    Set newModulesComparison = New EModulesComparison
End Function
