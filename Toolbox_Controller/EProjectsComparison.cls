VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EProjectsComparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "EProjectsComparison"
'----------------------------------------------------------------------------------------------------------
Private pBaseProject As EVbaProject
Private pCompareProject As EVbaProject
'----------------------------------------------------------------------------------------------------------
Private pModuleComparisons As Scripting.Dictionary
Private pBaseOnlyModules As Scripting.Dictionary
Private pCompareOnlyModules As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------



'[GETTERS]
Public Function getBaseProject() As EVbaProject
    Set getBaseProject = pBaseProject
End Function

Public Function getCompareProject() As EVbaProject
    Set getCompareProject = pCompareProject
End Function

Public Function getModuleComparisons() As Scripting.Dictionary
    Set getModuleComparisons = pModuleComparisons
End Function

Public Function getBaseOnlyModules() As Scripting.Dictionary
    Set getBaseOnlyModules = pBaseOnlyModules
End Function

Public Function getCompareOnlyModules() As Scripting.Dictionary
    Set getCompareOnlyModules = pCompareOnlyModules
End Function

Public Function countModulesByStatus(countEqual As Boolean) As Long
    Dim varKey As Variant
    Dim moduleComparison As EModulesComparison
    Dim moduleStatus As Boolean
    '------------------------------------------------------------------------------------------------------
    
    For Each varKey In pModuleComparisons.Keys
        Set moduleComparison = pModuleComparisons.item(varKey)
        moduleStatus = moduleComparison.getStatus
        If moduleStatus = countEqual Then
            countModulesByStatus = countModulesByStatus + 1
        End If
    Next varKey
    
End Function


Public Function run(baseProject As EVbaProject, compareProject As EVbaProject) As EProjectsComparison
    Dim baseComponents As Scripting.Dictionary
    Dim compareComponents As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim key As Variant
    Dim baseModule As EVbaModule
    Dim compareModule As EVbaModule
    '------------------------------------------------------------------------------------------------------
    Dim counter As Long
    '------------------------------------------------------------------------------------------------------
    
    Set run = Me
    
    'Assign projects to variables.
    Set pBaseProject = baseProject
    Set pCompareProject = compareProject
    Set baseComponents = pBaseProject.getComponents
    Set compareComponents = pCompareProject.getComponents
    
    'Reset result containers.
    Set pModuleComparisons = f.dictionaries.create(False)
    Set pBaseOnlyModules = f.dictionaries.create(False)
    Set pCompareOnlyModules = f.dictionaries.create(False)
    
    
Call ProgressBar.setCaptionLocked(False)

    For Each key In baseComponents
        If compareComponents.Exists(key) Then
            Set baseModule = baseComponents.item(key)
            Set compareModule = compareComponents.item(key)
            Call f.dictionaries.AddItem(pModuleComparisons, key, _
                                            newModulesComparison.run(baseModule, compareModule))
        Else
            Call f.dictionaries.AddItem(pBaseOnlyModules, key, Nothing)
        End If
        
counter = counter + 1
Call ProgressBar.setProgress(counter / baseComponents.Count)
    Next key

    
    '[Find components existing only in compare project]
    For Each key In compareComponents
        If Not baseComponents.Exists(key) Then
            Call f.dictionaries.AddItem(pCompareOnlyModules, key, Nothing)
        End If
    Next key
    
End Function
