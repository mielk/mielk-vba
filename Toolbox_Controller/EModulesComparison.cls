VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EModulesComparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "EModulesComparison"
'----------------------------------------------------------------------------------------------------------
Private Const PROGRESS_BAR_CAPTION As String = "ComparingCode.ProgressBar.ComparingModules"
'----------------------------------------------------------------------------------------------------------
Private pBaseModule As EVbaModule
Private pCompareModule As EVbaModule
'----------------------------------------------------------------------------------------------------------
Private pDeclarationEqual As Boolean
Private pEqualMethods As Scripting.Dictionary
Private pDifferentMethods As Scripting.Dictionary
Private pBaseOnlyMethods As Scripting.Dictionary
Private pCompareOnlyMethods As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------





'[GETTERS]
Public Function getBaseModule() As EVbaModule
    Set getBaseModule = pBaseModule
End Function

Public Function getCompareModule() As EVbaModule
    Set getCompareModule = pCompareModule
End Function

Public Function getEqualMethods() As Scripting.Dictionary
    If pEqualMethods Is Nothing Then
        Set getEqualMethods = f.dictionaries.create(False)
    Else
        Set getEqualMethods = pEqualMethods
    End If
End Function

Public Function getDifferentMethods() As Scripting.Dictionary
    If pDifferentMethods Is Nothing Then
        Set getDifferentMethods = f.dictionaries.create(False)
    Else
        Set getDifferentMethods = pDifferentMethods
    End If
End Function

Public Function getBaseOnlyMethods() As Scripting.Dictionary
    If pBaseOnlyMethods Is Nothing Then
        Set getBaseOnlyMethods = f.dictionaries.create(False)
    Else
        Set getBaseOnlyMethods = pBaseOnlyMethods
    End If
End Function

Public Function getCompareOnlyMethods() As Scripting.Dictionary
    If pCompareOnlyMethods Is Nothing Then
        Set getCompareOnlyMethods = f.dictionaries.create(False)
    Else
        Set getCompareOnlyMethods = pCompareOnlyMethods
    End If
End Function

Public Function countMethods() As Long
    countMethods = getEqualMethods.Count + _
                   getDifferentMethods.Count + _
                   getBaseOnlyMethods.Count + _
                   getCompareOnlyMethods.Count
End Function

Public Function getMethodsByStatus(status As CodeComparisonStatusEnum) As Scripting.Dictionary
    Select Case status
        Case CodeComparisonStatus_Equal:            Set getMethodsByStatus = pEqualMethods
        Case CodeComparisonStatus_Different:        Set getMethodsByStatus = pDifferentMethods
        Case CodeComparisonStatus_BaseOnly:         Set getMethodsByStatus = pBaseOnlyMethods
        Case CodeComparisonStatus_CompareOnly:      Set getMethodsByStatus = pCompareOnlyMethods
    End Select
    
    If getMethodsByStatus Is Nothing Then
        Set getMethodsByStatus = f.dictionaries.create(False)
    End If
    
End Function

Public Function isDeclarationEqual() As Boolean
    isDeclarationEqual = pDeclarationEqual
End Function





Public Function run(baseModule As EVbaModule, compareModule As EVbaModule) As EModulesComparison
    Set run = Me
    Set pBaseModule = baseModule
    Set pCompareModule = compareModule
    
Call ProgressBar.setCaption(f.Strings.Format(Msg.getText(PROGRESS_BAR_CAPTION), pBaseModule.getName))
DoEvents
Call sleep(50)

    Call compareDeclarations
    Call compareMethods
    
End Function

Private Sub compareDeclarations()
    pDeclarationEqual = f.Strings.compareStrings(pBaseModule.getDeclarations, pCompareModule.getDeclarations)
End Sub

Private Sub compareMethods()
    Dim baseMethods As Scripting.Dictionary
    Dim compareMethods As Scripting.Dictionary
    Dim comparison As Boolean
    '------------------------------------------------------------------------------------------------------
    Dim key As Variant
    Dim baseMethod As EVbaMethod
    Dim compareMethod As EVbaMethod
    '------------------------------------------------------------------------------------------------------
    
    '[Reset result containers]
    Set pEqualMethods = f.dictionaries.create(False)
    Set pDifferentMethods = f.dictionaries.create(False)
    Set pBaseOnlyMethods = f.dictionaries.create(False)
    Set pCompareOnlyMethods = f.dictionaries.create(False)
    
    Set baseMethods = pBaseModule.getMethods
    Set compareMethods = pCompareModule.getMethods
    
    For Each key In baseMethods.Keys
        Set baseMethod = baseMethods.item(key)
        If compareMethods.Exists(key) Then
            Set compareMethod = compareMethods.item(key)
            comparison = baseMethod.compare(compareMethod)
            Call f.dictionaries.AddItem(VBA.IIf(comparison, pEqualMethods, pDifferentMethods), key, baseMethod)
        Else
            Call f.dictionaries.AddItem(pBaseOnlyMethods, key, baseMethod)
        End If
    Next key
    
    '[Check for methods existing only in compare project]
    For Each key In compareMethods.Keys
        If Not baseMethods.Exists(key) Then
            Call f.dictionaries.AddItem(pCompareOnlyMethods, key, compareMethods.item(key))
        End If
    Next key
    
End Sub


Public Function getStatus()
    If Not pDeclarationEqual Then
    ElseIf f.dictionaries.isNonEmptyDictionary(pBaseOnlyMethods) Then
        getStatus = False
    ElseIf f.dictionaries.isNonEmptyDictionary(pCompareOnlyMethods) Then
        getStatus = False
    ElseIf f.dictionaries.isNonEmptyDictionary(pDifferentMethods) Then
        getStatus = False
    Else
        getStatus = True
    End If
End Function
