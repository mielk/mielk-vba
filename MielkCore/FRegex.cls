VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FRegex"
'----------------------------------------------------------------------------------------------------------


Public Function Create( _
                    Optional pattern As String = vbNullString, _
                    Optional searchGlobal As Boolean = True, _
                    Optional multiline As Boolean = True, _
                    Optional ignoreCase As Boolean = True) As Object
    Set Create = VBA.CreateObject("VBScript.RegExp")
    With Create
        .Global = searchGlobal
        .multiline = multiline
        .ignoreCase = ignoreCase
        .pattern = pattern
    End With
End Function

Public Function checkIfMatch(text As String, regexPattern As String, _
                    Optional ignoreCase As Boolean = True, _
                    Optional searchGlobal As Boolean = True, _
                    Optional multiline As Boolean = True) As Boolean
    Static objRegex As Object
    '------------------------------------------------------------------------------------------------------
    If objRegex Is Nothing Then Set objRegex = f.regex.Create
    With objRegex
        .Global = searchGlobal
        .multiline = multiline
        .ignoreCase = ignoreCase
        .pattern = regexPattern
        checkIfMatch = .test(text)
    End With
End Function

Public Function checkIfAnyMatch(text As String, patterns As Variant, _
                    Optional ignoreCase As Boolean = True, _
                    Optional searchGlobal As Boolean = True, _
                    Optional multiline As Boolean = True) As Boolean
    Static objRegex As Object
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim colPatterns As VBA.Collection
    Dim varPattern As Variant
    '------------------------------------------------------------------------------------------------------
    
    If objRegex Is Nothing Then Set objRegex = f.regex.Create
    
    If f.arrays.countDimensions(patterns) = 1 Then
        Set colPatterns = f.arrays.toCollection(patterns)
    ElseIf f.Collections.isNonEmptyCollection(patterns) Then
        Set colPatterns = patterns
    End If
    
    For Each varPattern In patterns
        With objRegex
            .pattern = VBA.CStr(varPattern)
            .Global = searchGlobal
            .multiline = multiline
            .ignoreCase = ignoreCase
            If .test(text) Then
                checkIfAnyMatch = True
                Exit For
            End If
        End With
    Next varPattern
    
End Function

Public Function replace(text As String, regexPattern As String, replacement As String, _
                                            Optional ignoreCase As Boolean = True, _
                                            Optional searchGlobal As Boolean = True, _
                                            Optional multiline As Boolean = True) As String
    Static objRegex As Object
    '------------------------------------------------------------------------------------------------------
    If objRegex Is Nothing Then Set objRegex = f.regex.Create
    With objRegex
        .ignoreCase = ignoreCase
        .Global = searchGlobal
        .multiline = multiline
        .pattern = regexPattern
        replace = .replace(text, replacement)
    End With
End Function
                                            

Public Function getFirstGroupMatch(text As String, pattern As String, _
                                            Optional ignoreCase As Boolean = True, _
                                            Optional searchGlobal As Boolean = True, _
                                            Optional multiline As Boolean = True) As String
    Static objRegex As Object           'VBScript_RegExp_55.RegExp
    '------------------------------------------------------------------------------------------------------
    Dim matches As Object               'IMatchCollection2
    Dim match As Object                 'IMatch2
    '------------------------------------------------------------------------------------------------------
    
    If objRegex Is Nothing Then
        Set objRegex = VBA.CreateObject("VBScript.RegExp")
    End If
    
    With objRegex
        .pattern = pattern
        .Global = searchGlobal
        .multiline = multiline
        .ignoreCase = ignoreCase
        Set matches = .Execute(text)
    End With
    
    If matches.count Then
        For Each match In matches
            If match.SubMatches.count Then
                getFirstGroupMatch = match.SubMatches(0)
                Exit For
            End If
        Next match
    End If
    
    Set matches = Nothing
    Set match = Nothing
    
End Function

Public Function getMatchArray(text As String, pattern As String, _
                                            Optional ignoreCase As Boolean = True, _
                                            Optional searchGlobal As Boolean = True, _
                                            Optional multiline As Boolean = True) As Variant
    Static objRegex As Object           'VBScript_RegExp_55.RegExp
    '------------------------------------------------------------------------------------------------------
    Dim matches As Object               'IMatchCollection2
    Dim match As Object                 'IMatch2
    '------------------------------------------------------------------------------------------------------
    Dim arr As Variant
    Dim Counter As Long
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If objRegex Is Nothing Then
        Set objRegex = VBA.CreateObject("VBScript.RegExp")
    End If
    
    With objRegex
        If .pattern <> pattern Then .pattern = pattern
        .Global = searchGlobal
        .multiline = multiline
        .ignoreCase = ignoreCase
        Set matches = .Execute(text)
    End With
    
    If matches.count Then
        For Each match In matches
            Counter = match.SubMatches.count
            If Counter Then
                ReDim arr(1 To Counter)
                For i = 0 To Counter - 1
                    arr(i + 1) = match.SubMatches(i)
                Next i
                getMatchArray = arr
            End If
            Exit For
        Next match
    End If
    
    Set matches = Nothing
    Set match = Nothing
    
End Function

Public Function getMatchArraysCollection(text As String, pattern As String, _
                                            Optional ignoreCase As Boolean = True, _
                                            Optional searchGlobal As Boolean = True, _
                                            Optional multiline As Boolean = True) As VBA.Collection
    Static objRegex As Object           'VBScript_RegExp_55.RegExp
    '------------------------------------------------------------------------------------------------------
    Dim matches As Object               'IMatchCollection2
    Dim match As Object                 'IMatch2
    '------------------------------------------------------------------------------------------------------
    Dim arr As Variant
    Dim Counter As Long
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    Set getMatchArraysCollection = New VBA.Collection
    
    If objRegex Is Nothing Then
        Set objRegex = VBA.CreateObject("VBScript.RegExp")
    End If
    
    With objRegex
        If .pattern <> pattern Then .pattern = pattern
        .Global = searchGlobal
        .multiline = multiline
        .ignoreCase = ignoreCase
        Set matches = .Execute(text)
    End With
    
    If matches.count Then
        For Each match In matches
            Counter = match.SubMatches.count
            If Counter Then
                ReDim arr(1 To Counter)
                For i = 0 To Counter - 1
                    arr(i + 1) = match.SubMatches(i)
                Next i
                Call getMatchArraysCollection.Add(arr)
            End If
        Next match
    End If
    
    Set matches = Nothing
    Set match = Nothing
    
End Function

