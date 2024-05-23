VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FStyling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FStyling"
'----------------------------------------------------------------------------------------------------------


Public Sub formatArray(rng As Excel.Range, dictConfig As Scripting.Dictionary, Optional arrayTag As String)
    Dim styles As Scripting.Dictionary
    Dim conditionalStyles As Scripting.Dictionary
    Dim headerConfig As Scripting.Dictionary
    Dim headerStyles As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim headerLast As Boolean
    Dim rowsFirst As Boolean
    Dim columns As VBA.Collection
    Dim rows As VBA.Collection
    Dim styleDto As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
        
    If F.dictionaries.isNonEmptyDictionary(dictConfig) Then
            
        '[Apply general styling]
        Set styles = F.dictionaries.GetObject(dictConfig, props.styles.getName)
        Call F.Ranges.applyFormatFromDictionary(rng, styles)
        Call F.Ranges.applyFormatFromDictionary(rng, getConditionalStyles(styles, arrayTag))
        
        '[Apply header styling]
        Stop
        Debug.Print "TM refactoring 2023-10-27"
        'headerLast = F.dictionaries.getItem(dictConfig, Props.headerLast.getName)
        'If Not headerLast Then Call applyHeaderConfig(rng, dictConfig)
        Call applyHeaderConfig(rng, dictConfig)
        
        '[Apply rows & columns styling]
        Stop 'TM - refactoring
        Debug.Print "TM refactoring 2023-10-27"
        'rowsFirst = F.dictionaries.getItem(dictConfig, Props.rowsFirst.getName)
        Set rows = F.dictionaries.GetObject(dictConfig, props.rows.getName)
        Set columns = F.dictionaries.GetObject(dictConfig, props.columns.getName)
        
        If rowsFirst Then
            Call applyRowsStyling(rng, rows)
            Call applyColumnsStyling(rng, columns)
        Else
            Call applyColumnsStyling(rng, columns)
            Call applyRowsStyling(rng, rows)
        End If
        
        'If headerLast Then Call applyHeaderConfig(rng, dictConfig)
        
    End If
    
End Sub

Private Function getConditionalStyles(styles As Scripting.Dictionary, arrayTag As String) As Scripting.Dictionary
    Dim dict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Set dict = F.dictionaries.GetObject(styles, props.StylesConditional.getName)
    If F.dictionaries.isNonEmptyDictionary(dict) Then
        Set getConditionalStyles = F.dictionaries.GetObject(dict, arrayTag)
    End If
End Function

Private Sub applyHeaderConfig(rng As Excel.Range, dictConfig As Scripting.Dictionary)
    Dim headerConfig As Scripting.Dictionary
    Dim styles As Scripting.Dictionary
    Dim rngHeader As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    Set headerConfig = F.dictionaries.GetObject(dictConfig, props.header.getName)
    If F.dictionaries.isNonEmptyDictionary(headerConfig) Then
        Set styles = F.dictionaries.GetObject(headerConfig, props.styles.getName)
        If F.dictionaries.isNonEmptyDictionary(styles) Then
            If F.Ranges.isValid(rng) Then
                Set rngHeader = rng.resize(1)
                Call F.Ranges.applyFormatFromDictionary(rngHeader, styles)
            End If
        End If
    End If

End Sub

Private Sub applyColumnsStyling(rng As Excel.Range, dtos As VBA.Collection)
    Dim styleDto As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    If F.Collections.isNonEmptyCollection(dtos) Then
        For Each styleDto In dtos
            Call applySingleColumnStyle(rng, styleDto)
        Next styleDto
    End If
End Sub

Private Sub applyRowsStyling(rng As Excel.Range, dtos As VBA.Collection)
    Dim styleDto As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    If F.Collections.isNonEmptyCollection(dtos) Then
        For Each styleDto In dtos
            Call applySingleRowStyle(rng, styleDto)
        Next styleDto
    End If
End Sub

Private Sub applySingleColumnStyle(rng As Excel.Range, columnDto As Scripting.Dictionary)
    Dim rngColumn As Excel.Range
    '------------------------------------------------------------------------------------------------------
    Dim index As Long
    Dim fromIndex As Long
    Dim toIndex As Long
    Dim styles As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim subranges As VBA.Collection
    Dim dtoSubrange As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------

    index = F.dictionaries.getItem(columnDto, props.index.getDbField)
    Debug.Print "TM refactoring 2023-10-27"
    'fromIndex = F.dictionaries.getItem(columnDto, Props.IndexFrom.getDbField)
    'toIndex = F.dictionaries.getItem(columnDto, Props.IndexTo.getDbField)
    
    Set styles = F.dictionaries.GetObject(columnDto, props.styles.getName)

    If index <> 0 Then
        Set rngColumn = rng.Offset(, calculateIndexNumber(rng, index, False) - 1).resize(, 1)
    ElseIf fromIndex <> 0 And toIndex <> 0 Then
        fromIndex = calculateIndexNumber(rng, fromIndex, False)
        toIndex = calculateIndexNumber(rng, toIndex, False)
        On Error Resume Next
        Set rngColumn = rng.Offset(, fromIndex - 1).resize(, toIndex - fromIndex + 1)
    End If
    
    If Not F.Ranges.isValid(rngColumn) Then
        ' skip - invalid range
    ElseIf F.Ranges.Intersect(rngColumn, rng) Is Nothing Then
        ' skip - subrange outside of main range
    Else
        Call F.Ranges.applyFormatFromDictionary(rngColumn, styles)
    End If
    
    '[Subranges]
    Debug.Print "TM refactoring 2023-10-27"
    'Set subranges = F.dictionaries.GetObject(columnDto, Props.subranges.getDbField)
    If Not subranges Is Nothing Then
        For Each dtoSubrange In subranges
            Call applySingleRowStyle(rngColumn, dtoSubrange)
        Next dtoSubrange
    End If
    
End Sub


Private Sub applySingleRowStyle(rng As Excel.Range, rowDto As Scripting.Dictionary)
    Dim rngRow As Excel.Range
    '------------------------------------------------------------------------------------------------------
    Dim index As Long
    Dim fromIndex As Long
    Dim toIndex As Long
    Dim styles As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim subranges As VBA.Collection
    Dim dtoSubrange As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------

    index = F.dictionaries.getItem(rowDto, props.index.getDbField)
    Debug.Print "TM refactoring 2023-10-27"
    'fromIndex = F.dictionaries.getItem(rowDto, Props.IndexFrom.getDbField)
    'toIndex = F.dictionaries.getItem(rowDto, Props.IndexTo.getDbField)
    
    Set styles = F.dictionaries.GetObject(rowDto, props.styles.getName)
    
    '[Find the range]
    If index <> 0 Then
        Set rngRow = rng.Offset(calculateIndexNumber(rng, index, True) - 1).resize(1)
    ElseIf fromIndex <> 0 And toIndex <> 0 Then
        fromIndex = calculateIndexNumber(rng, fromIndex, True)
        toIndex = calculateIndexNumber(rng, toIndex, True)
        On Error Resume Next
        Set rngRow = rng.Offset(fromIndex - 1).resize(toIndex - fromIndex + 1)
    End If
    
    If Not F.Ranges.isValid(rngRow) Then
        ' skip - invalid range
    ElseIf F.Ranges.Intersect(rngRow, rng) Is Nothing Then
        ' skip - subrange outside of main range
    Else
        Call F.Ranges.applyFormatFromDictionary(rngRow, styles)
    End If
    
    '[Subranges]
    Debug.Print "TM refactoring 2023-10-27"
    'Set subranges = F.dictionaries.GetObject(rowDto, Props.subranges.getDbField)
    If Not subranges Is Nothing Then
        For Each dtoSubrange In subranges
            Call applySingleColumnStyle(rngRow, dtoSubrange)
        Next dtoSubrange
    End If
    
End Sub

Private Function calculateIndexNumber(rng As Excel.Range, index As Long, Optional checkRows As Boolean = False) As Long
    If F.Ranges.isValid(rng) Then
        If index > 0 Then
            calculateIndexNumber = index
        ElseIf index < 0 Then
            If checkRows Then
                calculateIndexNumber = rng.rows.Count + index + 1
            Else
                calculateIndexNumber = rng.columns.Count + index + 1
            End If
        End If
    End If
End Function





Public Sub applyFreezing(wks As Excel.Worksheet, dictConfig As Scripting.Dictionary)
    Const FREEZING_TAG As String = "freeze"
    '------------------------------------------------------------------------------------------------------
    Dim dict As Scripting.Dictionary
    Dim cell As Excel.Range
    Dim rows As Long
    Dim columns As Long
    '------------------------------------------------------------------------------------------------------
    
    Set dict = F.dictionaries.GetObject(dictConfig, FREEZING_TAG)
    If Not dict Is Nothing Then
        rows = F.dictionaries.getItem(dict, props.rows.getName)
        columns = F.dictionaries.getItem(dict, props.columns.getName)
        
        If rows > 0 Or columns > 0 Then
            On Error Resume Next
            Call wks.Activate
            Set cell = wks.Cells(rows + 1, columns + 1)
            Call cell.Select
            Excel.ActiveWindow.FreezePanes = True
        End If
    End If
    
End Sub

