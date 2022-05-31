VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "ExcelFilter"
'----------------------------------------------------------------------------------------------------------
Private Const FILTER_OPERATOR As String = "operator"
Private Const CRITERIA_1 As String = "criteria1"
Private Const CRITERIA_2 As String = "criteria2"
Private Const FROM_SHEET As String = "fromSheet"
Private Const IS_DATE As String = "isDate"
'----------------------------------------------------------------------------------------------------------
Private pSheet As Excel.Worksheet
Private pHasFilter As Boolean
Private pIsActive As Boolean
Private pRange As Excel.Range
Private pFilters As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------



'[SETTERS]

Public Function setSheet(value As Excel.Worksheet) As ExcelFilter
    Set setSheet = Me
    Set pSheet = value
End Function

Public Sub readFromSheet(Optional wks As Excel.Worksheet)
    Dim filter As Excel.filter
    Dim counter As Long
    '------------------------------------------------------------------------------------------------------
    
    If Not wks Is Nothing Then Set pSheet = wks
    
    If F.Sheets.isValid(pSheet) Then
        With pSheet
            '[Meta properties]
            pHasFilter = .AutoFilterMode
            pIsActive = .FilterMode
            
            '[Filters properties]
            If pIsActive Then
                With .AutoFilter
                    Set pRange = .Range
                    Set pFilters = F.Dictionaries.Create(False)
                    For Each filter In .Filters
                        counter = counter + 1
                        If filter.On Then
                            Call pFilters.Add(counter, getFilterPropsDictionary(filter, counter))
                        End If
                    Next filter
                End With
            End If
            
        End With
    End If
    
End Sub

Private Function getFilterPropsDictionary(filter As Excel.filter, index As Long) As Scripting.Dictionary
    Dim filterActive As Boolean
    Dim dict As Scripting.Dictionary
    Dim criteria As Scripting.Dictionary
    Dim key As Variant
    Dim criteriaValue As Variant
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    filterActive = filter.On
    '###if not errHandler
    
    If filterActive Then
        Set getFilterPropsDictionary = F.Dictionaries.Create(False)
        
        With filter
            Call getFilterPropsDictionary.Add(FILTER_OPERATOR, F.Math.coalesce(.operator, xlFilterValues))
            '[Criteria1]
            Set criteria = F.Dictionaries.Create(False)
            With criteria
                On Error Resume Next
                Call .Add(CRITERIA_1, filter.criteria1)
                Call .Add(CRITERIA_2, filter.criteria2)
                '###if not errHandler
                
                If F.Dictionaries.isNonEmptyDictionary(criteria) Then
                    For Each key In .Keys
                        If Not F.Variables.isUndefined(.item(key)) Then
                            Call F.Variables.assign(criteriaValue, getFilterCriteriaValue(.item(key), filter))
                            If Not F.Variables.isUndefined(criteriaValue) Then
                                Call getFilterPropsDictionary.Add(key, F.Dictionaries.createWithItems(False, KeyValue(VALUE_____, criteriaValue)))
                            End If
                        End If
                    Next key
                Else
                    Call F.Variables.assign(criteriaValue, getCriteriaValueFromSheet(filter, index))
                    If Not F.Variables.isUndefined(criteriaValue) Then
                        Call getFilterPropsDictionary.Add(FROM_SHEET, criteriaValue)
                    End If
                End If
                
            End With
        End With
    End If
    
End Function

Private Function getFilterCriteriaValue(value As Variant, filter As Excel.filter) As Variant
    Select Case filter.operator
        Case xlFilterValues:
            getFilterCriteriaValue = value
        Case xlFilterCellColor:
            getFilterCriteriaValue = value.color
        Case Else:
            getFilterCriteriaValue = value
    End Select
End Function

Private Function getCriteriaValueFromSheet(filter As Excel.filter, index As Long) As Scripting.Dictionary
    Dim parent As AutoFilter
    Dim rng As Excel.Range
    Dim visibleCells As Excel.Range
    Dim cell As Excel.Range
    Dim dict As Scripting.Dictionary
    Dim isDateValue As Boolean
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    Set parent = filter.parent
    
    Set getCriteriaValueFromSheet = F.Dictionaries.Create(False)
    
    If Not parent Is Nothing Then
        Set rng = parent.Range
        If F.ranges.isValid(rng) Then
            If rng.columns.count >= index Then
                On Error Resume Next
                Set visibleCells = rng.columns(index).SpecialCells(xlCellTypeVisible)
                '### errHandler
            End If
        End If
        
        Set dict = F.Dictionaries.Create(False)
        If F.ranges.isValid(visibleCells) Then
            For Each cell In visibleCells
                If cell.row <> rng.row Then
                    isDateValue = VBA.IIf(isDateValue, isDateValue, VBA.isDate(cell.value))
                    If Not dict.exists(cell.value2) Then
                        Call dict.Add(cell.value2, vbNullString)
                    End If
                End If
            Next cell
        End If
        
    End If
    
    If F.Dictionaries.isNonEmptyDictionary(dict) Then
        Call getCriteriaValueFromSheet.Add(VALUE_____, dict.Keys)
        Call getCriteriaValueFromSheet.Add(IS_DATE, isDateValue)
    End If
    
End Function

Public Sub removeFilter()
    If F.Sheets.isValid(pSheet) Then
        If pSheet.FilterMode Then
            pSheet.AutoFilterMode = False
        End If
    End If
End Sub

Public Sub applyFilter()
    Dim index As Variant
    Dim dictCriteria As Scripting.Dictionary
    Dim operator As XlAutoFilterOperator
    Dim criteria1 As Scripting.Dictionary
    Dim criteria2 As Scripting.Dictionary
    Dim fromSheet As Scripting.Dictionary
    Dim isDate As Boolean
    '------------------------------------------------------------------------------------------------------
    
    If F.Sheets.isValid(pSheet) Then
        If pHasFilter Then
            If pIsActive Then
                pSheet.AutoFilterMode = False
                Call pRange.AutoFilter
                
                Call pRange.AutoFilter
                If F.Dictionaries.isNonEmptyDictionary(pFilters) Then
                    For Each index In pFilters.Keys
                        If index > 0 And index < pRange.columns.count Then
                            Set dictCriteria = pFilters.item(index)
                            operator = F.Dictionaries.getItem(dictCriteria, FILTER_OPERATOR)
                            Set criteria1 = F.Dictionaries.getObject(dictCriteria, CRITERIA_1)
                            Set criteria2 = F.Dictionaries.getObject(dictCriteria, CRITERIA_2)
                            Set fromSheet = F.Dictionaries.getObject(dictCriteria, FROM_SHEET)
                            isDate = F.Dictionaries.getItem(fromSheet, IS_DATE)
                            
                            If Not criteria2 Is Nothing Then
                                Call pRange.AutoFilter(Field:=index, criteria1:=getCriteriaValueFromJson(F.Dictionaries.getItem(dictCriteria, CRITERIA_1)), operator:=operator, criteria2:=getCriteriaValueFromJson(F.Dictionaries.getItem(dictCriteria, CRITERIA_2)))
                            ElseIf Not criteria1 Is Nothing Then
                                Call pRange.AutoFilter(Field:=index, criteria1:=getCriteriaValueFromJson(F.Dictionaries.getItem(dictCriteria, CRITERIA_1)), operator:=operator)
                            ElseIf Not fromSheet Is Nothing Then
                                Call pRange.AutoFilter(Field:=index, criteria2:=getCriteriaValueFromJson(F.Dictionaries.getItem(dictCriteria, CRITERIA_1), isDate), operator:=operator)
                            End If
                        End If
                    Next index
                    
                End If
                
            End If
        End If
    End If
    
End Sub

Public Function getCriteriaValueFromJson(dict As Scripting.Dictionary, Optional isDate As Boolean = False) As Variant
    If F.Dictionaries.isNonEmptyDictionary(dict) Then
        If Not dict.exists(VALUE_____) Then
        ElseIf isDate Then
            getCriteriaValueFromJson = createDatesFilterArray(dict.item(VALUE_____))
        ElseIf VBA.IsObject(dict.item(VALUE_____)) Then
            Set getCriteriaValueFromJson = dict.item(VALUE_____)
        Else
            getCriteriaValueFromJson = dict.item(VALUE_____)
            If F.Variables.isString(getCriteriaValueFromJson) Then
                If F.strings.startsWith(VBA.CStr(getCriteriaValueFromJson), "=") Then
                    getCriteriaValueFromJson = VBA.mid$(getCriteriaValueFromJson, 2)
                End If
            End If
        End If
    End If
End Function

Public Function createDatesFilterArray(arr As Variant) As Variant
    Dim i As Long
    Dim col As VBA.Collection
    Dim d As Date
    '------------------------------------------------------------------------------------------------------
    
    If F.arrays.countDimensions(arr) = 1 Then
        Set col = New VBA.Collection
        For i = LBound(arr) To UBound(arr)
            d = arr(i)
            Call col.Add(2)
            Call col.Add(VBA.Replace(VBA.format(d, "m/d/yyyy"), "-", "/"))
        Next i
    End If
    
    createDatesFilterArray = F.Collections.toArray(col)
    
End Function
