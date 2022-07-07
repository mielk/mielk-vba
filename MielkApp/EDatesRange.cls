VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EDatesRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IEditItem

Private Const CLASS_NAME As String = "EDatesRange"
'----------------------------------------------------------------------------------------------------------
Private Const ERROR_EMPTY_FIELD As String = "DatesRange.ValidationErrors.EmptyField"
Private Const ERROR_INVALID_DATE As String = "DatesRange.ValidationErrors.InvalidDateFormat"
Private Const ERROR_INVALID_MONTH As String = "DatesRange.ValidationErrors.InvalidMonthFormat"
Private Const ERROR_START_DATE_EARLIER_THAN_END_DATE As String = "DatesRange.ValidationErrors.StartDateLaterThanEndDate"
Private Const ERROR_END_DATE_EARLIER_THAN_START_DATE As String = "DatesRange.ValidationErrors.EndDateEarlierThanStartDate"
Private Const ERROR_MAX_MONTHS_RANGE_EXCEEDED As String = "DatesRange.ValidationErrors.MaxMonthsRangeExceeded"
Private Const ERROR_MAX_DAYS_RANGE_EXCEEDED As String = "DatesRange.ValidationErrors.MaxDaysRangeExceeded"
Private Const EDIT_DATES_RANGE_WINDOW_TITLE As String = "DatesRange.Edit.Window.DatesTitle"
Private Const EDIT_MONTHS_RANGE_WINDOW_TITLE As String = "DatesRange.Edit.Window.MonthsTitle"
'----------------------------------------------------------------------------------------------------------
Private Const TAG_START_DATE_STR As String = "StartDateStr"
Private Const TAG_START_DATE As String = "StartDate"
Private Const TAG_END_DATE_STR As String = "EndDateStr"
Private Const TAG_END_DATE As String = "EndDate"
'----------------------------------------------------------------------------------------------------------
Private pByMonth As Boolean
Private pStartDate As Date
Private pEndDate As Date
Private pMaxRange As Long
'----------------------------------------------------------------------------------------------------------




'[Setters]
Public Function setByMonths(value As Boolean) As EDatesRange
    Set setByMonths = Me
    pByMonth = value
End Function

Public Function setStartDate(value As Date) As EDatesRange
    Set setStartDate = Me
    pStartDate = value
End Function

Private Function setStartDateByString(value As String) As EDatesRange
    Set setStartDateByString = Me
    If pByMonth Then
        pStartDate = F.Dates.getDateFromMonthYearString(value)
    Else
        pStartDate = F.Dates.fromYYYYMMDDString(value)
    End If
End Function

Public Function setEndDate(value As Date) As EDatesRange
    Set setEndDate = Me
    pEndDate = value
End Function

Private Function setEndDateByString(value As String) As EDatesRange
    Set setEndDateByString = Me
    If pByMonth Then
        pEndDate = F.Dates.getDateFromMonthYearString(value)
    Else
        pEndDate = F.Dates.fromYYYYMMDDString(value)
    End If
End Function

Public Function setMaxRange(value As Long) As EDatesRange
    Set setMaxRange = Me
    pMaxRange = value
End Function

Public Function setValue(propName As String, value As Variant) As EDatesRange
    Select Case propName
        Case TAG_START_DATE:            Call setStartDate(VBA.CDate(value))
        Case TAG_START_DATE_STR:        Call setStartDateByString(VBA.CStr(value))
        Case TAG_END_DATE:              Call setEndDate(VBA.CDate(value))
        Case TAG_END_DATE_STR:          Call setEndDateByString(VBA.CStr(value))
    End Select
End Function



'[Getters]
Public Function getStartDateStr() As String
    If pStartDate <= 0 Then
        getStartDateStr = vbNullString
    ElseIf pByMonth Then
        getStartDateStr = F.Strings.addLeadingZeros(VBA.Month(pStartDate), 2) & "/" & VBA.Year(pStartDate)
    Else
        getStartDateStr = VBA.format(pStartDate, "YYYY-MM-DD")
    End If
End Function

Public Function getStartDate() As String
    getStartDate = pStartDate
End Function

Public Function getEndDateStr() As String
    If pEndDate <= 0 Then
        getEndDateStr = vbNullString
    ElseIf pByMonth Then
        getEndDateStr = F.Strings.addLeadingZeros(VBA.Month(pEndDate), 2) & "/" & VBA.Year(pEndDate)
    Else
        getEndDateStr = VBA.format(pEndDate, "YYYY-MM-DD")
    End If
End Function

Public Function getEndDate() As String
    getEndDate = pEndDate
End Function

Public Function getValue(propName As String) As Variant
    Select Case propName
        Case TAG_START_DATE:          getValue = getStartDate
        Case TAG_START_DATE_STR:      getValue = getStartDateStr
        Case TAG_END_DATE:            getValue = getEndDate
        Case TAG_END_DATE_STR:        getValue = getEndDateStr
    End Select
End Function

Public Function isDeclared() As Boolean
    If pStartDate <= 0 Then
        isDeclared = False
    ElseIf pEndDate <= 0 Then
        isDeclared = False
    Else
        isDeclared = True
    End If
End Function



'[Editing]
Public Function clone() As EDatesRange
    Set clone = New EDatesRange
    With clone
        Call .setStartDate(pStartDate)
        Call .setEndDate(pEndDate)
        Call .setByMonths(pByMonth)
        Call .setMaxRange(pMaxRange)
    End With
End Function

Public Function edit() As ValidationObject
    Dim editForm As WEditEntity
    Dim editedItem As EDatesRange
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Set editForm = getEditForm
    If Not editForm Is Nothing Then
        With editForm
            Call .run(vbModal, UI.Create.LocationPropsForScreenCenter)
            If Not .isCancelledByUser Then
                Set editedItem = .getEditItem
                Call saveChanges(editedItem)
            End If
            Set edit = F.Create.ValidationObject(Not .isCancelledByUser)
        End With
    End If
    
'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then
        Set edit = F.Create.ValidationObject(False, errDescription)
    End If
    Exit Function
    
'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = Err.Number: errDescription = Err.Description
    'call ErrorManager.addError(CLASS_NAME, metho
    GoTo ExitPoint
    
End Function

Private Sub saveChanges(editedItem As EDatesRange)
    With editedItem
        Call setStartDate(.getStartDateStr)
        Call setEndDate(.getEndDateStr)
    End With
End Sub


Private Function getEditForm() As WEditEntity
    Set getEditForm = MApp.Create.EditEntityWindow
    With getEditForm
        Call .setWindowTitle(getWindowTitle)
        Call .setItem(Me)
        Call .setLinesDefinitions(getEditLinesDefinitions(.getEditItem))
    End With
End Function

Private Function getWindowTitle() As String
    getWindowTitle = MsgService.getText(EDIT_DATES_RANGE_WINDOW_TITLE)
End Function

Private Function getEditLinesDefinitions(editItem As EDatesRange) As VBA.Collection
    Dim validationParams As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim startLine As Scripting.Dictionary
    Dim endLine As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    Set getEditLinesDefinitions = New VBA.Collection
    Set startLine = F.dictionaries.createWithItems(False, _
                        KeyValue(TAG_CONTROL_TYPE, ControlType_TextBox), _
                        KeyValue(TAG_CAPTION, MsgService.getText("DatesRange.Labels.StartDate")), _
                        KeyValue(TAG_PROPERTY, TAG_START_DATE_STR), _
                        KeyValue(TAG_VALIDATION_FUNCTION, "validateStartDate"), _
                        KeyValue(TAG_VALIDATION_PARAMS, validationParams), _
                        KeyValue(TAG_VALIDATION_LINKS, F.Collections.Create(TAG_END_DATE_STR)), _
                        KeyValue(TAG_CONFIRM_WITH_ENTER, True))
    Set endLine = F.dictionaries.createWithItems(False, _
                        KeyValue(TAG_CONTROL_TYPE, ControlType_TextBox), _
                        KeyValue(TAG_CAPTION, MsgService.getText("DatesRange.Labels.EndDate")), _
                        KeyValue(TAG_PROPERTY, TAG_END_DATE_STR), _
                        KeyValue(TAG_VALIDATION_FUNCTION, "validateEndDate"), _
                        KeyValue(TAG_VALIDATION_PARAMS, validationParams), _
                        KeyValue(TAG_VALIDATION_LINKS, F.Collections.Create(TAG_START_DATE_STR)), _
                        KeyValue(TAG_CONFIRM_WITH_ENTER, True))
    
    '[Adjust lines definitions if this should by dates range instead of months range]
    If Not pByMonth Then
        '[Start line]
        Call F.dictionaries.addItem(startLine, TAG_CONTROL_TYPE, ControlType_LabelWithActionButton)
        Call F.dictionaries.addItem(startLine, TAG_PICTURE, UI.Icons.icoCalendar.Picture)
        Call F.dictionaries.addItem(startLine, TAG_ACTION_PARAMS, F.dictionaries.createWithItems(False, _
                                                    KeyValue(TAG_PARENT, editItem), _
                                                    KeyValue(TAG_METHOD_NAME, "selectDate"), _
                                                    KeyValue(TAG_METHOD_PARAMS, F.dictionaries.createSingleEntryDict(TAG_PROPERTY, TAG_START_DATE_STR))))
        '[End line]
        Call F.dictionaries.addItem(endLine, TAG_CONTROL_TYPE, ControlType_LabelWithActionButton)
        Call F.dictionaries.addItem(endLine, TAG_PICTURE, UI.Icons.icoCalendar.Picture)
        Call F.dictionaries.addItem(endLine, TAG_ACTION_PARAMS, F.dictionaries.createWithItems(False, _
                                                    KeyValue(TAG_PARENT, editItem), _
                                                    KeyValue(TAG_METHOD_NAME, "selectDate"), _
                                                    KeyValue(TAG_METHOD_PARAMS, F.dictionaries.createSingleEntryDict(TAG_PROPERTY, TAG_END_DATE_STR))))
    End If
    
                            

    Set getEditLinesDefinitions = F.Collections.Create(startLine, endLine)
    
End Function

Public Function selectDate(Optional params As Scripting.Dictionary) As Variant
    Dim window As MielkUI.WDatePicker
    Dim propName As String
    Dim initialValue As String
    Dim initialDate As Date
    '------------------------------------------------------------------------------------------------------
    propName = F.dictionaries.getItem(params, TAG_PROPERTY)
    initialValue = VBA.CStr(getValue(propName))
    If F.Dates.isYYYYMMDDString(initialValue) Then initialDate = F.Dates.fromYYYYMMDDString(initialValue)
    Set window = UI.Create.DatePicker(initialValue)
    With window
        Call .display(vbModal, UI.Create.LocationPropsForCurrentCursor)
        selectDate = .getDate
    End With
End Function



'[Validation]
Public Function getValidation(validationFunction As String, Optional validationParams As Scripting.Dictionary) As ValidationObject
    Set getValidation = VBA.CallByName(Me, validationFunction, VbMethod, validationParams)
End Function

Public Function validateStartDate(Optional validationParams As Scripting.Dictionary) As ValidationObject
    Static names As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim value As String
    '------------------------------------------------------------------------------------------------------
    
    value = F.dictionaries.getItem(validationParams, VALUE_____)
    
    'If VBA.Len(value) = 0 Then
    '    Set validateStartDate = F.Create.ValidationObject(False, MsgService.getText(ERROR_EMPTY_FIELD))
    If pByMonth And Not F.regex.checkIfMatch(value, MONTH_YEAR_REGEX_PATTERN) Then
        Set validateStartDate = F.Create.ValidationObject(False, MsgService.getText(ERROR_INVALID_MONTH))
    ElseIf Not pByMonth And Not F.Dates.isYYYYMMDDString(value) Then
        Set validateStartDate = F.Create.ValidationObject(False, MsgService.getText(ERROR_INVALID_DATE))
    ElseIf Not isEarlierThanEndDate(value) Then
        Set validateStartDate = F.Create.ValidationObject(False, MsgService.getText(ERROR_START_DATE_EARLIER_THAN_END_DATE))
    Else
        Set validateStartDate = F.Create.ValidationObject(True)
    End If
    
End Function

Public Function validateEndDate(Optional validationParams As Scripting.Dictionary) As ValidationObject
    Dim value As String
    '------------------------------------------------------------------------------------------------------
    
    value = F.dictionaries.getItem(validationParams, VALUE_____)
    
    'If VBA.Len(value) = 0 Then
    '    Set validateEndDate = F.Create.ValidationObject(False, MsgService.getText(ERROR_EMPTY_FIELD))
    If pByMonth And Not F.regex.checkIfMatch(value, MONTH_YEAR_REGEX_PATTERN) Then
        Set validateEndDate = F.Create.ValidationObject(False, MsgService.getText(ERROR_INVALID_MONTH))
    ElseIf Not pByMonth And Not F.Dates.isYYYYMMDDString(value) Then
        Set validateEndDate = F.Create.ValidationObject(False, MsgService.getText(ERROR_INVALID_DATE))
    ElseIf Not isLaterThanStartDate(value) Then
        Set validateEndDate = F.Create.ValidationObject(False, MsgService.getText(ERROR_END_DATE_EARLIER_THAN_START_DATE))
    ElseIf pMaxRange > 0 Then
        If isMaxRangeExceeded(value) Then
            If pByMonth Then
                Set validateEndDate = F.Create.ValidationObject(False, F.Strings.format(MsgService.getText(ERROR_MAX_MONTHS_RANGE_EXCEEDED), pMaxRange))
            Else
                Set validateEndDate = F.Create.ValidationObject(False, F.Strings.format(MsgService.getText(ERROR_MAX_DAYS_RANGE_EXCEEDED), pMaxRange))
            End If
        Else
            Set validateEndDate = F.Create.ValidationObject(True)
        End If
    Else
        Set validateEndDate = F.Create.ValidationObject(True)
    End If
    
End Function

Private Function isEarlierThanEndDate(value As String) As Boolean
    Dim startDate As Date
    '------------------------------------------------------------------------------------------------------
    If pEndDate > 0 Then
        If pByMonth Then
            startDate = F.Dates.getDateFromMonthYearString(value)
        Else
            startDate = F.Dates.fromYYYYMMDDString(value)
        End If
        isEarlierThanEndDate = (startDate <= pEndDate)
    Else
        isEarlierThanEndDate = True
    End If
End Function

Private Function isLaterThanStartDate(value As String) As Boolean
    Dim endDate As Date
    '------------------------------------------------------------------------------------------------------
    If pStartDate > 0 Then
        If pByMonth Then
            endDate = F.Dates.getDateFromMonthYearString(value)
        Else
            endDate = F.Dates.fromYYYYMMDDString(value)
        End If
        isLaterThanStartDate = (pStartDate <= endDate)
    Else
        isLaterThanStartDate = True
    End If
End Function

Private Function isMaxRangeExceeded(strEndDate As String) As Boolean
    Dim endDate As Date
    Dim interval As String
    '------------------------------------------------------------------------------------------------------
    
    If pMaxRange = 0 Then
        isMaxRangeExceeded = False
    ElseIf pStartDate <= 0 Then
        isMaxRangeExceeded = False
    Else

        If pByMonth Then
            endDate = F.Dates.getDateFromMonthYearString(strEndDate)
        Else
            endDate = F.Dates.fromYYYYMMDDString(strEndDate)
        End If
        
        If endDate <= 0 Then
            isMaxRangeExceeded = False
        Else
            interval = VBA.IIf(pByMonth, "m", "d")
            isMaxRangeExceeded = VBA.DateDiff(interval, pStartDate, endDate) > pMaxRange
        End If
    End If
    
End Function



'[Implementation of IEditItem]
Private Function IEditItem_clone() As MielkCore.IEditItem
    Set IEditItem_clone = clone
End Function

Private Sub IEditItem_delete()
    'N/A
End Sub

Private Function IEditItem_edit(Optional params As Scripting.Dictionary) As MielkCore.ValidationObject
    Set IEditItem_edit = edit
End Function

Private Function IEditItem_getId() As Long
    'N/A
End Function

Private Function IEditItem_getName() As String
    'N/A
End Function

Private Function IEditItem_getObjectsAffected() As Collection
    Set IEditItem_getObjectsAffected = New VBA.Collection
End Function

Private Function IEditItem_getValidation(validationFunction As String, Optional validationParams As Scripting.IDictionary) As MielkCore.ValidationObject
    Set IEditItem_getValidation = getValidation(validationFunction, validationParams)
End Function

Private Function IEditItem_getValue(valueKey As String) As Variant
    Call F.Variables.assign(IEditItem_getValue, getValue(valueKey))
End Function

Private Sub IEditItem_runAction(actionName As String, Optional params As Scripting.IDictionary)
    Call VBA.CallByName(Me, actionName, VbMethod, params)
End Sub

Private Sub IEditItem_updateValue(valueKey As String, value As Variant)
    Call setValue(valueKey, value)
End Sub
