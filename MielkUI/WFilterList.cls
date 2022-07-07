VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WFilterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IFilter

Private Const CLASS_NAME As String = "WFilterList"
'----------------------------------------------------------------------------------------------------------
Private Const WINDOW_TITLE_TAG As String = "FilterFormTitle"
'----------------------------------------------------------------------------------------------------------
Private pParent As IFilterable
Attribute pParent.VB_VarHelpID = -1
Private pKey As String
Private pIsActive As Boolean
Private pHasSearchField As Boolean
'[UI components] ------------------------------------------------------------------------------------------
Private pForm As WSelectItem
'[State] --------------------------------------------------------------------------------------------------
Private pColumnName As String
Private pSelectedItems As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------







'[SETTERS]
Public Function setParent(value As IFilterable) As WFilterList
    Set setParent = Me
    Set pParent = value
End Function

Public Function setKey(value As String) As WFilterList
    Set setKey = Me
    pKey = value
End Function

Public Function setIsActive(value As Boolean) As WFilterList
    Set setIsActive = Me
    pIsActive = value
End Function

Public Function setSelection(value As Variant) As WFilterList
    Set setSelection = Me
    If F.dictionaries.isDictionary(value) Then
        Set pSelectedItems = value
    End If
End Function

Public Function setColumnName(value As String) As WFilterList
    Set setColumnName = Me
    pColumnName = value
End Function

Public Function setHasSearchField(value As Boolean) As WFilterList
    Set setHasSearchField = Me
    pHasSearchField = value
End Function





'[GETTERS]
Public Function getKey() As String
    getKey = pKey
End Function

Public Function getSelectionString() As String
    If pIsActive Then
        getSelectionString = VBA.Join(pSelectedItems.keys, " | ")
    End If
End Function

Public Function getSelection() As Scripting.Dictionary
    Set getSelection = pSelectedItems
End Function

Public Function isActive() As Boolean
    isActive = pIsActive
End Function





Public Function match(value As Variant) As Boolean
    Dim key As Variant
    Dim regex As String
    Dim col As VBA.Collection
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    
    'Apply regex if it is set for this column ------------------------------------------------------------|
    regex = pParent.getFilterRegex(pKey)
    If VBA.Len(regex) = 0 Then
    ElseIf F.Variables.isContainer(value) Then
        Set col = F.Collections.createFrom(value)
        For Each item In value
            Call col.Add(F.regex.getFirstGroupMatch(VBA.CStr(item), regex))
        Next item
        Set value = col
    Else
        value = F.regex.getFirstGroupMatch(VBA.CStr(value), regex)
    End If
    '-----------------------------------------------------------------------------------------------------|
    
    If F.Variables.isContainer(value) Then
        If F.Collections.isNonEmptyCollection(value) Then
            match = F.Collections.hasAnyCommonItemWithDictionary(F.Cast.toCollection(value), pSelectedItems)
        End If
    ElseIf pSelectedItems.Exists(value) Then
        match = True
    End If
End Function

Public Sub edit()
    Set pForm = UI.Create.SelectItemWindow(getItems). _
                    setInitialSelection(pSelectedItems). _
                    setTitle(F.Strings.format(MsgService.GetText(WINDOW_TITLE_TAG), pColumnName)). _
                    setMultiSelectAllowed(True). _
                    setHasSearchPanel(True). _
                    setHasTitlePanel(False)
    With pForm
        Call .display(UI.Create.LocationPropsForCurrentCursor)
        If Not .isCancelledByUser Then
            Set pSelectedItems = .getSelectionDictionary
            If F.dictionaries.isNonEmptyDictionary(pSelectedItems) Then
                Call update
            Else
                Call clear
            End If
        End If
    End With
End Sub


Public Function getItems() As Scripting.Dictionary
    Dim dict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Set dict = pParent.getAvailableValues(pKey)
    Set getItems = F.dictionaries.sortByKey(dict, True)
End Function

Public Sub clear()
    pIsActive = False
    Call pParent.removeFilter(Me)
End Sub

Public Sub update()
    pIsActive = F.dictionaries.isNonEmptyDictionary(pSelectedItems)
    Call pParent.applyFilter(Me)
End Sub




'[Implementation of IFilter]
Private Function IFilter_setKey(value As String) As IFilter
    Set IFilter_setKey = setKey(value)
End Function

Private Function IFilter_setParent(value As IFilterable) As IFilter
    Set IFilter_setParent = setParent(value)
End Function

Private Function IFilter_setColumnName(value As String) As IFilter
    Set IFilter_setColumnName = setColumnName(value)
End Function

Private Function IFilter_setIsActive(value As Boolean) As IFilter
    Set IFilter_setIsActive = setIsActive(value)
End Function

Private Function IFilter_setSelection(value As Variant) As IFilter
    Set IFilter_setSelection = setSelection(value)
End Function






Private Function IFilter_match(value As Variant) As Boolean
    IFilter_match = match(value)
End Function

Private Sub IFilter_clear()
    Call clear
End Sub

Private Sub IFilter_edit()
    Call edit
End Sub

Private Function IFilter_getKey() As String
    IFilter_getKey = getKey
End Function

Private Function IFilter_isActive() As Boolean
    IFilter_isActive = isActive
End Function

Private Function IFilter_getSelectionString() As String
    IFilter_getSelectionString = getSelectionString
End Function

Private Function IFilter_getSelection() As Variant
    Call F.Variables.assign(IFilter_getSelection, getSelection)
End Function
