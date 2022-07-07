VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Standarizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "Standarizer"
'----------------------------------------------------------------------------------------------------------
Private Const DEFAULT_STANDARD_NAMES_LIST_TITLE As String = "NamesStandarizer.Labels.StandardNamesListTitle"
Private Const DEFAULT_ALIASES_LIST_TITLE As String = "NamesStandarizer.Labels.AliasesListTitle"
Private Const DEFAULT_UNASSIGNED_TAG As String = "NamesStandarizer.Labels.UnassignedAliases"
Private Const DEFAULT_JUNK_TAG As String = "NamesStandarizer.Labels.JunkAliases"
Private Const DEFAULT_UNASSIGNED_WARNING As String = "NamesStandarizer.Warnings.UnassignedItems"
'----------------------------------------------------------------------------------------------------------
Private pParent As IStandarizerParent
Private pDisplayProperty As String
Private pObjects As Scripting.Dictionary            'IAliasable objects
Private pAliases As Scripting.Dictionary            'Dictionary alias -> IAliasable
Private pNames As VBA.Collection                    'All names to be assigned
'----------------------------------------------------------------------------------------------------------
Private pUnassigned As Scripting.Dictionary
Private pJunk As Scripting.Dictionary
'[View] ---------------------------------------------------------------------------------------------------
Private pViewer As WStandarizer
Attribute pViewer.VB_VarHelpID = -1
'[Properties] ---------------------------------------------------------------------------------------------
Private pHasJunk As Boolean
Private pAliasesListTitle As String
Private pAliasesListWidth As Single
Private pStandardNamesListTitle As String
Private pCreateNewItemButtonCaption As String
Private pUnassignedItemsWarning As String
Private pUnassignedGroupName As String
Private pJunkGroupName As String
'[State] --------------------------------------------------------------------------------------------------
Private pStandarizerContainerType As StandarizerContainerTypeEnum
Private pPreviewObject As IAliasable
Private pPreviewKey As Variant
Private pHistory As VBA.Collection
'[Events] -------------------------------------------------------------------------------------------------
Event PreviewChange()
Event AssignmentChange(previewKey As Variant, newKey As Variant)
Event AssignedToNewItem(item As IAliasable, previewKey As Variant)
Event ItemAdded(item As IAliasable)
Event AfterConfirm(objects As Scripting.Dictionary, aliases As Scripting.Dictionary, unassigned As VBA.Collection, junk As VBA.Collection)
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    
    '[Containers]
    Set pUnassigned = F.dictionaries.Create(False)
    Set pJunk = F.dictionaries.Create(False)
    Set pHistory = New VBA.Collection
    
    '[Default values]
    pHasJunk = True
    pAliasesListTitle = MsgService.getText(DEFAULT_ALIASES_LIST_TITLE)
    pStandardNamesListTitle = MsgService.getText(DEFAULT_STANDARD_NAMES_LIST_TITLE)
    pUnassignedGroupName = MsgService.getText(DEFAULT_UNASSIGNED_TAG)
    pJunkGroupName = MsgService.getText(DEFAULT_JUNK_TAG)
    pUnassignedItemsWarning = MsgService.getText(DEFAULT_UNASSIGNED_WARNING)
    
End Sub



'[Setup]
Public Function setup(parent As IStandarizerParent, displayProperty As String, _
                                    aliases As Scripting.Dictionary) As Standarizer
    Set setup = Me
    Set pParent = parent
    Set pObjects = pParent.getItems
    pDisplayProperty = displayProperty
    Set pAliases = aliases
    Call createJunkDictionary
End Function

Public Function setAliasesToBeAssigned(ByVal values As Variant) As Standarizer
    Dim value As Variant
    '----------------------------------------------------------------------------------------------------------
    
    Set setAliasesToBeAssigned = Me
    Set pUnassigned = F.dictionaries.Create(False)
    
    If F.Variables.isContainer(values) Then
        Set values = F.Collections.createFrom(values)
        If F.Collections.isNonEmptyCollection(values) Then
            For Each value In values
                If Not pAliases.Exists(value) Then
                    Call pUnassigned.Add(value, value)
                End If
            Next value
        End If
    End If
    
End Function

Private Sub createJunkDictionary()
    Dim key As Variant
    Dim id As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set pJunk = F.dictionaries.Create(False)
    If F.dictionaries.isNonEmptyDictionary(pAliases) Then
        For Each key In pAliases.Keys
            id = pAliases.item(key)
            If F.Variables.isUndefined(id) Then
                Call F.dictionaries.addItem(pJunk, key, key)
            End If
        Next key
    End If
End Sub

Private Sub reprocessJunkDictionary()
    If Not pHasJunk Then
        If F.dictionaries.isNonEmptyDictionary(pJunk) Then
            Call F.dictionaries.appendDictionary(pUnassigned, pJunk)
        End If
    End If
End Sub




'[Setters]
Public Function setHasJunkContainer(value As Boolean) As Standarizer
    Set setHasJunkContainer = Me
    pHasJunk = value
    Call reprocessJunkDictionary
End Function

Public Function setAliasesListTitle(value As String) As Standarizer
    Set setAliasesListTitle = Me
    pAliasesListTitle = value
End Function

Public Function setStandardNamesListTitle(value As String) As Standarizer
    Set setStandardNamesListTitle = Me
    pStandardNamesListTitle = value
End Function

Public Function setCreateNewItemButtonCaption(value As String) As Standarizer
    Set setCreateNewItemButtonCaption = Me
    pCreateNewItemButtonCaption = value
End Function

Public Function setUnassignedItemsGroupName(value As String) As Standarizer
    Set setUnassignedItemsGroupName = Me
    pUnassignedGroupName = value
End Function

Public Function setJunkItemsGroupName(value As String) As Standarizer
    Set setJunkItemsGroupName = Me
    pJunkGroupName = value
End Function

Public Function setUnassignedItemsWarning(value As String) As Standarizer
    Set setUnassignedItemsWarning = Me
    pUnassignedItemsWarning = value
End Function

Public Function setAliasesListWidth(value As Single) As Standarizer
    Set setAliasesListWidth = Me
    pAliasesListWidth = value
End Function


'[Getters]
Friend Function hasJunkContainer() As Boolean
    hasJunkContainer = pHasJunk
End Function

Friend Function getUnassignedContainer() As Scripting.Dictionary
    Set getUnassignedContainer = pUnassigned
End Function

Friend Function getJunkContainer() As Scripting.Dictionary
    Set getJunkContainer = pJunk
End Function

Friend Function getAliasesListHeaders() As VBA.Collection
    Set getAliasesListHeaders = New VBA.Collection
    With getAliasesListHeaders
        Call .Add(UI.Create.ListHeaderDefinition(pAliasesListTitle, TAG_ALIAS))
    End With
End Function

Friend Function getStandardNamesListHeaders() As VBA.Collection
    Set getStandardNamesListHeaders = New VBA.Collection
    With getStandardNamesListHeaders
        Call .Add(UI.Create.ListHeaderDefinition(pStandardNamesListTitle, TAG_NAME))
    End With
End Function

Friend Function getAttachedStandardNames() As Scripting.Dictionary
    Set getAttachedStandardNames = F.dictionaries.Create(False)
    With getAttachedStandardNames
        Call .Add(StandarizerContainerType_Unassigned, vbNullString)
        Call .Add(StandarizerContainerType_Junk, vbNullString)
    End With
End Function

Friend Function getStandardNamesListItems() As Scripting.Dictionary
    Dim key As Variant
    Dim obj As IAliasable
    Dim listLabel As String
    '------------------------------------------------------------------------------------------------------
    Set getStandardNamesListItems = F.dictionaries.Create(False)
    With getStandardNamesListItems
        Call .Add(StandarizerContainerType_Unassigned, pUnassigned)
        If pHasJunk Then Call .Add(StandarizerContainerType_Junk, pJunk)
        For Each key In pObjects.Keys
            Call .Add(key, pObjects.item(key))
        Next key
    End With
End Function

Friend Function getStandardNamesDisplayeValues() As Scripting.Dictionary
    Dim key As Variant
    Dim obj As IAliasable
    '------------------------------------------------------------------------------------------------------
    Set getStandardNamesDisplayeValues = F.dictionaries.Create(False)
    With getStandardNamesDisplayeValues
        Call .Add(StandarizerContainerType_Unassigned, _
                            getStandardNameListLabel(pUnassignedGroupName, pUnassigned.Count))
        If pHasJunk Then Call .Add(StandarizerContainerType_Junk, _
                            getStandardNameListLabel(pJunkGroupName, pJunk.Count))
        For Each key In pObjects.Keys
            Set obj = pObjects.item(key)
            Call .Add(key, getCaptionForStandardName(obj))
        Next key
    End With
End Function

Private Function getStandardNameListLabel(groupName As String, aliasesCounter As Long) As String
    getStandardNameListLabel = groupName & " (" & aliasesCounter & ")"
End Function

Friend Function getCaptionForStandardName(value As Variant) As String
    Dim item As IAliasable
    Dim label As String
    '------------------------------------------------------------------------------------------------------
    If F.Variables.isUndefined(value) Then
        'N/A
    ElseIf Not VBA.IsObject(value) Then
        'N/A
    ElseIf value Is pUnassigned Then
        getCaptionForStandardName = getStandardNameListLabel(pUnassignedGroupName, pUnassigned.Count)
    ElseIf value Is pJunk Then
        getCaptionForStandardName = getStandardNameListLabel(pJunkGroupName, pJunk.Count)
    Else
        Set item = value
        label = item.getLabel
        getCaptionForStandardName = getStandardNameListLabel(label, item.countAliases)
    End If
End Function

Friend Function getCaptionForUnassigned() As String
    getCaptionForUnassigned = getCaptionForStandardName(pUnassigned)
End Function

Friend Function getCaptionForJunk() As String
    getCaptionForJunk = getCaptionForStandardName(pJunk)
End Function

Friend Function getAliasesListTitle() As String
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    If pStandarizerContainerType = StandarizerContainerType_Unassigned Then
        getAliasesListTitle = pUnassignedGroupName
    ElseIf pStandarizerContainerType = StandarizerContainerType_Junk Then
        getAliasesListTitle = pJunkGroupName
    Else
        If Not pPreviewObject Is Nothing Then
            key = pPreviewObject.getLabel
            getAliasesListTitle = F.Strings.format(pAliasesListTitle, VBA.UCase$(key))
        End If
    End If
End Function

Friend Function getPreviewItemAliases() As Scripting.Dictionary
    Dim aliases As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If pStandarizerContainerType = StandarizerContainerType_Unassigned Then
        Set aliases = pUnassigned
    ElseIf pStandarizerContainerType = StandarizerContainerType_Junk Then
        Set aliases = pJunk
    Else
        If Not pPreviewObject Is Nothing Then
            Set aliases = F.Collections.toDictionary(pPreviewObject.getAliases)
        Else
            Set aliases = F.dictionaries.Create(False)
        End If
    End If
    
    Set getPreviewItemAliases = F.dictionaries.sortByKey(aliases, True)
    
End Function

Friend Function hasAnyUnassigned() As Boolean
    hasAnyUnassigned = F.dictionaries.isNonEmptyDictionary(pUnassigned)
End Function

Friend Function getWarningMessage() As String
    getWarningMessage = F.Strings.format(pUnassignedItemsWarning, pUnassigned.Count)
End Function

Friend Function getPreviewKey() As Variant
    Call F.Variables.assign(getPreviewKey, pPreviewKey)
End Function

Public Function getAliasesDictionary() As Scripting.Dictionary
    Set getAliasesDictionary = pAliases
End Function




'[Actions]
Public Sub run(Optional windowMode As FormShowConstants = vbModeless)
    Set pViewer = New WStandarizer
    With pViewer
        Call .setParent(Me)
        Call .setCreateNewItemButtonCaption(pCreateNewItemButtonCaption)
        If pAliasesListWidth Then Call .setAliasesListWidth(pAliasesListWidth)
        Call .run(windowMode, UI.Create.LocationPropsForScreenCenter)
        RaiseEvent PreviewChange
    End With
End Sub

Friend Sub changePreviewStandardName(value As Variant, key As Variant)
    If F.Variables.isUndefined(value) Then
        'N/A
    ElseIf Not VBA.IsObject(value) Then
        'N/A
    ElseIf value Is pUnassigned Then
        pStandarizerContainerType = StandarizerContainerType_Unassigned
    ElseIf value Is pJunk Then
        pStandarizerContainerType = StandarizerContainerType_Junk
    Else
        pStandarizerContainerType = StandarizerContainerType_AliasableObject
        Set pPreviewObject = value
    End If
    Call F.Variables.assign(pPreviewKey, key)
    RaiseEvent PreviewChange
End Sub

Friend Sub changeAssignment(aliases As VBA.Collection, value As Variant, key As Variant)
    Dim item As IAliasable
    '------------------------------------------------------------------------------------------------------
    
    If F.Variables.isUndefined(value) Then
        'N/A
    ElseIf Not VBA.IsObject(value) Then
        'N/A
    ElseIf value Is pUnassigned Then
        Call moveToPredefinedDictionary(aliases, pUnassigned)
    ElseIf value Is pJunk Then
        Call moveToPredefinedDictionary(aliases, pJunk)
    Else
        Set item = value
        Call assignTo(aliases, item)
    End If
    
    RaiseEvent AssignmentChange(pPreviewKey, key)
    
End Sub

Private Sub moveToPredefinedDictionary(aliases As VBA.Collection, dict As Scripting.Dictionary)
    Dim alias As Variant
    '------------------------------------------------------------------------------------------------------

    'Remove aliases from previous assignment
    If pStandarizerContainerType = StandarizerContainerType_Unassigned Then
        Call removeFromPredefinedDictionary(aliases, pUnassigned)
    ElseIf pStandarizerContainerType = StandarizerContainerType_Junk Then
        Call removeFromPredefinedDictionary(aliases, pJunk)
    Else
        Call removeFromObject(aliases, pPreviewObject)
    End If
    
    'Add aliases to new assignment
    For Each alias In aliases
        Call F.dictionaries.addItem(dict, alias, alias)
        Call addItemToHistory(VBA.CStr(alias), pPreviewObject, dict)
    Next alias
    
End Sub

Private Sub assignTo(aliases As VBA.Collection, item As IAliasable)
    Dim alias As Variant
    '------------------------------------------------------------------------------------------------------

    'Remove aliases from previous assignment
    If pStandarizerContainerType = StandarizerContainerType_Unassigned Then
        Call removeFromPredefinedDictionary(aliases, pUnassigned)
    ElseIf pStandarizerContainerType = StandarizerContainerType_Junk Then
        Call removeFromPredefinedDictionary(aliases, pJunk)
    Else
        Call removeFromObject(aliases, pPreviewObject)
    End If
    
    'Add aliases to new assignment
    For Each alias In aliases
        Call item.addAlias(VBA.CStr(alias))
        Call F.dictionaries.addItem(pAliases, alias, item.getId)
        Call addItemToHistory(VBA.CStr(alias), pPreviewObject, item)
    Next alias
    
End Sub

Private Sub removeFromPredefinedDictionary(aliases As VBA.Collection, dict As Scripting.Dictionary)
    Dim alias As Variant
    '------------------------------------------------------------------------------------------------------
    For Each alias In aliases
        Call F.dictionaries.removeItem(dict, alias)
    Next alias
End Sub

Private Sub removeFromObject(aliases As VBA.Collection, item As IAliasable)
    Dim alias As Variant
    '------------------------------------------------------------------------------------------------------
    For Each alias In aliases
        Call item.removeAlias(VBA.CStr(alias))
    Next alias
End Sub

Friend Sub addNewItem()
    Dim item As IAliasable
    '------------------------------------------------------------------------------------------------------
    Set item = pParent.createNewItem
    If Not item Is Nothing Then
        Call F.dictionaries.addItem(pObjects, item.getId, item)
        RaiseEvent ItemAdded(item)
    End If
End Sub

Friend Sub assignToNewItem(aliases As VBA.Collection)
    Dim item As IAliasable
    '------------------------------------------------------------------------------------------------------
    Set item = pParent.createNewItem
    If Not item Is Nothing Then
        Call F.dictionaries.addItem(pObjects, item.getId, item)
        Call assignTo(aliases, item)
        RaiseEvent AssignedToNewItem(item, pPreviewKey)
    End If
End Sub




Private Sub addItemToHistory(alias As String, previousItem As Variant, newItem As Variant)
    Dim action As StandarizingAction
    '------------------------------------------------------------------------------------------------------
    Set action = MApp.Create.StandarizingAction(alias, previousItem, newItem)
    Call pHistory.Add(action)
    Debug.Print "History items: " & pHistory.Count
End Sub




Friend Sub confirmChanges()
    RaiseEvent AfterConfirm(pObjects, pAliases, _
                                F.dictionaries.toCollection(pUnassigned, DictPart_KeyOnly), _
                                F.dictionaries.toCollection(pJunk, DictPart_KeyOnly))
End Sub
