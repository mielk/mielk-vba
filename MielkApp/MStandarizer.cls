VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MStandarizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
'
'Implements MielkCore.ITranslator
'
'Private Const CLASS_NAME As String = "MStandarizer"
''----------------------------------------------------------------------------------------------------------
'Private Const UNASSIGNED_GROUP_NAME As String = "[Unassigned]"
'Private Const SAVE_BUTTON_TAG As String = "[Save]"
'Private Const CANCEL_BUTTON_TAG As String = "[Cancel]"
''----------------------------------------------------------------------------------------------------------
'Private pTranslator As MielkCore.ITranslator
'Private pView As MStandarizerView
''[Standard names] -----------------------------------------------------------------------------------------
'Private pStandardNames As MielkCore.HashMap
'Private pUnassignedStandardName As MStandardName
'Private pRemovedStandardNames As MielkCore.ArrayList
'Private pAddedStandardNames As MielkCore.ArrayList
'Private pRenamedStandardNames As MielkCore.ArrayList
''[Mapping] ------------------------------------------------------------------------------------------------
'Private pMapping As Variant
'Private pChangeLogs() As Variant
''----------------------------------------------------------------------------------------------------------
'Private pStandardNamesListHeadersDefinitions As MielkCore.ArrayList
'Private pRawNamesListHeadersDefinitions As MielkCore.ArrayList
''[State] --------------------------------------------------------------------------------------------------
'Private pActiveStandardName As MStandardName
''[Events] -------------------------------------------------------------------------------------------------
'Event afterConfirm()
'Event afterCancel()
'Event AfterActiveStandardNameChanged(standardName As MStandardName)
'Event AfterNewStandardNameAdded(standardName As MStandardName)
'Event AfterAssignmentChanged()
''----------------------------------------------------------------------------------------------------------
'
'
'
''[INITIALIZER]
'Private Sub Class_Initialize()
'    Call initializeContainers
'End Sub
'
'Private Sub initializeContainers()
'    Set pStandardNames = MielkCore.createHashMap
'    Set pRemovedStandardNames = MielkCore.createArrayList
'    Set pAddedStandardNames = MielkCore.createArrayList
'    Set pRenamedStandardNames = MielkCore.createArrayList
'End Sub
'
'
'Public Sub initialize()
'    Call createUnassignedStandardName
'End Sub
'
'Private Sub createUnassignedStandardName()
'    Set pUnassignedStandardName = createStandardName(0, VBA.UCase$(pTranslator.getText(UNASSIGNED_GROUP_NAME))).setIsUnassigned(True)
'    Call pStandardNames.Add(pUnassignedStandardName.getName, pUnassignedStandardName)
'End Sub
'
'
'
''[SETTERS]
'Public Function setTranslator(value As MielkCore.ITranslator) As MStandarizer
'    Set setTranslator = Me
'    Set pTranslator = value
'End Function
'
'
'
''[GETTERS]
'Public Function getText(tag As String) As String
'    getText = pTranslator.getText(tag)
'End Function
'
'Public Function getActiveStandardName() As MStandardName
'    Set getActiveStandardName = pActiveStandardName
'End Function
'
'Public Function getStandardNames() As Scripting.Dictionary
'    Set getStandardNames = pStandardNames.toDictionary
'End Function
'
'Public Function isUnassignedEmpty() As Boolean
'    If Not pUnassignedStandardName Is Nothing Then
'        isUnassignedEmpty = (pUnassignedStandardName.getRawNames.Count = 0)
'    End If
'End Function
'
'Public Function getAddedStandardNames() As MielkCore.ArrayList
'    Set getAddedStandardNames = pAddedStandardNames
'End Function
'
'Public Function getChangesArray() As Variant
'    getChangesArray = pChangeLogs
'End Function
'
'
'
'
''[STANDARD NAMES]
'Public Function loadStandardNames(values As MielkCore.HashMap) As MStandarizer
'    Set loadStandardNames = Me
'    Call pStandardNames.addAll(convertToStandardNamesHashMap(values))
'End Function
'
'Public Sub addStandardName(name As String)
'    Dim standardName As MStandardName
'    '------------------------------------------------------------------------------------------------------
'    Set standardName = createStandardName(0, name)
'    Call pStandardNames.Add(name, standardName)
'    Call pAddedStandardNames.Add(standardName)
'    RaiseEvent AfterNewStandardNameAdded(standardName)
'End Sub
'
'Private Function convertToStandardNamesHashMap(values As MielkCore.HashMap) As MielkCore.HashMap
'    Dim dict As Scripting.Dictionary
'    Dim varValue As Variant
'    Dim name As String
'    Dim standardName As MStandardName
'    '------------------------------------------------------------------------------------------------------
'    Set convertToStandardNamesHashMap = MielkCore.createHashMap
'    Set dict = values.toDictionary
'    For Each varValue In dict
'        name = dict.item(varValue)
'        Set standardName = createStandardName(VBA.CLng(varValue), name)
'        Call convertToStandardNamesHashMap.Add(standardName.getName, standardName)
'    Next varValue
'End Function
'
'Public Function setMappingBy2DArray(values As Variant) As MStandarizer
'    Set setMappingBy2DArray = Me
'    pMapping = values
'End Function
'
'Public Function setUnassignedRawNames(values As MielkCore.ArrayList) As MStandarizer
'    Set setUnassignedRawNames = Me
'    With pUnassignedStandardName
'        Call .ClearContent
'        Call .addRawNames(values)
'    End With
'End Function
'
'Public Function setRawNamesListHeaderDefinitions(value As Variant) As MStandarizer
'    Set setRawNamesListHeaderDefinitions = Me
'    Set pRawNamesListHeadersDefinitions = MielkCore.createArrayList(value)
'End Function
'
'Public Function setStandardNamesListHeaderDefinitions(value As Variant) As MStandarizer
'    Set setStandardNamesListHeaderDefinitions = Me
'    Set pStandardNamesListHeadersDefinitions = MielkCore.createArrayList(value)
'End Function
'
'Public Sub setActiveStandardName(value As MStandardName)
'    Set pActiveStandardName = value
'    RaiseEvent AfterActiveStandardNameChanged(value)
'End Sub
'
'Public Sub setActiveStandardNameByCaption(caption As String)
'    Set pActiveStandardName = pStandardNames.getItem(caption)
'    RaiseEvent AfterActiveStandardNameChanged(pActiveStandardName)
'End Sub
'
'Public Function getRawNamesForActiveStandardName() As VBA.Collection
'    If Not pActiveStandardName Is Nothing Then
'        Set getRawNamesForActiveStandardName = pActiveStandardName.getRawNames
'    End If
'End Function
'
'Public Function getStandardNamesCaptions() As VBA.Collection
'    Set getStandardNamesCaptions = pStandardNames.getKeys(ContainerType_Collection)
'End Function
'
'
'
'Public Sub reload()
'    Dim standardName As MStandardName
'    Dim i As Long
'    '------------------------------------------------------------------------------------------------------
'    For Each standardName In pStandardNames.toCollection(DictEntry_Item)
'        If Not standardName.isUnassigned Then Call standardName.ClearContent
'    Next standardName
'    For i = LBound(pMapping, 2) + 1 To UBound(pMapping, 2)
'        Call assignRawName(pMapping(1, i), pMapping(2, i))
'    Next i
'End Sub
'
'Private Sub assignRawName(rawName As Variant, standardItemKey As Variant)
'    Dim standardName As MStandardName
'    '------------------------------------------------------------------------------------------------------
'    Set standardName = pStandardNames.GetObject(standardItemKey)
'    If Not standardName Is Nothing Then
'        Call standardName.addRawName(rawName)
'    End If
'End Sub
'
'
'
''[ACTIONS]
'Public Sub display(Optional mode As VBA.FormShowConstants = VBA.vbModeless)
'    If pView Is Nothing Then
'        Set pView = createStandarizerView(Me). _
'                        setSaveButtonCaption(pTranslator.getText(SAVE_BUTTON_TAG)). _
'                        setCancelButtonCaption(pTranslator.getText(CANCEL_BUTTON_TAG)). _
'                        setStandardNamesListHeaderDefinitions(pStandardNamesListHeadersDefinitions). _
'                        setRawNamesListHeaderDefinitions(pRawNamesListHeadersDefinitions)
'    End If
'    Call setActiveStandardName(pUnassignedStandardName)
'    Call pView.show(mode)
'
'End Sub
'
'Public Sub changeAssignment(rawNames As VBA.Collection, destinationStandardNameCaption As String)
'    Dim varRawName As Variant
'    Dim sourceStandardName As MStandardName
'    Dim destinationStandardName As MStandardName
'    '------------------------------------------------------------------------------------------------------
'
'    Set sourceStandardName = pActiveStandardName
'    Set destinationStandardName = pStandardNames.getItem(destinationStandardNameCaption)
'    Call sourceStandardName.deleteRawNames(rawNames)
'    Call destinationStandardName.addRawNames(rawNames)
'
'    For Each varRawName In rawNames
'        Call addChangeLog(MielkCore.stringify(varRawName), sourceStandardName, destinationStandardName)
'    Next varRawName
'
'    RaiseEvent AfterAssignmentChanged
'
'End Sub
'
'Private Sub addChangeLog(rawName As String, sourceStandardName As MStandardName, destinationStandardName As MStandardName)
'    Dim index As Long
'    '------------------------------------------------------------------------------------------------------
'
'    If MielkCore.isDefinedArray(pChangeLogs) Then
'        index = UBound(pChangeLogs, 2) + 1
'    Else
'        index = 1
'    End If
'
'    ReDim Preserve pChangeLogs(1 To 3, 1 To index)
'    pChangeLogs(1, index) = rawName
'    Set pChangeLogs(2, index) = sourceStandardName
'    Set pChangeLogs(3, index) = destinationStandardName
'
'End Sub
'
'
'
''[CONFIRM & CANCEL]
'Public Sub triggerCancel()
'    RaiseEvent afterCancel
'End Sub
'
'Public Sub triggerSave()
'    RaiseEvent afterConfirm
'End Sub
'
'
'
'
''[Implementation of ITranslator]
'Private Function ITranslator_getText(tag As String) As String
'    ITranslator_getText = getText(tag)
'End Function
