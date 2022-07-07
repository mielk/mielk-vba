VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Repository2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "Repository2"
'----------------------------------------------------------------------------------------------------------
Private pParent As IRepoParent
Private pMappings As Scripting.Dictionary
Private pDictionaries As Scripting.Dictionary
Private pItems As VBA.Collection
'----------------------------------------------------------------------------------------------------------
Private pManager As IRepoManager
'----------------------------------------------------------------------------------------------------------
Event ItemRemoved(item As Object)
'----------------------------------------------------------------------------------------------------------




Private Sub Class_Terminate()
    Set pDictionaries = Nothing
End Sub



'[Setters]
Public Function setParent(value As IRepoParent) As Repository2
    Set setParent = Me
    Set pParent = value
End Function



'[Getters]
Public Function getParent() As IRepoParent
    Set getParent = pParent
End Function

Public Function getItemsCollection(Optional activeOnly As Boolean = True) As VBA.Collection
    Dim item As Object
    Dim isActive As Boolean
    '------------------------------------------------------------------------------------------------------
    
    If activeOnly Then
        Set getItemsCollection = New VBA.Collection
        For Each item In pItems
            isActive = True
            On Error Resume Next
            isActive = item.isActive
            On Error GoTo ErrHandler
            If isActive Then Call getItemsCollection.Add(item)
        Next item
    Else
        Set getItemsCollection = pItems
    End If
    
    
ErrHandler:

End Function
    

Public Function getItemsDictionary(propName As String, Optional activeOnly As Boolean = True) As Scripting.Dictionary
    Dim key As Variant
    Dim item As Object
    Dim isActive As Boolean
    '------------------------------------------------------------------------------------------------------
    
    Set getItemsDictionary = F.dictionaries.clone(F.dictionaries.GetObject(pDictionaries, propName))
    
    If Not getItemsDictionary Is Nothing Then
        If activeOnly Then
            With getItemsDictionary
                For Each key In .Keys
                    Set item = .item(key)
                    isActive = True
                    On Error Resume Next
                    isActive = item.isActive
                    On Error GoTo ErrHandler
                    
                    If Not isActive Then Call .Remove(key)
                    
                Next key
            End With
        End If
    End If
        
        
ErrHandler:

End Function



Public Sub loadValues(items As VBA.Collection, mappings As Scripting.Dictionary)
    Dim propName As Variant
    Dim method As String
    Dim dict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim key As Variant
    Dim item As Object
    '------------------------------------------------------------------------------------------------------
    
    Set pItems = items
    Set pMappings = mappings
    Call createItemsDictionaries
    
End Sub


Private Sub createItemsDictionaries()
    Dim propName As Variant
    Dim method As String
    Dim dict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim key As Variant
    Dim item As Object
    '------------------------------------------------------------------------------------------------------
    
    Set pDictionaries = F.dictionaries.Create(False)
    
    For Each propName In pMappings.Keys
        method = pMappings.item(propName)
        Set dict = F.dictionaries.Create(False)
        Call F.dictionaries.addItem(pDictionaries, propName, dict)
        
        For Each item In pItems
            Call F.Variables.assign(key, item.getValue(VBA.CStr(propName)))
            'key = VBA.CallByName(item, method, VbMethod)
            Call F.dictionaries.addItem(dict, key, item)
        Next item
    Next propName
    
End Sub


Public Function getItem(propName As String, key As Variant) As Object
    Dim dict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    Set dict = F.dictionaries.GetObject(pDictionaries, propName)
    If Not dict Is Nothing Then
        Set getItem = F.dictionaries.GetObject(dict, key)
    End If
    
End Function



Public Sub updateItemKeys(editedItem As IEditItem)
    Dim item As Object
    Dim key As Variant
    Dim dict As Scripting.Dictionary
    Dim subkey As Variant
    '------------------------------------------------------------------------------------------------------
    
    'Find the given object in collection - required to cast it to the proper data type ------|
    For Each item In pItems                                                                 '|
        If item Is editedItem Then                                                          '|
            Exit For                                                                        '|
        End If                                                                              '|
    Next item                                                                               '|
    '----------------------------------------------------------------------------------------|
    
    If Not item Is Nothing Then
        '[Remove old-keys entries of the given item]
        For Each key In pDictionaries.Keys
            Set dict = pDictionaries.item(key)
            For Each subkey In dict.Keys
                If dict.item(subkey) Is item Then
                    Call dict.Remove(subkey)
                End If
            Next subkey
        Next key
        
        '[Add item to dictionaries with new keys]\
        Call addItemToDictionaries(item)
        
    End If
    
End Sub



Public Sub addItem(item As Object)
    Call pItems.Add(item)
    Call addItemToDictionaries(item)
End Sub

Private Sub addItemToDictionaries(item As Object)
    Dim propName As Variant
    Dim dict As Scripting.Dictionary
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    For Each propName In pMappings.Keys
        Set dict = F.dictionaries.GetObject(pDictionaries, propName)
        If Not dict Is Nothing Then
            key = item.getValue(VBA.CStr(propName))
            Call F.dictionaries.addItem(dict, key, item)
        Else
            Stop
        End If
    Next propName
    
End Sub



Public Sub removeItem(item As Object)
    Dim i As Long
    Dim colItem As Object
    '------------------------------------------------------------------------------------------------------
    
    '[Remove from main collection]
    For i = 1 To pItems.Count
        Set colItem = pItems.item(i)
        If colItem Is item Then
            Call pItems.Remove(i)
            Exit For
        End If
    Next i
    
    Call createItemsDictionaries
    
    RaiseEvent ItemRemoved(item)
    
End Sub



Public Function selectItems(labelsPropertyName As String, _
                            Optional multipleSelectionAllowed As Boolean = False, _
                            Optional ByVal defaultItem As Variant, _
                            Optional hasSearchPanel As Boolean = False, _
                            Optional addingAllowed As Boolean = True, _
                            Optional ByVal windowTitle As String, _
                            Optional ByVal headers As VBA.Collection, _
                            Optional minItems As Long = 0, _
                            Optional maxItems As Long = 0) As VBA.Collection
    Const DEFAULT_WINDOW_TITLE_TAG As String = "SelectItem.WindowTitle"
    '------------------------------------------------------------------------------------------------------
    Dim items As Scripting.Dictionary
    Dim captions As Scripting.Dictionary
    Dim window As WSelectItem
    Dim dictAction As Scripting.Dictionary
    Dim initialSelection As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(windowTitle) = 0 Then windowTitle = DEFAULT_WINDOW_TITLE_TAG
    
    If Not F.Collections.isNonEmptyCollection(headers) Then
        Set headers = F.Collections.Create(UI.Create.ListHeaderDefinition("Item", labelsPropertyName))
    End If
    Set dictAction = F.dictionaries.createWithItems(False, _
                            KeyValue(TAG_METHOD_INVOKER, Me), _
                            KeyValue(TAG_METHOD_NAME, "createNewItem"), _
                            KeyValue(TAG_PROPERTY, labelsPropertyName))
    

    Set items = getItemsDictionary(pParent.getIdField)
    Set captions = createCaptionsDictionary(items, labelsPropertyName)
    Set initialSelection = createInitialSelectionDictionary(defaultItem)
    
    If Not items Is Nothing Then
        Set window = UI.Dialogs.selectWindow(items, captions). _
                                        setTitle(MsgService.getText(windowTitle)). _
                                        setMultiSelectAllowed(multipleSelectionAllowed). _
                                        setInitialSelection(initialSelection). _
                                        setLocationProps(UI.Create.LocationPropsForScreenCenter). _
                                        setHasSearchPanel(hasSearchPanel). _
                                        setMinItemsSelected(minItems). _
                                        setMaxItemsSelected(maxItems). _
                                        setAddingAllowed(addingAllowed). _
                                        setAddingAction(dictAction). _
                                        setListHeaders(headers)
        Set selectItems = window.Execute
    End If
    
End Function

Private Function createCaptionsDictionary(items As Scripting.Dictionary, captionPropertyName As String) As Scripting.Dictionary
    Dim key As Variant
    Dim obj As IEditItem
    '------------------------------------------------------------------------------------------------------
    Set createCaptionsDictionary = F.dictionaries.Create(False)
    For Each key In items
        Set obj = items.item(key)
        If Not obj Is Nothing Then
            Call F.dictionaries.addItem(createCaptionsDictionary, key, obj.getValue(captionPropertyName))
        End If
    Next key
End Function

Private Function createInitialSelectionDictionary(defaultValue As Variant) As Scripting.Dictionary
    Dim col As VBA.Collection
    Dim item As Variant
    Dim editItem As IEditItem
    Dim idField As String:                  idField = pParent.getIdField
    '------------------------------------------------------------------------------------------------------
    
    Set createInitialSelectionDictionary = F.dictionaries.Create(False)
    If F.Variables.isContainer(defaultValue) Then
        Set col = F.Collections.createFrom(defaultValue)
    Else
        Set col = F.Collections.Create(defaultValue)
    End If
    
    If F.Collections.isNonEmptyCollection(col) Then
        For Each item In col
            If Not VBA.IsObject(item) Then
            ElseIf Not TypeOf item Is IEditItem Then
            Else
                Set editItem = item
                Call F.dictionaries.addItem(createInitialSelectionDictionary, editItem.getValue(idField), item)
            End If
        Next item
    End If
    
End Function

Public Function createNewItem(Optional params As Scripting.Dictionary) As Variant
    Set createNewItem = pParent.createNewItem
End Function



'[Managing items]
Public Sub manage(Optional params As Scripting.Dictionary)
    Set pManager = New RepoManager_Sheet
    With pManager
        Call .setRepo(Me)
        Call MApp.registerRepoManager(pManager)
        Call .show
    End With
End Sub

Public Function updateView()
    If Not pManager Is Nothing Then
        Call pManager.updateValues
    End If
End Function
