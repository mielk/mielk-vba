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
Private pWindow As RepoManager
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
            Stop
            'On Error GoTo errHandler
            If isActive Then Call getItemsCollection.Add(item)
        Next item
    Else
        Set getItemsCollection = pItems
    End If
    
End Function
    

Public Function getItemsDictionary(propName As String, Optional activeOnly As Boolean = True) As Scripting.Dictionary
    Dim key As Variant
    Dim item As Object
    Dim isActive As Boolean
    '------------------------------------------------------------------------------------------------------
    
    Set getItemsDictionary = F.Dictionaries.Clone(F.Dictionaries.GetObject(pDictionaries, propName))
    
    If activeOnly Then
        With getItemsDictionary
            For Each key In .Keys
                Set item = .item(key)
                isActive = True
                On Error Resume Next
                isActive = item.isActive
                Stop
                'On Error GoTo errhandler
                
                If Not isActive Then Call .Remove(key)
                
            Next key
        End With
    End If
        
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
    
    Set pDictionaries = F.Dictionaries.Create(False)
    
    For Each propName In pMappings.Keys
        method = pMappings.item(propName)
        Set dict = F.Dictionaries.Create(False)
        Call F.Dictionaries.addItem(pDictionaries, propName, dict)
        
        For Each item In pItems
            Call F.Variables.assign(key, item.getValue(VBA.CStr(propName)))
            key = VBA.CallByName(item, method, VbMethod)
            Call F.Dictionaries.addItem(dict, key, item)
        Next item
    Next propName
    
End Sub


Public Function getItem(propName As String, key As Variant) As Object
    Dim dict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    Set dict = F.Dictionaries.GetObject(pDictionaries, propName)
    If Not dict Is Nothing Then
        Set getItem = F.Dictionaries.GetObject(dict, key)
    End If
    
End Function



Public Sub updateItemKeys(listItem As IListItem)
    Dim item As Object
    Dim key As Variant
    Dim dict As Scripting.Dictionary
    Dim subkey As Variant
    '------------------------------------------------------------------------------------------------------
    
    'Find the given object in collection - required to cast it to the proper data type ------|
    For Each item In pItems                                                                 '|
        If item Is listItem Then                                                            '|
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
        Set dict = F.Dictionaries.GetObject(pDictionaries, propName)
        If Not dict Is Nothing Then
            key = item.getValue(VBA.CStr(propName))
            Call F.Dictionaries.addItem(dict, key, item)
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
    
End Sub



Public Function selectItem(labelsPropertyName As String, Optional defaultItem As Object, _
                            Optional hasSearchPanel As Boolean = False, _
                            Optional ByVal windowTitle As String) As Object
    Const DEFAULT_WINDOW_TITLE_TAG As String = "SelectItem.WindowTitle"
    '------------------------------------------------------------------------------------------------------
    Dim items As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(windowTitle) = 0 Then windowTitle = DEFAULT_WINDOW_TITLE_TAG
    
    Set items = getItemsDictionary(labelsPropertyName)
    If Not items Is Nothing Then
        Set items = F.Dictionaries.sortByKey(items)
        Set selectItem = F.cast.toObject( _
                            UI.Dialogs.selectItem( _
                                items, _
                                windowTitle:=MsgService.getText(windowTitle), _
                                defaultItem:=defaultItem, _
                                locationProps:=UI.Create.LocationPropsForScreenCenter, _
                                hasSearchPanel:=hasSearchPanel))
    End If
    
End Function



Public Sub manage(Optional params As Scripting.Dictionary)
    Set pWindow = New RepoManager
    With pWindow
        Call .setRepo(Me)
        Call .show
    End With
End Sub
