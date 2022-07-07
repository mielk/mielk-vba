VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Repository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'The purpose of this class is to store the set of objects and provide developers with
'a few methods to access those objects and manage them (.add | .remove).

Private Const CLASS_NAME As String = "Repository"
'----------------------------------------------------------------------------------------------------------
Private pItems As Scripting.Dictionary
Private pNameToKeyMap As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------


Private Sub Class_Initialize()
    '[Initialize containers]
    Set pItems = f.dictionaries.Create(False)
    Set pNameToKeyMap = f.dictionaries.Create(False)
End Sub

Private Sub Class_Terminate()
    
    Call pItems.RemoveAll
    Set pItems = Nothing
    
    Call pNameToKeyMap.RemoveAll
    Set pNameToKeyMap = Nothing
    
End Sub


Public Sub addItem(Key As Variant, Item As Variant, Optional name As String)
    Call f.dictionaries.addItem(pItems, Key, Item)
    If VBA.Len(name) Then
        Call f.dictionaries.addItem(pNameToKeyMap, name, Key)
    End If
End Sub

Public Sub removeItem(Key As Variant, Optional name As String)
    Dim varNameKey As Variant
    Dim subkey As Variant
    '------------------------------------------------------------------------------------------------------
    
    Call f.dictionaries.removeItem(pItems, Key)
    
    For Each varNameKey In pNameToKeyMap.Keys
        subkey = pNameToKeyMap.Item(varNameKey)
        If subkey = Key Then
            Call pNameToKeyMap.Remove(varNameKey)
        End If
    Next varNameKey
    
End Sub

Public Function exists(Key As Variant) As Boolean
    exists = pItems.exists(Key)
End Function

Public Function getItem(Key As Variant) As Variant
    Call f.Variables.assign(getItem, f.dictionaries.getObject(pItems, Key))
End Function

Public Function getItemByName(name As String) As Variant
    Dim Key As Variant
    '------------------------------------------------------------------------------------------------------
    Key = f.dictionaries.getItem(pNameToKeyMap, name)
    Call f.Variables.assign(getItemByName, f.dictionaries.getObject(pItems, Key))
End Function

Public Function getItemsCollection() As VBA.Collection
    Set getItemsCollection = f.dictionaries.toCollection(pItems)
End Function

Public Function getNameKeyDictionary() As Scripting.Dictionary
    Set getNameKeyDictionary = f.dictionaries.clone(pNameToKeyMap)
End Function

Public Function getItemsDictionaryById() As Scripting.Dictionary
    Set getItemsDictionaryById = f.dictionaries.clone(pItems)
End Function

Public Function getItemsDictionaryByName() As Scripting.Dictionary
    Dim Key As Variant
    Dim Id As Variant
    '------------------------------------------------------------------------------------------------------
    Set getItemsDictionaryByName = f.dictionaries.Create(False)
    For Each Key In pNameToKeyMap.Keys
        Id = pNameToKeyMap.Item(Key)
        If pItems.exists(Id) Then
            Call getItemsDictionaryByName.add(Key, pItems.Item(Id))
        End If
    Next Key
End Function
