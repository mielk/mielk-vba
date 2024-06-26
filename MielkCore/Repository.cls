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


Public Sub addItem(key As Variant, item As Variant, Optional name As String)
    Call f.dictionaries.addItem(pItems, key, item)
    If VBA.Len(name) Then
        Call f.dictionaries.addItem(pNameToKeyMap, name, key)
    End If
End Sub

Public Sub removeItem(key As Variant, Optional name As String)
    Dim varNameKey As Variant
    Dim subkey As Variant
    '------------------------------------------------------------------------------------------------------
    
    Call f.dictionaries.removeItem(pItems, key)
    
    For Each varNameKey In pNameToKeyMap.Keys
        subkey = pNameToKeyMap.item(varNameKey)
        If subkey = key Then
            Call pNameToKeyMap.Remove(varNameKey)
        End If
    Next varNameKey
    
End Sub

Public Function exists(key As Variant) As Boolean
    exists = pItems.exists(key)
End Function

Public Function getItem(key As Variant) As Variant
    Call f.Variables.assign(getItem, f.dictionaries.getObject(pItems, key))
End Function

Public Function getItemByName(name As String) As Variant
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    key = f.dictionaries.getItem(pNameToKeyMap, name)
    Call f.Variables.assign(getItemByName, f.dictionaries.getObject(pItems, key))
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
    Dim key As Variant
    Dim id As Variant
    '------------------------------------------------------------------------------------------------------
    Set getItemsDictionaryByName = f.dictionaries.Create(False)
    For Each key In pNameToKeyMap.Keys
        id = pNameToKeyMap.item(key)
        If pItems.exists(id) Then
            Call getItemsDictionaryByName.Add(key, pItems.item(id))
        End If
    Next key
End Function
