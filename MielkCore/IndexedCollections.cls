VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IndexedCollections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "IndexedCollections"
'----------------------------------------------------------------------------------------------------------
Private pDict As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------


Private Sub Class_Initialize()
    Set pDict = f.dictionaries.Create(False)
End Sub



Public Function addItem(groupKey As Variant, item As Variant) As IndexedCollections
    Dim group As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    Set addItem = Me
    
    If Not pDict.exists(groupKey) Then
        Call pDict.Add(groupKey, New VBA.Collection)
    End If
    
    Set group = pDict.item(groupKey)
    Call group.Add(item)

End Function

Public Function getGroup(groupKey As Variant) As VBA.Collection
    Set getGroup = f.dictionaries.getObject(pDict, groupKey)
End Function
