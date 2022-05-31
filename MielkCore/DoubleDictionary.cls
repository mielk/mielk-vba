VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DoubleDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'The purpose of this class is to have a container with fast access to keys as well as values.
'It is used to store sets of data with unique keys and unique values that should be accessed both ways

Private Const CLASS_NAME As String = "DoubleDictionary"
'----------------------------------------------------------------------------------------------------------
Private pDictById As Scripting.Dictionary
Private pDictByValue As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------


Private Sub Class_Initialize()
    Set pDictById = F.Dictionaries.Create(False)
    Set pDictByValue = F.Dictionaries.Create(False)
End Sub

Private Sub Class_Terminate()
    Set pDictById = Nothing
    Set pDictByValue = Nothing
End Sub



'[GETTERS]
Public Function getId(value As Variant) As Variant
    If VBA.IsObject(value) Then
        Call F.Variables.assign(getId, F.Dictionaries.getItem(pDictByValue, VBA.CLng(VBA.ObjPtr(value))))
    Else
        Call F.Variables.assign(getId, F.Dictionaries.getItem(pDictByValue, value))
    End If
End Function

Public Function getValue(id As Variant) As Variant
    If VBA.IsObject(id) Then
        Call F.Variables.assign(getValue, F.Dictionaries.getItem(pDictById, VBA.CLng(VBA.ObjPtr(id))))
    Else
        Call F.Variables.assign(getValue, F.Dictionaries.getItem(pDictById, id))
    End If
End Function


Public Function getDictionaryById() As Scripting.Dictionary
    Set getDictionaryById = F.Dictionaries.clone(pDictById)
End Function

Public Function getDictionaryByValue() As Scripting.Dictionary
    Set getDictionaryByValue = F.Dictionaries.clone(pDictByValue)
End Function






'[ADDING ITEMS]
Public Sub addItem(key As Variant, value As Variant)
    If VBA.IsObject(key) Then
        Call F.Dictionaries.addItem(pDictById, VBA.CLng(VBA.ObjPtr(key)), value)
    Else
        Call F.Dictionaries.addItem(pDictById, key, value)
    End If
    
    If VBA.IsObject(value) Then
        Call F.Dictionaries.addItem(pDictByValue, VBA.CLng(VBA.ObjPtr(value)), key)
    Else
        Call F.Dictionaries.addItem(pDictByValue, value, key)
    End If
End Sub
