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
    Set pDictById = f.dictionaries.Create(False)
    Set pDictByValue = f.dictionaries.Create(False)
End Sub

Private Sub Class_Terminate()
    Set pDictById = Nothing
    Set pDictByValue = Nothing
End Sub



'[GETTERS]
Public Function getId(value As Variant) As Variant
    If VBA.IsObject(value) Then
        Call f.Variables.assign(getId, f.dictionaries.getItem(pDictByValue, VBA.CLng(VBA.ObjPtr(value))))
    Else
        Call f.Variables.assign(getId, f.dictionaries.getItem(pDictByValue, value))
    End If
End Function

Public Function getValue(id As Variant) As Variant
    If VBA.IsObject(id) Then
        Call f.Variables.assign(getValue, f.dictionaries.getItem(pDictById, VBA.CLng(VBA.ObjPtr(id))))
    Else
        Call f.Variables.assign(getValue, f.dictionaries.getItem(pDictById, id))
    End If
End Function


Public Function getDictionaryById() As Scripting.Dictionary
    Set getDictionaryById = f.dictionaries.clone(pDictById)
End Function

Public Function getDictionaryByValue() As Scripting.Dictionary
    Set getDictionaryByValue = f.dictionaries.clone(pDictByValue)
End Function






'[ADDING ITEMS]
Public Sub addItem(key As Variant, value As Variant)
    If VBA.IsObject(key) Then
        Call f.dictionaries.addItem(pDictById, VBA.CLng(VBA.ObjPtr(key)), value)
    Else
        Call f.dictionaries.addItem(pDictById, key, value)
    End If
    
    If VBA.IsObject(value) Then
        Call f.dictionaries.addItem(pDictByValue, VBA.CLng(VBA.ObjPtr(value)), key)
    Else
        Call f.dictionaries.addItem(pDictByValue, value, key)
    End If
End Sub


Public Sub loadFromDbTable(connection As String, tableName As String, idField As String, _
                            valueField As String, Optional orderByField As String)
    Const SQL_PATTERN As String = "SELECT {1}, {2} FROM {0} ORDER BY {3}"
    '------------------------------------------------------------------------------------------------------
    Dim rs As Object        'ADODB.Recordset
    Dim sql As String
    Dim id As Variant
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set pDictById = f.dictionaries.Create(False)
    Set pDictByValue = f.dictionaries.Create(False)
    
    sql = f.strings.format(SQL_PATTERN, tableName, idField, valueField, _
                                        VBA.IIf(VBA.Len(orderByField), orderByField, idField))
    Set rs = VBA.CreateObject("ADODB.Recordset")
    Call rs.Open(sql, connection, adOpenStatic, adLockOptimistic)
    
    With rs
        Do Until .EOF
            id = .fields(idField)
            value = .fields(valueField)
            Call pDictById.Add(id, value)
            Call pDictByValue.Add(value, id)
            Call .MoveNext
        Loop
    End With
    
End Sub
