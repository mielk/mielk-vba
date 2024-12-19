VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WindowsCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "WindowsCache"
'----------------------------------------------------------------------------------------------------------
Private pContainer As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------

Private Sub Class_Initialize()
    Set pContainer = F.dictionaries.Create(True)
End Sub

Public Sub add(item As Object)
    Stop
    'If TypeOf item Is IUserForm Then
        Call F.dictionaries.addItem(pContainer, item, item)
        Debug.Print "Item " & ObjPtr(item) & " added to GlobalCache | GlobalCacheCounter: " & pContainer.Count
    'End If
End Sub

Public Sub remove(item As Object)
    Call F.dictionaries.removeKey(pContainer, item)
    Debug.Print "Item " & ObjPtr(item) & " removed from GlobalCache | GlobalCacheCounter: " & pContainer.Count
End Sub

Public Function getForms() As Scripting.Dictionary
    Set getForms = F.dictionaries.Clone(pContainer)
End Function

Public Sub showAll()
    Dim varKey As Variant
    
    Stop
    
    'Dim item As IUserForm
    Dim item As Object
    '----------------------------------------------------------------------------------------------------------
    
    For Each varKey In pContainer.keys
        On Error Resume Next
        Set item = pContainer.item(varKey)
        Call item.show
        On Error GoTo 0
    Next varKey
End Sub

Public Sub clear()
    Dim varKey As Variant
    Stop
    'Dim item As IUserForm
    Dim item As Object
    '----------------------------------------------------------------------------------------------------------
    For Each varKey In pContainer.keys
        On Error Resume Next
        Set item = pContainer.item(varKey)
        Call item.unload
        On Error GoTo 0
    Next varKey
    Call pContainer.removeAll
    Debug.Print "GlobalCache cleared | GlobalCacheCounter: " & pContainer.Count
End Sub
