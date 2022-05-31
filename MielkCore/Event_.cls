VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Event_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "Event_"
'----------------------------------------------------------------------------------------------------------
Private pId As Long
Private pName As String
Private pAliases As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------


Private Sub Class_Initialize()
    Set pAliases = F.Dictionaries.Create(False)
End Sub


'[SETTERS]

Public Function setId(value As Long) As Event_
    Set setId = Me
    pId = value
End Function

Public Function setName(value As String) As Event_
    Set setName = Me
    pName = value
End Function

Public Function addAlias(value As String) As Event_
    Set addAlias = Me
    With pAliases
        If Not .exists(value) Then
            Call .Add(value, value)
        End If
    End With
End Function



'[GETTERS]

Public Function getId() As Long
    getId = pId
End Function

Public Function getName() As String
    getName = pName
End Function
