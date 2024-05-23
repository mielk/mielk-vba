VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Language"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "language"
'----------------------------------------------------------------------------------------------------------
Private pKey As String
Private pName As String
Private pIsDefault As Boolean
Private pTags As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------


Private Sub Class_Initialize()
    Set pTags = f.dictionaries.Create(False)
End Sub


'[SETTERS]
Friend Function setKey(value As String) As Language
    Set setKey = Me
    pKey = value
End Function

Friend Function setName(value As String) As Language
    Set setName = Me
    pName = value
End Function

Friend Function setIsDefault(value As Boolean) As Language
    Set setIsDefault = Me
    pIsDefault = value
End Function

Friend Sub addText(tag As String, text As String)
    Call f.dictionaries.addItem(pTags, tag, text)
End Sub





'[GETTERS]
Public Function getKey() As String
    getKey = pKey
End Function

Public Function getName() As String
    getName = pName
End Function

Public Function getNameForCurrentLanguage() As String
    Const NAME_TAG As String = "languages."
    '------------------------------------------------------------------------------------------------------
    getNameForCurrentLanguage = MsgService.getText(NAME_TAG & pKey)
End Function

Public Function isDefault() As Boolean
    isDefault = pIsDefault
End Function

Public Function getText(tag As String) As String
    With pTags
        If .exists(tag) Then
            getText = .item(tag)
        Else
            Debug.Print "_________________ missing tag: " & tag
            getText = tag
        End If
    End With
End Function


