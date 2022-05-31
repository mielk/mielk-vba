VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParentApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The purpose of this class is to store information about parent application that uses this library file.

Private Const CLASS_NAME As String = "ParentApp"
    '----------------------------------------------------------------------------------------------------------
Private pName As String
Private pVersion As String
Private pPath As String
    '----------------------------------------------------------------------------------------------------------

'[Setters]
Public Function setName(value As String) As ParentApp
    Set setName = Me
    pName = value
End Function

Public Function setVersion(value As String) As ParentApp
    Set setVersion = Me
    pVersion = value
End Function

Public Function setPath(value As String) As ParentApp
    Set setPath = Me
    pPath = value
End Function


'[Getters]
Public Function getName() As String
    getName = pName
End Function

Public Function getVersion() As String
    getVersion = pVersion
End Function

Public Function getNameVersion() As String
    getNameVersion = pName & " (" & pVersion & ")"
End Function

Public Function getPath() As String
    getPath = pPath
End Function

Public Function getUncPath() As String
    getUncPath = F.files.getUncPath(pPath)
End Function
