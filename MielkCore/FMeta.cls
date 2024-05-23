VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FMeta"
'----------------------------------------------------------------------------------------------------------

Public Function getFullName() As String
    getFullName = Excel.ThisWorkbook.fullName
End Function

Public Function getPath() As String
    getPath = Excel.ThisWorkbook.path
End Function
 
Public Function getFolder() As Scripting.folder
    Static fso As Scripting.FileSystemObject
    '------------------------------------------------------------------------------------------------------
    Dim path As String
    Dim file As Scripting.file
    '------------------------------------------------------------------------------------------------------
    If fso Is Nothing Then Set fso = New Scripting.FileSystemObject
    path = Excel.ThisWorkbook.fullName
    Set file = fso.getFile(path)
    Set getFolder = file.parentFolder
End Function
