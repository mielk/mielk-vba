VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "EProject"
'[Default values] -----------------------------------------------------------------------------------------
Private Const DEFAULT_OVERRIDE_IF_EXISTS As Boolean = True
'----------------------------------------------------------------------------------------------------------
Private pName As String
Private pCodeName As String
Private pParentFolder As String
Private pLibFolderPath As String
Private pIsAddIn As Boolean
Private pOverrideIfExists As Boolean
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    Call setInitialValues
End Sub

Private Sub setInitialValues()
    pName = "Effective Yields"
    pCodeName = "ey"
    pParentFolder = "D:\vba-tests\apps\"
    pLibFolderPath = "D:\Dropbox\tm\mielk\mielk-vba\code\"
    pIsAddIn = False
    pOverrideIfExists = DEFAULT_OVERRIDE_IF_EXISTS
End Sub



Public Sub create(sheets As Scripting.Dictionary)
    Dim creator As ProjectCreator
    '----------------------------------------------------------------------------------------------------------
    
    Set creator = New ProjectCreator
    With creator
        Call .setSheets(sheets)
        Call .run(Me)
    End With

End Sub





'[GETTERS]
Public Function getName() As String
    getName = pName
End Function

Public Function getCodeName() As String
    getCodeName = pCodeName
End Function

Public Function getParentFolderPath() As String
    getParentFolderPath = pParentFolder
End Function

Public Function getFolderPath() As String
    getFolderPath = f.files.concatenatePath(pParentFolder, pName) & "\"
End Function

Public Function getViewFilePath() As String
    Dim fileName As String
    '----------------------------------------------------------------------------------------------------------
    fileName = pCodeName & VIEW_FILE_SUFFIX & VBA.IIf(pIsAddIn, EXTENSION_EXCEL_ADDIN, EXTENSION_EXCEL_MACRO_FILE)
    getViewFilePath = f.files.concatenatePath(getFolderPath, fileName)
End Function

Public Function getLibFolderPath() As String
    getLibFolderPath = pLibFolderPath
End Function

Public Function isAddIn() As Boolean
    isAddIn = pIsAddIn
End Function

Public Function isUniqueName() As Boolean
    Dim destinationPath As String
    '----------------------------------------------------------------------------------------------------------
    destinationPath = f.files.concatenatePath(pParentFolder, pName)
    isUniqueName = Not f.files.FolderExists(destinationPath)
End Function

Public Function overrideIfExists() As Boolean
    overrideIfExists = pOverrideIfExists
End Function

Public Function getValue(propName As String) As Variant
    Select Case propName
        Case props.projectName.getName:                     getValue = pName
        Case props.ProjectCodeName.getName:                 getValue = pCodeName
        Case Props_Project.ProjectCodeNameUCase.getName:    getValue = VBA.UCase$(pCodeName)
        Case Props_Project.ProjectLibFolderPath.getName:    getValue = pLibFolderPath
        Case Props_Project.ProjectToolFolderPath.getName:   getValue = getFolderPath
    End Select
End Function

Public Function getReplacements() As Scripting.Dictionary
    Dim key As Variant
    Dim prop As enumProperty
    Dim value As Variant
    '----------------------------------------------------------------------------------------------------------
    Set getReplacements = f.dictionaries.create(True)
    With getAllPlaceholders
        For Each key In .Keys
            Set prop = .item(key)
            value = getValue(prop.getName)
            Call getReplacements.Add(key, value)
        Next key
    End With
End Function

Public Function isValid() As Boolean
    Stop
End Function
