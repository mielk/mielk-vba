Attribute VB_Name = "modConst"
Option Explicit

Private Const CLASS_NAME As String = "modConst"
'[Tags] ---------------------------------------------------------------------------------------------------
Public Const PROJECT_NAME_PLACEHOLDER As String = "YourProjectName"
Public Const PROJECT_CODE_NAME_PLACEHOLDER As String = "YourProjectCodeName"
Public Const PROJECT_LIB_FOLDER_PLACEHOLDER As String = "YourProjectLibFolderPath"
Public Const PROJECT_TOOL_FOLDER_PLACEHOLDER As String = "YourProjectToolFolderPath"
'[File patterns] ------------------------------------------------------------------------------------------
Public Const VBS_FILE_PATTERN As String = "\.vbs$"
Public Const EXCEL_ADDIN_NAME_PATTERN As String = "\\([^\\]*)\.\w+$"
Public Const RIBBON_XML_FILE_NAME = "ribbon.xml"
'[Generic message tags] -----------------------------------------------------------------------------------
Public Const FAILED_BECAUSE_OF_PREDECESSORS As String = "CreatingNewProject.Errors.FailedBecauseOfPredecessors"
'----------------------------------------------------------------------------------------------------------


Public Function getAllPlaceholders() As Scripting.Dictionary
    Static instance As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then
        Set instance = f.dictionaries.createWithItems(True, _
                                        KeyValue(PROJECT_NAME_PLACEHOLDER, props.ProjectName), _
                                        KeyValue(PROJECT_CODE_NAME_PLACEHOLDER, props.ProjectCodeName), _
                                        KeyValue(PROJECT_LIB_FOLDER_PLACEHOLDER, Props_Project.ProjectLibFolderPath), _
                                        KeyValue(PROJECT_TOOL_FOLDER_PLACEHOLDER, Props_Project.ProjectToolFolderPath), _
                                        KeyValue(VBA.UCase$(PROJECT_CODE_NAME_PLACEHOLDER), Props_Project.ProjectCodeNameUCase) _
                                        )
    End If
    Set getAllPlaceholders = instance
End Function

Public Function getAllPlaceholdersRegex() As String
    Static regex As String
    '------------------------------------------------------------------------------------------------------
    If VBA.Len(regex) Then
        regex = f.Collections.toString( _
                    f.dictionaries.toCollection(getAllPlaceholders, DictPart_KeyOnly), _
                    StringifyMode_Normal, "|")
    End If
    getAllPlaceholdersRegex = regex
End Function
