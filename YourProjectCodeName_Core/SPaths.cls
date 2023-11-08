VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SPaths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "SPaths"
'----------------------------------------------------------------------------------------------------------
'Private Const MAIN_FOLDER_PATH As String = "_______________________"
Private Const MAIN_FOLDER_PATH As String = "C:\Users\Tomek\Dropbox\tm\mielk\UTD"
'[Folder name constants] --------------------------------------------------------------------------------------
Private Const FOLDER_NAME_BACKUPS As String = "nackups"
Private Const FOLDER_NAME_CONFIG As String = "config"
Private Const FOLDER_NAME_DB As String = "db"
Private Const FOLDER_NAME_ERRORS As String = "errors"
Private Const FOLDER_NAME_LOGS As String = "logs"
Private Const FOLDER_NAME_MSGS As String = "msgs"
Private Const FOLDER_NAME_QUERIES As String = "queries"
Private Const FOLDER_NAME_RESOURCES As String = "resources"
Private Const FOLDER_NAME_RIBBON As String = "ribbon"
Private Const FOLDER_NAME_TEMPLATES As String = "templates"
Private Const FOLDER_NAME_TEMPORARY_FILES As String = "temp"
'[File name constants] ------------------------------------------------------------------------------------
Private Const RIBBON_JSON_FILE As String = "ribbon.json"
Private Const ENVIRONMENT_INFO_FILE As String = "env.txt"
Private Const DB_PROXY_NAME_PATTERN As String = "BE_LOCAL.accdb$"
Private Const ERRORS_DB_NAME As String = "ERR.accdb"
'[Passwords] ----------------------------------------------------------------------------------------------
Private Const BACK_END_PASSWORD As String = vbNullString
'[Placeholders] -------------------------------------------------------------------------------------------
Private Const TOOL_FOLDER_PLACEHOLDER As String = "{tool}"
Private Const RESOURCES_FOLDER_PLACEHOLDER As String = "{resources}"
'----------------------------------------------------------------------------------------------------------





'[Specific folders]
Public Function ToolFolderPath() As String
    ToolFolderPath = ThisWorkbook.path
End Function

Public Function MainFolderPath() As String
    MainFolderPath = MAIN_FOLDER_PATH
End Function

Public Function ResourcesFolderPath() As String
    ResourcesFolderPath = F.Files.concatenatePath(ToolFolderPath, FOLDER_NAME_RESOURCES)
End Function

Public Function TemplatesFolderPath() As String
    TemplatesFolderPath = F.Files.concatenatePath(ToolFolderPath, FOLDER_NAME_TEMPLATES)
End Function

Public Function RibbonFolderPath() As String
    RibbonFolderPath = F.Files.concatenatePath(ResourcesFolderPath, FOLDER_NAME_RIBBON)
End Function

Public Function MessagesFolderPath() As String
    MessagesFolderPath = F.Files.concatenatePath(ResourcesFolderPath, FOLDER_NAME_MSGS)
End Function

Public Function QueriesFolderPath() As String
    QueriesFolderPath = F.Files.concatenatePath(ResourcesFolderPath, FOLDER_NAME_QUERIES)
End Function

Public Function ConfigFolderPath() As String
    ConfigFolderPath = F.Files.concatenatePath(ResourcesFolderPath, FOLDER_NAME_CONFIG)
End Function

Public Function TempFolderPath() As String
    TempFolderPath = F.Files.concatenatePath(ResourcesFolderPath, FOLDER_NAME_TEMPORARY_FILES)
End Function

Public Function BackupsFolderPath() As String
    BackupsFolderPath = F.Files.concatenatePath(ToolFolderPath, FOLDER_NAME_BACKUPS)
End Function



'[Static paths]
Public Function RibbonConfigFilePath() As String
    RibbonConfigFilePath = F.Files.concatenatePath(RibbonFolderPath, RIBBON_JSON_FILE)
End Function

Public Function EnvironmentNameFilePath() As String
    EnvironmentNameFilePath = F.Files.concatenatePath(ResourcesFolderPath, ENVIRONMENT_INFO_FILE)
End Function



'[Dynamic paths]
Public Function getResourcesFilePath(ByVal filepath As String) As String
    getResourcesFilePath = VBA.Replace(filepath, RESOURCES_FOLDER_PLACEHOLDER, ResourcesFolderPath)
End Function

Public Function getTempFilePath(ByVal fileTag As String) As String
    getTempFilePath = F.Files.concatenatePath(TemplatesFolderPath, fileTag)
End Function

Public Function getConfigFilePath(ByVal fileTag As String) As String
    getConfigFilePath = F.Files.concatenatePath(ConfigFolderPath, fileTag & ".json")
End Function



'[Database]
Public Function BackEndPath() As String
    Static path As String
    '----------------------------------------------------------------------------------------------------------
    Dim dbProxyFolder As String
    Dim dbProxy As Scripting.File
    '----------------------------------------------------------------------------------------------------------
    
    If VBA.Len(path) = 0 Then
        dbProxyFolder = F.Files.concatenatePath(ToolFolderPath, FOLDER_NAME_DB)
        Set dbProxy = F.Files.getFolderNewestFile(dbProxyFolder, False, DB_PROXY_NAME_PATTERN)
        If Not dbProxy Is Nothing Then
            path = dbProxy.path
        End If
    End If
    
    BackEndPath = path
    
End Function

Public Function BackEndPassword() As String
    BackEndPassword = BACK_END_PASSWORD
End Function




'[Errors]
Public Function ErrorsDbPath() As String
    ErrorsDbPath = F.Files.concatenatePath(MainFolderPath, FOLDER_NAME_ERRORS, ERRORS_DB_NAME)
End Function

Public Function ErrorsDbPassword() As String
    ErrorsDbPassword = vbNullString
End Function

Public Function ErrorLogsFolderPath() As String
    ErrorLogsFolderPath = F.Files.concatenatePath(MainFolderPath, FOLDER_NAME_ERRORS)
End Function



'[Action logs]
Public Function LogsFolderPath() As String
    LogsFolderPath = F.Files.concatenatePath(MainFolderPath, FOLDER_NAME_LOGS)
End Function



'[Deployment]
Public Function DeploySettingsJson() As String
    Stop
End Function
