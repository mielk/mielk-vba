Attribute VB_Name = "modConstructors"
Option Explicit

Private Const CLASS_NAME As String = "modConstructors"
'----------------------------------------------------------------------------------------------------------



'[Singletons]
Public Function Msg() As MsgService
    Static instance As MsgService
    '------------------------------------------------------------------------------------------------------
    
    If instance Is Nothing Then
        Set instance = MielkCore.MsgService
        Call instance.loadJsonsFromFolder(Paths.MessagesFolderPath)
        Call AppSettings.loadLanguageFromRegistry
    End If
    
    Set Msg = instance
    
End Function

Public Function Config() As SConfig
    Static instance As SConfig
    '------------------------------------------------------------------------------------------------------
    
    If instance Is Nothing Then
        Set instance = New SConfig
        Call instance.loadJsonsFromFolder
    End If
    
    Set Config = instance
    
End Function

Public Function Paths(Optional inject As SPaths) As SPaths
    Static instance As SPaths
    '------------------------------------------------------------------------------------------------------
    
    If Not inject Is Nothing Then Set instance = inject
    If instance Is Nothing Then Set instance = New SPaths
    Set Paths = instance
    
End Function

Public Function Props_Project() As CProperties
    Static instance As CProperties
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then Set instance = New CProperties
    Set Props_Project = instance
End Function

Public Function DataTypes() As CDataTypes
    Static instance As CDataTypes
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then Set instance = New CDataTypes
    Set DataTypes = instance
End Function

Public Function ErrorManager() As ErrorManager
    Static instance As ErrorManager
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then
        Set instance = MielkCore.ErrorManager
        With instance
            Call .setConnectionUrl(Paths.ErrorsDbPath, Paths.ErrorsDbPassword)
            Call .setLogFolderPath(Paths.ErrorLogsFolderPath)
        End With
    End If
    Set ErrorManager = instance
End Function


Public Function State(Optional inject As SState) As SState
    Static instance As SState
    '------------------------------------------------------------------------------------------------------
    If Not inject Is Nothing Then Set instance = inject
    If instance Is Nothing Then Set instance = New SState
    Set State = instance
End Function


Public Function TextfileActionLogger() As STextfileActionLogger
    Static instance As STextfileActionLogger
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then Set instance = New STextfileActionLogger
    Set TextfileActionLogger = instance
End Function





Public Function ProgressBar() As WProgressBar
    Static instance As WProgressBar
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then Set instance = MielkUI.ProgressBar
    Set ProgressBar = instance
End Function

Public Function AppSettings() As SAppSettings
    Static instance As SAppSettings
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then Set instance = New SAppSettings
    Set AppSettings = instance
End Function

Public Function App() As ParentApp
    Static instance As ParentApp
    '------------------------------------------------------------------------------------------------------
    
    If instance Is Nothing Then
        Set instance = MielkCore.App
        With instance
            Call .setName(APPLICATION_NAME)
            Call .setVersion(APPLICATION_VERSION)
            Call .setPath(F.files.getUncPath(Excel.Workbooks(VIEW_WORKBOOK_NAME).FullName))
        End With
    End If
    Set App = instance
    
End Function



Public Function RibbonManager(Optional inject As RibbonManager) As RibbonManager
    Static instance As RibbonManager
    '------------------------------------------------------------------------------------------------------
    If Not inject Is Nothing Then
        Set instance = inject
    End If
    Set RibbonManager = instance
End Function
