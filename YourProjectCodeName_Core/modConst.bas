Attribute VB_Name = "modConst"
Option Explicit

Private Const CLASS_NAME As String = "modConst"
'[Registry key] -------------------------------------------------------------------------------------------
Public Const REG_KEY_NAME As String = "Software\mielk\YourProjectCodeName\"
'[Application constants] ----------------------------------------------------------------------------------
Public Const APPLICATION_NAME As String = "YourProjectName"
Public Const APPLICATION_CODE_NAME As String = "YourProjectCodeName"
Public Const APPLICATION_VERSION As String = "0.0.1"
Public Const VIEW_WORKBOOK_NAME As String = "YourProjectCodeName-view.xl[as]m"
'[Context menu] -------------------------------------------------------------------------------------------
Public Const CONTEXT_MENU_PREFIX As String = "YourProjectCodeName_"
'----------------------------------------------------------------------------------------------------------



Public Function IsDevMode() As Boolean
    IsDevMode = DEV_MODE
End Function

Public Function IsLoggingOn() As Boolean
    IsLoggingOn = LOGGING_MODE
End Function

