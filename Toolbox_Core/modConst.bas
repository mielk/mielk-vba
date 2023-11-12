Attribute VB_Name = "modConst"
Option Explicit

Private Const CLASS_NAME As String = "modConst"
'[Registry key] -------------------------------------------------------------------------------------------
Public Const REG_KEY_NAME As String = "Software\mielk\toolbox\"
'[Application constants] ----------------------------------------------------------------------------------
Public Const APPLICATION_NAME As String = "VBA Toolbox"
Public Const APPLICATION_CODE_NAME As String = "toolbox"
Public Const APPLICATION_VERSION As String = "0.0.1"
Public Const VIEW_WORKBOOK_NAME As String = "toolbox.xlsm"
'[Context menu] -------------------------------------------------------------------------------------------
Public Const CONTEXT_MENU_PREFIX As String = "toolbox_"
'[Config names] -------------------------------------------------------------------------------------------
Public Const CONFIG_NEW_PROJECT As String = "new-project"
'[File names] ---------------------------------------------------------------------------------------------
Public Const VIEW_FILE_SUFFIX As String = "-view"
'----------------------------------------------------------------------------------------------------------

Public Function IsDevMode() As Boolean
    IsDevMode = DEV_MODE
End Function

Public Function IsLoggingOn() As Boolean
    IsLoggingOn = LOGGING_MODE
End Function

