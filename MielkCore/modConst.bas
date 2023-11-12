Attribute VB_Name = "modConst"
Option Explicit

Private Const CLASS_NAME As String = "modConst"
'----------------------------------------------------------------------------------------------------------
Public Const LIBRARY_NAME As String = "mielk"
'----------------------------------------------------------------------------------------------------------
Public Const EXCEL_APPLICATION As String = "Excel.Application"
'[Registry] -----------------------------------------------------------------------------------------------
Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
'[Methods] ------------------------------------------------------------------------------------------------
Public Const ESCAPE_CHARACTER As String = "\"
Public Const METHOD_START_TAG As String = "{"
Public Const METHOD_END_TAG As String = "}"
Public Const GET_NAME As String = "getName"
Public Const NEW_LINE_SYMBOL As String = "\n"
'[Validation constants]
Public Const VALUE_____ As String = "*%Value%*_;M}(H;C'M+?.>'#bx{pzk}2@.y%4Pr$z"
Public Const WARNING_CODE As Long = 123456789
'[Reflection] ---------------------------------------------------------------------------------------------
Public Const LOCAL_METHOD As String = "$."
'[Extensions] ---------------------------------------------------------------------------------------------
Public Const EXTENSION_TXT As String = ".txt"
Public Const EXTENSION_EXCEL_ADDIN As String = ".xlam"
Public Const EXTENSION_EXCEL_MACRO_FILE As String = ".xlsm"
Public Const EXTENSION_JSON As String = ".json"
Public Const EXTENSION_ZIP As String = ".zip"
'[Constants from other libraries - used with late binding] ------------------------------------------------
'[Outlook]
Public Const olMailItem As Long = 0
'[VBIDE]
Public Const vbext_pk_Proc As Long = 0
Public Const vbext_pk_Let As Long = 1
Public Const vbext_pk_Set As Long = 2
Public Const vbext_pk_Get As Long = 3
Public Const vbext_ct_StdModule As Long = 1
Public Const vbext_ct_ClassModule As Long = 2
Public Const vbext_ct_MSForm As Long = 3
Public Const vbext_ct_Document As Long = 100
'[View] ---------------------------------------------------------------------------------------------------
Public Const PIXEL_SIZE As Single = 0.75

'[Control keys] -------------------------------------------------------------------------------------------
Public Const SHIFT_MASK As Long = 1
Public Const CTRL_MASK As Long = 2
Public Const ALT_MASK As Long = 4



'#FORCHECK
'[Xml tags] -----------------------------------------------------------------------------------------------
Public Const XML_EMPTY As String = "#Empty"
Public Const XML_MISSING As String = "#Missing"
Public Const XML_NOTHING As String = "#Nothing"
Public Const XML_NULL As String = "#Null"
Public Const XML_ARRAY As String = "<array>{0}</array>"
Public Const XML_COLLECTION As String = "<collection>{0}</collection>"
Public Const XML_DICTIONARY As String = "<dictionary>{0}</dictionary>"
'[Db tags] ------------------------------------------------------------------------------------------------
Public Const DB_NULL As String = "NULL"
'[Context menu] -------------------------------------------------------------------------------------------
Public Const CONTEXT_MENU_TAG_METHOD_NAME As String = "methodName"
Public Const CONTEXT_MENU_TAG_METHOD_BOOK As String = "methodBook"
Public Const CONTEXT_MENU_TAG_CAPTION As String = "caption"
Public Const CONTEXT_MENU_TAG_TAG As String = "tag"
Public Const CONTEXT_MENU_TAG_FACE_ID As String = "faceId"
Public Const CONTEXT_MENU_TAG_SEPARATOR As String = "separator"
Public Const CONTEXT_MENU_TAG_IS_GROUP As String = "isGroup"
Public Const CONTEXT_MENU_TAG_ITEMS As String = "items"
Public Const CONTEXT_MENU_TAG_PARAM As String = "param"
'[Action logs] --------------------------------------------------------------------------------------------
Public Const ACTION_LOG_START As String = "Start app"
Public Const ACTION_LOG_CLOSE As String = "Close app"
'----------------------------------------------------------------------------------------------------------
