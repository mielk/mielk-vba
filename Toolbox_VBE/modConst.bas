Attribute VB_Name = "modConst"
Option Explicit

Private Const CLASS_NAME As String = "modConst"
'----------------------------------------------------------------------------------------------------------
Public Const CUSTOM_MENU_ITEM_TAG As String = "CustomContextMenuItem"
Public Const MENU_BAR_ITEM As String = "Menu Bar"
Public Const CODE_WINDOW_NAME As String = "Code Window"
'----------------------------------------------------------------------------------------------------------
Public Const CUSTOM_MENU_CAPTION As String = "Mielk"
'[VBE tags] -----------------------------------------------------------------------------------------------
Public Const META_TAG_NAME As String = "name"
Public Const META_TAG_DECLARATIONS As String = "declarations"
Public Const META_TAG_METHODS As String = "methods"
Public Const META_TAG_PROC_TYPE As String = "procType"
Public Const META_TAG_DESCRIPTION As String = "description"
Public Const META_TAG_BODY As String = "body"
'[Keywords] -----------------------------------------------------------------------------------------------
'[Scope type]
Public Const VBA_PUBLIC As String = "Public"
Public Const VBA_PRIVATE As String = "Private"
Public Const VBA_FRIEND As String = "Friend"
'[Var types]
Public Const VBA_DIM As String = "Dim"
Public Const VBA_CONST As String = "Const"
Public Const VBA_STATIC As String = "Static"
'[Method type]
Public Const VBA_SUB As String = "Sub"
Public Const VBA_FUNCTION As String = "Function"
Public Const VBA_PROPERTY As String = "Property"
Public Const VBA_PROPERTY_LET As String = "Property Let"
Public Const VBA_PROPERTY_SET As String = "Property Set"
Public Const VBA_PROPERTY_GET As String = "Property Get"
'[ByRef/ByVal]
Public Const VBA_BY_REF As String = "ByRef"
Public Const VBA_BY_VAL As String = "ByVal"
'[Other keywords]
Public Const VBA_OPTIONAL As String = "Optional"
Public Const VBA_PARAM_ARRAY As String = "ParamArray"
Public Const VBA_LINE_BREAK As String = " _"
Public Const VBA_ARRAY_BRACKETS As String = "()"
Public Const VBA_OPTION_EXPLICIT As String = "Option Explicit"
'[Common regex patterns]
Public Const DBO_TABLE_REGEX_PATTERN As String = "^\[dbo\]\.\[(\w+)\]$"
Public Const DBO_TABLE_BUILD_PATTERN As String = "[dbo].[{0}]"
'[Repository builder]
Public Const STRING_PROPERTY_SUFFIX As String = "Str"
'----------------------------------------------------------------------------------------------------------
Public Const MODULE_LEVEL_SEPARATOR_LENGTH As Long = 106
Public Const MODULE_LEVEL_SEPARATOR_INDENT As Long = 0
Public Const METHOD_LEVEL_SEPARATOR_LENGTH As Long = 102
Public Const METHOD_LEVEL_SEPARATOR_INDENT As Long = 4
'----------------------------------------------------------------------------------------------------------


Public Function filterComponentByName(componentName As String) As Boolean
    Static dict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If dict Is Nothing Then
        Set dict = F.dictionaries.Create(False)
        With dict
            'List of VB components to be processed.
            'If left empty, all components are processed.
'            Call .Add("FBooks", 0)
        End With
    End If
    
    If dict.count Then
        filterComponentByName = dict.Exists(componentName)
    Else
        filterComponentByName = True
    End If
    
End Function
