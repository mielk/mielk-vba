VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

' By Chip Pearson, www.cpearson.com, chip@cpearson.com

Private Const CLASS_NAME As String = "FRegistry"
    '----------------------------------------------------------------------------------------------------------
Private Const C_REG_ERR_NO_ERROR = 0
Private Const C_REG_ERR_INVALID_BASE_KEY = vbObjectError + 1
Private Const C_REG_ERR_INVALID_DATA_TYPE = vbObjectError + 2
Private Const C_REG_ERR_KEY_NOT_FOUND = vbObjectError + 3
Private Const C_REG_ERR_VALUE_NOT_FOUND = vbObjectError + 4
Private Const C_REG_ERR_DATA_TYPE_MISMATCH = vbObjectError + 5
Private Const C_REG_ERR_ENTRY_LOCKED = vbObjectError + 6
Private Const C_REG_ERR_INVALID_KEYNAME = vbObjectError + 7
Private Const C_REG_ERR_UNABLE_TO_OPEN_KEY = vbObjectError + 8
Private Const C_REG_ERR_UNABLE_TO_READ_KEY = vbObjectError + 9
Private Const C_REG_ERR_UNABLE_TO_CREATE_KEY = vbObjectError + 10
Private Const C_REG_ERR_UBABLE_TO_READ_VALUE = vbObjectError + 11
Private Const C_REG_ERR_UNABLE_TO_UDPATE_VALUE = vbObjectError + 12
Private Const C_REG_ERR_UNABLE_TO_CREATE_VALUE = vbObjectError + 13
Private Const C_REG_ERR_UNABLE_TO_DELETE_KEY = vbObjectError + 14
Private Const C_REG_ERR_UNABLE_TO_DELETE_VALUE = vbObjectError + 15
'--- [API Constants] --------------------------------------------------------------------------------------
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_USERS = &H80000003
Private Const HKCU = HKEY_CURRENT_USER
Private Const HKLM = HKEY_LOCAL_MACHINE
Private Const REGSTR_MAX_VALUE_LENGTH  As Long = &H100
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_OPENED_EXISTING_KEY = &H2
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF
Private Const REG_OPTION_NON_VOLATILE = 0&
Private Const REG_OPTION_VOLATILE = &H1
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_ACCESS_DENIED = 5
Private Const ERROR_INVALID_DATA = 13&
Private Const ERROR_MORE_DATA = 234 '  dderror
Private Const ERROR_NO_MORE_ITEMS = 259
Private Const S_OK = &H0
Private Const MAX_DATA_BUFFER_SIZE = 1024
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY  As Long = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE  As Long = &H800
Private Const FORMAT_MESSAGE_FROM_STRING  As Long = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM  As Long = &H1000
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK  As Long = &HFF
Private Const FORMAT_MESSAGE_IGNORE_INSERTS  As Long = &H200
Private Const FORMAT_MESSAGE_TEXT_LEN  As Long = &HA0  ' from VC++ ERRORS.H file
Private Const REG_INVALID As Long = -1&
Private Const REG_SZ As Long = 1&
Private Const REG_DWORD As Long = 4&
'--- [API Types] ------------------------------------------------------------------------------------------

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type

Private Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        control As Long
        Owner As Long
        group As Long
        Sacl As ACL
        Dacl As ACL
End Type

'--- [API Functions] --------------------------------------------------------------------------------------
#If VBA7 Then

    Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" ( _
        ByVal hKey As LongPtr) As LongPtr

    Private Declare PtrSafe Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" ( _
        ByVal hKey As LongPtr, _
        ByVal lpSubKey As String, _
        ByVal Reserved As Long, _
        ByVal lpClass As String, _
        ByVal dwOptions As Long, _
        ByVal samDesired As Long, _
        lpSecurityAttributes As SECURITY_ATTRIBUTES, _
        phkResult As Long, _
        lpdwDisposition As Long) As Long
    
    Private Declare PtrSafe Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" ( _
        ByVal hKey As LongPtr, _
        ByVal lpSubKey As String) As Long
    
    Private Declare PtrSafe Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" ( _
        ByVal hKey As LongPtr, _
        ByVal lpSubKey As String, _
        phkResult As Long) As Long
    
    Private Declare PtrSafe Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" ( _
        ByVal hKey As LongPtr, _
        ByVal lpValueName As String) As Long
    
    Private Declare PtrSafe Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" ( _
        ByVal hKey As LongPtr, _
        ByVal dwIndex As LongPtr, _
        ByVal lpName As String, _
        ByVal cbName As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" ( _
        ByVal hKey As LongPtr, _
        ByVal dwIndex As LongPtr, _
        ByVal lpName As String, _
        lpcbName As LongPtr, _
        ByVal lpReserved As LongPtr, _
        ByVal lpClass As String, _
        lpcbClass As LongPtr, _
        lpftLastWriteTime As FILETIME) As LongPtr
    
    Private Declare PtrSafe Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" ( _
        ByVal hKey As LongPtr, _
        ByVal dwIndex As LongPtr, _
        ByVal lpValueName As String, _
        lpcbValueName As LongPtr, _
        ByVal lpReserved As LongPtr, _
        lpType As LongPtr, _
        lpData As Byte, _
        lpcbData As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function RegFlushKey Lib "advapi32.dll" ( _
        ByVal hKey As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function RegGetKeySecurity Lib "advapi32.dll" ( _
        ByVal hKey As LongPtr, _
        ByVal SecurityInformation As LongPtr, _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        lpcbSecurityDescriptor As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" ( _
        ByVal hKey As LongPtr, _
        ByVal lpClass As String, _
        lpcbClass As LongPtr, _
        ByVal lpReserved As LongPtr, _
        lpcSubKeys As LongPtr, _
        lpcbMaxSubKeyLen As LongPtr, _
        lpcbMaxClassLen As LongPtr, _
        lpcValues As LongPtr, _
        lpcbMaxValueNameLen As LongPtr, _
        lpcbMaxValueLen As LongPtr, _
        lpcbSecurityDescriptor As LongPtr, _
        lpftLastWriteTime As FILETIME) As LongPtr
    
    Private Declare PtrSafe Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" ( _
        ByVal hKey As LongPtr, _
        ByVal lpSubKey As String, _
        ByVal lpValue As String, _
        lpcbValue As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
        ByVal hKey As LongPtr, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        lpData As Any, _
        lpcbData As LongPtr) As Long
    
    Private Declare PtrSafe Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
        ByVal hKey As LongPtr, _
        ByVal lpValueName As String, _
        ByVal Reserved As LongPtr, _
        ByVal dwType As LongPtr, _
        lpData As Any, _
        ByVal cbData As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function RegSetValueExStr Lib "advapi32" Alias "RegSetValueExA" ( _
        ByVal hKey As LongPtr, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        ByVal szData As String, _
        ByVal cbData As Long) As Long
    
    Private Declare PtrSafe Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" ( _
        ByVal hKey As LongPtr, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        szData As Long, _
        ByVal cbData As Long) As Long
    
    Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" ( _
        ByVal hKey As LongPtr, _
        ByVal lpSubKey As String, _
        ByVal ulOptions As Long, _
        ByVal samDesired As Long, _
        phkResult As Long) As Long
    
    Private Declare PtrSafe Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" ( _
        ByVal hKey As LongPtr, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        ByRef lpType As Long, _
        ByVal szData As String, _
        ByRef lpcbData As Long) As Long
    
    Private Declare PtrSafe Function FormatMessage Lib "kernel32" _
        Alias "FormatMessageA" ( _
        ByVal dwFlags As LongPtr, _
        ByVal lpSource As Any, _
        ByVal dwMessageId As Long, _
        ByVal dwLanguageId As Long, _
        ByVal lpBuffer As String, _
        ByVal nSize As Long, _
        ByRef Arguments As Long) As Long

#Else

    Private Declare Function RegCloseKey Lib "advapi32.dll" ( _
        ByVal hKey As Long) As Long
    
    Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" ( _
        ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal Reserved As Long, _
        ByVal lpClass As String, _
        ByVal dwOptions As Long, _
        ByVal samDesired As Long, _
        lpSecurityAttributes As SECURITY_ATTRIBUTES, _
        phkResult As Long, _
        lpdwDisposition As Long) As Long
    
    Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" ( _
        ByVal hKey As Long, _
        ByVal lpSubKey As String) As Long
    
    Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" ( _
        ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        phkResult As Long) As Long
    
    Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" ( _
        ByVal hKey As Long, _
        ByVal lpValueName As String) As Long
    
    Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" ( _
        ByVal hKey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpName As String, _
        ByVal cbName As Long) As Long
    
    Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" ( _
        ByVal hKey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpName As String, _
        lpcbName As Long, _
        ByVal lpReserved As Long, _
        ByVal lpClass As String, _
        lpcbClass As Long, _
        lpftLastWriteTime As FILETIME) As Long
    
    Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" ( _
        ByVal hKey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpValueName As String, _
        lpcbValueName As Long, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        lpData As Byte, _
        lpcbData As Long) As Long
    
    Private Declare Function RegFlushKey Lib "advapi32.dll" ( _
        ByVal hKey As Long) As Long
    
    Private Declare Function RegGetKeySecurity Lib "advapi32.dll" ( _
        ByVal hKey As Long, _
        ByVal SecurityInformation As Long, _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        lpcbSecurityDescriptor As Long) As Long
    
    Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" ( _
        ByVal hKey As Long, _
        ByVal lpClass As String, _
        lpcbClass As Long, _
        ByVal lpReserved As Long, _
        lpcSubKeys As Long, _
        lpcbMaxSubKeyLen As Long, _
        lpcbMaxClassLen As Long, _
        lpcValues As Long, _
        lpcbMaxValueNameLen As Long, _
        lpcbMaxValueLen As Long, _
        lpcbSecurityDescriptor As Long, _
        lpftLastWriteTime As FILETIME) As Long
    
    Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" ( _
        ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal lpValue As String, _
        lpcbValue As Long) As Long
    
    Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
        ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        lpData As Any, _
        lpcbData As Long) As Long
    
    Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
        ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        lpData As Any, _
        ByVal cbData As Long) As Long
    
    Private Declare Function RegSetValueExStr Lib "advapi32" Alias "RegSetValueExA" ( _
        ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        ByVal szData As String, _
        ByVal cbData As Long) As Long
    
    Private Declare Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" ( _
        ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        szData As Long, _
        ByVal cbData As Long) As Long
    
    Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" ( _
        ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal ulOptions As Long, _
        ByVal samDesired As Long, _
        phkResult As Long) As Long
    
    Private Declare Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" ( _
        ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        ByRef lpType As Long, _
        ByVal szData As String, _
        ByRef lpcbData As Long) As Long
    
    Private Declare Function FormatMessage Lib "kernel32" _
        Alias "FormatMessageA" ( _
        ByVal dwFlags As Long, _
        ByVal lpSource As Any, _
        ByVal dwMessageId As Long, _
        ByVal dwLanguageId As Long, _
        ByVal lpBuffer As String, _
        ByVal nSize As Long, _
        ByRef Arguments As Long) As Long

#End If

'--- [Application constants] ------------------------------------------------------------------------------
Private Type RegValue
    valueName As String
    ValueValue As Variant
End Type
'--- [Private variables] ----------------------------------------------------------------------------------
#If VBA7 Then
    Private G_Reg_AppErrNum As Long
    Private G_Reg_SysErrNum As Long
#Else
    Private G_Reg_AppErrNum As Long
    Private G_Reg_SysErrNum As Long
#End If
Private G_Reg_AppErrText As String
Private G_Reg_SysErrText As String
    '----------------------------------------------------------------------------------------------------------






Public Function getValue(baseKey As Long, keyName As String, valueName As String) As Variant
    Dim hKey As Long
'#If VBA7 Then
'    Dim res As LongPtr
'    Dim regDataType As LongPtr
'#Else
    Dim res As Long
    Dim regDataType As Long
'#End If
    
    Dim lenData As Long
    Dim longData As Long
    Dim lenStringData As Long
    Dim stringData As String
    Dim intArr(0 To 1024) As Integer
    '------------------------------------------------------------------------------------------------------

    Call resetErrorVariables

    If Not isValidBaseKey(baseKey) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        getValue = Null
        Exit Function
    End If

    If Not isValidKeyName(keyName) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        getValue = Null
        Exit Function
    End If

    If Not keyExists(baseKey, keyName) Then
        G_Reg_AppErrNum = C_REG_ERR_KEY_NOT_FOUND
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        getValue = Null
        Exit Function
    End If

    regDataType = getValueType(baseKey, keyName, valueName)
    hKey = openRegistryKey(baseKey, keyName)
    If hKey = 0 Then
        G_Reg_SysErrNum = VBA.CLng(res)
        G_Reg_SysErrText = GetSystemErrorMessageText(ErrorNumber:=VBA.CLng(res))
        G_Reg_AppErrNum = C_REG_ERR_UNABLE_TO_OPEN_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        getValue = Null
        Exit Function
    End If


    If regDataType = REG_DWORD Then
        res = RegQueryValueEx(hKey, valueName, 0&, regDataType, longData, VBA.Len(longData))
        If res = ERROR_SUCCESS Then
            getValue = longData
            Exit Function
        Else
            G_Reg_SysErrNum = VBA.CLng(res)
            G_Reg_SysErrText = GetSystemErrorMessageText(ErrorNumber:=VBA.CLng(res))
            G_Reg_AppErrNum = C_REG_ERR_UBABLE_TO_READ_VALUE
            G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
            Call RegCloseKey(hKey)
            getValue = Null
            Exit Function
        End If
    ElseIf regDataType = REG_SZ Then
        stringData = VBA.String$(MAX_DATA_BUFFER_SIZE, VBA.vbNullChar)
        lenStringData = VBA.Len(stringData)
        res = RegQueryValueExStr(hKey, valueName, 0&, regDataType, stringData, lenStringData)
        If res <> ERROR_SUCCESS Then
            G_Reg_SysErrNum = VBA.CLng(res)
            G_Reg_SysErrText = GetSystemErrorMessageText(VBA.CLng(res))
            G_Reg_AppErrNum = C_REG_ERR_UBABLE_TO_READ_VALUE
            G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
            Call RegCloseKey(hKey)
            getValue = Null
            Exit Function
        End If
        stringData = trimToNull(stringData)
        getValue = stringData
    Else
        G_Reg_AppErrNum = C_REG_ERR_INVALID_DATA_TYPE
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        getValue = Null
    End If

End Function

Public Function keyExists(baseKey As Long, keyName As String, Optional createIfNotExists As Boolean = False) As Boolean
    Dim hKey As Long
    Dim res As Long
    '------------------------------------------------------------------------------------------------------
    
    Call resetErrorVariables
    
    If Not isValidBaseKey(baseKey) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        keyExists = False
    End If

    If Not isValidKeyName(keyName) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        keyExists = False
    End If

    res = RegOpenKey(baseKey, keyName, hKey)
    If res = ERROR_SUCCESS Then
        keyExists = True
    Else
        keyExists = False
        If createIfNotExists = True Then
            res = createKey(baseKey, keyName)
            keyExists = VBA.CBool(res)
        End If
    End If

    Call RegCloseKey(hKey)

End Function

Public Function valueExists(baseKey As Long, keyName As String, valueName As String, _
                            Optional createIfNotExists As Boolean = False, _
                            Optional createType As Long = REG_DWORD) As Boolean
    Dim hKey As Long
    Dim res As Long
    '------------------------------------------------------------------------------------------------------

    Call resetErrorVariables
    
    If Not isValidBaseKey(baseKey) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        valueExists = False
    End If

    If Not isValidKeyName(keyName) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        valueExists = False
    End If

    hKey = openRegistryKey(baseKey, keyName)
    If hKey = 0 Then
        G_Reg_AppErrNum = C_REG_ERR_UNABLE_TO_OPEN_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        valueExists = False
    End If

    res = RegQueryValueEx(hKey, valueName, 0&, 0&, 0&, 0&)
    If (res = ERROR_SUCCESS) Or (res = ERROR_MORE_DATA) Then
        valueExists = True
    Else
        If createIfNotExists Then
            If createType = REG_DWORD Then
                res = createValue(baseKey, keyName, valueName, 0&, True)
            Else
                res = createValue(baseKey, keyName, valueName, vbNullString, True)
            End If
            If VBA.CBool(res) = True Then
                valueExists = True
            Else
                valueExists = False
            End If
        End If
    End If

    Call RegCloseKey(hKey)

End Function

Public Function getValueType(baseKey As Long, keyName As String, valueName As String) As Long
    Dim res As Long
    Dim hKey As Long
    Dim DataType As Long
    '------------------------------------------------------------------------------------------------------

    Call resetErrorVariables

    If Not isValidBaseKey(baseKey) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        getValueType = False
    End If

    If Not isValidKeyName(keyName) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        getValueType = False
    End If

    res = RegOpenKey(baseKey, keyName, hKey)
    If res <> ERROR_SUCCESS Then
        G_Reg_AppErrNum = C_REG_ERR_UNABLE_TO_OPEN_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        getValueType = REG_INVALID
        Exit Function
    End If

    res = RegQueryValueEx(hKey, valueName, 0&, DataType, 0&, 0&)
    If (res <> ERROR_SUCCESS) And (res <> ERROR_MORE_DATA) Then
        G_Reg_AppErrNum = C_REG_ERR_UBABLE_TO_READ_VALUE
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        getValueType = REG_INVALID
        Call RegCloseKey(hKey)
        Exit Function
    End If

    Select Case DataType
        Case REG_SZ
            getValueType = REG_SZ
        Case REG_DWORD
            getValueType = REG_DWORD
        Case Else
            getValueType = REG_INVALID
    End Select

    Call RegCloseKey(hKey)

End Function

Public Function createValue(baseKey As Long, keyName As String, valueName As String, ValueValue As Variant, _
                            Optional CreateKeyIfNotExists As Boolean = False) As Boolean
    Dim hKey As Long
    Dim res As Long
    Dim DataType As Long
    Dim LongValue As Long
    Dim StringValue As String
    '------------------------------------------------------------------------------------------------------

    Call resetErrorVariables

    If Not isValidBaseKey(baseKey:=baseKey) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        createValue = False
        Exit Function
    End If

    If Not isValidKeyName(keyName:=keyName) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        createValue = False
        Exit Function
    End If

    If Not keyExists(baseKey, keyName, CreateKeyIfNotExists) Then
        G_Reg_AppErrNum = C_REG_ERR_KEY_NOT_FOUND
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        createValue = False
        Exit Function
    End If

    If Not isCompatibleValueValue(var:=ValueValue) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_DATA_TYPE
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        createValue = False
        Exit Function
    End If

    If Not keyExists(baseKey, keyName, False) Then
        If CreateKeyIfNotExists = True Then
            If Not keyExists(baseKey, keyName, True) Then
                G_Reg_AppErrNum = C_REG_ERR_UNABLE_TO_CREATE_KEY
                G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
                createValue = False
                Exit Function
            End If
        Else
            G_Reg_AppErrNum = C_REG_ERR_KEY_NOT_FOUND
            G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
            createValue = False
            Exit Function
        End If
    End If

    If valueExists(baseKey, keyName, valueName) Then
        DataType = getValueType(baseKey, keyName, valueName)
        If DataType = REG_SZ Then
            If VBA.varType(ValueValue) <> VBA.vbString Then
                G_Reg_AppErrNum = C_REG_ERR_DATA_TYPE_MISMATCH
                G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
                createValue = False
                Exit Function
            End If
        End If
    Else
        If VBA.varType(ValueValue) = VBA.vbString Then
            DataType = REG_SZ
        Else
            DataType = REG_DWORD
        End If
    End If

    If DataType = REG_DWORD Then
        LongValue = VBA.CLng(ValueValue)
        hKey = openRegistryKey(baseKey, keyName)
        If hKey = 0 Then
            G_Reg_SysErrNum = err.LastDllError
            G_Reg_SysErrText = GetSystemErrorMessageText(G_Reg_SysErrNum)
            G_Reg_AppErrNum = C_REG_ERR_UNABLE_TO_OPEN_KEY
            G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
            Call RegCloseKey(hKey)
            createValue = False
            Exit Function
        End If

        res = RegSetValueExLong(hKey, valueName, 0&, REG_DWORD, LongValue, VBA.Len(LongValue))
        If res <> ERROR_SUCCESS Then
            G_Reg_AppErrNum = C_REG_ERR_UNABLE_TO_UDPATE_VALUE
            G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
            Call RegCloseKey(hKey)
            createValue = False
            Exit Function
        End If
    Else
        StringValue = VBA.CStr(ValueValue)
        hKey = openRegistryKey(baseKey, keyName)
        If hKey = 0 Then
            G_Reg_SysErrNum = err.LastDllError
            G_Reg_SysErrText = GetSystemErrorMessageText(G_Reg_SysErrNum)
            G_Reg_AppErrNum = C_REG_ERR_UNABLE_TO_OPEN_KEY
            G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
            Call RegCloseKey(hKey)
            createValue = False
            Exit Function
        End If
        res = RegSetValueExStr(hKey, valueName, 0&, REG_SZ, StringValue, VBA.Len(StringValue))
        If res <> ERROR_SUCCESS Then
            G_Reg_AppErrNum = C_REG_ERR_UNABLE_TO_UDPATE_VALUE
            G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
            createValue = False
            Call RegCloseKey(hKey)
            Exit Function
        End If
    End If

    Call RegCloseKey(hKey)
    createValue = True

End Function

Public Function createKey(baseKey As Long, keyName As String) As Boolean
    Dim res As Long
    Dim hKey As Long
    Dim DataType As Long
    Dim Disposition As Long
    Dim SecAttrib As SECURITY_ATTRIBUTES
    '------------------------------------------------------------------------------------------------------

    Call resetErrorVariables

    If Not isValidBaseKey(baseKey) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        createKey = False
    End If

    If Not isValidKeyName(keyName) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        createKey = False
    End If

    If keyExists(baseKey, keyName) Then
        createKey = True
        Exit Function
    End If

    res = RegCreateKeyEx(baseKey, keyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SecAttrib, hKey, Disposition)
    If res <> ERROR_SUCCESS Then
        G_Reg_SysErrNum = res
        G_Reg_SysErrText = GetSystemErrorMessageText(G_Reg_SysErrNum)
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        createKey = False
        Exit Function
    End If

    createKey = True

End Function

Public Function deleteValue(baseKey As Long, keyName As String, valueName As String) As Boolean
    Dim res As Long
    Dim hKey As Long
    Dim DataType As Long
    Dim Disposition As Long
    Dim SecAttrib As SECURITY_ATTRIBUTES
    '------------------------------------------------------------------------------------------------------

    Call resetErrorVariables

    If Not isValidBaseKey(baseKey) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        deleteValue = False
        Exit Function
    End If

    If Not isValidKeyName(keyName) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        deleteValue = False
        Exit Function
    End If

    If Not keyExists(baseKey, keyName, False) Then
        G_Reg_AppErrNum = C_REG_ERR_KEY_NOT_FOUND
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        deleteValue = False
        Exit Function
    End If

    hKey = openRegistryKey(baseKey, keyName)
    If hKey = 0 Then
        deleteValue = False
        Exit Function
    End If
    If Not valueExists(baseKey, keyName, valueName) Then
        RegCloseKey hKey
        deleteValue = True
        Exit Function
    End If

    res = RegDeleteValue(hKey, valueName)
    If res <> ERROR_SUCCESS Then
        G_Reg_SysErrNum = res
        G_Reg_SysErrText = GetSystemErrorMessageText(ErrorNumber:=res)
        G_Reg_AppErrNum = C_REG_ERR_UNABLE_TO_DELETE_VALUE
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        Call RegCloseKey(hKey)
        deleteValue = False
        Exit Function
    End If

    Call RegCloseKey(hKey)
    deleteValue = True

End Function

Public Function deleteKey(baseKey As Long, keyName As String) As Boolean
    Dim res As Long
    Dim hKey As Long
    Dim DataType As Long
    Dim Disposition As Long
    Dim SecAttrib As SECURITY_ATTRIBUTES
    '------------------------------------------------------------------------------------------------------

    Call resetErrorVariables

    If Not isValidBaseKey(baseKey) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        deleteKey = False
        Exit Function
    End If

    If Not isValidKeyName(keyName) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        deleteKey = False
        Exit Function
    End If

    If Not keyExists(baseKey, keyName, False) Then
        deleteKey = True
        Exit Function
    End If

    hKey = openRegistryKey(baseKey, keyName)
    If hKey = 0 Then
        deleteKey = False
        Exit Function
    End If

    res = RegDeleteKey(baseKey, keyName)
    Call RegCloseKey(hKey)
    If res <> ERROR_SUCCESS Then
        G_Reg_SysErrNum = res
        G_Reg_SysErrText = GetSystemErrorMessageText(res)
        G_Reg_AppErrNum = C_REG_ERR_UNABLE_TO_DELETE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        deleteKey = False
        Exit Function
    End If

    deleteKey = True

End Function

Public Function updateValue(baseKey As Long, keyName As String, _
                            valueName As String, newValue As Variant, _
                            Optional createIfNotExists As Boolean = True) As Boolean
    Dim hKey As Long
    Dim res As Boolean
    '------------------------------------------------------------------------------------------------------
    
    Call resetErrorVariables

    If Not isValidBaseKey(baseKey) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        updateValue = False
        Exit Function
    End If

    If Not isValidKeyName(keyName) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        updateValue = False
        Exit Function
    End If

    If Not isCompatibleValueValue(newValue) Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_DATA_TYPE
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        updateValue = False
        Exit Function
    End If

    res = keyExists(baseKey, keyName, True)
    If res = False Then
        G_Reg_AppErrNum = C_REG_ERR_KEY_NOT_FOUND
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        updateValue = False
        Exit Function
    End If

    If VBA.varType(newValue) = vbString Then
        res = valueExists(baseKey, keyName, valueName, createIfNotExists, REG_DWORD)
    Else
        res = valueExists(baseKey, keyName, valueName, createIfNotExists, REG_SZ)
    End If

    If res = False Then
        G_Reg_AppErrNum = C_REG_ERR_VALUE_NOT_FOUND
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        updateValue = False
        Exit Function
    End If

    res = deleteValue(baseKey, keyName, valueName)
    res = createValue(baseKey, keyName, valueName, newValue, True)

    updateValue = res

End Function





'[PRIVATE Functions]
Private Function GetSystemErrorMessageText(ErrorNumber As Long) As String
    Dim ErrorText As String
    Dim TextLen As Long
    Dim FormatMessageResult As Long
    Dim LangID As Long
    '------------------------------------------------------------------------------------------------------
    
    LangID = 0&   ' Default language
    ErrorText = VBA.String$(FORMAT_MESSAGE_TEXT_LEN, vbNullChar)
    TextLen = FORMAT_MESSAGE_TEXT_LEN
    
    FormatMessageResult = FormatMessage( _
                            dwFlags:=FORMAT_MESSAGE_FROM_SYSTEM Or _
                                     FORMAT_MESSAGE_IGNORE_INSERTS, _
                            lpSource:=0&, _
                            dwMessageId:=ErrorNumber, _
                            dwLanguageId:=LangID, _
                            lpBuffer:=ErrorText, _
                            nSize:=TextLen, _
                            Arguments:=0&)
    
    If FormatMessageResult = 0& Then
        Call VBA.MsgBox("An error occurred with the FormatMessage" & _
                        " API function call." & vbCrLf & _
                        "Error: " & VBA.CStr(err.LastDllError) & _
                        " Hex(" & VBA.Hex(err.LastDllError) & ").")
        GetSystemErrorMessageText = "An internal system error occurred with the" & vbCrLf & _
                        "FormatMessage API function: " & CStr(err.LastDllError) & ". No futher information" & vbCrLf & _
                        "is available."
        Exit Function
    End If

    ErrorText = VBA.left$(ErrorText, VBA.CLng(FormatMessageResult))
    If VBA.Len(ErrorText) >= 2 Then
        If VBA.right$(ErrorText, 2) = vbCrLf Then
            ErrorText = VBA.left$(ErrorText, Len(ErrorText) - 2)
        End If
    End If
    
    GetSystemErrorMessageText = ErrorText

End Function

Private Function openRegistryKey(baseKey As Long, keyName As String) As Long
    Dim res As Long
    Dim hKey As Long
    '------------------------------------------------------------------------------------------------------

    Call resetErrorVariables
    
    If isValidBaseKey(baseKey) = False Then
        openRegistryKey = 0
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        Exit Function
    End If

    res = RegOpenKeyEx(baseKey, keyName, 0&, KEY_ALL_ACCESS, hKey)
    If res <> ERROR_SUCCESS Then
        openRegistryKey = 0
        G_Reg_SysErrNum = res
        G_Reg_SysErrText = GetSystemErrorMessageText(res)
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = getAppErrText(G_Reg_AppErrNum)
        Exit Function
    End If

    openRegistryKey = VBA.CLng(hKey)

End Function

Private Function trimToNull(text As String, Optional reverse As Boolean = False) As String
    Dim pos As Long
    '------------------------------------------------------------------------------------------------------
    
    If reverse = False Then
        pos = VBA.InStr(1, text, vbNullChar, vbTextCompare)
    Else
        pos = VBA.InStrRev(text, vbNullChar, -1, vbTextCompare)
    End If
    
    If pos Then
        trimToNull = VBA.left(text, pos - 1)
    Else
        trimToNull = text
    End If

End Function

Private Function isValidBaseKey(baseKey As Long) As Boolean
    Select Case baseKey
        Case HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, _
                HKEY_CLASSES_ROOT, HKEY_CURRENT_CONFIG, HKEY_DYN_DATA, _
                HKEY_PERFORMANCE_DATA, HKEY_USERS
            isValidBaseKey = True
        Case Else
            isValidBaseKey = False
    End Select
End Function

Private Sub resetErrorVariables()
    G_Reg_AppErrNum = C_REG_ERR_NO_ERROR
    G_Reg_AppErrText = vbNullString
    G_Reg_SysErrNum = C_REG_ERR_NO_ERROR
    G_Reg_SysErrText = vbNullString
End Sub

Private Function getAppErrText(errNum As Variant) As String
    Select Case errNum
        Case C_REG_ERR_NO_ERROR
            getAppErrText = vbNullString
        Case C_REG_ERR_INVALID_BASE_KEY
            getAppErrText = "Invalid Base Key Value."
        Case C_REG_ERR_INVALID_DATA_TYPE
            getAppErrText = "Invalid Data Type."
        Case C_REG_ERR_KEY_NOT_FOUND
            getAppErrText = "Key Not Found."
        Case C_REG_ERR_VALUE_NOT_FOUND
            getAppErrText = "Value Not Found."
        Case C_REG_ERR_DATA_TYPE_MISMATCH
            getAppErrText = "Value Data Type Mismatch."
        Case C_REG_ERR_ENTRY_LOCKED
            getAppErrText = "Registry Entry Locked."
        Case C_REG_ERR_INVALID_KEYNAME
            getAppErrText = "The Specified Key Is Invalid."
        Case C_REG_ERR_UNABLE_TO_OPEN_KEY
            getAppErrText = "Unable To Open Key."
        Case C_REG_ERR_UNABLE_TO_READ_KEY
            getAppErrText = "Unable To Read Key."
        Case C_REG_ERR_UNABLE_TO_CREATE_KEY
            getAppErrText = "Unable To Create Key."
        Case C_REG_ERR_UBABLE_TO_READ_VALUE
            getAppErrText = "Unable To Read Value."
        Case C_REG_ERR_UNABLE_TO_UDPATE_VALUE
            getAppErrText = "Unable To Update Value."
        Case C_REG_ERR_UNABLE_TO_CREATE_VALUE
            getAppErrText = "Unable To Create Value."
        Case C_REG_ERR_UNABLE_TO_DELETE_KEY
            getAppErrText = "Unable To Delete Key."
        Case C_REG_ERR_UNABLE_TO_DELETE_VALUE
            getAppErrText = "Unable To Delete Value."
        Case Else
            getAppErrText = "Undefined Error."
    End Select
End Function

Private Function isStringValidLength(text As String) As Boolean
    isStringValidLength = (VBA.Len(text) <= REGSTR_MAX_VALUE_LENGTH)
End Function

Private Function isValidKeyName(keyName As String) As Boolean
    isValidKeyName = (VBA.Len(keyName) <= REGSTR_MAX_VALUE_LENGTH) And (VBA.Len(trim(keyName)) > 0)
End Function


Private Function isValidDataType(DataType As Long) As Boolean
    Select Case DataType
        Case REG_SZ, REG_DWORD
            isValidDataType = True
        Case Else
            isValidDataType = False
    End Select
End Function

Private Function isCompatibleValueValue(var As Variant) As Boolean

    If VBA.varType(var) >= vbArray Then
        isCompatibleValueValue = False
        Exit Function
    End If
    
    If VBA.IsArray(var) = True Then
        isCompatibleValueValue = False
        Exit Function
    End If
    
    If VBA.IsObject(var) = True Then
        isCompatibleValueValue = False
        Exit Function
    End If

    Select Case VBA.varType(var)
        Case vbBoolean, vbByte, vbCurrency, vbDate, vbDouble, vbInteger, vbLong, vbSingle, vbString
            isCompatibleValueValue = True
        Case Else
            isCompatibleValueValue = False
    End Select

End Function
