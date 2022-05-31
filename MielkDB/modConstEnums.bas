Attribute VB_Name = "modConstEnums"
Option Explicit

Private Const CLASS_NAME As String = "modConstEnums"
'[Default connection properties] --------------------------------------------------------------------------
'[SQL Server]
Public Const DEFAULT_OLEDB_PROVIDER As String = "SQLOLEDB"
Public Const DEFAULT_INTEGRATED_SECURITY As String = "SSPI"
'[MS Access]
Public Const DEFAULT_ACCESS_PROVIDER As String = "Microsoft.ACE.OLEDB.12.0"
'[Object typenames]
Public Const ADODB_CONNECTION As String = "ADODB.Connection"
Public Const ADODB_RECORDSET As String = "ADODB.Recordset"
'[ADO enums] ----------------------------------------------------------------------------------------------
Public Const adOpenForwardOnly As Long = 0
Public Const adOpenKeyset As Long = 1
Public Const adOpenDynamic As Long = 2
Public Const adOpenStatic As Long = 3
Public Const adLockReadOnly As Long = 1
Public Const adLockPessimistic As Long = 2
Public Const adLockOptimistic As Long = 3
Public Const adLockBatchOptimistic As Long = 1
'[ADO Data types]
Public Const adInteger = 3
Public Const adDecimal = 14
Public Const adVarWChar = 202
'----------------------------------------------------------------------------------------------------------
Public Const GET_DB_FIELD_FUNCTION_NAME As String = "getDbField"
'----------------------------------------------------------------------------------------------------------

Public Enum DbSortOrderEnum
    DbSortOrder_None = 0
    DbSortOrder_Asc = 1
    DbSortOrder_Desc = 2
End Enum

Public Enum ReadWriteModeEnum
    ReadWriteMode_ReadWrite = 0
    ReadWriteMode_ReadOnly = 1
End Enum

Public Enum ComparisonModeEnum
    ComparisonMode_Equal = 0
    ComparisonMode_NotEqual = 1
    ComparisonMode_GreaterThan = 2
    ComparisonMode_LessThan = 3
    ComparisonMode_In = 4
End Enum

Public Enum ConnectionTypeEnum
    ConnectionType_SqlServer = 1
    ConnectionType_MsAccess = 2
End Enum




'[Converters]
Public Function getReadWriteModeString(mode As ReadWriteModeEnum) As String
    Select Case mode
        Case ReadWriteMode_ReadWrite:       getReadWriteModeString = "ReadWrite"
        Case ReadWriteMode_ReadOnly:        getReadWriteModeString = "ReadOnly"
    End Select
End Function
