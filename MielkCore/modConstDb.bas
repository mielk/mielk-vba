Attribute VB_Name = "modConstDb"
Option Explicit

Private Const CLASS_NAME As String = "modConstDb"

'[Objects signatures] -------------------------------------------------------------------------------------
Public Const DEFAULT_ACCESS_PROVIDER As String = "Microsoft.ACE.OLEDB.12.0"
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
Public Const adLockBatchOptimistic As Long = 4
'----------------------------------------------------------------------------------------------------------
Public Const FIELD_IS_ACTIVE As String = "is_active"
'----------------------------------------------------------------------------------------------------------
Public Const SQL_SELECT_ALL As String = "SELECT * FROM {0}"
Public Const SQL_SELECT_ALL_WITH_SORTING As String = "SELECT * FROM {0} ORDER BY {1} {2}"
Public Const SQL_SELECT_RECORD As String = "SELECT * FROM {0} WHERE {1} = {2}"
Public Const SQL_DELETE_ALL As String = "DELETE FROM {0}"
Public Const SQL_DELETE As String = "DELETE FROM {0} WHERE {1} = {2}"
Public Const SQL_DEACTIVATE As String = "UPDATE {0} SET " & FIELD_IS_ACTIVE & " = False WHERE {1} = {2}"
'----------------------------------------------------------------------------------------------------------
Public Const DB_NULL As String = "NULL"
'----------------------------------------------------------------------------------------------------------
