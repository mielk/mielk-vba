VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "IDbConnection"
'----------------------------------------------------------------------------------------------------------



'[Connecting]
Public Sub setConnectionString(conStr As String)
End Sub

Public Sub setConnectionDetails(provider As String, dataSource As String, dbName As String, Optional userName As String, Optional password As String)
End Sub

Public Function getConnectionString() As String
End Function

Public Function getConnection() As Object ' ADODB.connection
End Function

Public Function getState() As Long
End Function

Public Sub openConnection()
End Sub

Public Sub closeConnection(Optional rollbackTransactionIfOpen As Boolean = False)
End Sub

Public Function toString() As String
End Function

Public Function cloneConnection() As IDbConnection
End Function



'[Transactions]
Public Function isTransactionOpen() As Boolean
End Function

Public Sub beginTransaction(uuid As String, Optional forceNewTransaction As Boolean = False)
End Sub

Public Sub commitTransaction(uuid As String)
End Sub

Public Sub rollbackTransaction(Optional uuid As String)
End Sub



'[Errors]
Public Function getErrors() As VBA.Collection
End Function

Public Function hasErrors() As Boolean
End Function



'[Fetching data and executing queries]
Public Function Execute(sqlString As String, Optional lockType As Long = adLockPessimistic, Optional handleErrors As Boolean = True) As Object ' ADODB.recordset
End Function

Public Function getRecordset(sqlString As String, Optional lockType As Long = adLockPessimistic) As Object 'ADODB.recordset
End Function

Public Function getRecordsetFromStoredProcedure(procName As String, params As Scripting.Dictionary) As Object 'ADODB.recordset
End Function

Public Function getDtos(sqlString As String, Optional lockType As Long = adLockPessimistic) As VBA.Collection
End Function

Public Function getArray(sqlString As String, Optional includeHeaderRow As Boolean = True, Optional lockType As Long = adLockPessimistic) As Variant()
End Function

Public Function getPrimitiveValue(sqlString As String, Optional lockType As Long = adLockPessimistic) As Variant
End Function

Public Function getScalarFunctionValue(sqlString As String, Optional lockType As Long = adLockPessimistic) As Variant
End Function

Public Function runStoredProcedure(procName As String, params As Scripting.Dictionary, Optional asynchronous As Boolean = True) As ValidationObject
End Function

Public Function runStoredProcedureWithCommand(procName As String, params As Scripting.Dictionary, Optional asynchronous As Boolean = True) As ValidationObject
End Function



'[Services]
Public Function Meta() As IMeta
End Function

Public Function Query() As IQuery
End Function

Public Function Delete() As IDelete
End Function

Public Function Insert() As IInsert
End Function

Public Function Update() As IUpdate
End Function
