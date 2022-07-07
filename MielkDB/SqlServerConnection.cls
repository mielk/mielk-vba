VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlServerConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDbConnection

Private Const CLASS_NAME As String = "SqlServerConnection"
'[Internal variables] -------------------------------------------------------------------------------------
Private WithEvents pConnection As ADODB.connection
Attribute pConnection.VB_VarHelpID = -1
Private pConnectionString As String
'[State] --------------------------------------------------------------------------------------------------
Private pErrors As VBA.Collection
Private pIsSuccessful As Boolean
Private pIsTransactionOpen As Boolean
Private pTransactionUuid As String
'---------------------------------------------------------------------------------------------------------
Event Complete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal command As ADODB.command, ByVal recordset As Object, ByVal connection As ADODB.connection)
Event BeginTransComplete()
Event CommitTransComplete()
Event RollbackTransComplete()
'---------------------------------------------------------------------------------------------------------






'[INITIALIZE]
Private Sub Class_Initialize()
    Call initializeContainers
End Sub

Private Sub initializeContainers()
    Set pErrors = New VBA.Collection
End Sub




'[SERVICES]
Public Function Meta() As IMeta
    Set Meta = New SqlServerMeta
    Call Meta.setConnection(Me)
End Function

Public Function Query() As IQuery
    Set Query = New SqlServerQuery
    Call Query.setConnection(Me)
End Function

Public Function Delete() As IDelete
    Set Delete = New SqlServerDelete
    Call Delete.setConnection(Me)
End Function

Public Function Insert() As IInsert
    Set Insert = New SqlServerInsert
    Call Insert.setConnection(Me)
End Function

Public Function Update() As IUpdate
    Set Update = New SqlServerUpdate
    Call Update.setConnection(Me)
End Function





'[CONNECTION]
Public Sub setConnectionString(conStr As String)
    pConnectionString = conStr
    Call recreateConnection
End Sub

Public Sub setConnectionDetails(dataSource As String, dbName As String, Optional provider As String = DEFAULT_OLEDB_PROVIDER, Optional userName As String, Optional password As String)
    pConnectionString = getConnectionString(provider, dataSource, dbName, userName, password)
    Call recreateConnection
End Sub

Private Function getConnectionString(provider As String, dataSource As String, dbName As String, Optional userName As String, Optional password As String) As String
    Const CONNECTION_PATTERN As String = "Provider={0};Data Source={1};Initial Catalog={2};Integrated Security={3}"
    Const CONNECTION_PATTERN_WITH_PASSWORD As String = "Provider={0};Data Source={1};Initial Catalog={2};User ID={3};Password={4};"
    '------------------------------------------------------------------------------------------------------
    If VBA.Len(userName) > 0 Then
        getConnectionString = F.Strings.Format(CONNECTION_PATTERN_WITH_PASSWORD, provider, dataSource, dbName, userName, password)
    Else
        getConnectionString = F.Strings.Format(CONNECTION_PATTERN, provider, dataSource, dbName, DEFAULT_INTEGRATED_SECURITY)
    End If
End Function

Public Function getConnection() As Object 'ADODB.connection
    Set getConnection = pConnection
End Function

Private Sub openConnection()
    Call ensureConnectionIsSet
    Call ensureConnectionIsOpen
End Sub

Private Function ensureConnectionIsSet()
    If pConnection Is Nothing Then
        Call reconnect
    End If
End Function

Private Function ensureConnectionIsOpen()
    Const METHOD_NAME As String = "ensureConnectionIsOpen"
    Const MAX_ATTEMPTS As Integer = 3
    '------------------------------------------------------------------------------------------------------
    Dim attempt As Integer
    '------------------------------------------------------------------------------------------------------

    For attempt = 1 To MAX_ATTEMPTS
        If pConnection.State >= 1 Then Exit For
        Call D.Functions.reopenConnection(pConnection)
    Next attempt

    DoEvents

    If pConnection.State = 0 Then
        Call Exceptions.DbConnectionException.Raise(pConnection.ConnectionString)
    End If

End Function

Public Sub reconnect()
    Call closeConnection
    Call recreateConnection
End Sub

Private Sub recreateConnection()
    Set pConnection = New ADODB.connection
    pConnection.ConnectionString = pConnectionString
End Sub

Public Function clone() As SqlServerConnection
    Set clone = New SqlServerConnection
    Call clone.setConnectionString(pConnectionString)
End Function

Private Sub closeConnection(Optional rollbackTransactionIfOpen As Boolean = False, _
                            Optional ignoreCloseWhileInTransactionError As Boolean = True)
                            
    If pConnection Is Nothing Then
        '
    ElseIf pConnection.State = 0 Then
        '
    Else
        If isTransactionOpen Then
            If rollbackTransactionIfOpen Then
                Call rollbackTransaction(vbNullString)
                Call pConnection.Close
            End If
        Else
            Call pConnection.Close
        End If
    End If
    
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Exit Sub

ErrHandler:
    errNumber = Err.Number: errDescription = Err.Description
    If ignoreCloseWhileInTransactionError And errNumber = Exceptions.DbClosingWhileInTransaction.getNumber Then
        If F.System.isDeveloper Then Stop
        Call VBA.Err.Clear
    Else
        Stop
        'Call ErrorManager.addError
    End If
    GoTo ExitPoint
    
End Sub



'[TRANSACTIONS]
Public Sub beginTransaction(uuid As String, Optional forceNewTransaction As Boolean = False)
    Set pErrors = New VBA.Collection
    
    If forceNewTransaction Then
        On Error Resume Next
        Call pConnection.RollbackTrans
        Call pConnection.Close
        pIsTransactionOpen = False
        If Not DEV_MODE Then On Error GoTo ErrHandler
    End If
    
    If Not pIsTransactionOpen Then
        If pConnection.State = 0 Then Call pConnection.Open
        Call pConnection.BeginTrans
        pTransactionUuid = uuid
        pIsTransactionOpen = True
    End If
    
ErrHandler:
End Sub

Public Sub rollbackTransaction(Optional uuid As String)
    If F.Strings.compareStrings(uuid, pTransactionUuid) Or VBA.Len(uuid) = 0 Then
        On Error Resume Next
        Call getConnection.RollbackTrans
        pIsTransactionOpen = False
        pTransactionUuid = vbNullString
    End If
End Sub

Public Sub commitTransaction(uuid As String)
    Dim errNumber As Long, errDescription As String
    '----------------------------------------------------------------------------------------------------------
    
    On Error GoTo Exception
    
    If pIsTransactionOpen And F.Strings.compareStrings(uuid, pTransactionUuid) Then
        If F.Collections.isNonEmptyCollection(pErrors) Then
            Call getConnection.RollbackTrans
            pIsSuccessful = False
        Else
            Call getConnection.CommitTrans
            pIsSuccessful = True
        End If
        pTransactionUuid = vbNullString
        pIsTransactionOpen = False
    End If

'==========================================================================================================
ExitPoint:
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
Exception:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    Call Exceptions.DbTransactionCommitException.Raise
    
End Sub

Public Function isTransactionOpen() As Boolean
    isTransactionOpen = pIsTransactionOpen
End Function




'[ERRORS]
Public Function getErrors() As VBA.Collection
    Set getErrors = pErrors
End Function

Public Function hasErrors() As Boolean
    hasErrors = (pErrors.Count > 0)
End Function





'[FETCHING DATA]
'[Basic operations]
Public Function Execute(sqlString As String, Optional lockType As Long = adLockOptimistic, Optional handleErrors As Boolean = True) As Object
    Set Execute = VBA.CreateObject(ADODB_RECORDSET)
    If Not pIsTransactionOpen Then Call openConnection
    Call Execute.Open(sqlString, getConnection, lockType:=lockType)
    If Not pIsTransactionOpen Then Call closeConnection
End Function

Public Function getRecordset(sqlString As String, Optional lockType As Long = adLockOptimistic) As Object
    Set getRecordset = VBA.CreateObject(ADODB_RECORDSET)
    Call openConnection
    Call getRecordset.Open(sqlString, getConnection, lockType:=lockType)
End Function



'[Wrappers]
Public Function getPrimitiveValue(sqlString As String, Optional lockType As Long = adLockOptimistic) As Variant
    Dim rs As Object
    '----------------------------------------------------------------------------------------------------------
    Set rs = getConnection.Execute(sqlString, lockType)
    Do Until rs.EOF
        getPrimitiveValue = rs.Fields(0).value
        Exit Do
    Loop
    Call cleanUp(rs)
End Function

Public Function getDtos(sqlString As String, Optional lockType As Long = adLockOptimistic) As VBA.Collection
    Dim rs As Object
    '------------------------------------------------------------------------------------------------------
    Set rs = getRecordset(sqlString, lockType)
    Set getDtos = D.Functions.convertRecordsetToDtosCollection(rs)
    Call cleanUp(rs)
End Function

Public Function getArray(sqlString As String, Optional includeHeaderRow As Boolean = True, Optional lockType As Long = adLockOptimistic) As Variant()
    Dim rs As Object
    '------------------------------------------------------------------------------------------------------
    Set rs = getRecordset(sqlString, lockType)
    getArray = D.Functions.convertRecordsetToArray(rs, includeHeaderRow)
    Call cleanUp(rs)
End Function

Public Function getScalarFunctionValue(sqlString As String, Optional lockType As Long = adLockOptimistic) As Variant
    Dim rs As Object
    '----------------------------------------------------------------------------------------------------------
    Set rs = getRecordset(sqlString, lockType)
    Do Until rs.EOF
        getScalarFunctionValue = rs.Fields(0).value
        Exit Do
    Loop
    Call cleanUp(rs)
End Function

Private Sub cleanUp(rs As Object)
    Call D.Connections.closeRecordset(rs, True, True)
End Sub



'[STORED PROCEDURES]
Public Sub runStoredProcedure(procedureName As String, params As Scripting.IDictionary, asynchronous As Boolean)
    Const METHOD_NAME As String = "runStoredProcedure"
    Const SQL_PATTERN As String = "EXEC {0} {1}"
    '------------------------------------------------------------------------------------------------------
    Dim sqlCommand As String
    Dim paramsText As String
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo ExceptionHandler
    
    If F.dictionaries.isNonEmptyDictionary(params) Then
        paramsText = F.dictionaries.toString(params, DictPart_KeyAndValue, StringifyMode_Db, ", ", " = ", "@", vbNullString)
    End If
    sqlCommand = F.Strings.Format(SQL_PATTERN, procedureName, paramsText)
    
    With getConnection
        If .State <> 1 Then Call ensureConnectionIsOpen
        .CommandTimeout = 0
        DoEvents
        Debug.Print "PROC: " & sqlCommand
        Call .Execute(sqlCommand) ',  options:=VBA.IIf(asynchronous, adAsyncExecute, adOptionUnspecified))
        DoEvents
    End With
    
'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then
        Call Exceptions.DbStoredProcedureError.Raise(sqlCommand, errNumber, errDescription)
    End If
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
ExceptionHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
    
End Sub

Public Sub runStoredProcedureWithCommand(procedureName As String, params As Scripting.IDictionary, asynchronous As Boolean)
    Const METHOD_NAME As String = "runStoredProcedureWithCommand"
    '------------------------------------------------------------------------------------------------------
    Dim command As Object ' ADODB.command
    Dim varKey As Variant
    Dim value As Variant
    Dim sql As String
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo ExceptionHandler
    
    Set command = VBA.CreateObject("ADODB.command")
    With command
        .ActiveConnection = getConnection
        .CommandText = procedureName
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        If F.dictionaries.isNonEmptyDictionary(params) Then
            For Each varKey In params.Keys
                value = params.Item(varKey)
                Call .Parameters.Append(.CreateParameter(VBA.CStr(varKey), D.Functions.getAdDataType(value), adParamInput, , value))
            Next varKey
        End If
        sql = .CommandText
        Call .Execute(options:=VBA.IIf(asynchronous, adAsyncExecute, adOptionUnspecified))
    End With

'==========================================================================================================
ExitPoint:
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
ExceptionHandler:
    Call Exceptions.SqlException.Raise(sql)
    GoTo ExitPoint
    
End Sub

Public Function getRecordsetFromStoredProcedure(procName As String, params As Scripting.Dictionary) As Object 'ADODB.Recordset
    Const METHOD_NAME As String = "runStoredProcedure"
    Const SQL_PATTERN As String = "EXEC {0} {1}"
    '------------------------------------------------------------------------------------------------------
    Dim sqlCommand As String
    Dim paramsText As String
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo ExceptionHandler
    
    If F.dictionaries.isNonEmptyDictionary(params) Then
        paramsText = F.dictionaries.toString(params, DictPart_KeyAndValue, StringifyMode_Db, ", ", " = ", "@", vbNullString)
    End If
    sqlCommand = F.Strings.Format(SQL_PATTERN, procName, paramsText)
    
    With getConnection
        If .State <> 1 Then Call ensureConnectionIsOpen
        .CommandTimeout = 0
        Debug.Print "PROC: " & sqlCommand
        Set getRecordsetFromStoredProcedure = .Execute(sqlCommand)
    End With
    
'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then
        Call Exceptions.DbStoredProcedureError.Raise(sqlCommand, errNumber, errDescription)
    End If
    Exit Function
    
'----------------------------------------------------------------------------------------------------------
ExceptionHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
    
End Function






'[CONNECTION EVENTS]
'Private Sub pConnection_BeginTransComplete(ByVal TransactionLevel As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.connection)
'    RaiseEvent BeginTransComplete
'End Sub
'
'Private Sub pConnection_CommitTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.connection)
'    RaiseEvent CommitTransComplete
'End Sub
'
'Private Sub pConnection_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal command As ADODB.command, ByVal recordset As Object, ByVal connection As ADODB.connection)
'    RaiseEvent Complete(RecordsAffected, pError, adStatus, command, recordset, connection)
'End Sub
'
'Private Sub pConnection_RollbackTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.connection)
'    RaiseEvent RollbackTransComplete
'End Sub






'[Implementation of IDbConnection interface]
Private Sub IDbConnection_setConnectionString(conStr As String)
    Call setConnectionString(conStr)
End Sub

Private Sub IDbConnection_openConnection()
    Call ensureConnectionIsOpen
End Sub

Private Function IDbConnection_getConnection() As Object 'ADODB.connection
    Set IDbConnection_getConnection = getConnection
End Function

Private Function IDbConnection_clone() As IDbConnection
    Set IDbConnection_clone = clone
End Function

Private Sub IDbConnection_closeConnection(Optional rollbackTransactionIfOpen As Boolean = False)
    Call closeConnection(rollbackTransactionIfOpen)
End Sub

Private Sub IDbConnection_beginTransaction(uuid As String, Optional forceNewTransaction As Boolean = False)
    Call beginTransaction(uuid, forceNewTransaction)
End Sub


Private Sub IDbConnection_commitTransaction(uuid As String)
    Call commitTransaction(uuid)
End Sub

Private Sub IDbConnection_rollbackTransaction(Optional uuid As String)
    Call rollbackTransaction(uuid)
End Sub

Private Function IDbConnection_isTransactionOpen() As Boolean
    IDbConnection_isTransactionOpen = isTransactionOpen
End Function

Private Function IDbConnection_getErrors() As VBA.Collection
    Set IDbConnection_getErrors = getErrors
End Function

Private Function IDbConnection_hasErrors() As Boolean
    IDbConnection_hasErrors = hasErrors
End Function

Private Function IDbConnection_execute(sqlString As String, Optional lockType As Long = adLockPessimistic, Optional handleErrors As Boolean = True) As Object
    Set IDbConnection_execute = Execute(sqlString, lockType, handleErrors)
End Function

Private Function IDbConnection_getRecordset(sqlString As String, Optional lockType As Long = adLockPessimistic) As Object
    Set IDbConnection_getRecordset = getRecordset(sqlString, lockType)
End Function

Private Function IDbConnection_getRecordsetFromStoredProcedure(procName As String, params As Scripting.IDictionary) As Object
    Set IDbConnection_getRecordsetFromStoredProcedure = getRecordsetFromStoredProcedure(procName, params)
End Function

Private Function IDbConnection_getDtos(sqlString As String, Optional lockType As Long = adLockPessimistic) As VBA.Collection
    Set IDbConnection_getDtos = getDtos(sqlString, lockType)
End Function

Private Function IDbConnection_getArray(sqlString As String, Optional includeHeaderRow As Boolean = True, Optional lockType As Long = adLockPessimistic) As Variant()
    IDbConnection_getArray = getArray(sqlString, includeHeaderRow, lockType)
End Function

Private Function IDbConnection_getPrimitiveValue(sqlString As String, Optional lockType As Long = adLockPessimistic) As Variant
    IDbConnection_getPrimitiveValue = getPrimitiveValue(sqlString, lockType)
End Function

Private Function IDbConnection_getScalarFunctionValue(sqlString As String, Optional lockType As Long = adLockPessimistic) As Variant
    IDbConnection_getScalarFunctionValue = getScalarFunctionValue(sqlString, lockType)
End Function

Private Sub IDbConnection_runStoredProcedure(procName As String, params As Scripting.Dictionary, Optional asynchronous As Boolean = True)
    Call runStoredProcedure(procName, params, asynchronous)
End Sub

Private Sub IDbConnection_runStoredProcedureWithCommand(procName As String, params As Scripting.Dictionary, Optional asynchronous As Boolean = True)
    Call runStoredProcedureWithCommand(procName, params, asynchronous)
End Sub

Private Function IDbConnection_Meta() As IMeta
    Set IDbConnection_Meta = Meta
End Function

Private Function IDbConnection_Query() As IQuery
    Set IDbConnection_Query = Query
End Function

Private Function IDbConnection_Delete() As IDelete
    Set IDbConnection_Delete = Delete
End Function

Private Function IDbConnection_Insert() As IInsert
    Set IDbConnection_Insert = Insert
End Function

Private Function IDbConnection_Update() As IUpdate
    Set IDbConnection_Update = Update
End Function



Private Function IDbConnection_getConnectionString() As String
    Stop
End Function

Private Function IDbConnection_toString() As String
    Stop
End Function

Private Function IDbConnection_getState() As Long
    If pConnection Is Nothing Then
        IDbConnection_getState = -1
    Else
        IDbConnection_getState = pConnection.State
    End If
End Function
