VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDbConnection

Private Const CLASS_NAME As String = "AccessConnection"
'[Internal variables] -------------------------------------------------------------------------------------
'Private WithEvents pConnection As ADODB.connection
Private pConnection As Object 'ADODB.connection
Private pConnectionString As String
'[State] --------------------------------------------------------------------------------------------------
Private pErrors As VBA.Collection
Private pIsSuccessful As Boolean
Private pIsTransactionOpen As Boolean
Private pTransactionUuid As String
'---------------------------------------------------------------------------------------------------------
'Event Complete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal command As ADODB.command, ByVal recordset As ADODB.recordset, ByVal connection As ADODB.connection)
'Event BeginTransComplete()
'Event CommitTransComplete()
'Event RollbackTransComplete()
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
    Set Meta = New AccessMeta
    Call Meta.setConnection(Me)
End Function

Public Function Query() As IQuery
    Set Query = New AccessQuery
    Call Query.setConnection(Me)
End Function

Public Function Delete() As IDelete
    Set Delete = New AccessDelete
    Call Delete.setConnection(Me)
End Function

Public Function Insert() As IInsert
    Set Insert = New AccessInsert
    Call Insert.setConnection(Me)
End Function

Public Function Update() As IUpdate
    Set Update = New AccessUpdate
    Call Update.setConnection(Me)
End Function





'[CONNECTION]
Public Sub setConnectionString(conStr As String)
    pConnectionString = conStr
    Call recreateConnection
End Sub

Public Sub setConnectionDetails(dataSource As String, readWriteMode As ReadWriteModeEnum, Optional password As String = vbNullString)
    Const CONN_STRING_PATTERN As String = "Provider={0};User ID=Admin;Data Source={1};Mode={2};Jet OLEDB:System database="""";" & _
                                          "Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password=""{3}"";Jet OLEDB:Engine Type=6;" & _
                                          "Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;" & _
                                          "Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";" & _
                                          "Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;" & _
                                          "Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;" & _
                                          "Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;" & _
                                          "Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False;"
    '------------------------------------------------------------------------------------------------------
    pConnectionString = F.Strings.Format(CONN_STRING_PATTERN, DEFAULT_ACCESS_PROVIDER, _
                                         dataSource, getReadWriteModeString(readWriteMode), password)
    Call recreateConnection
End Sub

Public Function getConnection() As Object 'ADODB.connection
    Set getConnection = pConnection
End Function

Friend Sub openConnection()
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
    Const MAX_ATTEMPTS As Integer = 50
    Const MAX_TIME As Integer = 30
    '------------------------------------------------------------------------------------------------------
    Dim attempt As Integer
    Dim startTime As Double
    '------------------------------------------------------------------------------------------------------

    startTime = VBA.Timer

    For attempt = 1 To MAX_ATTEMPTS
        Call D.Functions.reopenConnection(pConnection)
        If pConnection.State >= 1 Then Exit For
        Call sleep(250)
        If (VBA.Timer - startTime) > MAX_TIME Then Exit For
        Call recreateConnection
    Next attempt

    DoEvents

    If pConnection.State = 0 Then
        Call Exceptions.DbConnectionException.Raise(getConnectionStringFromAdodbConnection)
    End If

End Function

Private Function getConnectionStringFromAdodbConnection(Optional maskPassword As Boolean = True) As String
    Const PWRD_REGEX As String = "(password|pwrd)=""?[^""]+?""?;"
    Const PWRD_REPLACEMENT As String = "Password=***;"
    '------------------------------------------------------------------------------------------------------
    Dim str As String
    '------------------------------------------------------------------------------------------------------
    str = VBA.CStr(pConnection.ConnectionString)
    If maskPassword Then str = F.regex.Replace(str, PWRD_REGEX, PWRD_REPLACEMENT)
    getConnectionStringFromAdodbConnection = str
End Function

Public Sub reconnect()
    Call closeConnection
    Call recreateConnection
End Sub

Private Sub recreateConnection()
    Set pConnection = Nothing
    Set pConnection = VBA.CreateObject(ADODB_CONNECTION) ' New ADODB.connection
    pConnection.ConnectionString = pConnectionString
End Sub

Public Function clone() As IDbConnection
    Set clone = New AccessConnection
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
    If forceNewTransaction Then
        On Error Resume Next
        Call pConnection.RollbackTrans
        Call pConnection.Close
        pIsTransactionOpen = False
        If Not DEV_MODE Then On Error GoTo ErrHandler
    End If
    
    If Not pIsTransactionOpen Then
        Set pErrors = New VBA.Collection
        Call openConnection
        Call getConnection.BeginTrans
        pTransactionUuid = uuid
        pIsTransactionOpen = True
    End If
    
    
ErrHandler:
End Sub

Public Sub rollbackTransaction(Optional uuid As String)
    If VBA.Len(uuid) = 0 Or F.Strings.compareStrings(uuid, pTransactionUuid) Then
        On Error Resume Next
        Call getConnection.RollbackTrans
        pIsTransactionOpen = False
        pTransactionUuid = vbNullString
        Call closeConnection
    End If
End Sub

Public Sub commitTransaction(uuid As String)
    Dim errNumber As Long, errDescription As String
    '----------------------------------------------------------------------------------------------------------
    Dim uuidMatched As Boolean
    Dim skipCommit As Boolean
    '----------------------------------------------------------------------------------------------------------
    Dim timeout As Long
    Dim maxAttempts As Long
    Dim startTime As Double
    Dim attempts As Long
    '----------------------------------------------------------------------------------------------------------
    
    If pIsTransactionOpen Then
        uuidMatched = F.Strings.compareStrings(uuid, pTransactionUuid)
    End If
    
    If uuidMatched Then
        If F.Collections.isNonEmptyCollection(pErrors) Then
            Call getConnection.RollbackTrans
            pIsSuccessful = True
        Else
            
            'Try to commit transaction as long as it is successful or timeout is exceeded. --------|
            timeout = D.getTimeout
            maxAttempts = D.getMaxAttempts
            startTime = VBA.Timer
            
            Do
                Call VBA.Err.Clear
                On Error Resume Next
                Call getConnection.CommitTrans
                errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
                If errNumber = 0 Then
                    pIsSuccessful = True
                    Exit Do
                ElseIf D.isConcurrencyError(errNumber, errDescription) Then
                    If VBA.Timer - startTime > timeout And attempts >= maxAttempts Then
                        If Not DEV_MODE Then On Error GoTo ErrHandler
                        Call VBA.Err.Raise(1, Description:="Timeout (" & timeout & "s) | " & _
                                                            errNumber & " | " & errDescription)
                    End If
                Else
                    If Not DEV_MODE Then On Error GoTo ErrHandler
                    Call VBA.Err.Raise(errNumber, , errDescription)
                End If
                
                attempts = attempts + 1
                Call sleep(150)
                
            Loop
            '--------------------------------------------------------------------------------------|
            
        End If
    
        pTransactionUuid = vbNullString
        pIsTransactionOpen = False
        Call closeConnection
    
    End If
    
    
'==========================================================================================================
ExitPoint:
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    Stop
    'Call ErrorManager.addError(CLASS_NAME, method_name, errNumber, err
    
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
Public Function Execute(sqlString As String, Optional lockType As Long = adLockPessimistic, _
                            Optional handleErrors As Boolean = True) As Object 'ADODB.recordset
    Dim errNumber As Long, errDescription As String
    '----------------------------------------------------------------------------------------------------------
    Dim timeout As Long
    Dim maxAttempts As Long
    Dim startTime As Double
    Dim attempts As Long
    '----------------------------------------------------------------------------------------------------------

    Set Execute = VBA.CreateObject(ADODB_CONNECTION) ' New ADODB.recordset
    Call openConnection

    'Try to execute query as long as it is successful or timeout is exceeded.----- --------|
    timeout = D.getTimeout
    maxAttempts = D.getMaxAttempts
    startTime = VBA.Timer
    
    Do
        On Error Resume Next
        Call VBA.Err.Clear
        Call Execute.Open(sqlString, getConnection, lockType:=lockType)
        errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
        If errNumber = 0 Then
            Exit Do
        ElseIf D.isConcurrencyError(errNumber, errDescription) Then
            If VBA.Timer - startTime > timeout And attempts >= maxAttempts Then
                If Not DEV_MODE Then On Error GoTo ErrHandler
                Call VBA.Err.Raise(1, Description:="Timeout (" & timeout & "s) | " & _
                                                    errNumber & " | " & errDescription)
            End If
        Else
            If Not DEV_MODE Then On Error GoTo ErrHandler
            Call VBA.Err.Raise(errNumber, , errDescription)
        End If
        attempts = attempts + 1
        Call sleep(150)
    Loop
    '--------------------------------------------------------------------------------------|

    '[Restore error handling]
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
'==========================================================================================================
ExitPoint:
    
    If Not pIsTransactionOpen Then
        Call closeConnection
    ElseIf errNumber And handleErrors Then
        Call closeConnection
    End If
    
    Stop
    'if errNumber then call vba.Err.Raise(errNumber
    Exit Function
    
'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    If handleErrors Then
        Stop
        'call ErrorManager.addError
    End If
    GoTo ExitPoint
    
End Function









Public Function getRecordset(sqlString As String, Optional lockType As Long = adLockPessimistic) As Object ' ADODB.recordset
    Dim errNumber As Long, errDescription As String
    '----------------------------------------------------------------------------------------------------------
    Dim timeout As Long
    Dim maxAttempts As Long
    Dim startTime As Double
    Dim attempts As Long
    '----------------------------------------------------------------------------------------------------------

    Set getRecordset = VBA.CreateObject(ADODB_RECORDSET)
    Call openConnection

    'Try to execute query as long as it is successful or timeout is exceeded.----- --------|
    timeout = D.getTimeout
    maxAttempts = D.getMaxAttempts
    startTime = VBA.Timer
    
    Do
        On Error Resume Next
        Call VBA.Err.Clear
        Call getRecordset.Open(sqlString, getConnection, lockType:=lockType)
        errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
        If errNumber = 0 Then
            Exit Do
        ElseIf D.isConcurrencyError(errNumber, errDescription) Then
            If VBA.Timer - startTime > timeout And attempts >= maxAttempts Then
                If Not DEV_MODE Then On Error GoTo ErrHandler
                Call VBA.Err.Raise(1, Description:="Timeout (" & timeout & "s) | " & _
                                                    errNumber & " | " & errDescription)
            End If
        Else
            If Not DEV_MODE Then On Error GoTo ErrHandler
            Call VBA.Err.Raise(errNumber, , errDescription)
        End If
        attempts = attempts + 1
        Call sleep(150)
    Loop
    '--------------------------------------------------------------------------------------|

'==========================================================================================================
ExitPoint:
    If errNumber Then
        Call closeConnection
        Stop
        'call vba.Err.Raise(errNumber,
    End If
    Exit Function
    
'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    
End Function






'[Wrappers]
Public Function getPrimitiveValue(sqlString As String, Optional lockType As ADODB.LockTypeEnum = adLockPessimistic) As Variant
    Dim rs As Object 'ADODB.recordset
    '----------------------------------------------------------------------------------------------------------
    
    Set rs = getRecordset(sqlString, lockType)
    Do Until rs.EOF
        getPrimitiveValue = rs.Fields(0).value
        Exit Do
    Loop
        
End Function



Public Function getDtos(sqlString As String, Optional lockType As Long = adLockPessimistic) As VBA.Collection
    Dim rs As Object 'ADODB.recordset
    '------------------------------------------------------------------------------------------------------
    
    Set rs = getRecordset(sqlString, lockType)
    Set getDtos = D.Functions.convertRecordsetToDtosCollection(rs)
    
ExitPoint:
    If Not pConnection Is Nothing Then Call D.Connections.closeRecordset(rs, Not pConnection.isTransactionOpen, True)

End Function



Public Function getArray(sqlString As String, Optional includeHeaderRow As Boolean = True, Optional lockType As Long = adLockPessimistic) As Variant()
    Dim rs As Object 'ADODB.recordset
    '------------------------------------------------------------------------------------------------------
    
    Set rs = getRecordset(sqlString, lockType)
    getArray = D.Functions.convertRecordsetToArray(rs, includeHeaderRow)
    
ExitPoint:
    If Not pConnection Is Nothing Then Call D.Connections.closeRecordset(rs, Not pConnection.isTransactionOpen, True)

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
'Private Sub pConnection_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal command As ADODB.command, ByVal recordset As ADODB.recordset, ByVal connection As ADODB.connection)
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

Private Sub IDbConnection_setConnectionDetails(provider As String, dataSource As String, dbName As String, Optional userName As String, Optional password As String)
    Stop
End Sub

Private Function IDbConnection_getConnectionString() As String
    IDbConnection_getConnectionString = pConnectionString
End Function

Private Function IDbConnection_getConnection() As Object ' ADODB.connection
    Set IDbConnection_getConnection = getConnection
End Function

Private Sub IDbConnection_openConnection()
    Call openConnection
End Sub

Private Function IDbConnection_cloneConnection() As IDbConnection
    Set IDbConnection_cloneConnection = clone
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

Private Function IDbConnection_execute(sqlString As String, Optional lockType As Long = adLockPessimistic, Optional handleErrors As Boolean = True) As Object 'ADODB.recordset
    Set IDbConnection_execute = Execute(sqlString, lockType, handleErrors)
End Function

Private Function IDbConnection_getRecordset(sqlString As String, Optional lockType As Long = adLockPessimistic) As Object 'ADODB.recordset
    Set IDbConnection_getRecordset = getRecordset(sqlString, lockType)
End Function

Private Function IDbConnection_getRecordsetFromStoredProcedure(procName As String, params As Scripting.IDictionary) As Object
    'N/A
End Function

Private Function IDbConnection_getState() As Long
    If pConnection Is Nothing Then
        IDbConnection_getState = -1
    Else
        IDbConnection_getState = pConnection.State
    End If
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
    IDbConnection_getScalarFunctionValue = getPrimitiveValue(sqlString, lockType)
End Function

Private Function IDbConnection_runStoredProcedure(procName As String, params As Scripting.Dictionary, Optional asynchronous As Boolean = True) As ValidationObject
    'Call runStoredProcedure(procName, params, asynchronous)
End Function

Private Function IDbConnection_runStoredProcedureWithCommand(procName As String, params As Scripting.Dictionary, Optional asynchronous As Boolean = True) As ValidationObject
    'Call runStoredProcedureWithCommand(procName, params, asynchronous)
End Function

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

Private Function IDbConnection_toString() As String
    Dim prop As Variant
    Dim str As String
    '------------------------------------------------------------------------------------------------------
    
    IDbConnection_toString = "==========================================================" & VBA.vbCrLf & _
                             VBA.Format(VBA.Now, "YYYY-MM-DD HH:MM:SS") & VBA.vbCrLf & _
                             "State: " & pConnection.State & VBA.vbCrLf & _
                             "Transaction: " & pTransactionUuid & VBA.vbCrLf
    
    For Each prop In pConnection.properties
        IDbConnection_toString = IDbConnection_toString & _
                            prop.Name & " | " & prop.value & VBA.vbCrLf
    Next prop
    
End Function


