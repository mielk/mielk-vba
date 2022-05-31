VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IInsert

Private Const CLASS_NAME As String = "AccessInsert"
'----------------------------------------------------------------------------------------------------------
Private pConnection As IDbConnection
Private pInto As String
Private pIdentityInsert As Boolean
Private pValues As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------




'[SETTERS]
Public Sub setConnection(value As IDbConnection)
    Set pConnection = value
End Sub






'[DEFINITION]
Public Function Into(tableName As String) As IInsert
    Set Into = Me
    pInto = tableName
End Function

Public Function SetIdentityInsert(value As Boolean) As IInsert
    Set SetIdentityInsert = Me
    pIdentityInsert = value
End Function

Public Function Values(properties As Scripting.Dictionary) As IInsert
    Set Values = Me
    Set pValues = properties
End Function




'[API]

Public Sub fromExcelToTempTable(filepath As String, sheetName As String, destinationTable As String)
    Const SQL_PATTERN As String = "SELECT * INTO [{2}] FROM " & _
                                  "[Excel 12.0;HDR=YES;IMEX=2;ACCDB=YES;DATABASE={0}].[{1}$]"
    '------------------------------------------------------------------------------------------------------
    Dim transactionUuid As String:              transactionUuid = F.Crypto.createUUID
    Dim sql As String
    '------------------------------------------------------------------------------------------------------
    
    sql = F.Strings.Format(SQL_PATTERN, filepath, sheetName, destinationTable)
    
    With pConnection
        Call .openConnection
        Call .Execute(sql)
    End With
    
End Sub
    

Public Sub ExecuteWithId(ByRef id As Long, ByVal idFieldName As String)
    Dim errNumber As Long, errDescription As String
    '------------------------------------------------------------------------------------------------------
    Dim rs As Object            'ADODB.Recordset
    Dim timeout As Long
    Dim maxAttempts As Long
    Dim startTime As Double
    Dim attempts As Long
    '------------------------------------------------------------------------------------------------------
    
    Call pConnection.openConnection
    Set rs = VBA.CreateObject(ADODB_RECORDSET)
    
    'Try to execute INSERT as long as it is succesful or timeout is exceeeded. --------------|
    timeout = D.getTimeout
    maxAttempts = D.getMaxAttempts
    startTime = VBA.Timer
    
    Do
        On Error Resume Next
        Call VBA.Err.Clear
        Call rs.Open(pInto, pConnection.getConnection, adOpenKeyset, adLockPessimistic)     '|
        Call rs.AddNew(pValues.Keys, pValues.Items)                                         '|
        
        errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
        If errNumber = 0 Then
            Exit Do
        Else
            If VBA.Timer - startTime > timeout And attempts >= maxAttempts Then
                If Not dev_mode Then On Error GoTo ErrHandler
                Call VBA.Err.Raise(1, Description:="Timeout (" & timeout & "s)")
            End If
        End If
        attempts = attempts + 1
        
    Loop
    '----------------------------------------------------------------------------------------|
    
    '[Restore error handling]
    If Not dev_mode Then On Error GoTo ErrHandler
    
    If VBA.Len(idFieldName) Then
        '------------------------------------------------------------------------------------|
        Do
            On Error Resume Next
            Call VBA.Err.Clear
            id = rs.Fields(idFieldName).value
            errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
            If errNumber = 0 Then
                Exit Do
            Else
                Stop
            End If
        Loop
        '------------------------------------------------------------------------------------|
    End If
    
    
ExitPoint:
    If Not pConnection Is Nothing And errNumber = 0 Then
        Call D.Connections.closeRecordset(rs, Not pConnection.isTransactionOpen)
    End If

ErrHandler:

End Sub



Public Sub Execute(Optional ByRef id As Long)
    Call pConnection.Execute(getSqlString)
End Sub

Private Function getSqlString() As String
    Const SQL_PATTERN As String = "INSERT INTO {0}({1}) {2}"
    '------------------------------------------------------------------------------------------------------
    getSqlString = F.Strings.Format(SQL_PATTERN, _
                            getIntoPart, _
                            getFieldsPart, _
                            getValuesPart)
End Function

Private Function getIntoPart() As String
    getIntoPart = pInto
End Function

Private Function getFieldsPart() As String
    getFieldsPart = F.Arrays.toString(pValues.Keys, StringifyMode_Normal, ", ", "[", "]")
End Function

Private Function getValuesPart() As String
    getValuesPart = "SELECT " & F.Arrays.toString(pValues.Items, StringifyMode_Db)
End Function











'[Implementation of IInsert]
Private Sub IInsert_setConnection(value As IDbConnection)
    Call setConnection(value)
End Sub

Private Function IInsert_Into(tableName As String) As IInsert
    Set IInsert_Into = Into(tableName)
End Function

Private Function IInsert_SetIdentityInsert(value As Boolean) As IInsert
    Stop
End Function

Private Function IInsert_Values(properties As Scripting.IDictionary) As IInsert
    Set IInsert_Values = Values(properties)
End Function

Private Sub IInsert_Execute()
    Call Execute
End Sub

Private Sub IInsert_ExecuteWithId(id As Long, ByVal idFieldName As String)
    Call ExecuteWithId(id, idFieldName)
End Sub

Private Sub IInsert_fromExcelToTempTable(filepath As String, sheetName As String, tableName As String)
    Call fromExcelToTempTable(filepath, sheetName, tableName)
End Sub

