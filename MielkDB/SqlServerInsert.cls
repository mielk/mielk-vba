VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlServerInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IInsert

Private Const CLASS_NAME As String = "SqlServerInsert"
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
Public Sub Execute(Optional ByRef id As Long)
    Dim rs As ADODB.recordset
    '------------------------------------------------------------------------------------------------------
    
    Set rs = New ADODB.recordset
    With rs
        .CursorLocation = adUseClient
        Call .Open(getSqlString, pConnection.getConnection, adCmdText)
    End With

    On Error Resume Next
    Set rs = rs.NextRecordset
    id = rs.Fields(0).value
    On Error GoTo 0
    
    '[Clean up]
    Call D.Connections.closeRecordset(rs)
    
End Sub




Private Function getSqlString() As String
    Const SQL_PATTERN As String = "{0} INSERT INTO {1}({2}) {3};{4}{5}"
    '------------------------------------------------------------------------------------------------------
    getSqlString = F.Strings.Format(SQL_PATTERN, _
                            getIndentityInsertOnPart, _
                            getIntoPart, _
                            getFieldsPart, _
                            getValuesPart, _
                            getSelectScopePart, _
                            getIndentityInsertOffPart)
End Function

Private Function getIntoPart() As String
    getIntoPart = pInto
End Function

Private Function getIndentityInsertOnPart() As String
    Const SQL_PATTERN_ON As String = "SET IDENTITY_INSERT {0} ON; "
    'Const SQL_PATTERN_OFF As String = "SET IDENTITY_INSERT {0} OFF; "
    '------------------------------------------------------------------------------------------------------
    If pIdentityInsert Then
        getIndentityInsertOnPart = F.Strings.Format(SQL_PATTERN_ON, pInto)
    'Else
    '    getIndentityInsertOnPart = F.Strings.format(SQL_PATTERN_OFF, pInto)
    End If
End Function

Private Function getIndentityInsertOffPart() As String
    Const SQL_PATTERN As String = "SET IDENTITY_INSERT {0} OFF; "
    '------------------------------------------------------------------------------------------------------
    If pIdentityInsert Then
        getIndentityInsertOffPart = F.Strings.Format(SQL_PATTERN, pInto)
    End If
End Function

Private Function getSelectScopePart() As String
    If Not pIdentityInsert Then
        getSelectScopePart = "SELECT SCOPE_IDENTITY(); "
    End If
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
    Set IInsert_SetIdentityInsert = SetIdentityInsert(value)
End Function

Private Function IInsert_Values(properties As Scripting.IDictionary) As IInsert
    Set IInsert_Values = Values(properties)
End Function

Private Sub IInsert_Execute()
    Stop
    'Call Execute(id)
End Sub

Private Sub IInsert_ExecuteWithId(id As Long, ByVal idFieldName As String)
    Stop
End Sub

Private Sub IInsert_fromExcelToTempTable(filepath As String, sheetName As String, tableName As String)
    Stop
End Sub
