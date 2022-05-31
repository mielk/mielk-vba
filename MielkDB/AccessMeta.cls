VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IMeta

Private Const CLASS_NAME As String = "AccessMeta"
'----------------------------------------------------------------------------------------------------------
Private pConnection As IDbConnection
'----------------------------------------------------------------------------------------------------------


'[SETTERS]
Public Sub setConnection(value As IDbConnection)
    Set pConnection = value
End Sub



'[DB OPERATIONS]
Public Sub dropTableIfExists(tableName As String)
    Const SQL_PATTERN As String = "IF OBJECT_ID('{0}', 'U') IS NOT NULL DROP TABLE {0};"
    '------------------------------------------------------------------------------------------------------
    Dim sqlString As String
    '------------------------------------------------------------------------------------------------------
    sqlString = F.Strings.Format(SQL_PATTERN, tableName)
    Call pConnection.Execute(sqlString)
End Sub




'[DB STATE]
Public Function checkIfTableExists(tableName As String) As Boolean
    Const SQL_PATTERN As String = "SELECT IIF (EXISTS (" & _
                                    "SELECT * FROM INFORMATION_SCHEMA.TABLES " & _
                                    "EWhere CONCAT('[', TABLE_SCHEMA, '].[', TABLE_NAME, ']') = '{0}' " & _
                                            "OR TABLE_NAME = '{0}'), 1, 0)"
    '------------------------------------------------------------------------------------------------------
    Dim sqlString As String
    '------------------------------------------------------------------------------------------------------
    sqlString = F.Strings.Format(SQL_PATTERN, tableName)
    checkIfTableExists = VBA.CBool(pConnection.getPrimitiveValue(sqlString))
End Function

Public Function checkIfFunctionExists(functionName As String) As Boolean
    Const SQL_PATTERN As String = "SELECT " & _
                                        "CASE WHEN EXISTS (SELECT * FROM sys.objects EWhere object_id = OBJECT_ID(N'{0}') AND type IN ( N'FN', N'IF', N'TF', N'FS', N'FT' )) " & _
                                        "THEN 1 ELSE 0 END AS [FunctionExists]"
    '------------------------------------------------------------------------------------------------------
    Dim sqlString As String
    '------------------------------------------------------------------------------------------------------
    sqlString = F.Strings.Format(SQL_PATTERN, functionName)
    checkIfFunctionExists = VBA.CBool(pConnection.getPrimitiveValue(sqlString))
End Function




'[Implementation of IMeta]
Private Sub IMeta_setConnection(value As IDbConnection)
    Call setConnection(value)
End Sub

Private Function IMeta_checkIfTableExists(tableName As String) As Boolean
    IMeta_checkIfTableExists = checkIfTableExists(tableName)
End Function

Private Function IMeta_checkIfFunctionExists(functionName As String) As Boolean
    IMeta_checkIfFunctionExists = checkIfFunctionExists(functionName)
End Function

Private Sub IMeta_dropTableIfExists(tableName As String)
    Call dropTableIfExists(tableName)
End Sub

Private Function IMeta_checkIfDatabaseExists(dbName As String) As Boolean
    Stop
    IMeta_checkIfDatabaseExists = False
End Function

Private Sub IMeta_createDatabase(dbName As String)
    Stop
End Sub
