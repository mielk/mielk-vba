VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "IMeta"
'----------------------------------------------------------------------------------------------------------

'[SETTERS]
Public Sub setConnection(value As IDbConnection)
End Sub


'[DB OPERATIONS]
Public Sub dropTableIfExists(tableName As String)
End Sub

Public Sub createDatabase(dbName As String)
End Sub


'[DB STATE]
Public Function checkIfTableExists(tableName As String) As Boolean
End Function

Public Function checkIfFunctionExists(functionName As String) As Boolean
End Function

Public Function checkIfDatabaseExists(dbName As String) As Boolean
End Function
