VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "IInsert"
'----------------------------------------------------------------------------------------------------------

'[SETTERS]
Public Sub setConnection(value As IDbConnection)
End Sub



'[DEFINITION]
Public Function Into(tableName As String) As IInsert
End Function

Public Function SetIdentityInsert(value As Boolean) As IInsert
End Function

Public Function Values(properties As Scripting.Dictionary) As IInsert
End Function




'[FETCHING DATA]
Public Sub Execute()
End Sub

Public Sub ExecuteWithId(ByRef id As Long, ByVal idFieldName As String)
End Sub

Public Sub fromExcelToTempTable(filepath As String, sheetName As String, tableName As String)
End Sub
