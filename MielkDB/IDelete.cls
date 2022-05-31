VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "IDelete"
'----------------------------------------------------------------------------------------------------------

'[SETTERS]
Public Sub setConnection(value As IDbConnection)
End Sub




'[DEFINITION]
Public Function From(tableName As String) As IDelete
End Function

Public Function Where(fieldName As String, comparisonMode As ComparisonModeEnum, comparedValue As Variant) As IDelete
End Function

Public Function WhereCondition(Where As EWhere) As IDelete
End Function



'[FETCHING DATA]
Public Sub Execute()
End Sub
