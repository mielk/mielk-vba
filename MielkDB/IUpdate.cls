VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "IUpdate"
'----------------------------------------------------------------------------------------------------------

'[SETTERS]
Public Sub setConnection(value As IDbConnection)
End Sub



'[DEFINITION]
Public Function From(tableName As String) As IUpdate
End Function

Public Function Where(fieldName As String, comparisonMode As ComparisonModeEnum, comparedValue As Variant) As IUpdate
End Function

Public Function WhereCondition(Where As EWhere) As IUpdate
End Function

Public Function Values(properties As Scripting.Dictionary) As IUpdate
End Function




'[FETCHING DATA]
Public Sub Execute()
End Sub

