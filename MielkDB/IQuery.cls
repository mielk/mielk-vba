VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "IQuery"
'----------------------------------------------------------------------------------------------------------


'[SETTERS]
Public Sub setConnection(value As IDbConnection)
End Sub



'[DEFINITION]
Public Function Fields(ParamArray fieldNames() As Variant) As IQuery
End Function

Public Function From(tableName As String) As IQuery
End Function

Public Function FromFunction(functionName As String, Optional params As Variant) As IQuery
End Function

Public Function Where(fieldName As String, comparisonMode As ComparisonModeEnum, comparedValue As Variant) As IQuery
End Function

Public Function WhereCondition(Where As EWhere) As IQuery
End Function

Public Function OrderBy(fieldName As String, Optional ascending As Boolean = True) As IQuery
End Function

Public Function sql(value As String) As IQuery
End Function



'[FETCHING DATA]
Public Function toRecordset() As Object ' ADODB.recordset
End Function

Public Function toDto() As Scripting.Dictionary
End Function

Public Function toDtosCollection() As VBA.Collection
End Function

Public Function toDtosDictionary(keyField As String) As Scripting.Dictionary
End Function

Public Function toArray(Optional includeHeaders As Boolean = True) As Variant
End Function

Public Function toPrimitive() As Variant
End Function

Public Function toScalarValue() As Variant
End Function
