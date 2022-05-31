VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlServerDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDelete

Private Const CLASS_NAME As String = "SqlServerDelete"
'----------------------------------------------------------------------------------------------------------
Private pConnection As IDbConnection
'----------------------------------------------------------------------------------------------------------
Private pWheres As VBA.Collection
Private pFrom As String
'----------------------------------------------------------------------------------------------------------


'[INITIALIZER]
Private Sub Class_Initialize()
    Call initializeContainers
End Sub

Private Sub initializeContainers()
    Set pWheres = New VBA.Collection
End Sub



'[SETTERS]
Public Sub setConnection(value As IDbConnection)
    Set pConnection = value
End Sub




'[DEFINING QUERY]
Public Function From(tableName As String) As IDelete
    Set From = Me
    pFrom = tableName
End Function

Public Function Where(fieldName As String, comparisonMode As ComparisonModeEnum, comparedValue As Variant) As IDelete
    Set Where = Me
    Call pWheres.Add(D.Create.Where(fieldName, comparisonMode, comparedValue))
End Function

Public Function WhereCondition(Where As EWhere) As IDelete
    Set WhereCondition = Me
    If Not Where Is Nothing Then
        Call pWheres.Add(Where)
    End If
End Function




'[RETURNING DATA]
Public Sub Execute()
    Dim sql As String
    '------------------------------------------------------------------------------------------------------
    sql = getSqlString
    Call pConnection.Execute(sql)
End Sub

Private Function getSqlString() As String
    Const SQL_PATTERN As String = "DELETE FROM {0} {1}"
    '------------------------------------------------------------------------------------------------------
    getSqlString = F.Strings.Format(SQL_PATTERN, getFromPart, getWherePart)
End Function

Private Function getFromPart() As String
    getFromPart = pFrom
End Function

Private Function getWherePart() As String
    getWherePart = D.Functions.convertWheresCollectionToWhereSqlPart(pWheres)
End Function







'[Implementation of IDelete]
Private Sub IDelete_setConnection(value As IDbConnection)
    Call setConnection(value)
End Sub

Private Function IDelete_From(tableName As String) As IDelete
    Set IDelete_From = From(tableName)
End Function

Private Function IDelete_Where(fieldName As String, comparisonMode As ComparisonModeEnum, comparedValue As Variant) As IDelete
    Set IDelete_Where = Where(fieldName, comparisonMode, comparedValue)
End Function

Private Function IDelete_WhereCondition(Where As EWhere) As IDelete
    Set IDelete_WhereCondition = WhereCondition(Where)
End Function

Private Sub IDelete_execute()
    Call Execute
End Sub
