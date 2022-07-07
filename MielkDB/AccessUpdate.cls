VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IUpdate

Private Const CLASS_NAME As String = "AccessUpdate"
'----------------------------------------------------------------------------------------------------------
Private pConnection As IDbConnection
Private pWheres As VBA.Collection
Private pFrom As String
Private pValues As Scripting.Dictionary
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




'[DEFINITION]
Public Function EWhere(fieldName As String, comparisonMode As ComparisonModeEnum, comparedValue As Variant) As IUpdate
    Set EWhere = Me
    Call pWheres.Add(D.Create.Where(fieldName, comparisonMode, comparedValue))
End Function

Public Function WhereCondition(Where As EWhere) As IUpdate
    Set WhereCondition = Me
    If Not Where Is Nothing Then
        Call pWheres.Add(Where)
    End If
End Function

Public Function Values(properties As Scripting.Dictionary) As IUpdate
    Set Values = Me
    Set pValues = properties
End Function

Public Function From(tableName As String) As IUpdate
    Set From = Me
    pFrom = tableName
End Function



'[API]
Public Sub Execute()
    Call pConnection.Execute(getSqlString)
End Sub

Private Function getSqlString() As String
    Const SQL_PATTERN As String = "UPDATE {0} SET {1} {2}"
    '------------------------------------------------------------------------------------------------------
    getSqlString = F.Strings.Format(SQL_PATTERN, getFromPart, getValuePart, getWherePart)
End Function

Private Function getFromPart() As String
    getFromPart = pFrom
End Function

Private Function getValuePart() As String
    getValuePart = F.dictionaries.toString(pValues, DictPart_KeyAndValue, StringifyMode_Db, ", ", " = ", "")
End Function

Private Function getWherePart() As String
    getWherePart = D.Functions.convertWheresCollectionToWhereSqlPart(pWheres)
End Function





'[Implementation of IUpdate]
Private Sub IUpdate_setConnection(value As IDbConnection)
    Call setConnection(value)
End Sub

Private Function IUpdate_Values(properties As Scripting.IDictionary) As IUpdate
    Set IUpdate_Values = Values(properties)
End Function

Private Function IUpdate_Where(fieldName As String, comparisonMode As ComparisonModeEnum, comparedValue As Variant) As IUpdate
    Set IUpdate_Where = EWhere(fieldName, comparisonMode, comparedValue)
End Function

Private Function IUpdate_WhereCondition(Where As EWhere) As IUpdate
    Set IUpdate_WhereCondition = WhereCondition(Where)
End Function

Private Function IUpdate_From(tableName As String) As IUpdate
    Set IUpdate_From = From(tableName)
End Function

Private Sub IUpdate_Execute()
    Call Execute
End Sub
