VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IQuery

Private Const CLASS_NAME As String = "AccessQuery"
'----------------------------------------------------------------------------------------------------------
Private pConnection As IDbConnection
'----------------------------------------------------------------------------------------------------------
Private pWheres As VBA.Collection
Private pFields As VBA.Collection
Private pFrom As String
Private pOrderFields As Scripting.Dictionary
Private pFunction As String
Private pFunctionParams As Variant
Private pSql As String
'----------------------------------------------------------------------------------------------------------



'[INITIALIZER]
Private Sub Class_Initialize()
    Call initializeContainers
End Sub

Private Sub initializeContainers()
    Set pWheres = New VBA.Collection
    Set pFields = New VBA.Collection
    Set pOrderFields = F.dictionaries.Create
End Sub




'[SETTERS]
Public Sub setConnection(value As IDbConnection)
    Set pConnection = value
End Sub




'[DEFINING QUERY]
Public Function Fields(ParamArray fieldNames() As Variant) As IQuery
    Dim arr As Variant:             arr = F.arrays.fromParamArray(fieldNames)
    Dim varItem As Variant
    Dim fieldName As String
    '------------------------------------------------------------------------------------------------------
    
    Set Fields = Me
    
    Set pFields = New VBA.Collection
    For Each varItem In arr
        If VBA.IsObject(varItem) Then
            fieldName = VBA.CallByName(F.Cast.toObject(varItem), GET_DB_FIELD_FUNCTION_NAME, VbMethod)
        Else
            fieldName = VBA.CStr(varItem)
        End If
        Call pFields.Add(fieldName)
    Next varItem

End Function

Public Function From(tableName As String) As IQuery
    Set From = Me
    pFrom = tableName
End Function

Public Function FromFunction(functionName As String, Optional params As Variant) As IQuery
    Set FromFunction = Me
    pFunction = functionName
    Call F.Variables.assign(pFunctionParams, params)
End Function

Public Function EWhere(fieldName As String, comparisonMode As ComparisonModeEnum, comparedValue As Variant) As IQuery
    Set EWhere = Me
    Call pWheres.Add(D.Create.Where(fieldName, comparisonMode, comparedValue))
End Function

Public Function WhereCondition(Where As EWhere) As IQuery
    Set WhereCondition = Me
    If Not Where Is Nothing Then
        Call pWheres.Add(Where)
    End If
End Function

Public Function OrderBy(fieldName As String, Optional ascending As Boolean = True) As IQuery
    Set OrderBy = Me
    If VBA.Len(fieldName) Then
        Call pOrderFields.Add(fieldName, ascending)
    End If
End Function

Public Function sql(value As String) As IQuery
    Set sql = Me
    pSql = value
End Function




'[RETURNING DATA]
Public Function toRecordset() As Object 'ADODB.recordset
    Dim sql As String
    '------------------------------------------------------------------------------------------------------
    If VBA.Len(pSql) Then sql = pSql Else sql = getSqlString
    Set toRecordset = pConnection.getRecordset(sql)
End Function

Public Function toDto() As Scripting.Dictionary
    Dim rs As Object 'ADODB.recordset
    '------------------------------------------------------------------------------------------------------
    Set rs = toRecordset
    Set toDto = D.Functions.convertRecordsetToDto(rs)
    Call cleanUp(rs)
End Function

Public Function toDtosCollection() As VBA.Collection
    Dim rs As Object 'ADODB.recordset
    '------------------------------------------------------------------------------------------------------
    Set rs = toRecordset
    Set toDtosCollection = D.Functions.convertRecordsetToDtosCollection(rs)
    Call cleanUp(rs)
End Function

Public Function toDtosDictionary(keyField As String) As Scripting.Dictionary
    Dim rs As Object 'ADODB.recordset
    '------------------------------------------------------------------------------------------------------
    Set rs = toRecordset
    Set toDtosDictionary = D.Functions.convertRecordsetToDtosDictionary(rs, keyField)
    Call cleanUp(rs)
End Function

Public Function toArray(Optional includeHeaders As Boolean = True) As Variant
    Dim rs As Object 'ADODB.recordset
    '------------------------------------------------------------------------------------------------------
    Set rs = toRecordset
    toArray = D.Functions.convertRecordsetToArray(rs, includeHeaders)
    Call cleanUp(rs)
End Function

Public Function toPrimitive() As Variant
    toPrimitive = pConnection.getPrimitiveValue(getSqlString)
    Call pConnection.closeConnection
End Function

Public Function toScalarValue() As Variant
    'toScalarValue = pConnection.getScalarFunctionValue(getScalarFunctionSqlString)
    'Not applies to MS Access.
End Function

Private Function getScalarFunctionSqlString() As String
    Const SQL_PATTERN As String = "SELECT {0}({1})"
    Dim params As String
    '------------------------------------------------------------------------------------------------------
    params = F.dictionaries.toString(F.Cast.toDictionary(pFunctionParams), DictPart_ValueOnly)
    getScalarFunctionSqlString = F.Strings.Format(SQL_PATTERN, pFunction, params)
End Function

Private Sub cleanUp(rs As Object)
    If Not pConnection Is Nothing Then
        Call D.Connections.closeRecordset(rs, Not pConnection.isTransactionOpen, True)
    End If
End Sub



Private Function getSqlString() As String
    Const SQL_PATTERN As String = "SELECT {0} FROM {1} {2} {3}"
    '------------------------------------------------------------------------------------------------------
    getSqlString = F.Strings.Format(SQL_PATTERN, getSelectPart, getFromPart, getWherePart, getOrderPart)
End Function

Private Function getSelectPart() As String
    If F.Collections.isNonEmptyCollection(pFields) Then
        getSelectPart = F.Collections.toString(pFields)
    Else
        getSelectPart = "*"
    End If
End Function

Private Function getFromPart() As String
    If VBA.Len(pFunction) > 0 Then
        'Not available in MS Access.
        'getFromPart = pFunction & "(" & getFunctionParamsAsString & ")"
    Else
        getFromPart = pFrom
    End If
End Function

Private Function getWherePart() As String
    getWherePart = D.Functions.convertWheresCollectionToWhereSqlPart(pWheres)
End Function

Private Function getOrderPart() As String
    Dim arr() As String
    Dim varKey As Variant
    Dim fieldName As String
    Dim orderDirection As String
    '------------------------------------------------------------------------------------------------------
    
    If F.dictionaries.isNonEmptyDictionary(pOrderFields) Then
        For Each varKey In pOrderFields.Keys
            fieldName = D.Functions.getSafeFieldName(VBA.CStr(varKey))
            orderDirection = VBA.IIf(pOrderFields.Item(varKey), "ASC ", "DESC ")
            Call F.arrays.addEntry(arr, fieldName & " " & orderDirection)
        Next varKey
    End If
    
    If F.arrays.isDefined(arr) Then
        getOrderPart = " ORDER BY " & VBA.Join(arr, ", ")
    End If
    
End Function





'[Implementation of IUpdate]
Private Sub IQuery_setConnection(value As IDbConnection)
    Call setConnection(value)
End Sub

Private Function IQuery_Fields(ParamArray fieldNames() As Variant) As IQuery
    Set IQuery_Fields = Fields(fieldNames)
End Function

Private Function IQuery_From(tableName As String) As IQuery
    Set IQuery_From = From(tableName)
End Function

Private Function IQuery_FromFunction(functionName As String, Optional params As Variant) As IQuery
    Set IQuery_FromFunction = FromFunction(functionName, params)
End Function

Private Function IQuery_FromStoredProcedure(procName As String, Optional params As Scripting.IDictionary) As IQuery
    'N/A
End Function

Private Function IQuery_sql(value As String) As IQuery
    Set IQuery_sql = sql(value)
End Function

Private Function IQuery_Where(fieldName As String, comparisonMode As ComparisonModeEnum, comparedValue As Variant) As IQuery
    Set IQuery_Where = EWhere(fieldName, comparisonMode, comparedValue)
End Function

Private Function IQuery_WhereCondition(Where As EWhere) As IQuery
    Set IQuery_WhereCondition = WhereCondition(Where)
End Function

Private Function IQuery_OrderBy(fieldName As String, Optional ascending As Boolean = True) As IQuery
    Set IQuery_OrderBy = OrderBy(fieldName, ascending)
End Function

Private Function IQuery_toRecordset() As Object ' Object 'ADODB.recordset
    Set IQuery_toRecordset = toRecordset
End Function

Private Function IQuery_toDto() As Scripting.IDictionary
    Set IQuery_toDto = toDto
End Function

Private Function IQuery_toDtosCollection() As VBA.Collection
    Set IQuery_toDtosCollection = toDtosCollection
End Function

Private Function IQuery_toDtosDictionary(keyField As String) As Scripting.Dictionary
    Set IQuery_toDtosDictionary = toDtosDictionary(keyField)
End Function

Private Function IQuery_toArray(Optional includeHeaders As Boolean = True) As Variant
    IQuery_toArray = toArray(includeHeaders)
End Function

Private Function IQuery_toPrimitive() As Variant
    Call F.Variables.assign(IQuery_toPrimitive, toPrimitive)
End Function

Private Function IQuery_toScalarValue() As Variant
    Call F.Variables.assign(IQuery_toScalarValue, toScalarValue)
End Function
