VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FFunctions"

'[Connections]
Public Function isOpenConnection(value As Variant) As Boolean
    If VBA.IsObject(value) Then
        If F.Strings.compareStrings(VBA.TypeName(value), ADODB_CONNECTION) Then
            isOpenConnection = (CConnection(value).State = 1)
        End If
    End If
End Function

Public Sub reopenConnection(connection As Object) 'ADODB.connection)
    On Error Resume Next
    If Not connection Is Nothing Then
        Call connection.Close
        Call sleep(100)
        DoEvents
    End If
    Call connection.Open
End Sub


Public Function CConnection(value As Variant, _
                            Optional createNewConnectionForUndefined As Boolean = False) As Object 'ADODB.connection
    Const METHOD_NAME As String = "CConnection"
    Const ERR_DESC As String = "ADODB.Connection was expected"
    '------------------------------------------------------------------------------------------------------
    
    If F.Variables.isUndefined(value, False, False) Then
        If createNewConnectionForUndefined Then
            Set CConnection = VBA.CreateObject(ADODB_CONNECTION)
        Else
            Set CConnection = Nothing
        End If
    ElseIf F.Strings.compareStrings(VBA.TypeName(value), ADODB_CONNECTION) Then
        Set CConnection = value
    Else
        Call Exceptions.IllegalDataType.Raise
    End If
    
End Function





'[Data types]
Public Function getAdDataType(value As Variant) As Long ' ADODB.DataTypeEnum
    Select Case VBA.VarType(value)
        Case vbByte, vbInteger, vbLong
            getAdDataType = adInteger
        Case vbSingle, vbDouble, vbDecimal
            getAdDataType = adDecimal
        Case vbString
            getAdDataType = adVarWChar
    End Select
End Function




'[Fetching data]
Public Function convertWheresCollectionToWhereSqlPart(wheres As VBA.Collection) As String
    Dim whereObj As EWhere
    Dim arr() As String
    '----------------------------------------------------------------------------------------------------------
    
    For Each whereObj In wheres
        Call F.arrays.addEntry(arr, whereObj.toString)
    Next whereObj
    
    If F.arrays.isDefined(arr) Then
        convertWheresCollectionToWhereSqlPart = " WHERE " & VBA.Join(arr, " AND ")
    End If
    
End Function





'[Data manipulation]
Public Function convertRecordsetToArray(rs As Object, Optional includeHeaders As Boolean = False) As Variant
    Const REDIM_STEP As Long = 1000
    '------------------------------------------------------------------------------------------------------
    Dim results() As Variant
    Dim row As Long
    '------------------------------------------------------------------------------------------------------
    
    'On Error Resume Next
    
    ReDim results(1 To rs.Fields.Count, 1 To REDIM_STEP)

    If includeHeaders Then
        Call insertFieldHeaders(rs, results)
        row = LBound(results, 2)
    End If

    '[Feed array with data]
    Do Until rs.EOF
        row = row + 1
        If row > UBound(results, 2) Then Call expandArray(results, REDIM_STEP)
        Call feedArrayRowWithData(rs, results, row)
        Call rs.MoveNext
    Loop

    Call trimArray(results, row)
    convertRecordsetToArray = results
    
End Function

'Private Function waitForRecordset

Private Sub expandArray(ByRef arr As Variant, redimStep As Long)
    ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2) + redimStep)
End Sub

Private Sub insertFieldHeaders(rs As Object, ByRef arr As Variant)
    Dim fld As Object 'ADODB.field
    Dim column As Long
    '------------------------------------------------------------------------------------------------------
    For Each fld In rs.Fields
        column = column + 1
        arr(column, LBound(arr, 2)) = fld.Name
    Next fld
End Sub

Private Sub feedArrayRowWithData(rs As Object, ByRef arr As Variant, row As Long)
    Dim column As Long
    '------------------------------------------------------------------------------------------------------
    For column = LBound(arr, 1) To UBound(arr, 1)
        arr(column, row) = rs.Fields(column - LBound(arr, 1)).value
    Next column
End Sub

Private Sub trimArray(ByRef arr As Variant, rowsUsed As Long)
    If rowsUsed > 0 Then
        ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To rowsUsed)
    Else
        Erase arr
    End If
End Sub



Public Function convertRecordsetToDtosDictionary(rs As Object, keyField As String) As Scripting.Dictionary
    Dim keyValue As Variant
    Dim dto As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Set convertRecordsetToDtosDictionary = F.dictionaries.Create(False)
    If rs.State = 1 Then
        Do Until rs.EOF
            Set dto = getCurrentRecordPropertiesMap(rs)
            keyValue = F.dictionaries.getItem(dto, keyField)
            Call convertRecordsetToDtosDictionary.Add(keyValue, dto)
            Call rs.MoveNext
        Loop
    End If
End Function

Public Function convertRecordsetToDtosCollection(rs As Object) As VBA.Collection
    Set convertRecordsetToDtosCollection = New VBA.Collection
    If rs.State = 1 Then
        Do Until rs.EOF
            Call convertRecordsetToDtosCollection.Add(getCurrentRecordPropertiesMap(rs))
            Call rs.MoveNext
        Loop
    End If
End Function

Public Function convertRecordsetToDto(rs As Object) As Scripting.Dictionary
    Set convertRecordsetToDto = F.dictionaries.Create(False)
    If rs.State = 1 Then
        On Error Resume Next
        Call rs.MoveFirst
        If Not DEV_MODE Then On Error GoTo ErrHandler
        
        Do Until rs.EOF
            Set convertRecordsetToDto = getCurrentRecordPropertiesMap(rs)
            Exit Do
        Loop
    End If
    
    
ErrHandler:

End Function

Private Function getCurrentRecordPropertiesMap(rs As Object) As Scripting.Dictionary
    Dim field As Variant
    Dim fieldValue As Variant
    '------------------------------------------------------------------------------------------------------
    Set getCurrentRecordPropertiesMap = F.dictionaries.Create(False)
    For Each field In rs.Fields
        fieldValue = rs.Fields(field.Name).value
        Call getCurrentRecordPropertiesMap.Add(field.Name, fieldValue)
    Next field
End Function



Public Function getSafeFieldName(value As String) As String
    getSafeFieldName = VBA.Replace(VBA.Replace("[" & value & "]", "[[", "["), "]]", "]")
End Function


