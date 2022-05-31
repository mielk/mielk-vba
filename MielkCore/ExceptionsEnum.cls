VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExceptionsEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "ExceptionsEnum"
'----------------------------------------------------------------------------------------------------------
'[Values]
Private Const ERR_VALUES_ILLEGAL_DATA_TYPE As Long = 13
Private Const ERR_VALUES_ILLEGAL_VALUE As Long = 2001
Private Const ERR_VALUES_OBJECT_EXPECTED As Long = 424
Private Const ERR_VALUES_NULL_NOT_ALLOWED As Long = 2002
'[Array]
Private Const ERR_ARRAYS_NOT_ARRAY As Long = 4001
Private Const ERR_ARRAYS_DIMENSIONS_NOT_SUPPORTED As Long = 4002
Private Const ERR_ARRAYS_DIMENSION_OUT_OF_BOUND As Long = 4003
Private Const ERR_ARRAYS_INDEX_OUT_OF_BOUND As Long = 4004
Private Const ERR_ARRAYS_DYNAMIC_ARRAY_EXPECTED As Long = 4005
Private Const ERR_ARRAYS_UNDEFINED_ARRAY_NOT_ALLOWED As Long = 4006
'[Database]
Private Const ERR_DB_CONNECTION_EXCEPTION As Long = 5001
Private Const ERR_SQL_EXCEPTION As Long = 5002
Private Const ERR_DB_TRANSACTION_COMMIT_EXCEPTION As Long = 5003
Private Const ERR_DB_CLOSING_WHILE_IN_TRANSACTION As Long = 3246
'[Dictionaries]
Private Const ERR_DICT_DUPLICATED_KEY As Long = 6001
'[File system]
Private Const ERR_FILES_INVALID_DIRECTORY As Long = 7001
Private Const ERR_FILES_FILE_NOT_FOUND As Long = 7002
Private Const ERR_NO_WRITE_PERMISSION_FOR_FOLDER As Long = 7003
'[Reflection]
Private Const ERR_METHOD_NOT_SUPPORTED As Long = 438
'[Ranges]
Private Const ERR_INVALID_RANGE As Long = 8001
'[Sheets]
Private Const ERR_INVALID_SHEET As Long = 9001
Private Const ERR_ILLEGAL_SHEET As Long = 9002
Private Const ERR_INVALID_BOOK As Long = 9003
'----------------------------------------------------------------------------------------------------------
Private pExceptionsByNumber As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------

Private Sub Class_Initialize()
    Call initializeExceptions
End Sub

Private Sub initializeExceptions()
    '[Initialize containers]
    Set pExceptionsByNumber = F.Dictionaries.Create(False)
    
    '[Values exceptions]
    Call addException(ERR_VALUES_ILLEGAL_DATA_TYPE, "Exception.IllegalDataType")
    Call addException(ERR_VALUES_ILLEGAL_VALUE, "Exception.IllegalValue")
    Call addException(ERR_VALUES_OBJECT_EXPECTED, "Exception.ObjectExpected")
    Call addException(ERR_VALUES_NULL_NOT_ALLOWED, "Exception.NullNotAllowed")
    
    '[Arrays]
    Call addException(ERR_ARRAYS_NOT_ARRAY, "Exception.NotArray")
    Call addException(ERR_ARRAYS_DIMENSIONS_NOT_SUPPORTED, "Exception.DimensionNotSupported")
    Call addException(ERR_ARRAYS_DIMENSION_OUT_OF_BOUND, "Exception.DimensionOutOfBound")
    Call addException(ERR_ARRAYS_INDEX_OUT_OF_BOUND, "Exception.IndexOutOfBound")
    Call addException(ERR_ARRAYS_DYNAMIC_ARRAY_EXPECTED, "Exception.DynamicArrayExpected")
    Call addException(ERR_ARRAYS_UNDEFINED_ARRAY_NOT_ALLOWED, "Exception.UndefinedArray")
    
    '[DB]
    Call addException(ERR_DB_CONNECTION_EXCEPTION, "Exception.DbConnectionException")
    Call addException(ERR_SQL_EXCEPTION, "Exception.SqlException")
    Call addException(ERR_DB_TRANSACTION_COMMIT_EXCEPTION, "Exception.TransactionCommitException")
    Call addException(ERR_DB_CLOSING_WHILE_IN_TRANSACTION, "Exception.ClosingWhileInTransaction")
    
    '[Dictionaries]
    Call addException(ERR_DICT_DUPLICATED_KEY, "Exception.DuplicateKey")
    
    '[Files]
    Call addException(ERR_FILES_INVALID_DIRECTORY, "Exception.InvalidDirectory")
    Call addException(ERR_FILES_FILE_NOT_FOUND, "Exception.FileNotFound")
    Call addException(ERR_NO_WRITE_PERMISSION_FOR_FOLDER, "Exception.NoWritePermission")
    
    '[Reflection]
    Call addException(ERR_METHOD_NOT_SUPPORTED, "Exception.MethodNotSupported")
    
    '[Ranges]
    Call addException(ERR_INVALID_RANGE, "Exception.InvalidRange")
    
    '[Sheets]
    Call addException(ERR_INVALID_SHEET, "Exception.InvalidSheet")
    Call addException(ERR_ILLEGAL_SHEET, "Exception.IllegalSheet")
    Call addException(ERR_INVALID_BOOK, "Exception.InvalidBook")
    
End Sub

Private Sub addException(number As Long, name As String)
    Dim item As Exception
    '----------------------------------------------------------------------------------------------------------
    
    Set item = New Exception
    With item
        Call .setNumber(number)
        Call .setName(name)
        
        '[Add to repositories]
        Call pExceptionsByNumber.Add(number, item)
    End With
    
End Sub



'[Quick access]
Public Function IllegalDataType() As Exception: Set IllegalDataType = F.Dictionaries.getObject(pExceptionsByNumber, ERR_VALUES_ILLEGAL_DATA_TYPE): End Function
Public Function IllegalValue() As Exception: Set IllegalValue = F.Dictionaries.getObject(pExceptionsByNumber, ERR_VALUES_ILLEGAL_VALUE): End Function
Public Function ObjectExpected() As Exception: Set ObjectExpected = F.Dictionaries.getObject(pExceptionsByNumber, ERR_VALUES_OBJECT_EXPECTED): End Function
Public Function NullNotAllowed() As Exception: Set NullNotAllowed = F.Dictionaries.getObject(pExceptionsByNumber, ERR_VALUES_NULL_NOT_ALLOWED): End Function
Public Function NotArray() As Exception: Set NotArray = F.Dictionaries.getObject(pExceptionsByNumber, ERR_ARRAYS_NOT_ARRAY): End Function
Public Function InvalidDimension() As Exception: Set InvalidDimension = F.Dictionaries.getObject(pExceptionsByNumber, ERR_ARRAYS_DIMENSIONS_NOT_SUPPORTED): End Function
Public Function DimensionOutOfBound() As Exception: Set DimensionOutOfBound = F.Dictionaries.getObject(pExceptionsByNumber, ERR_ARRAYS_DIMENSION_OUT_OF_BOUND): End Function
Public Function IndexOutOfBound() As Exception: Set IndexOutOfBound = F.Dictionaries.getObject(pExceptionsByNumber, ERR_ARRAYS_INDEX_OUT_OF_BOUND): End Function
Public Function DynamicArrayExpected() As Exception: Set DynamicArrayExpected = F.Dictionaries.getObject(pExceptionsByNumber, ERR_ARRAYS_DYNAMIC_ARRAY_EXPECTED): End Function
Public Function UndefinedArrayNotAllowed() As Exception: Set UndefinedArrayNotAllowed = F.Dictionaries.getObject(pExceptionsByNumber, ERR_ARRAYS_UNDEFINED_ARRAY_NOT_ALLOWED): End Function
Public Function DbConnectionException() As Exception: Set DbConnectionException = F.Dictionaries.getObject(pExceptionsByNumber, ERR_DB_CONNECTION_EXCEPTION): End Function
Public Function SqlException() As Exception: Set SqlException = F.Dictionaries.getObject(pExceptionsByNumber, ERR_SQL_EXCEPTION): End Function
Public Function DbTransactionCommitException() As Exception: Set DbTransactionCommitException = F.Dictionaries.getObject(pExceptionsByNumber, ERR_DB_TRANSACTION_COMMIT_EXCEPTION): End Function
Public Function DbClosingWhileInTransaction() As Exception: Set DbClosingWhileInTransaction = F.Dictionaries.getObject(pExceptionsByNumber, ERR_DB_CLOSING_WHILE_IN_TRANSACTION): End Function
Public Function DuplicateKey() As Exception: Set DuplicateKey = F.Dictionaries.getObject(pExceptionsByNumber, ERR_DICT_DUPLICATED_KEY): End Function
Public Function InvalidDirectory() As Exception: Set InvalidDirectory = F.Dictionaries.getObject(pExceptionsByNumber, ERR_FILES_INVALID_DIRECTORY): End Function
Public Function FileNotFound() As Exception: Set FileNotFound = F.Dictionaries.getObject(pExceptionsByNumber, ERR_FILES_FILE_NOT_FOUND): End Function
Public Function NoWritePermission() As Exception: Set NoWritePermission = F.Dictionaries.getObject(pExceptionsByNumber, ERR_NO_WRITE_PERMISSION_FOR_FOLDER): End Function
Public Function MethodNotSupported() As Exception: Set MethodNotSupported = F.Dictionaries.getObject(pExceptionsByNumber, ERR_METHOD_NOT_SUPPORTED): End Function
Public Function InvalidRange() As Exception: Set InvalidRange = F.Dictionaries.getObject(pExceptionsByNumber, ERR_INVALID_RANGE): End Function
Public Function IllegalSheet() As Exception: Set IllegalSheet = F.Dictionaries.getObject(pExceptionsByNumber, ERR_INVALID_SHEET): End Function
Public Function InvalidSheet() As Exception: Set InvalidSheet = F.Dictionaries.getObject(pExceptionsByNumber, ERR_ILLEGAL_SHEET): End Function
Public Function InvalidBook() As Exception: Set InvalidBook = F.Dictionaries.getObject(pExceptionsByNumber, ERR_INVALID_BOOK): End Function
