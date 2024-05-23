VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "ErrorManager"
'----------------------------------------------------------------------------------------------------------
Private Const DEFAULT_TABLE_ERROR_STACKS As String = "tbl_error_stacks"
Private Const DEFAULT_TABLE_ERROR_DETAILS As String = "tbl_errors"
Private Const FIELD_CALL_STACK_ID As String = "stack_id"
Private Const FIELD_ERROR_ID As String = "error_id"
Private Const FIELD_APP_NAME As String = "app_name"
Private Const FIELD_APP_VERSION As String = "app_version"
Private Const FIELD_APP_PATH As String = "app_path"
Private Const FIELD_USER_NAME As String = "user_name"
Private Const FIELD_CLASS_NAME As String = "class_name"
Private Const FIELD_METHOD_NAME As String = "method_name"
Private Const FIELD_ERROR_NUMBER As String = "error_number"
Private Const FIELD_ERROR_DESCRIPTION As String = "error_description"
Private Const FIELD_TIMESTAMP As String = "error_timestamp"
Private Const FIELD_PARAMS As String = "params"
Private Const FIELD_DECIMAL_SEPARATOR As String = "decimal_separator"
Private Const FIELD_DATE_FORMAT As String = "date_format"
Private Const FIELD_VBA_ERROR_BREAK_MODE As String = "vba_error_break_mode"
Private Const FIELD_PARENT_APP As String = "parent_app"
Private Const FIELD_PARENT_APP_VERSION As String = "parent_app_version"
Private Const FIELD_WINDOWS_BITS As String = "windows_bits"
'----------------------------------------------------------------------------------------------------------
Private Const INTERFACE_REGULAR_ERROR As String = "Application-defined or object-defined error"
'----------------------------------------------------------------------------------------------------------
Private pTableErrorStacks As String
Private pTableErrorDetails As String
'----------------------------------------------------------------------------------------------------------
Private pLogFolder As String
Private pDbPath As String
Private pDbPassword As String
Private pConnection As Object               'ADODB.Connection
Private pErrors As VBA.Collection
Private pFixedMessage As String
'----------------------------------------------------------------------------------------------------------
Private pLogToDbMode As Boolean
Private pLogToTextFileMode As Boolean
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    Set pErrors = New VBA.Collection
    pLogToDbMode = True
    pLogToTextFileMode = True
End Sub




'[SETTERS]

'<simple setters>

Public Function setTableNames(Optional tableErrorStacks As String, Optional tableErrorDetails As String) As ErrorManager
    Set setTableNames = Me
    If VBA.Len(tableErrorStacks) Then pTableErrorStacks = tableErrorStacks
    If VBA.Len(tableErrorDetails) Then pTableErrorDetails = tableErrorDetails
End Function

Public Function setLogFolderPath(value As String) As ErrorManager
    Set setLogFolderPath = Me
    pLogFolder = value
End Function

Public Function setConnection(value As Object) As ErrorManager
    Set setConnection = Me
    Set pConnection = value
End Function

Public Function setConnectionUrl(url As String, Optional password As String) As ErrorManager
    Set setConnectionUrl = Me
    pDbPath = url
    pDbPassword = password
    Set pConnection = createConnection
End Function

Public Function setLogToDbMode(value As Boolean) As ErrorManager
    Set setLogToDbMode = Me
    pLogToDbMode = value
End Function

Public Function setLogToTextFileMode(value As Boolean) As ErrorManager
    Set setLogToTextFileMode = Me
    pLogToTextFileMode = value
End Function

Public Function setFixedMessage(value As String) As ErrorManager
    Set setFixedMessage = Me
    pFixedMessage = value
End Function


'</simple setters>




'[GETTERS]

'<simple getters>
Public Function getLogFolderPath() As String
    getLogFolderPath = pLogFolder
End Function

Public Function getErrors() As VBA.Collection
    Set getErrors = pErrors
End Function

Public Function countErrors() As Long
    countErrors = pErrors.count
End Function

'</simple getters>




'[DB CONNECTION]

Private Function createConnection() As Object
    Set createConnection = VBA.CreateObject("ADODB.Connection")
    With createConnection
        .ConnectionString = getConnectionString(pDbPath, pDbPassword)
    End With
End Function

Private Function getConnectionString(url As String, Optional password As String) As String
    Const CONN_STRING_PATTERN As String = "Provider={0};User ID=Admin;Data Source={1};Mode={2}"
    Const CONN_WITH_PWRD_STRING_PATTERN As String = "Provider={0};User ID=Admin;Data Source={1};Mode={2};" & _
                                                    "Jet OLEDB:Database Password={3}"
    '----------------------------------------------------------------------------------------------------------
    If VBA.Len(password) Then
        getConnectionString = f.strings.format(CONN_WITH_PWRD_STRING_PATTERN, DEFAULT_ACCESS_PROVIDER, url, _
                                            "ReadWrite", password)
    Else
        getConnectionString = f.strings.format(CONN_STRING_PATTERN, DEFAULT_ACCESS_PROVIDER, url, "ReadWrite")
    End If
End Function

Private Function getConnection() As Object
    If pConnection Is Nothing Then
        On Error Resume Next
        Set pConnection = createConnection
    End If
    Set getConnection = pConnection
End Function




'[MESSAGES]
Public Function getAddErrorErrorMessage() As String
    getAddErrorErrorMessage = "Adding new error failed"
End Function

Public Function getLogToDatabaseErrorMessage() As String
    getLogToDatabaseErrorMessage = "Error when trying to add errors to the database"
End Function

Public Function getToImmediateErrorMessage() As String
    getToImmediateErrorMessage = "Error when trying to print errors stack to Immediate window"
End Function

Public Function getDisplayErrorMessage() As String
    getDisplayErrorMessage = "Error when trying to display errors stack"
End Function

Public Function getLogToTextfileErrorMessage() As String
    getLogToTextfileErrorMessage = "Error when trying to log errors to textfile"
End Function





'[ADDING/REMOVING ERRORS]

Public Sub clear()
    Set pErrors = New VBA.Collection
    pFixedMessage = vbNullString
End Sub

Public Sub addError(ByVal className As String, ByVal methodName As String, ByVal number As Long, ByVal description As String, _
                    ParamArray params() As Variant)
    Dim errLocal As VBA.errObject, errNumber As Long, errDescription As String, errSource As String
    '------------------------------------------------------------------------------------------------------
    Dim item As ErrorObject
    Dim paramsDict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    
    Set paramsDict = createDictionaryFromParamArray(f.arrays.fromParamArray(params))
    Set item = New ErrorObject
    Call item.setClassName(className). _
              setMethodName(methodName). _
              setErrorNumber(number). _
              setErrorDescription(description). _
              setParams(paramsDict)
    Call pErrors.Add(item)
    
    
'==========================================================================================================
ExitPoint:
    If errNumber > 0 Then GoTo ErrHandler
    Exit Sub

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim errMessage As String, paramsCounter As Long
    '------------------------------------------------------------------------------------------------------
    errNumber = VBA.err.number: errDescription = VBA.err.description
    If Not paramsDict Is Nothing Then paramsCounter = paramsDict.count
    errMessage = getAddErrorErrorMessage & VBA.vbCrLf & VBA.vbCrLf & _
                 "----------- [Stack] -------------------------------------" & VBA.vbCrLf & _
                 " * ErrorManager.addError" & VBA.vbCrLf & _
                        VBA.vbTab & "className: " & className & VBA.vbCrLf & _
                        VBA.vbTab & "methodName: " & methodName & VBA.vbCrLf & _
                        VBA.vbTab & "number: " & number & VBA.vbCrLf & _
                        VBA.vbTab & "description: " & description & VBA.vbCrLf & _
                        VBA.vbTab & "params counter: " & paramsCounter & VBA.vbCrLf & _
                        VBA.vbCrLf & VBA.vbCrLf & _
                    errDescription
    
    Call VBA.MsgBox(errMessage, vbCritical, app.getName)
    Set paramsDict = Nothing

End Sub

Private Function createDictionaryFromParamArray(arr As Variant) As Scripting.Dictionary
    Dim firstItem As Variant
    Dim strVariablesNames As String
    Dim arrVariablesNames() As String
    Dim namesFound As Long
    Dim counter As Long
    Dim i As Long
    Dim arrSize As Long: arrSize = -1
    '------------------------------------------------------------------------------------------------------
    
    Set createDictionaryFromParamArray = f.dictionaries.Create(False)
    
    On Error Resume Next
    arrSize = UBound(arr) - LBound(arr) + 1
    On Error GoTo 0
    
    If arrSize > 0 Then
        strVariablesNames = VBA.replace(VBA.replace(arr(LBound(arr)), ";", ","), " ", vbNullString)
        arrVariablesNames = VBA.Split(strVariablesNames, ",")
        On Error Resume Next
        namesFound = UBound(arrVariablesNames) - LBound(arrVariablesNames)
        If namesFound Then
            counter = LBound(arr) + 1
            For i = LBound(arrVariablesNames) To UBound(arrVariablesNames)
                If counter <= UBound(arr) Then
                    Call createDictionaryFromParamArray.Add(arrVariablesNames(i), arr(counter))
                    counter = counter + 1
                End If
            Next i
        End If
        
        On Error GoTo 0
        
    End If
    
End Function





'[ACTIONS]
Public Sub save(Optional connection As Object, Optional ByVal customErrorMessage As String, Optional errorsFilePath As String)
    If pLogToDbMode Then Call logToDatabase(connection)
    If pLogToTextFileMode Then Call logToTextFile(errorsFilePath)
    Call display(customErrorMessage)
    Call toImmediate
End Sub



Public Sub display(Optional ByVal message As String)
    Const METHOD_NAME As String = "display"
    '------------------------------------------------------------------------------------------------------
    Dim errNumber As Long, errDescription As String, errSource As String
    '------------------------------------------------------------------------------------------------------
    Dim errObj As ErrorObject
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If VBA.Len(pFixedMessage) > 0 Then
        message = pFixedMessage
    ElseIf VBA.Len(message) = 0 Then
        For i = pErrors.count To 1 Step -1
            Set errObj = pErrors.item(i)
            If errObj Is Nothing Then
            ElseIf errObj.getErrorDescription = INTERFACE_REGULAR_ERROR Then
            Else
                message = errObj.getErrorDescription
                Exit For
            End If
        Next i
    End If

    Call VBA.MsgBox(message, vbCritical, app.getNameVersion)
    
    '[Clear pFixedMessage to avoid case that this is displayed again later for completely different error]
    pFixedMessage = vbNullString

'==========================================================================================================
ExitPoint:
    Exit Sub

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim errMessage As String
    '------------------------------------------------------------------------------------------------------
    errNumber = VBA.err.number: errDescription = VBA.err.description: errSource = VBA.err.source
    errMessage = getDisplayErrorMessage & VBA.vbCrLf & VBA.vbCrLf & _
                "----------- [Stack] -------------------------------------" & VBA.vbCrLf & _
                " * ErrorManager.display" & VBA.vbCrLf & errDescription
    Call VBA.MsgBox(errMessage, vbCritical, app.getName)
    GoTo ExitPoint

End Sub



Public Sub logToTextFile(Optional errorsFilePath As String)
    Const METHOD_NAME As String = "logToTextFile"
    '------------------------------------------------------------------------------------------------------
    Dim fileName As String
    Dim filepath As String
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If VBA.Len(errorsFilePath) Then
        filepath = errorsFilePath
    ElseIf fso.folderExists(pLogFolder) Then
        fileName = VBA.format(VBA.Date, "YYYY-MM-DD") & ".txt"
        filepath = f.files.concatenatePath(pLogFolder, fileName)
    End If
    
    If VBA.Len(filepath) Then
        Call f.TextFiles.printToTextFile(getStack, filepath, False)
    Else
        Call VBA.MsgBox(getLogToTextfileErrorMessage & VBA.vbCrLf & "Empty [filepath]", _
                    vbCritical, app.getName)
    End If

'==========================================================================================================
ExitPoint:
    Exit Sub

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim errNumber As Long, errDescription As String, errSource As String, errMessage As String
    '------------------------------------------------------------------------------------------------------
    errNumber = VBA.err.number: errDescription = VBA.err.description: errSource = VBA.err.source
    errMessage = getLogToTextfileErrorMessage & VBA.vbCrLf & VBA.vbCrLf & _
                "----------- [Stack] -------------------------------------" & VBA.vbCrLf & _
                " * ErrorManager.logToTextFile" & VBA.vbCrLf & _
                        VBA.vbTab & "errorsFilePath: " & errorsFilePath & VBA.vbCrLf & _
                        VBA.vbTab & "pLogFolder: " & pLogFolder & VBA.vbCrLf & _
                errDescription
    Call VBA.MsgBox(errMessage, vbCritical, app.getName)
    GoTo ExitPoint

End Sub

Private Function getStack() As String
    Const METHOD_NAME As String = "getStack"
    '------------------------------------------------------------------------------------------------------
    Const STACK_INFO_PATTERN As String = "\n================================================================\n" & _
                                         "AppName: {0}\nAppVersion: {1}\nUser: {2}\nMacro path: {3}\n" & _
                                         "Decimal separator: {4}\nDate format: {5}\nWindows bits: {6}\n" & _
                                         "VBA error break mode: {7}\n\n*** STACK ***:\n"
    Dim error As ErrorObject
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    getStack = f.strings.format(STACK_INFO_PATTERN, app.getName, app.getVersion, f.System.getWindowsUserName, _
                                app.getUncPath, f.System.getDecimalSeparator, f.System.getDateFormat, _
                                f.System.getSystemBits, f.System.getErrorHandlingSetting)
    
    For Each error In pErrors
        getStack = getStack & error.getStringForTextFile & VBA.vbCrLf
    Next error
    getStack = getStack & "    ------------------------------------------------------------" & VBA.vbCrLf & VBA.vbCrLf
    
    
'==========================================================================================================
ExitPoint:
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim errNumber As Long, errDescription As String, errSource As String
    Dim errMessage As String, isSourceError As Boolean
    '------------------------------------------------------------------------------------------------------
    errNumber = VBA.err.number: errDescription = VBA.err.description: errSource = VBA.err.source
    isSourceError = Not VBA.left(errDescription, 3) = " * "
    errMessage = " * " & CLASS_NAME & "." & METHOD_NAME & VBA.IIf(isSourceError, _
                            VBA.vbCrLf & " * " & errSource & ": " & errDescription & " (" & errNumber & ")", _
                            VBA.vbCrLf & errDescription)
    Call VBA.err.raise(1, description:=errMessage)
    
End Function


Public Sub logToDatabase(Optional ByVal connection As Object)
    Const METHOD_NAME As String = "logToDatabase"
    '------------------------------------------------------------------------------------------------------
    Dim callStackId As Long:            callStackId = 0
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If connection Is Nothing Then Set connection = getConnection
    If Not connection Is Nothing Then Call connection.Open
    
    If connection Is Nothing Then
        Call VBA.err.raise(1, CLASS_NAME & "." & METHOD_NAME, "[Connection] is empty")
    ElseIf connection.state = 0 Then
        Call VBA.err.raise(1, CLASS_NAME & "." & METHOD_NAME, "[Connection] is closed")
    End If
    
    With connection
        Call .BeginTrans
        Call logErrorStack(connection, callStackId)
        Call logErrorDetails(connection, callStackId)
        Call .CommitTrans
    End With
    
'==========================================================================================================
ExitPoint:
    Stop
    'Call F.db.closeConnection, true)
    Exit Sub

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim errNumber As Long, errDescription As String, errMessage As String
    '------------------------------------------------------------------------------------------------------
    errNumber = VBA.err.number: errDescription = VBA.err.description
    errMessage = getLogToDatabaseErrorMessage & VBA.vbCrLf & VBA.vbCrLf & _
                "----------- [Stack] -------------------------------------" & VBA.vbCrLf & _
                " * ErrorManager.logToDatabase" & VBA.vbCrLf & errDescription
    Call VBA.MsgBox(errMessage, vbCritical, app.getName)
    GoTo ExitPoint
    
End Sub


Private Sub logErrorStack(connection As Object, ByRef callStackId As Long)
    Const METHOD_NAME As String = "logErrorStack"
    '------------------------------------------------------------------------------------------------------
    Const SQL_PATTERN As String = "SELECT * FROM {0}"
    '------------------------------------------------------------------------------------------------------
    Dim dictProps As Scripting.Dictionary
    Dim sql As String
    Dim rs As Object                'ADODB.Recordset
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    Set dictProps = createGeneralPropertiesDictionary
    sql = f.strings.format(SQL_PATTERN, pTableErrorStacks)
    
    Set rs = VBA.CreateObject(ADODB_RECORDSET)
    With rs
        Call .Open(sql, connection, adOpenKeyset, adLockPessimistic)
        Call .addNew(dictProps.Keys, dictProps.items)
        callStackId = .fields(FIELD_CALL_STACK_ID).value
    End With
    
    On Error Resume Next
    Stop
    'call F.db.closeRecordset(rs, False, true)
    
'==========================================================================================================
ExitPoint:
    Exit Sub

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim errNumber As Long, errDescription As String, errSource As String
    Dim errMessage As String, isSourceError As Boolean
    '------------------------------------------------------------------------------------------------------
    errNumber = VBA.err.number: errDescription = VBA.err.description: errSource = VBA.err.source
    isSourceError = Not VBA.left(errDescription, 3) = " * "
    errMessage = " * " & CLASS_NAME & "." & METHOD_NAME & VBA.IIf(isSourceError, _
                            VBA.vbCrLf & " * " & errSource & ": " & errDescription & " (" & errNumber & ")", _
                            VBA.vbCrLf & errDescription)
    Call VBA.err.raise(1, description:=errMessage)
    
End Sub


Private Function createGeneralPropertiesDictionary() As Scripting.Dictionary
    Const METHOD_NAME As String = "createGeneralPropertiesDictionary"
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    Set createGeneralPropertiesDictionary = f.dictionaries.Create(False)
    With createGeneralPropertiesDictionary
        Call .Add(FIELD_APP_NAME, app.getName)
        Call .Add(FIELD_APP_VERSION, app.getVersion)
        Call .Add(FIELD_APP_PATH, app.getUncPath)
        Call .Add(FIELD_USER_NAME, f.System.getWindowsUserName)
        Call .Add(FIELD_DECIMAL_SEPARATOR, f.System.getDecimalSeparator)
        Call .Add(FIELD_DATE_FORMAT, f.System.getDateFormat)
        Call .Add(FIELD_VBA_ERROR_BREAK_MODE, f.System.getErrorHandlingSetting)
        Call .Add(FIELD_PARENT_APP, Application.name)
        Call .Add(FIELD_PARENT_APP_VERSION, Application.Version)
        Call .Add(FIELD_WINDOWS_BITS, f.System.getSystemBits)
    End With
    
'==========================================================================================================
ExitPoint:
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim errNumber As Long, errDescription As String, errSource As String
    Dim errMessage As String, isSourceError As Boolean
    '------------------------------------------------------------------------------------------------------
    errNumber = VBA.err.number: errDescription = VBA.err.description: errSource = VBA.err.source
    isSourceError = Not VBA.left(errDescription, 3) = " * "
    errMessage = " * " & CLASS_NAME & "." & METHOD_NAME & VBA.IIf(isSourceError, _
                            VBA.vbCrLf & " * " & errSource & ": " & errDescription & " (" & errNumber & ")", _
                            VBA.vbCrLf & errDescription)
    Call VBA.err.raise(1, description:=errMessage)
    
End Function



Private Sub logErrorDetails(connection As Object, callStackId As Long)
    Const METHOD_NAME As String = "logErrorDetails"
    '------------------------------------------------------------------------------------------------------
    Dim err As errObject
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    For Each err In pErrors
        Call logSingleError(err, connection, callStackId)
    Next err
    
'==========================================================================================================
ExitPoint:
    Exit Sub

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim errNumber As Long, errDescription As String, errSource As String
    Dim errMessage As String, isSourceError As Boolean
    '------------------------------------------------------------------------------------------------------
    errNumber = VBA.err.number: errDescription = VBA.err.description: errSource = VBA.err.source
    isSourceError = Not VBA.left(errDescription, 3) = " * "
    errMessage = " * " & CLASS_NAME & "." & METHOD_NAME & VBA.IIf(isSourceError, _
                            VBA.vbCrLf & " * " & errSource & ": " & errDescription & " (" & errNumber & ")", _
                            VBA.vbCrLf & errDescription)
    Call VBA.err.raise(1, description:=errMessage)
    
End Sub



Private Sub logSingleError(err As ErrorObject, connection As Object, callStackId As Long)
    Const METHOD_NAME As String = "logSingleError"
    '------------------------------------------------------------------------------------------------------
    Dim sql As String
    Dim dict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    Set dict = createErrorPropertiesDictionary(err, callStackId)
    'sql = F.Db.createSqlInsertCommand(TABLE_ERROR_DETAILS, dict)
    Call connection.Execute(sql)
    
'==========================================================================================================
ExitPoint:
    Exit Sub

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim errNumber As Long, errDescription As String, errSource As String
    Dim errMessage As String, isSourceError As Boolean
    '------------------------------------------------------------------------------------------------------
    errNumber = VBA.err.number: errDescription = VBA.err.description: errSource = VBA.err.source
    isSourceError = Not VBA.left(errDescription, 3) = " * "
    errMessage = " * " & CLASS_NAME & "." & METHOD_NAME & VBA.IIf(isSourceError, _
                            VBA.vbCrLf & " * " & errSource & ": " & errDescription & " (" & errNumber & ")", _
                            VBA.vbCrLf & errDescription)
    Call VBA.err.raise(1, description:=errMessage)
    
End Sub


Private Function createErrorPropertiesDictionary(err As ErrorObject, callStackId As Long) As Scripting.Dictionary
    Const METHOD_NAME As String = "createErrorPropertiesDictionary"
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    Set createErrorPropertiesDictionary = f.dictionaries.Create(False)
    With createErrorPropertiesDictionary
        Call .Add(FIELD_CALL_STACK_ID, callStackId)
        Call .Add(FIELD_CLASS_NAME, err.getClassName)
        Call .Add(FIELD_METHOD_NAME, err.getMethodName)
        Call .Add(FIELD_ERROR_NUMBER, err.getErrorNumber)
        Call .Add(FIELD_ERROR_DESCRIPTION, err.getErrorDescription)
        Call .Add(FIELD_TIMESTAMP, err.getTimestamp)
        Call .Add(FIELD_PARAMS, err.getParamsAsString)
    End With
    
'==========================================================================================================
ExitPoint:
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim errNumber As Long, errDescription As String, errSource As String
    Dim errMessage As String, isSourceError As Boolean
    '------------------------------------------------------------------------------------------------------
    errNumber = VBA.err.number: errDescription = VBA.err.description: errSource = VBA.err.source
    isSourceError = Not VBA.left(errDescription, 3) = " * "
    errMessage = " * " & CLASS_NAME & "." & METHOD_NAME & VBA.IIf(isSourceError, _
                            VBA.vbCrLf & " * " & errSource & ": " & errDescription & " (" & errNumber & ")", _
                            VBA.vbCrLf & errDescription)
    Call VBA.err.raise(1, description:=errMessage)
    
End Function






Public Sub toImmediate()
    Const METHOD_NAME As String = "toImmediate"
    '------------------------------------------------------------------------------------------------------
    Dim errObj As ErrorObject
    Dim text As String
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    text = "========================================" & VBA.vbCrLf & _
           "[Errors] | " & VBA.Now & VBA.vbCrLf & getStack
    
    Debug.Print text
    
'==========================================================================================================
ExitPoint:
    Exit Sub

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim errNumber As Long, errDescription As String, errSource As String
    Dim errMessage As String
    '------------------------------------------------------------------------------------------------------
    errNumber = VBA.err.number: errDescription = VBA.err.description: errSource = VBA.err.source
    errMessage = VBA.vbCrLf & VBA.vbCrLf & getToImmediateErrorMessage & VBA.vbCrLf & VBA.vbCrLf & _
                "----------- [Stack] -------------------------------------" & VBA.vbCrLf & _
                " * ErrorManager.toImmediate" & VBA.vbCrLf & errDescription
    Debug.Print errMessage
    
End Sub



Public Function getErrorString(errNumber As Long, errDescription As String) As String
    Const ERR_DESCRIPTION_PATTERN As String = "({0}) {1}"
    '------------------------------------------------------------------------------------------------------
    getErrorString = f.strings.format(ERR_DESCRIPTION_PATTERN, errNumber, errDescription)
End Function
