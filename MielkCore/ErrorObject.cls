VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "Error"
'----------------------------------------------------------------------------------------------------------
Private pClassName As String
Private pMethodName As String
Private pParams As Scripting.Dictionary
Private pErrorNumber As Long
Private pErrorDescription As String
Private pTimestamp As Date
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    pTimestamp = VBA.Now
End Sub



'[SETTERS]
Public Function setClassName(value As String) As ErrorObject
    Set setClassName = Me
    pClassName = value
End Function

Public Function setMethodName(value As String) As ErrorObject
    Set setMethodName = Me
    pMethodName = value
End Function

Public Function setErrorNumber(value As Long) As ErrorObject
    Set setErrorNumber = Me
    pErrorNumber = value
End Function

Public Function setErrorDescription(value As String) As ErrorObject
    Set setErrorDescription = Me
    pErrorDescription = value
End Function

Public Function setParams(value As Scripting.Dictionary) As ErrorObject
    Set setParams = Me
    Set pParams = value
End Function



'[GETTERS]
Public Function getClassName() As String
    getClassName = pClassName
End Function

Public Function getMethodName() As String
    getMethodName = pMethodName
End Function

Public Function getErrorNumber() As Long
    getErrorNumber = pErrorNumber
End Function

Public Function getErrorDescription() As String
    getErrorDescription = pErrorDescription
End Function

Public Function getParams() As Scripting.Dictionary
    Set getParams = pParams
End Function

Public Function getTimestamp() As Date
    getTimestamp = pTimestamp
End Function




Public Function getParamsAsString() As String
    Const METHOD_NAME As String = "getParamsAsString"
    '------------------------------------------------------------------------------------------------------
    Dim key As Variant
    Dim arr() As String
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If Not pParams Is Nothing Then
        If pParams.count Then
            ReDim arr(1 To pParams.count)
            For Each key In pParams.Keys
                i = i + 1
                arr(i) = "{" & key & ":" & F.strings.stringify(pParams.item(key), StringifyMode_Normal, True) & "}"
            Next key
            getParamsAsString = VBA.join(arr, ", ")
        End If
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim errNumber As Long, errDescription As String, errSource As String
    Dim isSourceError As Boolean, errMessage As String
    '------------------------------------------------------------------------------------------------------
    errNumber = VBA.err.number: errDescription = VBA.err.description: errSource = VBA.err.source
    isSourceError = Not VBA.left(errDescription, 3) = " * "
    errMessage = " * " & CLASS_NAME & "." & METHOD_NAME & VBA.IIf(isSourceError, _
                            VBA.vbCrLf & " * " & errSource & ": " & errDescription & " (" & errNumber & ")", _
                            VBA.vbCrLf & errDescription)
    Call VBA.err.raise(1, description:=errMessage)
    
End Function



Public Function getStringForTextFile() As String
    Const METHOD_NAME As String = "getStringForTextFile"
    '------------------------------------------------------------------------------------------------------
    Const STRING_PATTERN As String = "* Class.method                  {0}" & vbCrLf & _
                                     "  Error number                  {1}" & vbCrLf & _
                                     "  Error description             {2}" & vbCrLf & _
                                     "  Timestamp                     {3}" & vbCrLf & _
                                     "  Params                        {4}" & vbCrLf
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    getStringForTextFile = F.strings.format(STRING_PATTERN, _
                            pClassName & "." & pMethodName, _
                            pErrorNumber, _
                            pErrorDescription, _
                            pTimestamp, _
                            getParamsStringForTextFile)

'==========================================================================================================
ExitPoint:
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim errNumber As Long, errDescription As String, errSource As String
    Dim isSourceError As Boolean, errMessage As String
    '------------------------------------------------------------------------------------------------------
    errNumber = VBA.err.number: errDescription = VBA.err.description: errSource = VBA.err.source
    isSourceError = Not VBA.left(errDescription, 3) = " * "
    errMessage = " * " & CLASS_NAME & "." & METHOD_NAME & VBA.IIf(isSourceError, _
                            VBA.vbCrLf & " * " & errSource & ": " & errDescription & " (" & errNumber & ")", _
                            VBA.vbCrLf & errDescription)
    Call VBA.err.raise(1, description:=errMessage)

End Function



Private Function getParamsStringForTextFile() As String
    Const METHOD_NAME As String = "getParamsStringForTextFile"
    '------------------------------------------------------------------------------------------------------
    Dim key As Variant
    Dim arr() As String
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If F.Dictionaries.isNonEmptyDictionary(pParams) Then
        ReDim arr(1 To pParams.count)
        For Each key In pParams.Keys
            i = i + 1
            arr(i) = "[" & key & "]: " & F.strings.stringify(pParams.item(key), StringifyMode_Normal, True)
        Next key
        getParamsStringForTextFile = VBA.join(arr, VBA.vbCrLf & VBA.String(32, " "))
    Else
        getParamsStringForTextFile = "(no params)"
    End If

'==========================================================================================================
ExitPoint:
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim errNumber As Long, errDescription As String, errSource As String
    Dim isSourceError As Boolean, errMessage As String
    '------------------------------------------------------------------------------------------------------
    errNumber = VBA.err.number: errDescription = VBA.err.description: errSource = VBA.err.source
    isSourceError = Not VBA.left(errDescription, 3) = " * "
    errMessage = " * " & CLASS_NAME & "." & METHOD_NAME & VBA.IIf(isSourceError, _
                            VBA.vbCrLf & " * " & errSource & ": " & errDescription & " (" & errNumber & ")", _
                            VBA.vbCrLf & errDescription)
    Call VBA.err.raise(1, description:=errMessage)

End Function



Public Function equals(errObject As ErrorObject, _
                            Optional compareClassName As Boolean = True, _
                            Optional compareMethodName As Boolean = True, _
                            Optional compareNumber As Boolean = True, _
                            Optional compareDescription As Boolean = True, _
                            Optional compareParams As Boolean = True) As Boolean
    If compareClassName And pClassName <> errObject.getClassName Then
        equals = False
    ElseIf compareMethodName And pMethodName <> errObject.getMethodName Then
        equals = False
    ElseIf compareNumber And pErrorNumber <> errObject.getErrorNumber Then
        equals = False
    ElseIf compareDescription And pErrorDescription <> errObject.getErrorDescription Then
        equals = False
    ElseIf compareParams And Not F.Dictionaries.compare(pParams, errObject.getParams) Then
        equals = False
    Else
        equals = True
    End If
End Function
