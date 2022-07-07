VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "DbEngine"
'----------------------------------------------------------------------------------------------------------
Private pTimeout As Long
Private pMaxAttempts As Long
Private pConcurrencyErrors As Scripting.Dictionary
'[Services] -----------------------------------------------------------------------------------------------
Private pConstructors As FConstructors
Private pConnections As FConnections
Private pFunctions As FFunctions
'----------------------------------------------------------------------------------------------------------


'[Initializer]
Private Sub Class_Initialize()
    Set pConstructors = New FConstructors
    Set pConnections = New FConnections
    Set pFunctions = New FFunctions
    Set pConcurrencyErrors = getConcurrencyErrorsDictionary
End Sub

Private Function getConcurrencyErrorsDictionary() As Scripting.Dictionary
    Set getConcurrencyErrorsDictionary = F.dictionaries.Create(False)
    With getConcurrencyErrorsDictionary
        Call .Add(-2147467259, "Operation must use an updateable query.")
        Call .Add(-2147217887, "Cannot be updated. Record is locked.")
    End With
End Function

'[Submodules]
Public Function Create() As FConstructors
    Set Create = pConstructors
End Function

Public Function Connections() As FConnections
    Set Connections = pConnections
End Function

Public Function Functions() As FFunctions
    Set Functions = pFunctions
End Function

'Public Function Recordsets() As FRecordsets
'    Set Recordsets = pRecordsets
'End Function



'[TIMEOUT]

'<simple setter & getter>

Public Sub setTimeout(value As Long)
    pTimeout = value
End Sub

Friend Function getTimeout() As Long
    getTimeout = pTimeout
End Function

'</simple setter & getter>





'[ATTEMPTS]

'<simple setter & getter>

Public Sub setMaxAttempts(value As Long)
    pMaxAttempts = value
End Sub

Friend Function getMaxAttempts() As Long
    getMaxAttempts = pMaxAttempts
End Function

'</simple setter & getter>


Friend Function isConcurrencyError(errNumber As Long, errDescription As String) As Boolean
    Const CONCURRENCY_ERROR_REGEX As String = "(locked|in use|multiple|too many active)"
    '------------------------------------------------------------------------------------------------------
    If Not getConcurrencyErrorsDictionary.Exists(errNumber) Then
        isConcurrencyError = False
    Else
        isConcurrencyError = F.regex.checkIfMatch(errDescription, CONCURRENCY_ERROR_REGEX)
    End If
End Function
