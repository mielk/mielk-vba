Attribute VB_Name = "modCreators"
Option Explicit

Private Const CLASS_NAME As String = "modCreators"
    '----------------------------------------------------------------------------------------------------------


Public Function F() As Functions
    Static instance As Functions
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then
        Set instance = New Functions
    End If
    Set F = instance
End Function

Public Function fso() As Scripting.FileSystemObject
    Static instance As Scripting.FileSystemObject
    If instance Is Nothing Then
        Set instance = New Scripting.FileSystemObject
    End If
    Set fso = instance
End Function

Public Function ErrorManager() As ErrorManager
    Static instance As ErrorManager
    If instance Is Nothing Then
        Set instance = New ErrorManager
    End If
    Set ErrorManager = instance
End Function

Public Function app() As ParentApp
    Static instance As ParentApp
    If instance Is Nothing Then
        Set instance = New ParentApp
    End If
    Set app = instance
End Function

Public Function KeyValue(key As Variant, value As Variant) As Variant
    Dim arr(1 To 2) As Variant
    '------------------------------------------------------------------------------------------------------
    Call F.Variables.assign(arr(1), key)
    Call F.Variables.assign(arr(2), value)
    KeyValue = arr
End Function

Public Function MsgService() As MsgService
    Static instance As MsgService
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then
        Set instance = New MsgService
    End If
    Set MsgService = instance
End Function

Public Function ActionLogger(Optional inject As IActionLogger) As IActionLogger
    Static instance As IActionLogger
    '------------------------------------------------------------------------------------------------------
    If Not inject Is Nothing Then
        Set instance = inject
    ElseIf instance Is Nothing Then
        Set instance = New DefaultActionLogger
    End If
    Set ActionLogger = instance
End Function

Public Function Events() As EventsEnum
    Static instance As EventsEnum
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then
        Set instance = New EventsEnum
    End If
    Set Events = instance
End Function

Public Function Exceptions() As ExceptionsEnum
    Static instance As ExceptionsEnum
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then
        Set instance = New ExceptionsEnum
    End If
    Set Exceptions = instance
End Function

Public Function UIProps() As UIPropsEnum
    Static instance As UIPropsEnum
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then
        Set instance = New UIPropsEnum
    End If
    Set UIProps = instance
End Function
