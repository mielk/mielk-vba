VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STextfileActionLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IActionLogger

Private Const CLASS_NAME As String = "STextfileActionLogger"
'----------------------------------------------------------------------------------------------------------
Private Const LOG_FILE_NAME As String = "logs.txt"
'----------------------------------------------------------------------------------------------------------


Public Function addLog(action As String, Optional params As Variant, Optional forceLog As Boolean = False) As ValidationObject
    Dim parts As VBA.Collection
    Dim filepath As String
    Dim text As String
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If Not forceLog And Not IsLoggingOn Then
        'skip - log not required in such case
    Else
        
        filepath = F.files.concatenatePath( _
                                Paths.LogsFolderPath, _
                                F.Dates.toString, _
                                LOG_FILE_NAME)
        Call F.files.CreateFolder(F.files.getFileFolderPath(filepath))
        
        Set parts = getPartsCollection(action, params)
        text = F.Collections.toString(parts, StringifyMode_Normal, ";") & VBA.vbCrLf
        Call F.TextFiles.printToTextFile_Scripting(text, filepath)
        
        Set addLog = F.Create.ValidationObject(True)
        
    End If
    
'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Exit Function
    
'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    Set addLog = F.Create.ValidationObject(False, ErrorManager.getErrorString(errNumber, errDescription))
    
End Function

Private Function getPartsCollection(action As String, params As Variant) As VBA.Collection
    Set getPartsCollection = New VBA.Collection
    With getPartsCollection
        Call .Add(F.System.getUserUid)
        Call .Add(F.Dates.toString(, DateVisibility_DateTime))
        Call .Add(action)
        
        If F.Variables.isUndefined(params) Then
            ' skip
        ElseIf F.Variables.isPrimitive(params) Then
            Call .Add(removeSensitiveData(VBA.CStr(params)))
        ElseIf F.dictionaries.isNonEmptyDictionary(params) Then
            Call .Add(removeSensitiveData(F.dictionaries.toString(F.cast.toDictionary(params))))
        End If
        
    End With
    
    
    
    '------------------------------------------------------------------------------------------------------
End Function



Private Function removeSensitiveData(text As String) As String
    Const REGEX_PATTERN As String = "Password=\""(.*?)\"";"
    '------------------------------------------------------------------------------------------------------
    Dim subtext As String
    '------------------------------------------------------------------------------------------------------
    
    subtext = F.regex.getFirstGroupMatch(text, REGEX_PATTERN)
    If VBA.Len(subtext) Then
        removeSensitiveData = VBA.Replace(text, subtext, "*")
    Else
        removeSensitiveData = text
    End If
    
End Function





'[Implementation of IActionLogger]
Private Function IActionLogger_addLog(action As String, Optional params As Variant, Optional forceLog As Boolean = False) As ValidationObject
    Set IActionLogger_addLog = addLog(action, params, forceLog)
End Function
