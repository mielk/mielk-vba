VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefaultActionLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IActionLogger

Private Const CLASS_NAME As String = "DefaultActionLogger"
'----------------------------------------------------------------------------------------------------------


Public Function addLog(action As String, Optional params As Variant, Optional forceLog As Boolean = False) As ValidationObject
    Dim entry As String
    '------------------------------------------------------------------------------------------------------

    If Not forceLog And Not LOGGING_MODE Then
        'skip - log not required in such case
    Else
        entry = f.System.getWindowsUserName & ";" & VBA.format(VBA.Now, "YYYY-MM-DD HH:MM:SS") & ";" & action & ";"
        If f.Variables.isUndefined(params) Then
            '
        ElseIf f.Variables.isPrimitive(params) Then
            entry = entry & VBA.CStr(params)
        ElseIf f.dictionaries.isNonEmptyDictionary(params) Then
            entry = entry & f.dictionaries.toString(f.cast.toDictionary(params))
        End If
        Debug.Print entry
    End If
    
    Set addLog = f.Create.ValidationObject(True)

End Function



'[Implementation of IActionLogger interface]
Private Function IActionLogger_addLog(action As String, Optional params As Variant, Optional forceLog As Boolean = False) As ValidationObject
    Set IActionLogger_addLog = addLog(action, params, forceLog)
End Function
