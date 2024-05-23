VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "ClassGenerator"
'----------------------------------------------------------------------------------------------------------
Private Const CLASS_NAME_PROMPT As String = "VBE.Actions.AddClass.EnterClassName"
'[Line patterns and constants] ----------------------------------------------------------------------------
Private Const CLASS_NAME_LINE_PATTERN As String = "Private Const CLASS_NAME As String = ""{0}"""
'----------------------------------------------------------------------------------------------------------

Public Sub addClass()
    Dim activeProject As VBIDE.VBProject
    Dim newClass As VBIDE.VBComponent
    Dim className As String
    '------------------------------------------------------------------------------------------------------
    
    Set activeProject = Fn.getActiveProject
    className = VBA.InputBox(Msg.getText(CLASS_NAME_PROMPT))
    
    Set newClass = activeProject.VBComponents.add(vbext_ct_ClassModule)
    With newClass
        .name = className
        Call F.Dev.overrideCodeModuleContent(.CodeModule, generateClassContent(className))
    End With
    
End Sub


Private Function generateClassContent(className As String) As String
    Dim lines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    Set lines = New VBA.Collection
    
    With lines
        Call .add(VBA_OPTION_EXPLICIT)
        Call .add(vbNullString)
        Call .add(F.Strings.Format(CLASS_NAME_LINE_PATTERN, className))
        Call .add(moduleLevelSeparatorLine)
    End With
    
    generateClassContent = F.Collections.toString(lines, , VBA.vbCrLf)
    
End Function

Private Function moduleLevelSeparatorLine() As String
    moduleLevelSeparatorLine = VBA.String(MODULE_LEVEL_SEPARATOR_INDENT, " ") & "'" & _
                               VBA.String(MODULE_LEVEL_SEPARATOR_LENGTH, "-")
End Function

Private Function methodLevelSeparatorLine() As String
    methodLevelSeparatorLine = VBA.String(METHOD_LEVEL_SEPARATOR_INDENT, " ") & "'" & _
                               VBA.String(METHOD_LEVEL_SEPARATOR_LENGTH, "-")
End Function
