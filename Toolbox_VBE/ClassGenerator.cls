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
Private Const PROPS_NAME_PROMPT As String = "VBE.Actions.AddProps.EnterPropsNames"
'[Line patterns and constants] ----------------------------------------------------------------------------
Private Const CLASS_NAME_LINE_PATTERN As String = "Private Const CLASS_NAME As String = ""{0}"""
Private Const REGEX_PATTERN_USER_INPUT As String = "^\s*(?:\[([^\[\]]*)\])?((?: *\w+ *(?:: *\w* *)?(?:,|$))+)"
Private Const REGEX_PATTERN_PROPS_NAMES As String = "(?:^|,)\s*(\w*)\s*(?::\s*(\w*))?"
'----------------------------------------------------------------------------------------------------------

Public Sub addClass()
    Dim activeProject As VBIDE.VBProject
    Dim newClass As VBIDE.VBComponent
    Dim className As String
    '------------------------------------------------------------------------------------------------------
    
    Set activeProject = Fn.getActiveProject
    className = VBA.InputBox(msg.getText(CLASS_NAME_PROMPT))
    
    Set newClass = activeProject.VBComponents.Add(vbext_ct_ClassModule)
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
        Call .Add(VBA_OPTION_EXPLICIT)
        Call .Add(vbNullString)
        Call .Add(F.Strings.Format(CLASS_NAME_LINE_PATTERN, className))
        Call .Add(moduleLevelSeparatorLine)
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



Public Sub addSettersAndGetters()
    Const PROPERTY_REGEX As String = "^(Private|Static) p(\w+) As (\w*)$"
    Const SETTER_PATTERN As String = "Public Function set{0}(value As {1}) As {2}\n\tSet set{0} = Me\n\tp{0} = value\nEnd Function\n"
    Const GETTER_PATTERN As String = "Public Function get{0}() As {1}\n\tget{0} = p{0}\nEnd Function\n"
    '------------------------------------------------------------------------------------------------------
    Dim selection As ECodeSelection
    Dim cm As VBIDE.CodeModule
    Dim properties As VBA.Collection
    Dim codeLines As VBA.Collection
    Dim property As Variant
    Dim code As String
    '------------------------------------------------------------------------------------------------------
    
    Set selection = Fn.Create.currentCodeSelection
    Set cm = selection.getCodeModule
    Set properties = F.regex.getMatchArraysCollection(F.Developer.getCodeModuleContent(cm), PROPERTY_REGEX)
    
    Set codeLines = New VBA.Collection
    
    '[Add setters]
    Call codeLines.Add(VBA.String(2, VBA.vbCrLf))
    Call codeLines.Add("'[SETTERS]")
    For Each property In properties
        code = F.Strings.Format(SETTER_PATTERN, VBA.CStr(property(2)), VBA.CStr(property(3)), cm.parent.name)
        Call codeLines.Add(code)
    Next property
    
    '[Add getters]
    Call codeLines.Add(VBA.String(1, VBA.vbCrLf))
    Call codeLines.Add("'[GETTERS]")
    For Each property In properties
        code = F.Strings.Format(GETTER_PATTERN, VBA.CStr(property(2)), VBA.CStr(property(3)))
        Call codeLines.Add(code)
    Next property
    
    Call cm.InsertLines(cm.CountOfLines + 1, F.Collections.toString(codeLines, , VBA.vbCrLf))
    
End Sub




Public Sub addProps()
    Dim project As VBIDE.VBProject
    Dim userInput As String
    '------------------------------------------------------------------------------------------------------
    Dim matches As Variant
    Dim category As String
    Dim names As String
    Dim propsMatches As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    Set project = Fn.getActiveProject
    userInput = VBA.InputBox(msg.getText(PROPS_NAME_PROMPT))
    matches = F.regex.getMatchArray(userInput, REGEX_PATTERN_USER_INPUT)
    
    If F.arrays.isDefined(matches) Then
        Call F.System.saveToClipboard(userInput)
        category = matches(LBound(matches))
        names = matches(LBound(matches) + 1)
        Set propsMatches = F.regex.getMatchArraysCollection(names, REGEX_PATTERN_PROPS_NAMES)
        
        Call PropsGenerator.addProps(project, propsMatches, category)
        
    End If
    
End Sub
