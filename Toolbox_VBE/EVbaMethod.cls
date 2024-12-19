VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EVbaMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "EVbaMethod"
'----------------------------------------------------------------------------------------------------------
Private pModule As EVbaModule
'[Metadata]
Private pName As String
Private pProcType As VBIDE.vbext_ProcKind
Private pFirstLine As Long
Private pTotalLines As Long
'----------------------------------------------------------------------------------------------------------
Private pMethodType As MethodTypeEnum
Private pScope As ScopeTypeEnum
Private pReturnedType As String
Private pInputParams As Scripting.Dictionary
Private pVariables As Scripting.Dictionary
'[Declarations] -------------------------------------------------------------------------------------------
Private pDeclarationLineIndex As Long
Private pDeclarationLinesCounter As Long
Private pDeclarationLine As String
'----------------------------------------------------------------------------------------------------------
Private pDescription As String
Private pBody As String
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    pProcType = -1
End Sub


'[SETTERS]
Public Function setModule(value As EVbaModule) As EVbaMethod
    Set setModule = Me
    Set pModule = value
End Function

Public Function readByLineNumber(line As Long) As EVbaMethod
    Set readByLineNumber = Me
    Call loadMetaData(line)
    Call loadDescription
    Call loadBody
End Function

Private Sub loadMetaData(line As Long)
    Dim cm As VBIDE.CodeModule:                 Set cm = getCodeModule
    '------------------------------------------------------------------------------------------------------
    
    '[Set metadata]
    pProcType = F.Developer.getProcType(cm, line)
    With cm
        pName = .ProcOfLine(line, pProcType)
        pFirstLine = .ProcStartLine(pName, pProcType)
        pTotalLines = .ProcCountLines(pName, pProcType)
        pDeclarationLineIndex = .ProcBodyLine(pName, pProcType)
    End With
    
End Sub

Private Sub loadDescription()
    If pDeclarationLineIndex > pFirstLine Then
        pDescription = getCodeModule.lines(pFirstLine, pDeclarationLineIndex - pFirstLine)
    End If
End Sub

Private Sub loadBody()
    pBody = getCodeModule.lines(pDeclarationLineIndex, pTotalLines - pDeclarationLineIndex + pFirstLine)
End Sub



'[GETTERS]
Public Function getComponent() As VBIDE.VBComponent
    Set getComponent = pModule.getComponent
End Function

Public Function getCodeModule() As VBIDE.CodeModule
    Set getCodeModule = pModule.getCodeModule
End Function

Public Function getName() As String
    getName = pName
End Function

Public Function getKey() As String
    getKey = pName & "|" & pProcType
End Function

Public Function getProcType() As VBIDE.vbext_ProcKind
    getProcType = pProcType
End Function

Public Function getLastLineIndex() As Long
    getLastLineIndex = pFirstLine + pTotalLines
End Function

Public Function getJsonDict() As Scripting.Dictionary
    Set getJsonDict = F.dictionaries.Create(False)
    With getJsonDict
        Call .Add(META_TAG_NAME, pName)
        Call .Add(META_TAG_PROC_TYPE, pProcType)
        Call .Add(META_TAG_DESCRIPTION, pDescription)
        Call .Add(META_TAG_BODY, pBody)
    End With
End Function

Public Function getBody() As String
    getBody = pBody
End Function



'[ANALYZING]
Public Sub analyzeBody()
    With VbaCodeParser
        pBody = cutCommentsAfterEnd(pBody)
        pDeclarationLinesCounter = .countDeclarationLines(pBody)
        pDeclarationLine = .getMethodDeclarationLine(pBody)
        pScope = .findMethodScope(pBody)
        pMethodType = .findMethodType(pBody)
        pReturnedType = .findMethodReturnedType(pBody)
        Set pInputParams = getVariablesFromDtos(.getMethodInputParamsDtos(pBody))
        Set pVariables = getVariablesFromDtos(.getVariablesDtos(pBody))
    End With
End Sub

Private Function cutCommentsAfterEnd(body As String) As String
    Const REGEX_PATTERN As String = "^([\s\u000A\u000D]*End (?:Sub|Function))(?:.|\u000A|\u000D)*"
    Const REPLACEMENT_PATTERN As String = "$1"
    '------------------------------------------------------------------------------------------------------
    cutCommentsAfterEnd = F.regex.Replace(body, REGEX_PATTERN, REPLACEMENT_PATTERN)
End Function

Private Function getVariablesFromDtos(dtos As VBA.Collection) As Scripting.Dictionary
    Dim dto As Scripting.Dictionary
    Dim variable As EVbaVariable
    '------------------------------------------------------------------------------------------------------
    Set getVariablesFromDtos = F.dictionaries.Create(False)
    For Each dto In dtos
        Set variable = createVbaVariable(Me).loadFromDto(dto)
        Call F.dictionaries.addItem(getVariablesFromDtos, variable.getName, variable)
    Next dto
End Function







Public Function toString(Optional ignoreDescription As Boolean = False, _
                         Optional ignoreErrorHandling As Boolean = False, _
                         Optional ignoreWhiteSpaces As Boolean = True) As String
                         
    If Not ignoreDescription Then toString = pDescription
    If ignoreErrorHandling Then
        toString = toString & VBA.vbLf & clearBody
    Else
        toString = toString & VBA.vbLf & pBody
    End If
    
    If ignoreWhiteSpaces Then
        toString = removeWhiteSpace(toString)
    End If

End Function

Private Function clearBody() As String
    Dim objRegex As Object
    Dim patterns As Scripting.Dictionary
    Dim pattern As Variant
    Dim replacement As String
    '------------------------------------------------------------------------------------------------------
    
    Set patterns = F.dictionaries.Create(False)
    With patterns
        Call .Add("(\n|\r)\s*Const METHOD_NAME As String = .*?(\n|\r)", VBA.vbCrLf)
        Call .Add("(\n|\r)\s*Dim errLocal.*?(\n|\r)", VBA.vbCrLf)
        Call .Add("(\n|\r)\s*\'-*?\s*?(\n|\r)", VBA.vbCrLf)
        Call .Add("(\n|\r)\s*If Not DEV_MODE Then On Error GoTo ErrHandler.*?(\n|\r)", VBA.vbCrLf)
        Call .Add("(?:\n|\r)\s*'=*(?:\n|\r)ExitPoint:(?:\n|\r|.)*(End (Function|Sub|Property))", "$1")
        Call .Add("((?:\n|\r)End Sub)(?:\n|\r|.)", "$1")
    End With
    
    clearBody = VBA.Replace(pBody, VBA.vbLf, VBA.vbCrLf)
    clearBody = removeAfterComments(clearBody)
    
    Set objRegex = F.regex.Create(MultiLine:=False)
    For Each pattern In patterns
        replacement = patterns.item(pattern)
        objRegex.pattern = VBA.CStr(pattern)
        clearBody = objRegex.Replace(clearBody, replacement)
    Next pattern
    
End Function


Public Function compare(comparedMethod As EVbaMethod, _
                            Optional ignoreDescriptions As Boolean = True, _
                            Optional ignoreErrorHandling As Boolean = True, _
                            Optional ignoreWhiteSpaces As Boolean = True) As Boolean
    Dim strBaseMethod As String
    Dim strComparedMethod As String
    '------------------------------------------------------------------------------------------------------
    Dim message As String
    '------------------------------------------------------------------------------------------------------
    
    strBaseMethod = removeErrorHandling(removeSeparators(toString(ignoreDescriptions, ignoreErrorHandling, ignoreWhiteSpaces)))
    strComparedMethod = removeErrorHandling(removeSeparators(comparedMethod.toString(ignoreDescriptions, ignoreErrorHandling, ignoreWhiteSpaces)))
    compare = F.Strings.compareStrings(strBaseMethod, strComparedMethod, False, True)
    
    If Not compare Then
        message = "-------------------------------------------------" & VBA.vbCrLf & _
                  pModule.getName & "." & comparedMethod.getName & VBA.vbCrLf & _
                  "[B]" & VBA.vbCrLf & _
                  strBaseMethod & VBA.vbCrLf & _
                  "[C]" & VBA.vbCrLf & _
                  strComparedMethod & VBA.vbCrLf & _
                  "================================================="
        Call F.TextFiles.printToTextFile(message, getLogTextFilePath)
    End If
    
End Function

Private Function removeErrorHandling(text As String) As String
    Const REGEX_PATTERN As String = "(?:'=*?)?ExitPoint:DimerrNumberAsLong,errDescriptionAsString" & _
                                                "IferrNumberThenCallVBA\.err\.raise\(errNumber,CLASS_NAME&""\.""&METHOD_NAME," & _
                                                "errDescription\)ExitFunctionErrHandler:errNumber=err\.number:errDescription=" & _
                                                "err\.descriptionCallErrorManager\.addError\(CLASS_NAME,METHOD_NAME,errNumber," & _
                                                "errDescription,_?""(?:\w+[,;]\s?)*?\w+"",(\w+,\s?)*?\w+\)GoToExitPoint"
    '------------------------------------------------------------------------------------------------------
    removeErrorHandling = F.regex.Replace(text, REGEX_PATTERN, vbNullString)
End Function






'[ERROR HANDLING]
Public Function hasAnyErrorHandling() As Boolean
    Const REGEX_PATTERN As String = "^([\t ]*On Error|\w+:)"
    '------------------------------------------------------------------------------------------------------
    hasAnyErrorHandling = F.regex.checkIfMatch(pBody, REGEX_PATTERN)
End Function

Public Function isInterfaceMethod() As Boolean
    Const REGEX_PATTERN As String = "Public (Sub|Function) .*( _ &\n(?:.* _ &\n)*.*)?\nEnd \1"
    '------------------------------------------------------------------------------------------------------
    isInterfaceMethod = F.regex.checkIfMatch(pBody, REGEX_PATTERN)
End Function

Public Function isSimpleMethod() As Boolean
    Const SIMPLE_GETTER_REGEX_PATTERN As String = "Function (\w*)\(\) As [\w\.]*[\u000a\u000d]+\s*(?:Set )?\1 = p\w*[\u000a\u000d]+End Function"
    Const SIMPLE_SETTER_REGEX_PATTERN As String = "Sub set(\w*)\((\w*) As [\w\.]*\)[\u000a\u000d]+\s*(?:Set )?p\1 = \2[\u000a\u000d]+End Sub"
    Const SIMPLE_SETTER_FN_REGEX_PATTERN As String = "Function set(\w*)\((\w*) As [\w\.]*\) As \w*[\u000a\u000d]+\s*Set set\1 = Me[\u000a\u000d]+\s*(?:Set )?p\1 = \2[\u000a\u000d]+End Function"
    '------------------------------------------------------------------------------------------------------
    Dim patterns As VBA.Collection
    Dim pattern As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set patterns = F.Collections.Create(SIMPLE_GETTER_REGEX_PATTERN, SIMPLE_SETTER_REGEX_PATTERN, SIMPLE_SETTER_FN_REGEX_PATTERN)
    For Each pattern In patterns
        If F.regex.checkIfMatch(pBody, VBA.CStr(pattern)) Then
            isSimpleMethod = True
            Exit For
        End If
    Next pattern
    
End Function

Public Function hasNoErrorHandlingAnnotation() As Boolean
    Const REGEX_PATTERN As String = "'@no-error-handling"
    '------------------------------------------------------------------------------------------------------
    hasNoErrorHandlingAnnotation = F.regex.checkIfMatch(pDescription, REGEX_PATTERN)
End Function

Public Function hasCustomErrorHandlingAnnotation() As Boolean
    Const REGEX_PATTERN As String = "'@custom-error-handling"
    '------------------------------------------------------------------------------------------------------
    hasCustomErrorHandlingAnnotation = F.regex.checkIfMatch(pDescription, REGEX_PATTERN)
End Function

Public Function isOneLiner() As Boolean
    Const REGEX_PATTERN As String = "^Public Function \w*\(\) As [\w\.]*: Set"
    '------------------------------------------------------------------------------------------------------
    isOneLiner = F.regex.checkIfMatch(pBody, REGEX_PATTERN)
End Function

Public Function getBodyWithErrorHandlingAppended(isEntryPoint As Boolean) As String

    Call analyzeBody

    If isEntryPoint Then
        getBodyWithErrorHandlingAppended = getPreBodyPart & VBA.vbCrLf & appendEntryPointErrorHandling
    Else
        getBodyWithErrorHandlingAppended = getPreBodyPart & VBA.vbCrLf & appendRegularErrorHandling
    End If
    
End Function

Private Function getPreBodyPart() As String
    Dim preBodyLinesCounter As Long
    '------------------------------------------------------------------------------------------------------
    preBodyLinesCounter = pDeclarationLineIndex - pFirstLine
    If preBodyLinesCounter Then
        getPreBodyPart = getCodeModule.lines(pFirstLine, preBodyLinesCounter)
    End If
End Function

Private Function appendRegularErrorHandling() As String
    Dim arr As Variant
    Dim lineAfterDeclarations As Long
    '------------------------------------------------------------------------------------------------------
    
    arr = VBA.Split(pBody, VBA.vbCrLf)
    
    'Append method name line.
    Call appendMethodNameLine(arr(pDeclarationLinesCounter + LBound(arr) - 1))
    
    'Append On Error GoTo command
    lineAfterDeclarations = getFirstLineAfterVariablesDeclaration(arr)
    Call appendOnErrorGoToLine(arr(lineAfterDeclarations))
    
    'Append exitPoint and errHandler section
    Call appendExitPointAndErrorHandler(arr(UBound(arr)))
    
    appendRegularErrorHandling = VBA.Join(arr, VBA.vbCrLf)
    
End Function

Private Function appendEntryPointErrorHandling() As String
    Dim arr As Variant
    Dim lineAfterDeclarations As Long
    '------------------------------------------------------------------------------------------------------
    
    arr = VBA.Split(pBody, VBA.vbCrLf)
    Call removeEndingEmptyLines(arr)
    
    'Append method name line.
    Call appendMethodNameLine(arr(pDeclarationLinesCounter + LBound(arr) - 1))
    
    'Append On Error GoTo command
    lineAfterDeclarations = getFirstLineAfterVariablesDeclaration(arr)
    Call appendOnErrorGoToLine(arr(lineAfterDeclarations), True)
    
    'Append exitPoint and errHandler section
    Call appendExitPointAndErrorHandler(arr(UBound(arr)), True)
    
    appendEntryPointErrorHandling = VBA.Join(arr, VBA.vbCrLf)
    
End Function

Private Sub removeEndingEmptyLines(ByRef arr As Variant)
    Dim lastNonEmptyLine As Long
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If F.arrays.isDefined(arr) Then
        For i = UBound(arr) To LBound(arr) Step -1
            If VBA.Len(VBA.Trim(VBA.CStr(arr(i)))) > 0 Then
                lastNonEmptyLine = i
                Exit For
            End If
        Next i
    End If
    
    ReDim Preserve arr(LBound(arr) To lastNonEmptyLine)
    
End Sub

Private Sub appendMethodNameLine(ByRef line As Variant)
    Const LINE_PATTERN As String = "    Const METHOD_NAME As String = ""{0}"""
    '------------------------------------------------------------------------------------------------------
    line = line & VBA.vbCrLf & _
                F.Strings.Format(LINE_PATTERN, pName) & VBA.vbCrLf & _
                VBA.String(METHOD_LEVEL_SEPARATOR_INDENT, " ") & "'" & VBA.String(METHOD_LEVEL_SEPARATOR_LENGTH, "-")
End Sub

Private Function getFirstLineAfterVariablesDeclaration(arr As Variant) As Long
    Const DECLARATION_OR_SEPARATOR_LINE_PATTERN As String = "^([\t ]*(?:Dim|Const|Static|'-))"
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If F.dictionaries.isNonEmptyDictionary(pVariables) Then
        If F.arrays.isDefined(arr) Then
            For i = LBound(arr) + pDeclarationLinesCounter To UBound(arr)
                If Not F.regex.checkIfMatch(VBA.CStr(arr(i)), DECLARATION_OR_SEPARATOR_LINE_PATTERN) Then
                    getFirstLineAfterVariablesDeclaration = i
                    Exit For
                End If
            Next i
        End If
    Else
        getFirstLineAfterVariablesDeclaration = LBound(arr) + pDeclarationLinesCounter
    End If
    
End Function

Private Sub appendOnErrorGoToLine(ByRef line As Variant, Optional isEntryPoint As Boolean = False)
    Const LINE_PATTERN As String = "    If Not DEV_MODE Then On Error GoTo ErrHandler"
    Const ENTRY_POINT_EXTRA_LINE As String = "    Call ErrorManager.clear"
    '------------------------------------------------------------------------------------------------------
    
    line = VBA.vbCrLf & _
            LINE_PATTERN & _
            VBA.IIf(isEntryPoint, VBA.vbCrLf & ENTRY_POINT_EXTRA_LINE, VBA.vbNullString) & _
            VBA.vbCrLf & _
            VBA.IIf(VBA.Len(VBA.Trim(line)) = 0, VBA.vbNullString, VBA.vbCrLf) & _
            line

End Sub

Private Sub appendExitPointAndErrorHandler(ByRef line As Variant, Optional isEntryPoint As Boolean = False)
    line = VBA.vbCrLf & _
           "'" & VBA.String(MODULE_LEVEL_SEPARATOR_LENGTH, "=") & VBA.vbCrLf & _
           "ExitPoint:" & VBA.vbCrLf & _
           "    Dim errNumber As Long, errDescription As String" & VBA.vbCrLf & _
           getReraiseErrorLine(isEntryPoint) & _
           "    Exit " & getMethodTypeName(pMethodType) & VBA.vbCrLf & _
                          VBA.vbCrLf & _
           "'" & VBA.String(MODULE_LEVEL_SEPARATOR_LENGTH, "-") & VBA.vbCrLf & _
           "ErrHandler:" & VBA.vbCrLf & _
           "    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description" & VBA.vbCrLf & _
           "    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription" & _
                          getErrorHandlingParams & VBA.vbCrLf & _
           "    GoTo ExitPoint" & VBA.vbCrLf & VBA.vbCrLf & _
           line
End Sub

Private Function getReraiseErrorLine(isEntryPoint As Boolean) As String
    If isEntryPoint Then
        getReraiseErrorLine = "    If errNumber > 0 Then Call ErrorManager.save" & VBA.vbCrLf
    Else
        getReraiseErrorLine = "    If errNumber > 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & ""."" & METHOD_NAME, errDescription)" & VBA.vbCrLf
    End If
End Function

Private Function getErrorHandlingParams() As String
    Dim params As Scripting.Dictionary
    Dim colParams As VBA.Collection
    Dim strParams As String
    '------------------------------------------------------------------------------------------------------
    
    Set params = F.dictionaries.Join(pInputParams, pVariables)

    If F.dictionaries.isNonEmptyDictionary(params) Then
        Set colParams = F.Collections.Map( _
                            F.Collections.Filter( _
                                F.dictionaries.toCollection(params), "$.isNotConst"), _
                            "$.getName")
        'Set colParams = F.dictionaries.toCollection(params, DictPart_KeyOnly)
        
        
        If colParams.count = 1 Then
            getErrorHandlingParams = ", """ & F.Collections.toString(colParams, , "; ") & """" & _
                                     ", " & F.Collections.toString(colParams, , ", ") & ")"
        ElseIf colParams.count > 1 Then
            getErrorHandlingParams = ", _" & VBA.vbCrLf & VBA.String(34, " ") & _
                                     """" & F.Collections.toString(colParams, , "; ") & """" & _
                                     ", _" & VBA.vbCrLf & VBA.String(34, " ") & _
                                     F.Collections.toString(colParams, , ", ") & ")"
        Else
            getErrorHandlingParams = ")"
        End If
    Else
        getErrorHandlingParams = ")"
    End If
    
End Function


Public Sub replaceWithNewVersion(newVersion As String)
    Dim selection As ECodeSelection
    '------------------------------------------------------------------------------------------------------
    
    Set selection = Fn.Create.codeSelection. _
                                setComponent(pModule.getComponent). _
                                setStartRow(pFirstLine).setStartCol(1). _
                                setEndRow(pFirstLine + pTotalLines - 1).setEndCol(1)
    Call selection.replaceWithGivenCode(newVersion)
    
End Sub



Public Function getErrorHandlingStatus() As String
    Const REGEX_HAS_METHOD_NAME As String = "Const METHOD_NAME As String = ""{0}"""
    Const REGEX_HAS_ON_ERR_GO_TO As String = "If Not DEV_MODE Then On Error GoTo ErrHandler"
    'Const REGEX_ERR_HANDLER As String = "ExitPoint:[\u000A\u000D](?:.|[\u000A\u000D])*[\u000A\u000D]\s*Exit (?:Function|Sub)[\u000A\u000D\s]*'-*[\u000A\u000D]+ErrHandler:\s*errNumber = VBA\.Err\.Number: errDescription = VBA\.Err\.Description[\u000A\u000D]\s*Call ErrorManager\.addError(?:.|[\u000A\u000D])*?GoTo ExitPoint"
    Const REGEX_EXIT_POINT As String = "ExitPoint:(?:.|\u000A|\u000D)*?Exit (?:Function|Sub)"
    Const REGEX_ERR_HANDLER As String = "ErrHandler:\s*errNumber = VBA\.Err\.Number: errDescription = VBA\.Err\.Description[\u000A\u000D]\s*Call ErrorManager\.addError(?:.|[\u000A\u000D])*?GoTo ExitPoint"
    Const REGEX_ERR_HANDLER_AFTER_EXIT_POINT As String = "ExitPoint:(?:.|\u000A|\u000D)*?ErrHandler:"
    Const REGEX_FREEZE_EXCEL As String = "Call F\.Excel\.freeze"
    Const REGEX_UNFREEZE_EXCEL As String = "ExitPoint:(?:.|[\u000A\u000D]\s)*Call F\.Excel\.unfreeze"
    Const REGEX_NOTHING_AFTER_ERR_HANDLING As String = "GoTo ExitPoint[\n\s\t]*End (?:Sub|Function)"
    '------------------------------------------------------------------------------------------------------
    Dim errors As VBA.Collection
    Dim regexPattern As String
    '------------------------------------------------------------------------------------------------------
    
    Set errors = New VBA.Collection
    
    Call analyzeBody
    
    'Has proper method name
    regexPattern = F.Strings.Format(REGEX_HAS_METHOD_NAME, pName)
    If Not F.regex.checkIfMatch(pBody, regexPattern) Then
        Call errors.Add("No method name constant")
    End If
    
    'Has On Error GoTo ErrHandler
    If Not F.regex.checkIfMatch(pBody, REGEX_HAS_ON_ERR_GO_TO) Then
        Call errors.Add("No On Error GoTo ErrHandler command")
    End If
    
    'Has proper ExitPoint section
    If Not F.regex.checkIfMatch(pBody, REGEX_EXIT_POINT) Then
        Call errors.Add("Invalid ExitPoint part")
    End If
    
    'Has proper ErrHandler section
    If Not F.regex.checkIfMatch(pBody, REGEX_ERR_HANDLER) Then
        Call errors.Add("Invalid ErrHandler part")
    End If
    
    'Has ErrHandler after ExitPoint
    If Not F.regex.checkIfMatch(pBody, REGEX_ERR_HANDLER_AFTER_EXIT_POINT) Then
        Call errors.Add("ErrHandler should be after ExitPoint")
    End If
    
    'Has correct variables in ErrManager.addError part
    If Not verifyVariablesInExitPoint Then
        Call errors.Add("Discrepancies in variables definitions in ErrorManager.add")
    End If
    
    'if xlsState.freeze used, has .unfreeze in exitpoint
    If Not F.regex.checkIfMatch(pBody, REGEX_FREEZE_EXCEL) Then
        ' no Excel freezing - no need to unfreeze it
    ElseIf Not F.regex.checkIfMatch(pBody, REGEX_UNFREEZE_EXCEL) Then
        Call errors.Add("No Excel unfreezing in ExitPoint")
    End If
    
    'Nothing after GoToExitPoint
    If Not F.regex.checkIfMatch(pBody, REGEX_NOTHING_AFTER_ERR_HANDLING) Then
        Call errors.Add("There is some code between ErrorHandler and end of function/sub")
    End If

    If errors.count Then
        getErrorHandlingStatus = "errors: " & F.Collections.toString(errors, StringifyMode_Normal, " | ")
    Else
        getErrorHandlingStatus = "correct"
    End If

End Function

Private Function verifyVariablesInExitPoint() As Boolean
    'Const REGEX_VARIABLES As String = "Call ErrorManager\.addError\(CLASS_NAME, METHOD_NAME, errNumber, errDescription, (?:_\n\s*)?""([\w;, ]*)"", (?:_\n\s*)?([\w, ]*)\)"
    Const REGEX_VARIABLES As String = "Call ErrorManager\.addError\(CLASS_NAME, METHOD_NAME, errNumber, errDescription, (?:_[\u000A\u000D]\s*)?""([\w;, ]*)"", (?:_[\u000A\u000D]\s*)?([\w, \(\)\.]*)\)"
    '------------------------------------------------------------------------------------------------------
    Dim matches As Variant
    Dim names As Variant
    Dim variables As Variant
    Dim i As Long
    Dim count As Long
    '------------------------------------------------------------------------------------------------------
    
    matches = F.regex.getMatchArray(pBody, REGEX_VARIABLES)
    
    If F.arrays.isDefined(matches) Then
        names = VBA.Split(VBA.Replace(matches(1), " ", vbNullString), ";")
        variables = VBA.Split(VBA.Replace(matches(2), " ", vbNullString), ",")
        count = F.arrays.getSize(names)
        
        For i = LBound(variables) To UBound(variables)
            If F.Strings.compareStrings(variables(i), "F.Excel.getExcelStateString(xlsState)", False) Then
                variables(i) = "xlsState"
            ElseIf F.Strings.compareStrings(variables(i), "areaToString(screenArea)", False) Then
                variables(i) = "screenArea"
            End If
        Next i
        
        If F.arrays.compare(names, variables) Then
            verifyVariablesInExitPoint = True
        Else
            verifyVariablesInExitPoint = False
        End If
    Else
        verifyVariablesInExitPoint = True
    End If
    
    If count <> countVariablesAndInputParams Then
        verifyVariablesInExitPoint = False
    End If
    
End Function

Private Function countVariablesAndInputParams() As Long
    Dim key As Variant
    Dim var As EVbaVariable
    '------------------------------------------------------------------------------------------------------
    
    For Each key In pInputParams.Keys
        Set var = pInputParams.item(key)
        If var.getName <> "errNumber" Then countVariablesAndInputParams = countVariablesAndInputParams + 1
    Next key
    
    For Each key In pVariables.Keys
        Set var = pVariables.item(key)
        If var.getVarType <> VarType_Const Then
            countVariablesAndInputParams = countVariablesAndInputParams + 1
        End If
    Next key
    
End Function
