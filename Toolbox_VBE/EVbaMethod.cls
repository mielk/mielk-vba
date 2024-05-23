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
        Call .add(META_TAG_NAME, pName)
        Call .add(META_TAG_PROC_TYPE, pProcType)
        Call .add(META_TAG_DESCRIPTION, pDescription)
        Call .add(META_TAG_BODY, pBody)
    End With
End Function

Public Function getBody() As String
    getBody = pBody
End Function



'[ANALYZING]
Public Sub analyzeBody()
    With VbaCodeParser
        pDeclarationLinesCounter = .countDeclarationLines(pBody)
        pDeclarationLine = .getMethodDeclarationLine(pBody)
        pScope = .findMethodScope(pBody)
        pMethodType = .findMethodType(pBody)
        pReturnedType = .findMethodReturnedType(pBody)
        Set pInputParams = getVariablesFromDtos(.getMethodInputParamsDtos(pBody))
        Set pVariables = getVariablesFromDtos(.getVariablesDtos(pBody))
    End With
    
    Stop
    
    Dim inputParams2 As VBA.Collection
    Set inputParams2 = F.dictionaries.toCollection(pInputParams, DictPart_ValueOnly)
    
    Stop
    
End Sub

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
        Call .add("(\n|\r)\s*Const METHOD_NAME As String = .*?(\n|\r)", VBA.vbCrLf)
        Call .add("(\n|\r)\s*Dim errLocal.*?(\n|\r)", VBA.vbCrLf)
        Call .add("(\n|\r)\s*\'-*?\s*?(\n|\r)", VBA.vbCrLf)
        Call .add("(\n|\r)\s*If Not DEV_MODE Then On Error GoTo ErrHandler.*?(\n|\r)", VBA.vbCrLf)
        Call .add("(?:\n|\r)\s*'=*(?:\n|\r)ExitPoint:(?:\n|\r|.)*(End (Function|Sub|Property))", "$1")
        Call .add("((?:\n|\r)End Sub)(?:\n|\r|.)", "$1")
    End With
    
    clearBody = VBA.Replace(pBody, VBA.vbLf, VBA.vbCrLf)
    clearBody = removeAfterComments(clearBody)
    
    Set objRegex = F.Regex.Create(MultiLine:=False)
    For Each pattern In patterns
        replacement = patterns.item(pattern)
        objRegex.pattern = VBA.CStr(pattern)
        clearBody = objRegex.Replace(clearBody, replacement)
    Next pattern
    
End Function


Public Function removeWhiteSpace(text As String) As String
    Const REGEX_PATTERN As String = "(\r|\n)(\s|\t)*"
    '------------------------------------------------------------------------------------------------------
    Dim objRegex As Object
    '------------------------------------------------------------------------------------------------------
    
    Set objRegex = F.Regex.Create(REGEX_PATTERN, MultiLine:=False)
    removeWhiteSpace = objRegex.Replace(F.Strings.trimFull(text), VBA.vbCrLf)
    removeWhiteSpace = VBA.Replace(removeWhiteSpace, VBA.Chr(13), vbNullString)
    removeWhiteSpace = VBA.Replace(removeWhiteSpace, VBA.Chr(10), vbNullString)
    removeWhiteSpace = VBA.Replace(removeWhiteSpace, VBA.Chr(9), vbNullString)
    removeWhiteSpace = VBA.Replace(removeWhiteSpace, VBA.Chr(32), vbNullString)
    
End Function
    
Private Function removeAfterComments(text As String) As String
    Dim keywords As Variant
    Dim keyword As Variant
    Dim comments As String
    '------------------------------------------------------------------------------------------------------
    
    removeAfterComments = text
    keywords = VBA.Array(VBA_SUB, VBA_FUNCTION, VBA_PROPERTY)
    
    For Each keyword In keywords
        comments = F.Strings.substring(text, VBA.Chr(13) & "End " & keyword, vbNullString)
        If VBA.Len(comments) Then removeAfterComments = VBA.Replace(removeAfterComments, comments, vbNullString)
    Next keyword
    
End Function



Public Function compare(comparedMethod As EVbaMethod, _
                            Optional ignoreDescriptions As Boolean = True, _
                            Optional ignoreErrorHandling As Boolean = True, _
                            Optional ignoreWhiteSpaces As Boolean = True) As Boolean
    Dim strBaseMethod As String
    Dim strComparedMethod As String
    '------------------------------------------------------------------------------------------------------
    
    strBaseMethod = removeErrorHandling(removeSeparators(toString(ignoreDescriptions, ignoreErrorHandling, ignoreWhiteSpaces)))
    strComparedMethod = removeErrorHandling(removeSeparators(comparedMethod.toString(ignoreDescriptions, ignoreErrorHandling, ignoreWhiteSpaces)))
    compare = F.Strings.compareStrings(strBaseMethod, strComparedMethod, False, True)
    
End Function

Private Function removeErrorHandling(text As String) As String
    Const REGEX_PATTERN As String = "(?:'=*?)?ExitPoint:DimerrNumberAsLong,errDescriptionAsString" & _
                                                "IferrNumberThenCallVBA\.err\.raise\(errNumber,CLASS_NAME&""\.""&METHOD_NAME," & _
                                                "errDescription\)ExitFunctionErrHandler:errNumber=err\.number:errDescription=" & _
                                                "err\.descriptionCallErrorManager\.addError\(CLASS_NAME,METHOD_NAME,errNumber," & _
                                                "errDescription,_?""(?:\w+[,;]\s?)*?\w+"",(\w+,\s?)*?\w+\)GoToExitPoint"
    '------------------------------------------------------------------------------------------------------
    removeErrorHandling = F.Regex.Replace(text, REGEX_PATTERN, vbNullString)
End Function

Private Function removeSeparators(text As String) As String
    Const REGEX_PATTERN As String = "'[=-]+"
    '------------------------------------------------------------------------------------------------------
    removeSeparators = F.Regex.Replace(text, REGEX_PATTERN, vbNullString)
End Function


