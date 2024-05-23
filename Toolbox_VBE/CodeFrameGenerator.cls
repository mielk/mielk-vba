VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeFrameGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "CodeFrameGenerator"
'[Helper constants] ---------------------------------------------------------------------------------------
Private Const FRAME_LENGTH As Long = 101
'----------------------------------------------------------------------------------------------------------
Private pSelection As ECodeSelection
Private pDescription As String
Private pCode As String
Private pStartIndex As Long
'----------------------------------------------------------------------------------------------------------


Public Sub addSeparatorLine()
    Dim code As String
    '------------------------------------------------------------------------------------------------------
    Set pSelection = Fn.Create.currentCodeSelection
    code = VBA.String(METHOD_LEVEL_SEPARATOR_INDENT, " ") & "'" & VBA.String(METHOD_LEVEL_SEPARATOR_LENGTH, "-")
    Call pSelection.insertCode(code)
End Sub


Public Sub addFrame()
    Dim frameDescription As String
    Dim code As String
    Dim convertedCode As String
    '------------------------------------------------------------------------------------------------------
    
    Set pSelection = Fn.Create.currentCodeSelection
    pDescription = VBA.InputBox("Enter frame description")
    pCode = pSelection.readContent
    
    convertedCode = convertCode(pCode)
    
    Call pSelection.replaceWithGivenCode(convertedCode)
    
End Sub

Private Function convertCode(code As String) As String
    Dim codeLines() As String
    Dim i As Long
    Dim codeLine As String
    '------------------------------------------------------------------------------------------------------
    Dim convertedLines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    codeLines = VBA.Split(code, VBA.vbCrLf)
    If F.arrays.isDefined(codeLines) Then
        pStartIndex = findStartIndex(codeLines)
        
        Set convertedLines = New VBA.Collection
        With convertedLines
            Call .add(getStartLines)
            For i = LBound(codeLines) To UBound(codeLines)
                codeLine = codeLines(i)
                Call .add(getConvertedCodeLine(codeLine))
            Next i
            Call .add(getBottomLine)
        End With
        
        convertCode = F.Collections.toString(convertedLines, , VBA.vbCrLf)
        
    Else
        convertCode = code
    End If
    
End Function

Private Function findStartIndex(codeLines() As String) As Long
    Const LEFT_INDENT_REGEX As String = "^([\s|\t]*)"
    Const IS_EMPTY_LINE_REGEX As String = "^[\s|\t]*$"
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim codeLine As String
    Dim spaces As String
    Dim isFirstLine As Boolean:             isFirstLine = True
    '------------------------------------------------------------------------------------------------------

    If F.arrays.isDefined(codeLines) Then
        For i = LBound(codeLines) To UBound(codeLines)
            codeLine = codeLines(i)
            
            If F.Regex.checkIfMatch(codeLine, IS_EMPTY_LINE_REGEX) Then
                ' Empty line - skip it
            Else
                spaces = F.Regex.getFirstGroupMatch(codeLine, LEFT_INDENT_REGEX)
                If isFirstLine Or VBA.Len(spaces) < findStartIndex Then
                    findStartIndex = VBA.Len(spaces)
                End If
                isFirstLine = False
            End If
        Next i
    End If
    
End Function

Private Function getStartLines() As String
    Const SINGLE_LINE_PATTERN As String = "{0}'[{1}] {2}|"
    '------------------------------------------------------------------------------------------------------
    getStartLines = F.Strings.Format(SINGLE_LINE_PATTERN, _
                        VBA.String(pStartIndex, " "), _
                        pDescription, _
                        VBA.String(FRAME_LENGTH - pStartIndex - VBA.Len(pDescription) - 5, "-"))
End Function

Private Function getConvertedCodeLine(text As String) As String
    Const LINE_PATTERN As String = "{0}{1}'|"
    '------------------------------------------------------------------------------------------------------
    getConvertedCodeLine = F.Strings.Format(LINE_PATTERN, _
                        text, _
                        VBA.String(FRAME_LENGTH - VBA.Len(text) - 2, " "))
End Function

Private Function getBottomLine() As String
    getBottomLine = VBA.String(pStartIndex, " ") & "'" & _
                    VBA.String(FRAME_LENGTH - pStartIndex - 2, "-") & "|"
End Function


