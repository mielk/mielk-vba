Attribute VB_Name = "modHelpers"
Option Explicit

Private Const CLASS_NAME As String = "modHelpers"
'----------------------------------------------------------------------------------------------------------

Public Function getLogTextFilePath() As String
    getLogTextFilePath = "C:\_projects\_lib\compare\comparison.txt"
End Function


Public Function removeSeparators(text As String) As String
    'Const REGEX_PATTERN As String = "'[=-]+"
    Const REGEX_PATTERN As String = "'(\[[\w\d\-\., ]*\] ?)?[=-]+"
    '------------------------------------------------------------------------------------------------------
    removeSeparators = F.regex.Replace(text, REGEX_PATTERN, vbNullString)
End Function

Public Function removeWhiteSpace(text As String) As String
    'Const REGEX_PATTERN As String = "(\r|\n)(\s|\t)*"
    Const REGEX_PATTERN As String = "(?: _)?(\r|\n)(\s|\t)*"
    '------------------------------------------------------------------------------------------------------
    Dim objRegex As Object
    '------------------------------------------------------------------------------------------------------
    
    Set objRegex = F.regex.Create(REGEX_PATTERN, MultiLine:=False)
    removeWhiteSpace = objRegex.Replace(F.Strings.trimFull(text), VBA.vbCrLf)
    removeWhiteSpace = VBA.Replace(removeWhiteSpace, VBA.Chr(13), vbNullString)
    removeWhiteSpace = VBA.Replace(removeWhiteSpace, VBA.Chr(10), vbNullString)
    removeWhiteSpace = VBA.Replace(removeWhiteSpace, VBA.Chr(9), vbNullString)
    removeWhiteSpace = VBA.Replace(removeWhiteSpace, VBA.Chr(32), vbNullString)
    
End Function

Public Function removeAfterComments(text As String) As String
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



