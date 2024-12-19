VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaCodeParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "VbaCodeParser"
'----------------------------------------------------------------------------------------------------------
Private Const TO_SINGLE_LINERS_REPLACE_PATTERN As String = " _[\s\t]*[\u000A\u000D]*"
'----------------------------------------------------------------------------------------------------------


Public Function convertToSingleLiners(body As String)
    convertToSingleLiners = F.regex.Replace(body, TO_SINGLE_LINERS_REPLACE_PATTERN, " ")
End Function

Public Function countDeclarationLines(body As String) As Long
    Dim arrBody() As String
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    arrBody = VBA.Split(body, VBA.vbCrLf)
    If F.arrays.isDefined(arrBody) Then
        For i = LBound(arrBody) To UBound(arrBody)
            countDeclarationLines = countDeclarationLines + 1
            If Not F.Strings.endsWith(arrBody(i), VBA_LINE_BREAK) Then
                Exit For
            End If
        Next i
    End If
End Function

Public Function getMethodDeclarationLine(body As String) As String
    Dim lines() As String
    Dim line As String
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    lines = VBA.Split(convertToSingleLiners(body), VBA.vbCrLf)
    If F.arrays.isDefined(lines) Then
        For i = LBound(lines) To UBound(lines)
            line = VBA.Trim(lines(i))
            If VBA.Len(line) = 0 Then
            ElseIf VBA.Left$(line, 1) = "'" Then
            Else
                getMethodDeclarationLine = line
                Exit For
            End If
        Next i
    End If
    
End Function



Public Function findMethodScope(body As String) As ScopeTypeEnum
    Const REGEX_PATTERN As String = "(Private|Public|Friend) (?:Sub|Function|Property)"
    '------------------------------------------------------------------------------------------------------
    Dim scopeTypeName As String
    '------------------------------------------------------------------------------------------------------
    scopeTypeName = F.regex.getFirstGroupMatch(body, REGEX_PATTERN)
    findMethodScope = getScopeTypeFromName(scopeTypeName)
End Function

Public Function findMethodType(body As String) As MethodTypeEnum
    Const REGEX_PATTERN As String = "(Sub|Function|Property (?:L|G|S)et) [a-zA-Z0-9_]*\("
    '------------------------------------------------------------------------------------------------------
    Dim methodTypeName As String
    '------------------------------------------------------------------------------------------------------
    methodTypeName = F.regex.getFirstGroupMatch(body, REGEX_PATTERN)
    findMethodType = getMethodTypeFromName(methodTypeName)
End Function

Public Function getMethodInputParamsDtos(body As String) As VBA.Collection
    Const REGEX_PATTERN As String = "(Dim)?[\(|,| ](?:(Optional|ParamArray) |)(?:(ByRef|ByVal) |)(\w+)(\(\)|)" & _
                                    "(?: As ([\w\.]+))(?:| = ([\w\.]+|\"".*\""))[\)|,|]"
    '------------------------------------------------------------------------------------------------------
    Dim col As VBA.Collection
    Dim dto As Scripting.Dictionary
    Dim arr As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set getMethodInputParamsDtos = New VBA.Collection
    
    Set col = F.regex.getMatchArraysCollection(body, REGEX_PATTERN)
    If F.Collections.isNonEmptyCollection(col) Then
        For Each arr In col
            If VBA.IsEmpty(arr(LBound(arr))) Then
                Set dto = F.dictionaries.Create(False)
                With dto
                    Call .Add(Props_Vbe.IsOptional.getName, arr(LBound(arr) + 1))
                    Call .Add(Props_Vbe.IsParamArray.getName, arr(LBound(arr) + 1))
                    Call .Add(Props_Vbe.PassMode.getName, arr(LBound(arr) + 2))
                    Call .Add(props.name.getName, arr(LBound(arr) + 3))
                    Call .Add(Props_Vbe.IsArray.getName, arr(LBound(arr) + 4))
                    Call .Add(Props_Vbe.dataType.getName, arr(LBound(arr) + 5))
                    Call .Add(Props_Vbe.DefaultValue.getName, arr(LBound(arr) + 6))
                End With
                Call getMethodInputParamsDtos.Add(dto)
            End If
        Next arr
    End If
    
End Function

Public Function findMethodReturnedType(body As String) As String
    Const REGEX_PATTERN As String = "\) As ([\w\.]+(?:\(\)|))(?:|\s*'.*)$"
    '------------------------------------------------------------------------------------------------------
    findMethodReturnedType = F.regex.getFirstGroupMatch(body, REGEX_PATTERN)
End Function

Public Function getVariablesDtos(body As String) As VBA.Collection
    Dim arrBody() As String
    Dim i As Long
    Dim line As String
    Dim dtos As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    Set getVariablesDtos = New VBA.Collection
    
    arrBody = VBA.Split(convertToSingleLiners(body), VBA.vbCrLf)
    If F.arrays.isDefined(arrBody) Then
        For i = LBound(arrBody) To UBound(arrBody)
            line = arrBody(i)
            Set dtos = getVariablesDtosFromCodeLine(line)
            If F.Collections.isNonEmptyCollection(dtos) Then
                Call F.Collections.appendCollection(getVariablesDtos, dtos)
            End If
        Next i
    End If
    
End Function

Public Function getVariablesDtosFromCodeLine(codeLine As String) As VBA.Collection
    Const VARIABLES_REGEX_PATTERN As String = "(?:dim |const |static |, )(\w+)(\(.*\))?(?: As ([\w\.]+))?(?: = (.*))?"
    '------------------------------------------------------------------------------------------------------
    Dim col As VBA.Collection
    Dim arr As Variant
    Dim dto As Scripting.Dictionary
    Dim variableType As VarTypeEnum
    '------------------------------------------------------------------------------------------------------
    
    Set getVariablesDtosFromCodeLine = New VBA.Collection
    
    variableType = getVariableType(codeLine)
    If variableType > VarType_Unknown Then
        Set col = F.regex.getMatchArraysCollection(removeAfterColonPart(codeLine), VARIABLES_REGEX_PATTERN)
        If F.Collections.isNonEmptyCollection(col) Then
            For Each arr In col
                Set dto = F.dictionaries.Create(False)
                With dto
                    Call .Add(props.name.getDbField, arr(LBound(arr)))
                    Call .Add(Props_Vbe.variableType.getDbField, variableType)
                    Call .Add(Props_Vbe.IsArray.getDbField, arr(LBound(arr) + 1))
                    Call .Add(Props_Vbe.ArraySize.getDbField, arr(LBound(arr) + 1))
                    Call .Add(Props_Vbe.dataType.getDbField, arr(LBound(arr) + 2))
                    Call .Add(Props_Vbe.DefaultValue.getDbField, arr(LBound(arr) + 3))
                End With
                Call getVariablesDtosFromCodeLine.Add(dto)
            Next arr
        End If
    End If
    
End Function

Public Function getVariableType(ByVal codeLine As String) As VarTypeEnum
    Const REGEX_PATTERN As String = "\s*\b(Dim \b(?!err(Number|Description))|Static \b|Const \b)"
    '------------------------------------------------------------------------------------------------------
    Dim keyword As String
    '------------------------------------------------------------------------------------------------------
    keyword = VBA.Trim$(F.regex.getFirstGroupMatch(codeLine, REGEX_PATTERN))
    getVariableType = getVarTypeFromName(keyword)
End Function

Public Function removeAfterColonPart(ByVal line As String) As String
    Dim colonPosition As Long
    '------------------------------------------------------------------------------------------------------
    colonPosition = VBA.InStr(1, line, ":")
    If colonPosition Then
        removeAfterColonPart = VBA.Left$(line, colonPosition + 1)
    Else
        removeAfterColonPart = line
    End If
End Function
