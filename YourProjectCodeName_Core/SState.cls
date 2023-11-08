VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "SState"
'----------------------------------------------------------------------------------------------------------
Private Const PRIMITIVE_SYMBOL As String = "P|"
Private Const DICT_SYMBOL As String = "D|"
Private Const COLLECTION_SYMBOL As String = "C|"
Private Const ARRAY_SYMBOL As String = "A|"
'----------------------------------------------------------------------------------------------------------
Private Const QUOTE_SYMBOL As String = "#Chr34#"
'----------------------------------------------------------------------------------------------------------


Public Sub setValue(prop As enumProperty, value As Variant)
    Dim serialization As String
    '------------------------------------------------------------------------------------------------------
    
    If F.Variables.isPrimitive(value) Then
        serialization = PRIMITIVE_SYMBOL & value
    ElseIf F.dictionaries.isDictionary(value) Then
        serialization = DICT_SYMBOL & F.json.ConvertToJson(value)
    ElseIf F.Collections.isCollection(value) Then
        serialization = COLLECTION_SYMBOL & F.json.ConvertToJson(value)
    ElseIf VBA.IsArray(value) Then
        serialization = ARRAY_SYMBOL & F.json.ConvertToJson(value)
    End If
    
    serialization = VBA.Replace(serialization, VBA.Chr(34), QUOTE_SYMBOL)
    Call F.ExcelNames.assignValue(Excel.ThisWorkbook, prop.getName, serialization, True)
    
End Sub



Public Function getValue(prop As enumProperty) As Variant
    Dim value As Variant
    Dim dataTypeSymbol As String
    Dim serialization As String
    '------------------------------------------------------------------------------------------------------
    
    value = F.ExcelNames.getValue(Excel.ThisWorkbook, prop.getName)
    dataTypeSymbol = getDataTypeSymbol(VBA.CStr(value))
    serialization = VBA.Replace(VBA.Mid$(value, VBA.Len(dataTypeSymbol) + 1), QUOTE_SYMBOL, VBA.Chr(34))
    
    Select Case dataTypeSymbol
        Case PRIMITIVE_SYMBOL
            If F.Math.IsNumeric(serialization) Then getValue = serialization * 1 Else getValue = serialization
        Case DICT_SYMBOL
            Set getValue = F.json.ParseJson(serialization)
        Case COLLECTION_SYMBOL
            Stop
        Case ARRAY_SYMBOL
            Stop
    End Select
    
End Function


Private Function getDataTypeSymbol(serialization As String) As String
    Dim arrSymbols As Variant
    Dim symbol As Variant
    '------------------------------------------------------------------------------------------------------
    
    arrSymbols = VBA.Array(PRIMITIVE_SYMBOL, DICT_SYMBOL, COLLECTION_SYMBOL, ARRAY_SYMBOL)
    
    For Each symbol In arrSymbols
        If F.Strings.startsWith(serialization, VBA.CStr(symbol)) Then
            getDataTypeSymbol = symbol
            Exit For
        End If
    Next symbol
    
End Function


Public Sub clear()
    Dim name As Excel.name
    '------------------------------------------------------------------------------------------------------
    For Each name In Excel.ThisWorkbook.Names
        Call name.Delete
    Next name
End Sub
