VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FNumerics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FNumerics"
'----------------------------------------------------------------------------------------------------------

Public Function getMin(ParamArray values() As Variant) As Variant
    Dim varItem As Variant
    '------------------------------------------------------------------------------------------------------
    
    For Each varItem In values
        If VBA.IsEmpty(getMin) Then
            getMin = varItem
        ElseIf getMin > varItem Then
            getMin = varItem
        End If
    Next varItem
    
End Function

Public Function getMax(ParamArray values() As Variant) As Variant
    Dim varValue As Variant
    '------------------------------------------------------------------------------------------------------
    
    For Each varValue In values
        If VBA.IsEmpty(getMax) Then
            getMax = varValue
        ElseIf getMax < varValue Then
            getMax = varValue
        End If
    Next varValue
    
End Function

Public Function roundTo(value As Double, factor As Double) As Double
    Dim multiplier As Variant
    '------------------------------------------------------------------------------------------------------
    
    multiplier = VBA.Int(value / factor)
    roundTo = multiplier * factor
    If (value - roundTo > factor / 2) Then roundTo = roundTo + factor
    
End Function

Public Function isInRange(value As Variant, minLimit As Variant, maxLimit As Variant, _
                    Optional includeMinLimit As Boolean = True, _
                    Optional includeMaxLimit As Boolean = True) As Boolean
    Dim dblValue As Double
    Dim dblMinLimit As Double
    Dim dblMaxLimit As Double
    '------------------------------------------------------------------------------------------------------
    
'    'Throw exceptions if any of value is not numeric.
'    If Not isNumeric(value) Then Call Exceptions.IllegalDataType.raise
'    If Not isNumeric(minLimit) Then Call Exceptions.IllegalDataType.raise
'    If Not isNumeric(maxLimit) Then Call Exceptions.IllegalDataType.raise
    
    'Convert parameters to Double type.
    dblValue = value ' toNumeric(value)
    dblMinLimit = minLimit ' toNumeric(minLimit)
    dblMaxLimit = maxLimit ' toNumeric(maxLimit)

    'Check low bound.
    If includeMinLimit Then
        If dblValue >= dblMinLimit Then isInRange = True
    Else
        If dblValue > dblMinLimit Then isInRange = True
    End If

    'Check upper bound
    If isInRange Then
        If includeMaxLimit Then
            isInRange = (dblValue <= dblMaxLimit)
        Else
            isInRange = (dblValue < dblMaxLimit)
        End If
    End If

End Function

Public Function increment(ByRef value As Long, Optional step As Long = 1) As Long
    increment = value
    value = value + step
End Function

Public Function preincrement(ByRef value As Long, Optional step As Long = 1) As Long
    value = value + step
    preincrement = value
End Function

Public Function isInteger(value As Variant) As Boolean
    Const REGEX_PATTERN As String = "^\s*-?\d+\s*$"
    '------------------------------------------------------------------------------------------------------
    If F.Variables.isUndefined(value) Then
    ElseIf VBA.IsArray(value) Then
    ElseIf VBA.IsObject(value) Then
    Else
        isInteger = F.regex.checkIfMatch(VBA.CStr(value), REGEX_PATTERN)
    End If
End Function












Public Function getRandomBetween(downLimit As Double, upLimit As Double) As Double
    Dim rand As Double
    '------------------------------------------------------------------------------------------------------
    Call VBA.Randomize(VBA.Timer)
    rand = VBA.Rnd(VBA.Timer + 1)
    getRandomBetween = (upLimit - downLimit) * rand + downLimit
End Function


Public Function decToBin(ByVal number As Long) As String
    Do
        decToBin = number Mod 2 & decToBin
        number = number \ 2
    Loop While number
End Function

Public Function binaryInclude(baseValue As Long, searchValue As Long) As Boolean
    binaryInclude = ((baseValue And searchValue) = searchValue)
End Function












Public Function isNumeric(value As Variant) As Boolean
    'Const REGEX_PATTERN As String = "^(-?\d+(?:\.|\,)\d+)$|^(-?\d+)$"
    Const REGEX_PATTERN As String = "^\s*-?(\d+[.,]{0,1}|)\d+%?\s*$"
    '------------------------------------------------------------------------------------------------------
    Static objRegex As Object
    '------------------------------------------------------------------------------------------------------
    
    If F.Variables.isUndefined(value) Then
    ElseIf VBA.IsArray(value) Then
    ElseIf VBA.IsObject(value) Then
    Else
        isNumeric = F.regex.checkIfMatch(VBA.CStr(value), REGEX_PATTERN)
    End If
    
End Function

Public Function toNumeric(ByVal value As Variant) As Variant
    If F.Variables.isPrimitive(value) Then
        value = VBA.Replace(value, " ", vbNullString)
        value = VBA.Replace(value, VBA.Chr(160), vbNullString)
        
        If F.Numerics.isNumeric(value) Then
            On Error Resume Next
            toNumeric = VBA.CDbl(VBA.Replace(value, ".", ","))
            toNumeric = VBA.CDbl(VBA.Replace(value, ",", "."))
        End If
        
    End If
End Function

Public Function coalesce(ParamArray values() As Variant) As String
    Dim varNumber As Variant
    '------------------------------------------------------------------------------------------------------
    
    For Each varNumber In values
        If F.Numerics.isNumeric(varNumber) Then
            If varNumber <> 0 Then
                coalesce = F.Numerics.toNumeric(varNumber)
                Exit For
            End If
        End If
    Next varNumber
    
End Function





'
'
'
'
'Public Function isPositive(value As Variant, Optional strictType As Boolean = False) As Boolean
'    Dim variableType As VBA.VbVarType
'    '------------------------------------------------------------------------------------------------------
'    If Not isNumeric(value) Then
'
'    ElseIf VBA.Val(value) <= 0 Then
'
'    ElseIf strictType Then
'        variableType = VBA.varType(value)
'        isPositive = (variableType <> vbString)
'    Else
'        isPositive = True
'    End If
'End Function
'
'Public Function isNotNegative(value As Variant, Optional strictType As Boolean = False) As Boolean
'    Dim variableType As VBA.VbVarType
'    '------------------------------------------------------------------------------------------------------
'    If Not isNumeric(value) Then
'
'    ElseIf VBA.Val(value) >= 0 Then
'        variableType = VBA.varType(value)
'        If variableType = vbBoolean Or variableType = vbEmpty Then
'            isNotNegative = False
'        ElseIf strictType Then
'            variableType = VBA.varType(value)
'            isNotNegative = (variableType <> vbString)
'        Else
'            isNotNegative = True
'        End If
'    End If
'End Function
'
'Public Function isNegative(value As Variant, Optional strictType As Boolean = False) As Boolean
'    Dim variableType As VBA.VbVarType
'    '------------------------------------------------------------------------------------------------------
'    If Not isNumeric(value) Then
'
'    ElseIf VBA.Val(value) < 0 Then
'        variableType = VBA.varType(value)
'        If variableType = vbBoolean Then
'
'        ElseIf strictType Then
'            isNegative = (variableType <> vbString)
'        Else
'            isNegative = True
'        End If
'    End If
'End Function
'
'Public Function isNotPositive(value As Variant, Optional strictType As Boolean = False) As Boolean
'    Dim variableType As VBA.VbVarType
'    '------------------------------------------------------------------------------------------------------
'    If Not isNumeric(value) Then
'
'    ElseIf VBA.Val(value) <= 0 Then
'        variableType = VBA.varType(value)
'        If variableType = vbBoolean Or variableType = vbEmpty Then
'            isNotPositive = False
'        ElseIf strictType Then
'            variableType = VBA.varType(value)
'            isNotPositive = (variableType <> vbString)
'        Else
'            isNotPositive = True
'        End If
'    End If
'End Function
'
'
'
'
'
'
''***** WITHOUT UNIT TESTS *****
'Public Function fromBoolean(value As Boolean) As Integer
'    If value Then fromBoolean = 1
'End Function
'
'
'Public Function getLimitedValue(value As Variant, minValue As Variant, maxValue As Variant) As Variant
'    If value < minValue Then
'        getLimitedValue = minValue
'    ElseIf value > maxValue Then
'        getLimitedValue = maxValue
'    Else
'        getLimitedValue = value
'    End If
'End Function
'
'Public Function isProperPercent(value As Variant) As Boolean
'    Const REGEX_PATTERN As String = "^[-|+]?([0-9]+(?:[\.][0-9]*)?|\.[0-9]+)%$"
'    Static objRegex As Object
'    '------------------------------------------------------------------------------------------------------
'    If F.Variables.isUndefined(value) Then
'    ElseIf VBA.IsArray(value) Then
'    ElseIf VBA.IsObject(value) Then
'    Else
'        If objRegex Is Nothing Then
'            Set objRegex = F.regex.Create(REGEX_PATTERN)
'        End If
'        isProperPercent = objRegex.test(value)
'    End If
'End Function
'
'Public Function fromPercent(value As String) As Double
'    fromPercent = VBA.CDbl(VBA.left$(value, VBA.Len(value) - 1)) / 100
'End Function
