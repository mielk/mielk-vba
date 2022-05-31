Attribute VB_Name = "modStyles"
Option Explicit

Private Const CLASS_NAME As String = "modStyles"
    '----------------------------------------------------------------------------------------------------------


'[ALIGNMENTS]
Public Function convertAlignTextToEnum(value As String) As Variant
    Select Case VBA.LCase$(value)
        Case "center":  convertAlignTextToEnum = xlCenter
        Case "left":    convertAlignTextToEnum = xlLeft
        Case "right":   convertAlignTextToEnum = xlRight
        Case "top":     convertAlignTextToEnum = xlTop
        Case "bottom":  convertAlignTextToEnum = xlBottom
        Case Else:      convertAlignTextToEnum = xlCenter
    End Select
End Function



'[RANGE BORDERS]
Public Function convertBorderIndexNameToEnum(value As String) As XlBordersIndex
    Select Case VBA.LCase$(value)
        Case "left":                convertBorderIndexNameToEnum = xlEdgeLeft
        Case "right":               convertBorderIndexNameToEnum = xlEdgeRight
        Case "top":                 convertBorderIndexNameToEnum = xlEdgeTop
        Case "bottom":              convertBorderIndexNameToEnum = xlEdgeBottom
        Case "inside-vertical":     convertBorderIndexNameToEnum = xlInsideVertical
        Case "inside-horizontal":   convertBorderIndexNameToEnum = xlInsideHorizontal
    End Select
End Function

Public Function convertBorderIndexEnumToName(value As XlBordersIndex) As String
    Select Case VBA.LCase$(value)
        Case xlEdgeLeft:            convertBorderIndexEnumToName = "left"
        Case xlEdgeRight:           convertBorderIndexEnumToName = "right"
        Case xlEdgeTop:             convertBorderIndexEnumToName = "top"
        Case xlEdgeBottom:          convertBorderIndexEnumToName = "bottom"
        Case xlInsideVertical:      convertBorderIndexEnumToName = "inside-vertical"
        Case xlInsideHorizontal:    convertBorderIndexEnumToName = "inside-horizontal"
    End Select
End Function

Public Function convertBorderStyleNameToEnum(value As String) As XlLineStyle
    Select Case VBA.LCase$(value)
        Case "continuous":          convertBorderStyleNameToEnum = xlContinuous
        Case "none":                convertBorderStyleNameToEnum = xlLineStyleNone
        Case Else:                  convertBorderStyleNameToEnum = xlContinuous
    End Select
End Function

Public Function convertBorderWeightNameToEnum(value As String) As XlBorderWeight
    Select Case VBA.LCase$(value)
        Case "thin":                convertBorderWeightNameToEnum = xlThin
        Case "medium":              convertBorderWeightNameToEnum = xlMedium
        Case "thick":               convertBorderWeightNameToEnum = xlThick
        Case Else:                  convertBorderWeightNameToEnum = xlThin
    End Select
End Function

Public Function isInsideBorder(borderIndex As XlBordersIndex)
    If borderIndex = xlInsideHorizontal Then
        isInsideBorder = True
    ElseIf borderIndex = xlInsideVertical Then
        isInsideBorder = True
    End If
End Function

Public Function isOutsideBorder(borderIndex As XlBordersIndex)
    isOutsideBorder = Not isInsideBorder(borderIndex)
End Function

Public Function convertRgbToLong(ByVal text As String) As Long
    convertRgbToLong = F.colors.convertCssRgbToLong(text)
End Function




'[FORMAT CONDITIONS]
Public Function convertFormatConditionTypeToEnum(value As String) As XlFormatConditionType
    Select Case VBA.LCase$(value)
        Case "expression", "formula":
                                    convertFormatConditionTypeToEnum = xlExpression
        Case "cellValue":           convertFormatConditionTypeToEnum = xlCellValue
        Case Else:                  convertFormatConditionTypeToEnum = xlExpression
    End Select
End Function

Public Function convertFormatConditionOperatorToEnum(value As String) As XlFormatConditionOperator
    Select Case VBA.LCase$(value)
        Case "greater":             convertFormatConditionOperatorToEnum = xlGreater
        Case "equal":               convertFormatConditionOperatorToEnum = xlEqual
        Case "less":                convertFormatConditionOperatorToEnum = xlLess
        Case Else:                  convertFormatConditionOperatorToEnum = xlEqual
    End Select
End Function
