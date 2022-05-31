VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FColors"
'[Windows API functions] ----------------------------------------------------------------------------------

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function ChooseColorDlg Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
#ElseIf VBA7 Then
    Private Declare PtrSafe Function ChooseColorDlg Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
#Else
    Private Declare Function ChooseColorDlg Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
#End If
'[Windows API types] --------------------------------------------------------------------------------------
Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
'[Windows API constants] ----------------------------------------------------------------------------------
Private Const CC_RGBINIT = &H1&
Private Const CC_FULLOPEN = &H2&
Private Const CC_ANYCOLOR = &H100&
'----------------------------------------------------------------------------------------------------------
    
    

Public Function convertCssRgbToLong(ByVal text As String) As Long
    Const REGEX_PATTERN As String = "^rgb\((\d{1,3}),(?: |)(\d{1,3}),(?: |)(\d{1,3})\)$"
    '------------------------------------------------------------------------------------------------------
    Dim colors As Variant
    Dim red As Long
    Dim green As Long
    Dim blue As Long
    '------------------------------------------------------------------------------------------------------
    
    If VBA.isNumeric(text) Then
        convertCssRgbToLong = VBA.CLng(text)
    Else
        colors = F.regex.getMatchArray(text, REGEX_PATTERN)
        If F.arrays.isDefined(colors) Then
            red = colors(LBound(colors))
            green = colors(LBound(colors) + 1)
            blue = colors(LBound(colors) + 2)
            convertCssRgbToLong = VBA.RGB(red, green, blue)
        End If
    End If
    
End Function


Public Function getColorFromScale(value As Double, greenBound As Double, redBound As Double, _
                    Optional bottomScale As Long = 0, Optional topScale As Long = 255) As Long
    Dim middleBound As Double
    Dim posOnScale As Double
    Dim green As Long
    Dim red As Long
    '------------------------------------------------------------------------------------------------------
    
    middleBound = (redBound - greenBound) / 2
    
    If redBound > greenBound Then
        If value >= redBound Then
            green = bottomScale
            red = topScale
        ElseIf value <= greenBound Then
            green = topScale
            red = bottomScale
        Else
            posOnScale = (value - greenBound) / (redBound - greenBound)
            green = VBA.IIf(posOnScale <= 0.5, topScale, bottomScale + (topScale - bottomScale) * (1 - posOnScale) / 0.5)
            red = VBA.IIf(posOnScale >= 0.5, topScale, bottomScale + (topScale - bottomScale) * posOnScale / 0.5)
        End If
    Else
        If value >= greenBound Then
            red = bottomScale
            green = topScale
        ElseIf value <= redBound Then
            red = topScale
            green = bottomScale
        Else
            posOnScale = (value - redBound) / (greenBound - redBound)
            red = VBA.IIf(posOnScale <= 0.5, topScale, bottomScale + (topScale - bottomScale) * (1 - posOnScale) / 0.5)
            green = VBA.IIf(posOnScale >= 0.5, topScale, bottomScale + (topScale - bottomScale) * posOnScale / 0.5)
        End If
    End If
    
    getColorFromScale = VBA.RGB(red, green, bottomScale)
    
End Function


Public Function convertColorLongToRgb(ByVal value As Long) As RgbArray
    With convertColorLongToRgb
        If isValidRgbColor(value) Then
            .red = value And vbRed
            .green = (value And vbGreen) \ &H100&
            .blue = (value And vbBlue) \ &H10000
        Else
            .red = -1
            .green = -1
            .blue = -1
        End If
    End With
End Function

Public Function isValidRgbColor(ByVal value As Long) As Boolean
    If value < vbBlack Then
        isValidRgbColor = False
    ElseIf value > vbWhite Then
        isValidRgbColor = False
    Else
        isValidRgbColor = True
    End If
End Function

Public Function pickColor(Optional ByVal defaultColor As Long) As Long
    Dim lpChoosecolor As CHOOSECOLOR
    Dim dwCustClrs(0 To 15) As Long
    '------------------------------------------------------------------------------------------------------
    
    If Not isValidRgbColor(defaultColor) Then defaultColor = 0
    
    With lpChoosecolor
        .lStructSize = Len(lpChoosecolor)
        .hwndOwner = GetActiveWindow
        .rgbResult = defaultColor
        .lpCustColors = VarPtr(dwCustClrs(0))
        .flags = CC_ANYCOLOR Or CC_RGBINIT Or CC_FULLOPEN
    End With
    
    If ChooseColorDlg(lpChoosecolor) Then
        pickColor = lpChoosecolor.rgbResult
    Else
        pickColor = -1
    End If
End Function
