VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xxOptionLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "xxOptionLine"
'----------------------------------------------------------------------------------------------------------
Private Const DEFAULT_FONT_FAMILY As String = APP_FONT_FAMILY
Private Const DEFAULT_FONT_SIZE As Single = 10
Private Const DEFAULT_FONT_COLOR As Long = vbBlack
Private Const DEFAULT_FONT_BOLD As Boolean = False
Private Const DEFAULT_LOCKED_FONT_COLOR As Long = 11447982
Private Const DEFAULT_ALIGNMENT As Variant = xlLeft
Private Const DEFAULT_SELECTION_BACK_COLOR As Long = 14286831
Private Const DEFAULT_SELECTION_BORDER_COLOR As Long = 40282
'----------------------------------------------------------------------------------------------------------
Private pParent As XOptionsPanel
Private pUuid As String
Private pKey As String
Private pItem As Variant
Private pState As Boolean
Private pIsFixed As Boolean
'----------------------------------------------------------------------------------------------------------
Private WithEvents pBackground As MSForms.label
Attribute pBackground.VB_VarHelpID = -1
Private WithEvents pCheckBox As MSForms.label   'checkbox
Attribute pCheckBox.VB_VarHelpID = -1
Private WithEvents pLabel As MSForms.label
Attribute pLabel.VB_VarHelpID = -1
'[UI properties] ------------------------------------------------------------------------------------------
Private pTop As Single
Private pLeft As Single
Private pFontFamily As String
Private pFontSize As Single
Private pFontColor As Long
Private pFontBold As Boolean
Private pBackgroundColor As Long
Private pBorderColor As Long
Private pLockedFontColor As Long
'----------------------------------------------------------------------------------------------------------
Private pIsRendered As Boolean
Private pEventsDisabled As Boolean
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    pUuid = F.Crypto.createUUID
    pFontFamily = DEFAULT_FONT_FAMILY
    pFontSize = DEFAULT_FONT_SIZE
    pFontColor = DEFAULT_FONT_COLOR
    pFontBold = DEFAULT_FONT_BOLD
    pBackgroundColor = DEFAULT_SELECTION_BACK_COLOR
    pBorderColor = DEFAULT_SELECTION_BORDER_COLOR
    pLockedFontColor = DEFAULT_LOCKED_FONT_COLOR
End Sub



'[SETTERS]
Friend Function setParent(value As XOptionsPanel) As xxOptionLine
    Set setParent = Me
    Set pParent = value
End Function

Friend Function setKey(value As String) As xxOptionLine
    Set setKey = Me
    pKey = value
End Function

Friend Function setItem(value As Variant) As xxOptionLine
    Set setItem = Me
    Call F.Variables.assign(pItem, value)
End Function

Friend Function setState(value As Boolean) As xxOptionLine
    Set setState = Me
    pState = value
    Call updateView
    Call pParent.updateValidation
End Function

Friend Function setTop(value As Single) As xxOptionLine
    Set setTop = Me
    pTop = value
    Call arrangeControls
End Function

Friend Function setLeft(value As Single) As xxOptionLine
    Set setLeft = Me
    pLeft = value
    Call arrangeControls
End Function

Friend Function setAsFixed(value As Boolean) As xxOptionLine
    Set setAsFixed = Me
    pIsFixed = value
    Call updateView
End Function




'[GETTERS]
Private Function getFrame() As MSForms.frame
    Set getFrame = pParent.getItemsFrame
End Function

Friend Function getRight() As Single
    getRight = UI.Forms.getRight(pLabel)
End Function

Friend Function getKey() As String
    getKey = pKey
End Function

Friend Function getItem() As Variant
    Call F.Variables.assign(getItem, pItem)
End Function

Friend Function isFixed() As Boolean
    isFixed = pIsFixed
End Function

Friend Function isSelected() As Boolean
    isSelected = pState
End Function









'[RENDERING]
Friend Sub render()
    If Not pIsRendered Then
        pEventsDisabled = True
        Call insertBackgroundLabel
        Call insertCheckBox
        Call insertLabel
        pIsRendered = True
        Call arrangeControls
        Call updateView
        pEventsDisabled = False
    End If
End Sub

Private Sub insertBackgroundLabel()
    Set pBackground = getFrame.controls.Add(CREATE_LABEL_ID)
    With pBackground
        .backColor = pBackgroundColor
        .BackStyle = fmBackStyleTransparent
        .borderColor = pBorderColor
        .borderStyle = fmBorderStyleNone
        .width = getFrame.InsideWidth
    End With
End Sub

Private Sub insertCheckBox()
    Dim size As Single
    '------------------------------------------------------------------------------------------------------
    
    size = Icons.icoCheckedActive.width
    
    Set pCheckBox = getFrame.controls.Add(CREATE_LABEL_ID)
    With pCheckBox
        .BackStyle = fmBackStyleTransparent
        .caption = vbNullString
        .MousePointer = fmMousePointerCustom
        Set .MouseIcon = Icons.getMousePointer
        .width = size
        .height = size
    End With
    
End Sub

Private Sub insertLabel()
    Set pLabel = getFrame.controls.Add(CREATE_LABEL_ID)
    With pLabel
        .BackStyle = fmBackStyleTransparent
        .caption = pKey
        .ForeColor = pFontColor
        With .font
            .name = pFontFamily
            .size = pFontSize
            .Bold = VBA.IIf(pFontBold, 1, False)
        End With
        .MousePointer = fmMousePointerCustom
        Set .MouseIcon = Icons.getMousePointer
        .width = pParent.getLabelWidth
        .AutoSize = True
        .AutoSize = False
        .width = pParent.getLabelWidth
    End With
End Sub

Private Sub updateView()
    pEventsDisabled = True
    
    If pIsRendered Then
        If pIsFixed Then
            With pCheckBox
                If pState Then
                    Set .picture = Icons.icoCheckedInactive.picture
                Else
                    Set .picture = Icons.icoUncheckedInactive.picture
                End If
                .MousePointer = fmMousePointerDefault
            End With
            With pLabel
                .ForeColor = VBA.IIf(pIsFixed, pLockedFontColor, pFontColor)
                .MousePointer = fmMousePointerDefault
            End With
        Else
            With pCheckBox
                If pState Then
                    Set .picture = Icons.icoCheckedActive.picture
                Else
                    Set .picture = Icons.icoUncheckedActive.picture
                End If
                .MousePointer = fmMousePointerCustom
                Set .MouseIcon = Icons.getMousePointer
            End With
            With pLabel
                .ForeColor = VBA.IIf(pIsFixed, pLockedFontColor, pFontColor)
                .MousePointer = fmMousePointerCustom
                Set .MouseIcon = Icons.getMousePointer
            End With
        End If
        
        With pParent.getItemsFrame.controls(pLabel.name)
            .font.Bold = pState
        End With
        
    End If
    
    
    
ExitPoint:
    pEventsDisabled = False
End Sub

Friend Sub Move(left As Single, top As Single)
    pLeft = left
    pTop = top
    Call arrangeControls
End Sub

Private Sub arrangeControls()
    Dim lineHeight As Single:               lineHeight = pParent.getLineHeight
    '------------------------------------------------------------------------------------------------------
    
    If pIsRendered Then
        Call UI.Forms.locateControl(pBackground, pLeft, pTop)
        Call UI.Forms.locateControl(pCheckBox, pLeft, pTop + (lineHeight - pCheckBox.height) / 2)
        Call UI.Forms.locateControl(pLabel, _
                            UI.Forms.getRight(pCheckBox) + pParent.getCheckboxValueSpace, _
                            pTop + (lineHeight - pLabel.height) / 2)
    End If
    
End Sub



'[EVENTS]
Private Sub pCheckBox_MouseUp(ByVal Button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal Y As Single)
    Call ErrorManager.clear
    Call handleClick(Button, shift)
End Sub

Private Sub pLabel_MouseUp(ByVal Button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal Y As Single)
    Call ErrorManager.clear
    Call handleClick(Button, shift)
End Sub

Private Sub handleClick(ByVal Button As Integer, ByVal shift As Integer)
    If Not pEventsDisabled And Not pIsFixed Then
        If shift = 1 Or shift = 2 Then
            Call pParent.changeSelectionForAllItems(Not pState)
        Else
            Call setState(Not pState)
        End If
    End If
End Sub
