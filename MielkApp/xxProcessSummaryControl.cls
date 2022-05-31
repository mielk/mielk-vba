VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xxProcessSummaryControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum ControlTypeEnum
    ControlType_Background = 1
    ControlType_StatusIcon = 2
    ControlType_NameLabel = 3
End Enum

Private Const CLASS_NAME As String = "xxProcessSummaryControl"
'[UI constants] -------------------------------------------------------------------------------------------
Private Const STATUS_ICON_LEFT_OFFSET As Single = 4.5
Private Const STATUS_ICON_TOP_OFFSET As Single = 4.5
Private Const NAME_LABEL_TO_ICON_OFFSET As Single = 3
Private Const DEFAULT_FONT_FAMILY As String = "Segoe UI" '"Century Gothic"
Private Const DEFAULT_FONT_SIZE As Long = 9
Private Const DEFAULT_LEFT_INDENT As Single = 0
Private Const DEFAULT_LINE_HEIGHT As Single = 18
Private Const DEFAULT_BORDER_COLOR As Long = 6579300
Private Const DEFAULT_ACTIVE_BACK_COLOR As Long = 12052479
Private Const DEFAULT_ACTIVE_BORDER_COLOR As Long = 1684223
'[UI components] ------------------------------------------------------------------------------------------
Private pParent As XProcessSummaryPanel
Private pFrame As MSForms.frame
Private WithEvents pBackgroundLabel As MSForms.Label
Private WithEvents pStatusIcon As MSForms.Label
Private WithEvents pNameLabel As MSForms.Label
'[UI variables] -------------------------------------------------------------------------------------------
Private pActiveItemBackColor As Long
Private pActiveItemBorderColor As Long
Private pTop As Single
Private pLeft As Single
Private pLineHeight As Single
Private pFontFamily As String
Private pFontSize As Single
'[State] --------------------------------------------------------------------------------------------------
Private pKey As Variant
Private pName As String
Private pColor As Long
Private pBorderColor As Long
Private pIsRendered As Boolean
Private pIsActive As Boolean
'----------------------------------------------------------------------------------------------------------




'[INITIALIZE]
Private Sub Class_Initialize()
    Call initializeDefaultValues
End Sub

Private Sub initializeDefaultValues()
    pLeft = DEFAULT_LEFT_INDENT
    pLineHeight = DEFAULT_LINE_HEIGHT
    pFontFamily = DEFAULT_FONT_FAMILY
    pFontSize = DEFAULT_FONT_SIZE
    pActiveItemBackColor = DEFAULT_ACTIVE_BACK_COLOR
    pActiveItemBorderColor = DEFAULT_ACTIVE_BORDER_COLOR
End Sub



'[SETTERS]
Public Function setParent(value As XProcessSummaryPanel) As xxProcessSummaryControl
    Set setParent = Me
    Set pParent = value
    Set pFrame = pParent.getContainerFrame
End Function

Public Function setKey(value As Variant) As xxProcessSummaryControl
    Set setKey = Me
    pKey = value
End Function

Public Function loadDto(dto As Scripting.Dictionary) As xxProcessSummaryControl
    Set loadDto = Me
    Call F.Dictionaries.assignValueFromDictToVariable(dto, pName, TAG_NAME)
    Call F.Dictionaries.assignValueFromDictToVariable(dto, pColor, TAG_MAIN_BACKGROUND)
    Call F.Dictionaries.assignValueFromDictToVariable(dto, pIsActive, TAG_IS_ACTIVE)
End Function

Public Function setIsActive(value As Boolean) As xxProcessSummaryControl
    Set setIsActive = Me
    pIsActive = value
    If pIsRendered Then Call updateActiveness
End Function


'[Font]
Public Function setFontSize(value As Single) As xxProcessSummaryControl
    Set setFontSize = Me
    pFontSize = value
    If pIsRendered Then Call refreshNameLabel
End Function

Public Function setFontFamily(value As String) As xxProcessSummaryControl
    Set setFontFamily = Me
    pFontFamily = value
    If pIsRendered Then Call refreshNameLabel
End Function


'[Position]
Public Function setTop(value As Single) As xxProcessSummaryControl
    Set setTop = Me
    pTop = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setBottom(value As Single) As xxProcessSummaryControl
    Set setBottom = Me
    pTop = value - pLineHeight
    If pIsRendered Then Call arrangeControls
End Function

Public Function setLeft(value As Single) As xxProcessSummaryControl
    Set setLeft = Me
    pLeft = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setRight(value As Single) As xxProcessSummaryControl
    Set setRight = Me
    pLeft = value - getWidth
    If pIsRendered Then Call arrangeControls
End Function




'[Size]
Public Function setWidth(value As Single) As xxProcessSummaryControl
    Set setWidth = Me
End Function

Public Function setHeight(value As Single) As xxProcessSummaryControl
    Set setHeight = Me
    pLineHeight = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setLineHeight(value As Single) As xxProcessSummaryControl
    Set setLineHeight = setHeight(value)
End Function



'[GETTERS]
Public Function getName() As String
    getName = pName
End Function

Public Function getKey() As Variant
    getKey = pKey
End Function

Public Function getColor() As Long
    getColor = pColor
End Function

Public Function isActive() As Boolean
    isActive = pIsActive
End Function


'[State]
Public Function isValid() As Boolean
    isValid = UI.Forms.isControlValid(pBackgroundLabel)
End Function


'[Position]
Public Function getLeft() As Single
    getLeft = pBackgroundLabel.left
End Function

Public Function getRight() As Single
    getRight = UI.Forms.getRight(pBackgroundLabel)
End Function

Public Function getTop() As Single
    getTop = pBackgroundLabel.top
End Function

Public Function getBottom() As Single
    getBottom = UI.Forms.getBottom(pBackgroundLabel)
End Function


'[Size]
Public Function getWidth() As Single
    getWidth = pParent.getWidth
End Function

Public Function getHeight() As Single
    getHeight = pBackgroundLabel.height
End Function




'[API]
Public Function render() As xxProcessSummaryControl
    Set render = Me
    If Not pIsRendered Then
        Call insertBackgroundLabel
        Call insertStatusIcon
        Call insertNameLabel
        Call updateActiveness
        Call arrangeControls
        pIsRendered = True
    End If
End Function

Public Sub move(Optional left As Variant, Optional top As Variant)
    If F.Variables.isOfNumericType(left) Then pLeft = left
    If F.Variables.isOfNumericType(top) Then pTop = top
    Call arrangeControls
End Sub

Public Sub resize(Optional width As Variant, Optional height As Variant)
    If height >= 0 Then pLineHeight = height
    Call arrangeControls
End Sub



'[RENDERING]
Private Function getLabelName(controlType As ControlTypeEnum) As String
    Dim tag As String
    '----------------------------------------------------------------------------------------------------------
    tag = CLASS_NAME & "|" & pName & "|" & controlType
    getLabelName = "ctrl_" & VBA.left$(F.Crypto.SHA1HASH(tag), 25)
End Function


'[Inserting controls]
Private Sub insertBackgroundLabel()
    Set pBackgroundLabel = pFrame.Controls.Add(CREATE_LABEL_ID, getLabelName(ControlType_Background), True)
    Call refreshBackgroundLabel
End Sub

Private Sub insertStatusIcon()
    Set pStatusIcon = pFrame.Controls.Add(CREATE_LABEL_ID, getLabelName(ControlType_StatusIcon), True)
    Call refreshStatusIcon
End Sub

Private Sub insertNameLabel()
    Set pNameLabel = pFrame.Controls.Add(CREATE_LABEL_ID, getLabelName(ControlType_NameLabel), True)
    Call refreshNameLabel
End Sub



'[Refreshing controls]
Private Sub refreshBackgroundLabel()
    With pBackgroundLabel
        .BackStyle = fmBackStyleTransparent
        .BackColor = pActiveItemBackColor
        .BorderStyle = fmBorderStyleNone
        .BorderColor = pActiveItemBorderColor
        .MousePointer = fmMousePointerCustom
        Set .MouseIcon = UI.Icons.getMousePointer
    End With
End Sub

Private Sub refreshStatusIcon()
    With pStatusIcon
        .BackStyle = fmBackStyleOpaque
        .BackColor = pColor
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = pBorderColor
        .MousePointer = fmMousePointerCustom
        Set .MouseIcon = UI.Icons.getMousePointer
    End With
End Sub

Private Sub refreshNameLabel()
    With pNameLabel
        .BackStyle = fmBackStyleTransparent
        .Font.name = pFontFamily
        .Font.size = UI.Normalize.size(pFontSize)
        .TextAlign = fmTextAlignLeft
        .caption = getCaption
        .MousePointer = fmMousePointerCustom
        Set .MouseIcon = UI.Icons.getMousePointer
    End With
End Sub



'[Size & location]
Private Sub arrangeControls()
    Dim left As Single
    Dim top As Single
    Dim buttons As VBA.Collection
    Dim button As MSForms.Label
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    Call resizeControls
    Call UI.Forms.locateControl(pBackgroundLabel, pLeft, pTop)
    Call UI.Forms.locateControl(pStatusIcon, pLeft + STATUS_ICON_LEFT_OFFSET, pTop + STATUS_ICON_TOP_OFFSET + 0 * PIXEL_SIZE)
    Call UI.Forms.locateControl(pNameLabel, UI.Forms.getRight(pStatusIcon) + NAME_LABEL_TO_ICON_OFFSET, _
                                        pTop + (pLineHeight - pNameLabel.height) / 2)

End Sub

Private Sub resizeControls()
    Call UI.Forms.resizeControl(pBackgroundLabel, getWidth, pLineHeight)
    Call resizeColorLabel
    Call resizeNameLabel
End Sub

Private Sub resizeColorLabel()
    Dim size As Single
    '----------------------------------------------------------------------------------------------------------
    size = pLineHeight - 2 * STATUS_ICON_TOP_OFFSET - 2 * PIXEL_SIZE
    Call UI.Forms.resizeControl(pStatusIcon, size, size)
End Sub

Private Sub resizeNameLabel()
    Dim caption As String
    '----------------------------------------------------------------------------------------------------------
    
    With pNameLabel
        caption = .caption
        .caption = "tj"
        .AutoSize = True
        .AutoSize = False
        .width = getWidth - 1 = 2 * STATUS_ICON_LEFT_OFFSET - NAME_LABEL_TO_ICON_OFFSET
        .caption = getCaption
    End With
    
End Sub

Private Function getCaption() As String
    getCaption = pName & " (" & pParent.getProcessor.getStatusCounter(pKey) & ")"
End Function

Private Sub updateActiveness()
    pBackgroundLabel.BackStyle = VBA.IIf(pIsActive, fmBackStyleOpaque, fmBackStyleTransparent)
    pBackgroundLabel.BorderStyle = VBA.IIf(pIsActive, fmBorderStyleSingle, fmBorderStyleNone)
    pNameLabel.Font.Bold = pIsActive
End Sub

Public Sub updateCaption()
    pNameLabel.caption = getCaption
End Sub



'[Clearing]
Public Sub clear()
    On Error Resume Next
    With pFrame.Controls
        Call .Remove(getLabelName(ControlType_Background))
        Call .Remove(getLabelName(ControlType_StatusIcon))
        Call .Remove(getLabelName(ControlType_NameLabel))
    End With
End Sub





'[Events]

Private Sub pBackgroundLabel_Click()
    Call ErrorManager.clear
    Call pParent.triggerControlClickedEvent(Me)
End Sub

Private Sub pNameLabel_Click()
    Call ErrorManager.clear
    Call pParent.triggerControlClickedEvent(Me)
End Sub

Private Sub pStatusIcon_Click()
    Call ErrorManager.clear
    Call pParent.triggerControlClickedEvent(Me)
End Sub
