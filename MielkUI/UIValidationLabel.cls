VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIValidationLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "UIValidationLabel"
'[Default UI properties] ----------------------------------------------------------------------------------
Private Const DEFAULT_ERROR_BACK_COLOR As Long = vbRed
Private Const DEFAULT_WARNING_BACK_COLOR As Long = 4244467       'rgb(243, 195, 64)
Private Const DEFAULT_FONT_COLOR As Long = vbWhite
Private Const DEFAULT_FONT_NAME As String = APP_FONT_FAMILY
Private Const DEFAULT_FONT_SIZE As Single = 10
Private Const DEFAULT_FONT_BOLD As Boolean = True
Private Const DEFAULT_PADDING As Single = 3
'----------------------------------------------------------------------------------------------------------
Private pFrame As MSForms.frame
Private pBackgroundLabel As MSForms.label
Attribute pBackgroundLabel.VB_VarHelpID = -1
Private pCaptionLabel As MSForms.label
Attribute pCaptionLabel.VB_VarHelpID = -1
'----------------------------------------------------------------------------------------------------------
Private pCaption As String
Private pErrorBackColor As Long
Private pWarningBackColor As Long
Private pFontName As String
Private pFontColor As Long
Private pFontSize As Single
Private pFontBold As Boolean
'----------------------------------------------------------------------------------------------------------
Private pPadding As Single
Private pLeft As Single
Private pTop As Single
Private pHeight As Single
Private pWidth As Single
Private pVisible As Boolean
'----------------------------------------------------------------------------------------------------------
Private pState As Variant
Private pIsRendered As Boolean
'----------------------------------------------------------------------------------------------------------




Private Sub Class_Initialize()
    pErrorBackColor = DEFAULT_ERROR_BACK_COLOR
    pWarningBackColor = DEFAULT_WARNING_BACK_COLOR
    pFontColor = DEFAULT_FONT_COLOR
    pFontBold = DEFAULT_FONT_BOLD
    pFontSize = DEFAULT_FONT_SIZE
    pFontName = DEFAULT_FONT_NAME
    pPadding = DEFAULT_PADDING
    pState = False
End Sub




'[SETTERS]
Public Function setFrame(value As MSForms.frame) As UIValidationLabel
    Set setFrame = Me
    Set pFrame = value
End Function

Public Function setCaption(value As String) As UIValidationLabel
    Set setCaption = Me
    pCaption = value
    Call updateView
    Call updateSize
End Function

Public Function setVisible(value As Boolean) As UIValidationLabel
    Set setVisible = Me
    pVisible = value
    Call updateVisibility
End Function

Public Function setState(value As Variant) As UIValidationLabel
    'Introduced with validation WARNING status.
    pState = value
    Call updateView
End Function




Public Function setErrorBackColor(value As Long) As UIValidationLabel
    Set setErrorBackColor = Me
    pErrorBackColor = value
    Call updateView
End Function

Public Function setWarningBackColor(value As Long) As UIValidationLabel
    Set setWarningBackColor = Me
    pWarningBackColor = value
    Call updateView
End Function

Public Function setFontColor(value As Long) As UIValidationLabel
    Set setFontColor = Me
    pFontColor = value
    Call updateView
End Function

Public Function setFontName(value As String) As UIValidationLabel
    Set setFontName = Me
    pFontName = value
    Call updateView
    Call updateSize
End Function

Public Function setFontSize(value As Single) As UIValidationLabel
    Set setFontSize = Me
    pFontSize = value
    Call updateView
    Call updateSize
End Function

Public Function setFontBold(value As Single) As UIValidationLabel
    Set setFontBold = Me
    pFontBold = value
    Call updateView
    Call updateSize
End Function



Public Function setLeft(value As Single) As UIValidationLabel
    Set setLeft = Me
    pLeft = value
    Call updatePosition
End Function

Public Function setTop(value As Single) As UIValidationLabel
    Set setTop = Me
    pTop = value
    Call updatePosition
End Function

Public Function setHeight(value As Single) As UIValidationLabel
    Set setHeight = Me
    pHeight = value
    Call updateSize
End Function

Public Function setWidth(value As Single) As UIValidationLabel
    Set setWidth = Me
    pWidth = value
    Call updateSize
End Function


Public Function setSize(width As Single, height As Single) As UIValidationLabel
    Set setSize = Me
    pWidth = width
    pHeight = height
    Call updateSize
End Function

Public Function setPosition(left As Single, top As Single) As UIValidationLabel
    Set setPosition = Me
    pLeft = left
    pTop = top
    Call updatePosition
End Function











'[GETTERS]

Public Function getLeft() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getLeft = pBackgroundLabel.left
    End If
End Function

Public Function getRight() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getRight = pBackgroundLabel.left + pBackgroundLabel.width
    End If
End Function

Public Function getTop() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getTop = pBackgroundLabel.top
    End If
End Function

Public Function getBottom() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getBottom = pBackgroundLabel.top + pBackgroundLabel.height
    End If
End Function

Public Function getHeight() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getHeight = pBackgroundLabel.height
    End If
End Function

Public Function getWidth() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getWidth = pBackgroundLabel.width
    End If
End Function

Private Function getBackColorByState() As Long
    If pState = WARNING_CODE Then
        getBackColorByState = pWarningBackColor
    Else
        getBackColorByState = pErrorBackColor
    End If
End Function








'[RENDERING]
Public Function render() As UIValidationLabel
    Set render = Me
    If Not pIsRendered Then
        If UI.Forms.isControlValid(pFrame) Then
            Call insertErrorBackground
            Call insertErrorCaptionLabel
            Call updateView
            Call updateSize
            Call updatePosition
        End If
        pIsRendered = True
    End If
End Function

Private Sub insertErrorBackground()
    Set pBackgroundLabel = pFrame.controls.Add(CREATE_LABEL_ID, visible:=False)
    With pBackgroundLabel
        .BackStyle = fmBackStyleOpaque
        .backColor = pErrorBackColor
        .height = pHeight
        .width = pFrame.InsideWidth
        .left = 0
    End With
End Sub

Private Sub insertErrorCaptionLabel()
    Set pCaptionLabel = pFrame.controls.Add(CREATE_LABEL_ID, visible:=False)
    With pCaptionLabel
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .caption = pCaption
        .ForeColor = pFontColor
        With .font
            .name = pFontName
            .size = pFontSize
            .Bold = VBA.IIf(pFontBold, 1, False)
        End With
        .height = pHeight
        .width = pBackgroundLabel.width - 2 * pPadding
        .left = pPadding
    End With
End Sub




Private Sub updateView()
    With pBackgroundLabel
        .BackStyle = fmBackStyleOpaque
        .backColor = getBackColorByState
    End With
    
    With pCaptionLabel
        .BackStyle = fmBackStyleTransparent
        .caption = pCaption
        .ForeColor = pFontColor
        With .font
            .Bold = VBA.IIf(pFontBold, 1, False)
            .size = pFontSize
            .name = pFontName
        End With
    End With
    
End Sub

Private Sub updatePosition()
    With pBackgroundLabel
        .left = pLeft
        .top = pTop
    End With
    
    With pCaptionLabel
        .left = pLeft + pPadding
        .top = pTop + (pBackgroundLabel.height - .height) / 2
    End With
End Sub

Private Sub updateSize()
    Const TEMP_WIDTH_FOR_AUTO_SIZE As Single = 1000
    '------------------------------------------------------------------------------------------------------
    Dim width As Single
    '------------------------------------------------------------------------------------------------------
    
    With pCaptionLabel
        .width = VBA.IIf(pWidth > 0, pWidth, TEMP_WIDTH_FOR_AUTO_SIZE)
        .height = pHeight
        .AutoSize = True
        .AutoSize = False
        If .height < 0 Then .height = 0
        .BackStyle = fmBackStyleTransparent
    End With
    
    With pBackgroundLabel
        .height = pHeight
        If pWidth Then
            .width = pWidth
        Else
            .width = pCaptionLabel.width + 2 * pPadding
        End If
    End With
    
    Call updatePosition
    
End Sub

Private Sub updateVisibility()
    pBackgroundLabel.visible = pVisible
    pCaptionLabel.visible = pVisible
End Sub

Public Sub delete()
    On Error Resume Next
    With pFrame.controls
        Call .remove(pBackgroundLabel.name)
        Call .remove(pCaptionLabel.name)
    End With
End Sub

