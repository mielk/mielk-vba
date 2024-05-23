VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "WProgressBar"
'[Caption tags] -------------------------------------------------------------------------------------------
Private Const TAG_INITIALIZATION_PROCESS As String = "ProgressBarInitializingTag"
'[UI default properties] ----------------------------------------------------------------------------------
Private Const DEFAULT_FORM_BORDER_COLOR As Long = 3289650
Private Const DEFAULT_PADDING As Single = 6
Private Const DEFAULT_BAR_BORDER_COLOR As Long = MIELK_COLOR_LIGHT
Private Const DEFAULT_BAR_BACK_COLOR As Long = vbWhite
Private Const DEFAULT_BAR_INSIDE_COLOR As Long = MIELK_COLOR_LIGHT
Private Const DEFAULT_BAR_FONT_COLOR As Long = 3289650
Private Const DEFAULT_BAR_FONT_SIZE As Single = 12
Private Const DEFAULT_BAR_HEIGHT As Single = 24
Private Const DEFAULT_BAR_WIDTH As Single = 200
Private Const DEFAULT_BAR_CAPTION_FORMAT As String = "#,##0%"
Private Const DEFAULT_CAPTION_FONT_COLOR As Long = 3289650
Private Const DEFAULT_CAPTION_FONT_SIZE As Single = 11
Private Const FORM_RIGHT_PADDING As Single = 7.5
'[UI components] ------------------------------------------------------------------------------------------
Private pForm As WTemplate
Private pFrame As MSForms.frame
Private pCaptionLabel As MSForms.label
Private pProgressBorderLabel As MSForms.label
Private pProgressBarFillLabel As MSForms.label
Private pProgressValueLabel As MSForms.label
'[UI properties] ------------------------------------------------------------------------------------------
Private pFormBorderColor As Long
Private pPadding As Single
Private pBarBorderColor As Long
Private pBarBackColor As Long
Private pBarInsideColor As Long
Private pBarFontColor As Long
Private pBarFontSize As Single
Private pBarHeight As Single
Private pBarWidth As Single
Private pBarCaptionFormat As String
Private pCaptionFontColor As Long
Private pCaptionFontSize As Single
'[State] --------------------------------------------------------------------------------------------------
Private pProgress As Double
Private pCaption As String
Private pIsRendered As Boolean
Private pCaptionLocked As Boolean
'[Actions data] -------------------------------------------------------------------------------------------
Private pMethodInvoker As Object
Private pMethodName As String
Private pMethodParams As Scripting.Dictionary
Private pMethodErrorMessage As String
'----------------------------------------------------------------------------------------------------------
Private pLevels As VBA.Collection
'----------------------------------------------------------------------------------------------------------




Private Sub Class_Initialize()
    '[Set default properties]
    pFormBorderColor = DEFAULT_FORM_BORDER_COLOR
    pPadding = DEFAULT_PADDING
    pBarBorderColor = DEFAULT_BAR_BORDER_COLOR
    pBarBackColor = DEFAULT_BAR_BACK_COLOR
    pBarInsideColor = DEFAULT_BAR_INSIDE_COLOR
    pBarFontColor = DEFAULT_BAR_FONT_COLOR
    pBarFontSize = DEFAULT_BAR_FONT_SIZE
    pBarHeight = DEFAULT_BAR_HEIGHT
    pBarWidth = DEFAULT_BAR_WIDTH
    pBarCaptionFormat = DEFAULT_BAR_CAPTION_FORMAT
    pCaptionFontColor = DEFAULT_CAPTION_FONT_COLOR
    pCaptionFontSize = DEFAULT_CAPTION_FONT_SIZE
    
    '[Containers]
    Set pLevels = New VBA.Collection
End Sub


Public Function reset() As WProgressBar
    Set reset = Me
    Set pLevels = New VBA.Collection
    pCaptionLocked = False
    Call setProgress(0)
    Call setCaption(MsgService.getText(TAG_INITIALIZATION_PROCESS))
    
End Function




'[SETTERS]
Public Function setWidth(value As Single) As WProgressBar
    Set setWidth = Me
    pBarWidth = value
    If pIsRendered Then Call arrangeLayout
End Function

Public Function setProgress(value As Double) As WProgressBar
    
    Set setProgress = Me
    pProgress = calculateRealProgress(value)
    
    Call updateProgress
    
    DoEvents
End Function

Public Function setProgressAndMessage(progress As Double, caption As String) As WProgressBar
    Set setProgressAndMessage = Me
    Call setCaption(caption)
    Call setProgress(progress)
End Function

Public Function setCaption(value As String) As WProgressBar
    Set setCaption = Me
    If Not pCaptionLocked Then
        pCaption = value
        Call updateCaption
    End If
End Function

Public Function setCaptionLocked(value As Boolean) As WProgressBar
    Set setCaptionLocked = Me
    pCaptionLocked = value
End Function

Public Function setAction(methodName As String, Optional methodInvoker As Object, _
                            Optional methodParams As Scripting.Dictionary, _
                            Optional methodErrorMessage As String) As WProgressBar
    Set setAction = Me
    
    pMethodName = methodName
    Set pMethodInvoker = methodInvoker
    Set pMethodParams = methodParams
    pMethodErrorMessage = methodErrorMessage
    
    If Not pForm Is Nothing Then
        Call pForm.getForm.setUnderlyingMethod(methodName, methodInvoker, methodParams, True, methodErrorMessage)
    End If
    
End Function



'[GETTERS]
Public Function getProgress() As Double
    getProgress = pProgress
End Function

Public Function isVisible() As Boolean
    If Not pForm Is Nothing Then
        isVisible = pForm.isVisible
    End If
End Function




'[RENDERING]
Public Sub run()
    If Not pIsRendered Then
        Call createForm
        Call render
        Call pForm.arrangeComponents
        Call pForm.locateRelativeTo(UI.Create.LocationPropsForScreenCenter)
    End If
    Call pForm.display(vbModal, False)
End Sub

Public Sub hide()
    If Not pForm Is Nothing Then
        If pForm.isVisible Then
            Call pForm.hide
        End If
    End If
End Sub

Public Sub show()
    If Not pForm Is Nothing Then
        Call pForm.display(vbModal, False)
    End If
End Sub

Private Sub createForm()
    If pForm Is Nothing Then
        Set pForm = UI.Create.WindowTemplate
        With pForm
            Call .setHasTitleBar(False)
            Call .setHasButtonsPanel(False)
            Call .setContentFrameHorizontalMargin(FORM_RIGHT_PADDING)
            Call .setBorderColor(pFormBorderColor)
            Call .render
            Call .getForm.setUnderlyingMethod(pMethodName, pMethodInvoker, pMethodParams, True, pMethodErrorMessage)
        End With
    End If
End Sub

Private Sub render()
    Set pFrame = pForm.getContentFrame
    Call insertCaptionLabel
    Call insertProgressBarControls
    Call arrangeLayout
    pIsRendered = True
End Sub



Private Sub insertCaptionLabel()
    If pCaptionLabel Is Nothing Then
        Set pCaptionLabel = pFrame.controls.Add(CREATE_LABEL_ID)
    End If
    With pCaptionLabel
        .BackStyle = fmBackStyleTransparent
        With .font
            .size = pCaptionFontSize
            .name = APP_FONT_FAMILY
            .Bold = False
        End With
        .caption = pCaption
    End With
End Sub

Private Sub insertProgressBarControls()

    '[Progress border]
    If pProgressBorderLabel Is Nothing Then
        Set pProgressBorderLabel = pFrame.controls.Add(CREATE_LABEL_ID)
    End If
    With pProgressBorderLabel
        .BackStyle = fmBackStyleOpaque
        .backColor = pBarBackColor
        .borderStyle = fmBorderStyleSingle
        .borderColor = pBarBorderColor
        .height = pBarHeight
        .width = pBarWidth
    End With
    
    '[Progress bar fill]
    If pProgressBarFillLabel Is Nothing Then
        Set pProgressBarFillLabel = pFrame.controls.Add(CREATE_LABEL_ID)
    End If
    With pProgressBarFillLabel
        .BackStyle = fmBackStyleOpaque
        .backColor = pBarInsideColor
        .borderStyle = fmBorderStyleNone
        .height = pBarHeight - 2 * PIXEL_SIZE
    End With
    
    '[Progress value label]
    If pProgressValueLabel Is Nothing Then
        Set pProgressValueLabel = pFrame.controls.Add(CREATE_LABEL_ID)
    End If
    With pProgressValueLabel
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .font.name = APP_FONT_FAMILY
        .font.size = pBarFontSize
        .ForeColor = pBarFontColor
        .caption = getProgressValueText
        .width = pBarWidth
        .AutoSize = True
        .AutoSize = False
        .width = pBarWidth
        .TextAlign = fmTextAlignCenter
    End With
    
End Sub

Private Sub arrangeLayout()
    With pCaptionLabel
        .width = pBarWidth
        .AutoSize = True
        .AutoSize = False
        .width = pBarWidth
        .left = pPadding
        .top = pPadding
    End With
    
    With pProgressBorderLabel
        .width = pBarWidth
        .top = UI.Forms.getBottom(pCaptionLabel) + pPadding
        .left = pPadding
    End With
    
    With pProgressBarFillLabel
        .width = pProgress * pProgressBorderLabel.width
        .left = pProgressBorderLabel.left + PIXEL_SIZE
        .top = pProgressBorderLabel.top
        .height = pProgressBorderLabel.height
    End With
    
    With pProgressValueLabel
        .width = pBarWidth
        .top = pProgressBorderLabel.top + (pProgressBorderLabel.height - .height) / 2
        .left = pPadding
    End With
    
    Call pForm.arrangeComponents
    
End Sub

Private Sub updateCaption()
    If pIsRendered Then
        pCaptionLabel.caption = pCaption
        Call arrangeLayout
    End If
End Sub

Private Sub updateProgress()
    If pIsRendered Then
        pProgressBarFillLabel.width = F.Numerics.roundTo(pProgress * pProgressBorderLabel.width, PIXEL_SIZE)
        pProgressValueLabel.caption = getProgressValueText
    End If
End Sub

Private Function getProgressValueText() As String
    getProgressValueText = VBA.format(pProgress, pBarCaptionFormat)
End Function




'[LEVELS]
Public Sub openSubProcess(ByVal initialProgress As Double, ByVal finalProgress As Double)
    Dim currentLevel As Limit
    Dim limitRange As Double
    Dim realStart As Double
    Dim realEnd As Double
    '------------------------------------------------------------------------------------------------------
    
    If F.Collections.isNonEmptyCollection(pLevels) Then
        Set currentLevel = pLevels.item(pLevels.Count)
    Else
        Set currentLevel = F.Create.Limit(0, 1)
    End If
    
    With currentLevel
        limitRange = VBA.CDbl(.getRange)
        realStart = .getLow + initialProgress * limitRange
        realEnd = .getLow + finalProgress * limitRange
    End With
    Call pLevels.Add(F.Create.Limit(realStart, realEnd))
    
    '[Update progress]
    pProgress = realStart
    Call updateProgress
    
End Sub


Public Sub closeCurrentSubProcess(Optional ByRef progress As Single)
    Dim currentProcess As Limit
    '------------------------------------------------------------------------------------------------------
    
    If F.Collections.isNonEmptyCollection(pLevels) Then
        Set currentProcess = pLevels.item(pLevels.Count)
        pProgress = currentProcess.getHigh
        progress = calculateCurrentProcessProgress
        Call updateProgress
        Call pLevels.remove(pLevels.Count)
    End If

End Sub

Private Function calculateCurrentProcessProgress() As Double
    Dim parentalProcess As Limit
    Dim currentProcess As Limit
    '------------------------------------------------------------------------------------------------------
    Set currentProcess = pLevels.item(pLevels.Count)
    If pLevels.Count > 1 Then
        Set parentalProcess = pLevels.item(pLevels.Count - 1)
        calculateCurrentProcessProgress = (currentProcess.getHigh - parentalProcess.getLow) / parentalProcess.getRange
    Else
        calculateCurrentProcessProgress = currentProcess.getHigh
    End If
End Function

Public Sub moveToSiblingProcess(ByVal offset As Double)
    Dim currentProcess As Limit
    Dim parentProcess As Limit
    '------------------------------------------------------------------------------------------------------
    
    If F.Collections.isNonEmptyCollection(pLevels) Then
        Set currentProcess = pLevels.item(pLevels.Count)
        If pLevels.Count > 1 Then
            Set parentProcess = pLevels.item(pLevels.Count - 1)
        Else
            Set parentProcess = F.Create.Limit(0, 1)
        End If
        
        With currentProcess
            Call .setLowValue(.getHigh)
            Call .setHighValue(.getHigh + offset * parentProcess.getRange)
            pProgress = .getLow
        End With
        
        Call updateProgress
    End If
    
End Sub

Private Function calculateRealProgress(ByVal value As Double) As Double
    Dim currentProcessBounds As Limit
    '------------------------------------------------------------------------------------------------------
    
    '[Adjust sub-process progress value]
    If value < 0 Then
        value = 0
    ElseIf value > 1 Then
        value = 1
    End If
    
    '[Adjust progress value to current sub-process]
    If F.Collections.isNonEmptyCollection(pLevels) Then
        Set currentProcessBounds = pLevels.item(pLevels.Count)
    Else
        Set currentProcessBounds = F.Create.Limit(0, 1)
    End If
    
    '[Actual calculation]
    With currentProcessBounds
        calculateRealProgress = .getLow + value * .getRange
    End With
    
End Function


Public Sub toFront()
    Call pForm.toFront
End Sub
