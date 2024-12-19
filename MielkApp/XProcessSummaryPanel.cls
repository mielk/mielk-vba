VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XProcessSummaryPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum ControlTypeEnum
    ControlType_ContainerFrame = 1
End Enum

Private Const CLASS_NAME As String = "XProcessSummaryPanel"
'[UI constants] -------------------------------------------------------------------------------------------
Private Const DEFAULT_WIDTH As Single = 402
Private Const DEFAULT_TOP_PADDING As Single = 3
Private Const DEFAULT_LEFT_PADDING As Single = 3
Private Const DEFAULT_BOTTOM_PADDING As Single = 3
Private Const DEFAULT_INSIDE_PADDING As Single = 3
Private Const DEFAULT_BACK_STYLE As Variant = MSForms.fmBackStyleTransparent
Private Const DEFAULT_BACK_COLOR As Long = 16316664 ' &H8000000F
Private Const DEFAULT_BORDER_STYLE As Variant = MSForms.fmBorderStyleSingle
Private Const DEFAULT_BORDER_COLOR As Long = &H80000006
Private Const DEFAULT_CONTAINER_MIN_HEIGHT As Single = 24
Private Const DEFAULT_CONTAINER_MAX_HEIGHT As Single = 96
Private Const DEFAULT_CONTROL_WIDTH As Single = 150
Private Const DEFAULT_LINE_HEIGHT As Single = 24
'[Other constants] ----------------------------------------------------------------------------------------
Private Const STATUS_ALL_TAG As String = "Status.All"
'[UI components] ------------------------------------------------------------------------------------------
Private WithEvents pProcessor As ItemsProcessor
Attribute pProcessor.VB_VarHelpID = -1
Private pController As IItemsProcessorViewer
Private pWindow As MielkUI.WTemplate
'[Internal UI components] ---------------------------------------------------------------------------------
Private pContainerFrame As MSForms.frame
'[UI properties] ------------------------------------------------------------------------------------------
Private pTop As Single
Private pLeft As Single
Private pWidth As Single
Private pContainerMaxHeight As Single
Private pContainerMinHeight As Single
Private pLeftPadding As Single
Private pTopPadding As Single
Private pBottomPadding As Single
Private pInsidePadding As Single
Private pBackStyle As MSForms.fmBackStyle
Private pBackColor As Long
Private pBorderStyle As MSForms.fmBorderStyle
Private pBorderColor As Long
Private pControlWidth As Single
Private pLineHeight As Single
'[State] --------------------------------------------------------------------------------------------------
Private pUuid As String
Private pIsRendered As Boolean
Private pControlsDictionary As Scripting.Dictionary
Private pActiveControl As xxProcessSummaryControl
'[Events] -------------------------------------------------------------------------------------------------
Event AfterClick(key As Variant)
'----------------------------------------------------------------------------------------------------------



'[INITIALIZE]
Private Sub Class_Initialize()
    Call initializeContainers
    Call initializeDefaultValues
End Sub

Private Sub initializeContainers()
    Set pControlsDictionary = F.dictionaries.Create(False)
End Sub

Private Sub initializeDefaultValues()
    pUuid = F.Crypto.createUUID
    pWidth = DEFAULT_WIDTH
    pContainerMinHeight = DEFAULT_CONTAINER_MIN_HEIGHT
    pContainerMaxHeight = DEFAULT_CONTAINER_MAX_HEIGHT
    pLeftPadding = DEFAULT_LEFT_PADDING
    pTopPadding = DEFAULT_TOP_PADDING
    pBottomPadding = DEFAULT_BOTTOM_PADDING
    pInsidePadding = DEFAULT_INSIDE_PADDING
    pBackStyle = DEFAULT_BACK_STYLE
    pBackColor = DEFAULT_BACK_COLOR
    pBorderStyle = DEFAULT_BORDER_STYLE
    pBorderColor = DEFAULT_BORDER_COLOR
    pLineHeight = DEFAULT_LINE_HEIGHT
    pControlWidth = DEFAULT_CONTROL_WIDTH
End Sub


'[SETTERS]
Public Function setWindow(value As MielkUI.WTemplate) As XProcessSummaryPanel
    Set setWindow = Me
    Set pWindow = value
End Function

Public Function setFrame(value As MSForms.frame) As XProcessSummaryPanel
    Set setFrame = Me
    Set pContainerFrame = value
End Function

Public Function setController(value As IItemsProcessorViewer) As XProcessSummaryPanel
    Set setController = Me
    Set pController = value
    Set pProcessor = pController.getProcessor
End Function




'[Position]
Public Function setTop(value As Single) As XProcessSummaryPanel
    Set setTop = Me
    pTop = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setBottom(value As Single) As XProcessSummaryPanel
    Set setBottom = Me
    pTop = value - getHeight
    If pIsRendered Then Call arrangeControls
End Function

Public Function setLeft(value As Single) As XProcessSummaryPanel
    Set setLeft = Me
    pLeft = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setRight(value As Single) As XProcessSummaryPanel
    Set setRight = Me
    pLeft = value - pWidth
    If pIsRendered Then Call arrangeControls
End Function


'[Size]
Public Function setWidth(value As Single) As XProcessSummaryPanel
    Set setWidth = Me
    pWidth = value
    If pIsRendered Then Call arrangeControls
End Function



'[Backcolor and borders]
Public Function setBackStyle(value As MSForms.fmBackStyle) As XProcessSummaryPanel
    Set setBackStyle = Me
    pBackStyle = value
    If pIsRendered Then Call refreshBackgroundFrame
End Function

Public Function setBackColor(value As Long) As XProcessSummaryPanel
    Set setBackColor = Me
    pBackColor = value
    If pIsRendered Then Call refreshBackgroundFrame
End Function

Public Function setBorderStyle(value As MSForms.fmBorderStyle) As XProcessSummaryPanel
    Set setBorderStyle = Me
    pBorderStyle = value
    If pIsRendered Then Call refreshBackgroundFrame
End Function

Public Function setBorderColor(value As Long) As XProcessSummaryPanel
    Set setBorderColor = Me
    pBorderColor = value
    If pIsRendered Then Call refreshBackgroundFrame
End Function



'[Layout]
Public Function setTopPadding(value As Single) As XProcessSummaryPanel
    Set setTopPadding = Me
    pTopPadding = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setBottomPadding(value As Single) As XProcessSummaryPanel
    Set setBottomPadding = Me
    pBottomPadding = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setInsidePadding(value As Single) As XProcessSummaryPanel
    Set setInsidePadding = Me
    pInsidePadding = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setContainerMaxHeight(value As Boolean) As XProcessSummaryPanel
    Set setContainerMaxHeight = Me
    pContainerMaxHeight = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setContainerMinHeight(value As Boolean) As XProcessSummaryPanel
    Set setContainerMinHeight = Me
    pContainerMinHeight = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setControlWidth(value As Single) As XProcessSummaryPanel
    Set setControlWidth = Me
    pControlWidth = value
End Function

Public Function setLineHeight(value As Single) As XProcessSummaryPanel
    Set setLineHeight = Me
    pLineHeight = value
End Function



'[GETTERS]
Public Function getProcessor() As ItemsProcessor
    Set getProcessor = pProcessor
End Function

Public Function isValid() As Boolean
    isValid = UI.Controls.isValid(pContainerFrame)
End Function

Public Function getContainerFrame() As MSForms.frame
    Set getContainerFrame = pContainerFrame
End Function



'[Position]
Public Function getLeft() As Single
    getLeft = pContainerFrame.left
End Function

Public Function getRight() As Single
    getRight = UI.Controls.getRight(pContainerFrame)
End Function

Public Function getTop() As Single
    getTop = pContainerFrame.top
End Function

Public Function getBottom() As Single
    getBottom = UI.Controls.getBottom(pContainerFrame)
End Function


'[Size]
Public Function getWidth() As Single
    getWidth = pWidth
End Function

Public Function getHeight() As Single
    getHeight = pContainerFrame.height
End Function

Public Function getControlWidth() As Single
    getControlWidth = pControlWidth
End Function


'[Other]
Friend Function getController() As IItemsProcessorViewer
    Set getController = pController
End Function





'[API]
Public Sub render()
    If Not pIsRendered Then
        Call refreshBackgroundFrame
        Call generateAndInsertItems
        Call arrangeControls
        pIsRendered = True
    End If
End Sub

Public Sub move(Optional left As Variant, Optional top As Variant)
    If F.Variables.isOfNumericType(left) Then pLeft = left
    If F.Variables.isOfNumericType(top) Then pTop = top
    Call arrangeControls
End Sub

Public Sub resize(Optional width As Variant, Optional height As Variant)
    If width >= 0 Then pWidth = width
    Call arrangeControls
End Sub




'[RENDERING]
Private Function getLabelName(controlType As ControlTypeEnum) As String
    Dim tag As String
    '----------------------------------------------------------------------------------------------------------
    tag = CLASS_NAME & "|" & pUuid & "|" & controlType
    getLabelName = "ctrl_" & VBA.left$(F.Crypto.SHA1HASH(tag), 25)
End Function



'[Refreshing controls]
Private Sub refreshBackgroundFrame()
    With pContainerFrame
        .BackColor = pBackColor
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = pBorderColor
        .SpecialEffect = fmSpecialEffectFlat
        .caption = vbNullString
    End With
End Sub




'[Size and layout]
Private Sub arrangeControls()
    On Error Resume Next
    pContainerFrame.width = pWidth
    Call arrangeItems
    Call UI.Controls.locate(pContainerFrame, pLeft, pTop)
End Sub


'[Items controls]
Private Sub generateAndInsertItems()
    Dim statuses As Scripting.Dictionary
    Dim key As Variant
    Dim dto As Scripting.Dictionary
    Dim control As xxProcessSummaryControl
    '----------------------------------------------------------------------------------------------------------
    
    Set pControlsDictionary = F.dictionaries.Create(False)
    Set statuses = getStatusesDictionary
    
    If F.dictionaries.isNonEmptyDictionary(statuses) Then
        For Each key In statuses.Keys
            Set dto = statuses.item(key)
            Set control = MApp.Create.processSummaryControl(Me).setKey(key).loadDto(dto). _
                                    setLineHeight(pLineHeight).setWidth(pControlWidth).render
            Call F.dictionaries.addItem(pControlsDictionary, key, control)
            If control.isActive Then Set pActiveControl = control
        Next key
    End If
    
    Call arrangeItems
    
End Sub

Private Function getStatusesDictionary() As Scripting.Dictionary
    Set getStatusesDictionary = F.dictionaries.Create(False)
    Call getStatusesDictionary.Add(-1, createStatusAllDto)
    Call F.dictionaries.appendDictionary(getStatusesDictionary, pController.getStatusesDictionary, duplicateBehaviour_Override)
End Function

Private Function createStatusAllDto() As Scripting.Dictionary
    Set createStatusAllDto = F.dictionaries.Create(False)
    With createStatusAllDto
        Call .Add(TAG_NAME, MsgService.getText(STATUS_ALL_TAG))
        Call .Add(TAG_MAIN_BACKGROUND, vbBlack)
        Call .Add(TAG_IS_ACTIVE, True)
    End With
End Function

Private Sub arrangeItems()
    Dim key As Variant
    Dim control As xxProcessSummaryControl
    Dim top As Single:                              top = pTopPadding
    Dim left As Single:                             left = pLeftPadding
    '------------------------------------------------------------------------------------------------------
    Dim columns As Long
    Dim column As Long
    '------------------------------------------------------------------------------------------------------
    
    If F.dictionaries.isNonEmptyDictionary(pControlsDictionary) Then
        columns = VBA.Int((pWidth - 2 * pLeftPadding) / pControlWidth)
        For Each key In pControlsDictionary.Keys
            If column = columns Then
                column = 1
                top = top + pLineHeight
            Else
                column = column + 1
            End If
            Set control = pControlsDictionary.item(key)
            Call control.move(left + pControlWidth * (column - 1), top)
        Next key
    End If
    
    Call updateContainerFrameHeight(top + pLineHeight + pTopPadding + pBottomPadding)
    
End Sub


Private Sub updateContainerFrameHeight(ByVal height As Single)
    If height > pContainerMaxHeight Then
        pContainerFrame.height = pContainerMaxHeight
    ElseIf height < pContainerMinHeight Then
        pContainerFrame.height = pContainerMinHeight
    Else
        pContainerFrame.height = height
    End If
End Sub


Friend Sub triggerControlClickedEvent(control As xxProcessSummaryControl)
    If pActiveControl Is control Then
    
    Else
        If Not pActiveControl Is Nothing Then Call pActiveControl.setIsActive(False)
        Call control.setIsActive(True)
        Set pActiveControl = control
        RaiseEvent AfterClick(control.getKey)
    End If
End Sub




Private Sub pProcessor_AfterStatusCountersChange(previousStatus As Variant, newStatus As Variant)
    Dim control As xxProcessSummaryControl
    '------------------------------------------------------------------------------------------------------
    
    Call ErrorManager.clear
    
    '[Previous status]
    Set control = F.dictionaries.GetObject(pControlsDictionary, previousStatus)
    If Not control Is Nothing Then Call control.updateCaption
    
    '[New status]
    Set control = F.dictionaries.GetObject(pControlsDictionary, newStatus)
    If Not control Is Nothing Then Call control.updateCaption
    
End Sub
