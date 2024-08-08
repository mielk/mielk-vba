VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XProcessItemsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum ControlTypeEnum
    ControlType_ContainerFrame = 1
    ControlType_BottomBorder = 2
End Enum

Private Const CLASS_NAME As String = "XProcessItemsList"
'[UI constants] -------------------------------------------------------------------------------------------
Private Const DEFAULT_WIDTH As Single = 402
Private Const DEFAULT_TOP_PADDING As Single = 0
Private Const DEFAULT_BOTTOM_PADDING As Single = 0
Private Const DEFAULT_INSIDE_PADDING As Single = 0
Private Const DEFAULT_BACK_STYLE As Variant = MSForms.fmBackStyleTransparent
Private Const DEFAULT_BACK_COLOR As Long = 16316664 ' &H8000000F
Private Const DEFAULT_BORDER_STYLE As Variant = MSForms.fmBorderStyleSingle
Private Const DEFAULT_BORDER_COLOR As Long = &H80000006
Private Const DEFAULT_CONTAINER_MIN_HEIGHT As Single = 24
Private Const DEFAULT_CONTAINER_MAX_HEIGHT As Single = 300
Private Const DEFAULT_HAS_BOTTOM_BORDER_ONLY As Boolean = False
Private Const DEFAULT_LINE_HEIGHT As Single = 19.5
'[UI components] ------------------------------------------------------------------------------------------
Private pProcessor As ItemsProcessor
Private pView As IItemsProcessorViewer
Private pWindow As MielkUI.WTemplate
'[Internal UI components] ---------------------------------------------------------------------------------
Private pContainerFrame As MSForms.frame
Private pBottomBorder As MSForms.label
'[UI properties] ------------------------------------------------------------------------------------------
Private pTop As Single
Private pLeft As Single
Private pWidth As Single
Private pContainerMaxHeight As Single
Private pContainerMinHeight As Single
Private pTopPadding As Single
Private pBottomPadding As Single
Private pInsidePadding As Single
Private pBackStyle As MSForms.fmBackStyle
Private pBackColor As Long
Private pBorderStyle As MSForms.fmBorderStyle
Private pBorderColor As Long
Private pLineHeight As Single
Private pHasBottomBorderOnly As Boolean
'[State] --------------------------------------------------------------------------------------------------
Private pUuid As String
Private pIsRendered As Boolean
Private pItemLinesDictionary As Scripting.Dictionary
Private pItemLinesArray As Variant
'[Events] -------------------------------------------------------------------------------------------------
Event AfterExpandCollapse(item As IProcessItem)
Event AfterPreviewFileClicked(item As IProcessItem)
Event AfterReprocessClicked(item As IProcessItem)
Event AfterRemoveClicked(item As IProcessItem)
'----------------------------------------------------------------------------------------------------------



'[INITIALIZE]
Private Sub Class_Initialize()
    Call initializeContainers
    Call initializeDefaultValues
End Sub

Private Sub initializeContainers()
    Set pItemLinesDictionary = F.dictionaries.Create(False)
End Sub

Private Sub initializeDefaultValues()
    pUuid = F.Crypto.createUUID
    pWidth = DEFAULT_WIDTH
    pLineHeight = DEFAULT_LINE_HEIGHT
    pContainerMinHeight = DEFAULT_CONTAINER_MIN_HEIGHT
    pContainerMaxHeight = DEFAULT_CONTAINER_MAX_HEIGHT
    pTopPadding = DEFAULT_TOP_PADDING
    pBottomPadding = DEFAULT_BOTTOM_PADDING
    pInsidePadding = DEFAULT_INSIDE_PADDING
    pBackStyle = DEFAULT_BACK_STYLE
    pBackColor = DEFAULT_BACK_COLOR
    pBorderStyle = DEFAULT_BORDER_STYLE
    pBorderColor = DEFAULT_BORDER_COLOR
    pHasBottomBorderOnly = DEFAULT_HAS_BOTTOM_BORDER_ONLY
End Sub


'[SETTERS]
Public Function setWindow(value As MielkUI.WTemplate) As XProcessItemsList
    Set setWindow = Me
    Set pWindow = value
End Function

Public Function setFrame(value As MSForms.frame) As XProcessItemsList
    Set setFrame = Me
    Set pContainerFrame = value
End Function

Public Function setView(value As IItemsProcessorViewer) As XProcessItemsList
    Set setView = Me
    Set pView = value
End Function

Public Function setProcessor(value As ItemsProcessor) As XProcessItemsList
    Set setProcessor = Me
    Set pProcessor = value
End Function




'[Position]
Public Function setTop(value As Single) As XProcessItemsList
    Set setTop = Me
    pTop = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setBottom(value As Single) As XProcessItemsList
    Set setBottom = Me
    pTop = value - getHeight
    If pIsRendered Then Call arrangeControls
End Function

Public Function setLeft(value As Single) As XProcessItemsList
    Set setLeft = Me
    pLeft = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setRight(value As Single) As XProcessItemsList
    Set setRight = Me
    pLeft = value - pWidth
    If pIsRendered Then Call arrangeControls
End Function


'[Size]
Public Function setWidth(value As Single) As XProcessItemsList
    Set setWidth = Me
    pWidth = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setHeight(value As Single) As XProcessItemsList
    Set setHeight = Me
    Call VBA.Err.Raise(1)
End Function

Public Function setLineHeight(value As Single) As XProcessItemsList
    Set setLineHeight = Me
    pLineHeight = value
End Function



'[Backcolor and borders]
Public Function setBackStyle(value As MSForms.fmBackStyle) As XProcessItemsList
    Set setBackStyle = Me
    pBackStyle = value
    If pIsRendered Then Call refreshBackgroundFrame
End Function

Public Function setBackColor(value As Long) As XProcessItemsList
    Set setBackColor = Me
    pBackColor = value
    If pIsRendered Then Call refreshBackgroundFrame
End Function

Public Function setBorderStyle(value As MSForms.fmBorderStyle) As XProcessItemsList
    Set setBorderStyle = Me
    pBorderStyle = value
    If pIsRendered Then Call refreshBackgroundFrame
End Function

Public Function setBorderColor(value As Long) As XProcessItemsList
    Set setBorderColor = Me
    pBorderColor = value
    If pIsRendered Then Call refreshBackgroundFrame
End Function



'[Layout]
Public Function setTopPadding(value As Single) As XProcessItemsList
    Set setTopPadding = Me
    pTopPadding = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setBottomPadding(value As Single) As XProcessItemsList
    Set setBottomPadding = Me
    pBottomPadding = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setInsidePadding(value As Single) As XProcessItemsList
    Set setInsidePadding = Me
    pInsidePadding = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setHasBottomBorderOnly(value As Boolean) As XProcessItemsList
    Set setHasBottomBorderOnly = Me
    pHasBottomBorderOnly = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setContainerMaxHeight(value As Boolean) As XProcessItemsList
    Set setContainerMaxHeight = Me
    pContainerMaxHeight = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setContainerMinHeight(value As Boolean) As XProcessItemsList
    Set setContainerMinHeight = Me
    pContainerMinHeight = value
    If pIsRendered Then Call arrangeControls
End Function





'[GETTERS]
Public Function isValid() As Boolean
    isValid = UI.Controls.isValid(pContainerFrame)
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


'[Other]
Friend Function getItemLinesContainerFrame() As MSForms.frame
    Set getItemLinesContainerFrame = pContainerFrame
End Function

Friend Function getView() As IItemsProcessorViewer
    Set getView = pView
End Function

Friend Function getProcessor() As ItemsProcessor
    Set getProcessor = pProcessor
End Function

Friend Function getStatusesDictionary() As Scripting.Dictionary
    Set getStatusesDictionary = pView.getStatusesDictionary
End Function





'[API]
Public Sub render()
    Call refreshBackgroundFrame
    Call generateAndInsertItemLines
    Call arrangeControls
    pIsRendered = True
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
    Dim TAG As String
    '----------------------------------------------------------------------------------------------------------
    TAG = CLASS_NAME & "|" & pUuid & "|" & controlType
    getLabelName = "ctrl_" & VBA.left$(F.Crypto.SHA1HASH(TAG), 25)
End Function


'[Inserting controls]
Private Sub insertBottomBorder()
    Set pBottomBorder = pContainerFrame.Controls.Add(CREATE_LABEL_ID, getLabelName(ControlType_BottomBorder), True)
End Sub



'[Refreshing controls]
Private Sub refreshBackgroundFrame()
    With pContainerFrame
        .BackColor = pBackColor
        .BorderStyle = VBA.IIf(pHasBottomBorderOnly, MSForms.fmBorderStyleNone, pBorderStyle)
        .BorderColor = pBorderColor
        .SpecialEffect = fmSpecialEffectFlat
        .caption = vbNullString
    End With
End Sub

Private Sub refreshBottomBorder()
    With pBottomBorder
        .BackStyle = MSForms.fmBackStyleOpaque
        .BackColor = pBorderColor
        .BorderStyle = MSForms.fmBorderStyleNone
    End With
End Sub




'[Size and layout]
Private Sub arrangeControls()
    On Error Resume Next
    Call resizeControls
    'Call locateControl(pBottomBorder, 0, 0)
    Call UI.Controls.locate(pContainerFrame, pLeft, pTop)
End Sub

Private Sub resizeControls()
    Call resizeContainerFrame
    Call UI.Controls.resize(pBottomBorder, pWidth, 1)
End Sub

Private Sub resizeContainerFrame()
    With pContainerFrame
        .width = pWidth
        .height = pContainerMaxHeight
    End With
End Sub

Public Sub updateScrollBarVisibility()
    Dim endPoint As Coordinate
    Dim bottom As Single
    '------------------------------------------------------------------------------------------------------
    
    endPoint = UI.Forms.calculateRightBottomPoint(pContainerFrame, True)
    bottom = endPoint.Y
    
    With pContainerFrame
        If bottom > .InsideHeight Then
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = bottom
        ElseIf bottom < .InsideHeight Then
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = 0
        End If
    End With
    
End Sub




'[Clearing]
Private Sub clear()
    Call UI.Forms.clearMsFrame(pContainerFrame)
End Sub

Public Sub removeItem(name As String)
    Dim itemLine As xxProcessItemLine
    '----------------------------------------------------------------------------------------------------------
    Set itemLine = getItemLineByName(name)
    If Not itemLine Is Nothing Then
        Call pItemLinesDictionary.Remove(name)
        Call itemLine.clear
        pItemLinesArray = pItemLinesDictionary.items
        Call resetLinesIndex
    End If
End Sub

Private Sub resetLinesIndex()
    Dim i As Long
    Dim index As Long
    Dim line As xxProcessItemLine
    '----------------------------------------------------------------------------------------------------------
    If F.arrays.isDefined(pItemLinesArray) Then
        For i = LBound(pItemLinesArray) To UBound(pItemLinesArray)
            Set line = pItemLinesArray(i)
            Call line.setIndex(index)
            index = index + 1
        Next i
    End If
End Sub



'[FILE LINES]
Private Sub generateAndInsertItemLines()
    Dim items As VBA.Collection
    Dim item As IProcessItem
    Dim index As Long
    Dim itemLine As xxProcessItemLine
    '----------------------------------------------------------------------------------------------------------
    
    Set items = pView.getItems
    For Each item In items
        index = index + 1
        Set itemLine = MApp.Create.processItemLine(Me, item, index).setLineHeight(pLineHeight).setStatusesDictionary(getStatusesDictionary)
        Call itemLine.render
        Call addItemLineToContainers(itemLine)
    Next item
    Call arrangeItemLines
    
End Sub

Private Sub addItemLineToContainers(itemLine As xxProcessItemLine)
    Call pItemLinesDictionary.Add(itemLine.getName, itemLine)
    Call F.arrays.addEntry(pItemLinesArray, itemLine)
End Sub

Private Function getItemLineByName(name As String) As xxProcessItemLine
    Set getItemLineByName = F.dictionaries.GetObject(pItemLinesDictionary, name)
End Function

Friend Sub arrangeItemLines(Optional arrangeProcessBeginItem As xxProcessItemLine)
    Dim i As Long
    Dim itemLine As xxProcessItemLine
    Dim top As Single:                              top = pTopPadding
    Dim initialIndex As Integer
    Dim tops() As Single
    '------------------------------------------------------------------------------------------------------
    
    If F.arrays.isDefined(pItemLinesArray) Then
        
        initialIndex = LBound(pItemLinesArray)
        
        If Not arrangeProcessBeginItem Is Nothing Then
            With arrangeProcessBeginItem
                top = .getTop
                initialIndex = .getIndex
            End With
        End If
        
        ReDim tops(initialIndex To UBound(pItemLinesArray))
        
        For i = initialIndex To UBound(pItemLinesArray)
            Set itemLine = pItemLinesArray(i)
            If Not itemLine.isHidden Then
                tops(i) = top
                top = top + itemLine.getHeight - PIXEL_SIZE
            End If
        Next i
        
        For i = UBound(pItemLinesArray) To initialIndex Step -1
            Set itemLine = pItemLinesArray(i)
            If Not itemLine.isHidden Then
                Call itemLine.move(top:=tops(i))
            End If
        Next i
        
    End If
    
End Sub

Private Sub adjustContainerFrameHeight()
    Stop
End Sub


Public Sub updateItemLine(itemName As String, status As ItemProcessStatusEnum, message As String)
    Dim itemLine As xxProcessItemLine
    '------------------------------------------------------------------------------------------------------
    Set itemLine = F.dictionaries.getItem(pItemLinesDictionary, itemName)
    If Not itemLine Is Nothing Then
        Call itemLine.updateInfo(status, message)
    End If
End Sub

Friend Sub triggerAfterExpandCollapseEvent(line As xxProcessItemLine)
    Call updateScrollBarVisibility
    RaiseEvent AfterExpandCollapse(line.getItem)
End Sub

Friend Sub triggerAfterPreviewItemEvent(line As xxProcessItemLine)
    RaiseEvent AfterPreviewFileClicked(line.getItem)
End Sub

Friend Sub triggerAfterReprocessItemEvent(line As xxProcessItemLine)
    RaiseEvent AfterReprocessClicked(line.getItem)
End Sub

Friend Sub triggerAfterRemoveItemEvent(item As IProcessItem)
    Call removeItem(item.getName)
    Call arrangeItemLines
End Sub



Friend Sub filter(status As Variant)
    Dim key As Variant
    Dim line As xxProcessItemLine
    '------------------------------------------------------------------------------------------------------
    
    For Each key In pItemLinesDictionary.Keys
        Set line = pItemLinesDictionary.item(key)
        If status < 0 Or line.getItem.getStatus = status Then
            Call line.setIsHidden(False)
        Else
            Call line.setIsHidden(True)
        End If
    Next key
    
    Call arrangeItemLines
    Call updateScrollBarVisibility
    
End Sub
