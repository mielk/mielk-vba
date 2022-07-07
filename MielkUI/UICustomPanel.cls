VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UICustomPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPanel

Private Const CLASS_NAME As String = "UICustomPanel"
'----------------------------------------------------------------------------------------------------------
Private Const DEFAULT_HORIZONTAL_PADDING As Single = 0
Private Const DEFAULT_VERTICAL_PADDING As Single = 9
Private Const DEFAULT_BACKGROUND_COLOR As Long = vbWhite
'----------------------------------------------------------------------------------------------------------
Private pListener As IListener
Private pWindow As WTemplate
Private pParentFrame As MSForms.frame
Private pInnerFrame As MSForms.frame
Private pTopBorder As MSForms.label
Private pBottomBorder As MSForms.label
Private pControls As VBA.Collection             'Collection to keep references to all controls added to this panel
'[UI properties] ------------------------------------------------------------------------------------------
Private pTop As Single
Private pLeft As Single
Private pBottom As Single
Private pRight As Single
Private pWidth As Single
Private pMinWidth As Single
Private pMaxWidth As Single
Private pHeight As Single
Private pMinHeight As Single
Private pMaxHeight As Single
Private pHorizontalPadding As Single
Private pVerticalPadding As Single
Private pBackgroundColor As Long
'[Properties] ---------------------------------------------------------------------------------------------
Private pName As String
Private pIndex As Long
Private pIsRendered As Boolean
Private pIsVisible As Boolean
Private pBordersVisibility As Boolean
Private pDisableEvents As Boolean
'----------------------------------------------------------------------------------------------------------
Event AfterResize()
Event AfterVisibilityChange()
Event RenderControls()
'----------------------------------------------------------------------------------------------------------





Private Sub Class_Initialize()
    '[Size]
    pWidth = -1
    pMinWidth = -1
    pMaxWidth = -1
    pHeight = -1
    pMinHeight = -1
    pMaxHeight = -1
    pHorizontalPadding = DEFAULT_HORIZONTAL_PADDING
    pVerticalPadding = DEFAULT_VERTICAL_PADDING
    pBackgroundColor = DEFAULT_BACKGROUND_COLOR
    pIsVisible = True
    '[Containers]
    Set pControls = New VBA.Collection
End Sub



'[SETTERS]
Public Function setListener(value As IListener) As UICustomPanel
    Set setListener = Me
    Set pListener = value
End Function

Public Function setWindow(value As WTemplate) As UICustomPanel
    Set setWindow = Me
    Set pWindow = value
End Function

Public Function setParentFrame(value As MSForms.frame) As UICustomPanel
    Set setParentFrame = Me
    Set pParentFrame = value
End Function


Public Function setName(value As String) As UICustomPanel
    Set setName = Me
    pName = value
End Function

Public Function setIndex(value As Long) As UICustomPanel
    Set setIndex = Me
    pIndex = value
End Function

Public Function setTop(value As Single) As UICustomPanel
    Set setTop = Me
    pTop = value
End Function

Public Function setLeft(value As Single) As UICustomPanel
    Set setLeft = Me
    pLeft = value
End Function

Public Function addControl(value As Variant) As UICustomPanel
    Set addControl = Me
    Call pControls.Add(value)
End Function

Public Function setIsVisible(value As Boolean) As UICustomPanel
    Set setIsVisible = Me
    pIsVisible = value
    Call updateVisibility
End Function

Public Function setUiProperty(prop As UIProp, value As Variant) As UICustomPanel
    Set setUiProperty = Me
    Select Case prop.getId
        Case UI.props.top.getId:                    Call setTop(VBA.CSng(value))
        Case UI.props.left.getId:                   Call setLeft(VBA.CSng(value))
        Case UI.props.width.getId:                  pWidth = VBA.CSng(value)
        Case UI.props.MinWidth.getId:               pMinWidth = VBA.CSng(value)
        Case UI.props.maxWidth.getId:               pMaxWidth = VBA.CSng(value)
        Case UI.props.height.getId:                 pHeight = VBA.CSng(value)
        Case UI.props.minHeight.getId:              pMinHeight = VBA.CSng(value)
        Case UI.props.maxHeight.getId:              pMaxHeight = VBA.CSng(value)
        Case UI.props.padding.getId:
                                                    pHorizontalPadding = VBA.CSng(value)
                                                    pVerticalPadding = VBA.CSng(value)
        Case UI.props.paddingHorizontal.getId:      pHorizontalPadding = VBA.CSng(value)
        Case UI.props.paddingVertical.getId:        pVerticalPadding = VBA.CSng(value)
        Case UI.props.backgroundColor.getId:        pBackgroundColor = VBA.CLng(value)
                                                    Call updateBackgroundColor
    End Select
    
    If pIsRendered Then
        'If prop.isLayoutProperty Then Call arrangeLines
        If prop.isPositionProperty Then Call updatePosition
        If prop.isSizeProperty Then Call updateSize
    End If
    
End Function


'[GETTERS]
Public Function getWindow() As WTemplate
    Set getWindow = pWindow
End Function

Public Function getInnerFrame() As MSForms.frame
    Set getInnerFrame = pInnerFrame
End Function

Public Function getName() As String
    getName = pName
End Function

Public Function getIndex() As Long
    getIndex = pIndex
End Function

Public Function getBottom() As Single
    getBottom = pBottom
End Function

Public Function getWidth() As Single
    getWidth = pInnerFrame.width
End Function

Public Function getVerticalPadding() As Single
    getVerticalPadding = pVerticalPadding
End Function

Public Function getHorizontalPadding() As Single
    getHorizontalPadding = pHorizontalPadding
End Function

Public Function isVisible() As Boolean
    isVisible = pIsVisible
End Function


Public Function getUiProperty(prop As UIProp) As Variant
    If UI.Forms.isControlValid(pInnerFrame) Then
        With pInnerFrame
            Select Case prop.getId
                Case UI.props.top.getId:                getUiProperty = .top
                Case UI.props.left.getId:               getUiProperty = .left
                Case UI.props.right.getId:              getUiProperty = .left + .width
                Case UI.props.bottom.getId:             getUiProperty = .top + .height
                Case UI.props.width.getId:              getUiProperty = .width
                Case UI.props.height.getId:             getUiProperty = .height
            End Select
        End With
    End If
End Function



'[RENDERING]
Public Sub render()
    If Not pIsRendered Then
        Call insertInnerFrame
        RaiseEvent RenderControls
        Call autoResize
        Call updatePosition
        Call updateVisibility
        pIsRendered = True
    End If
End Sub

Private Sub insertInnerFrame()
    Set pInnerFrame = pParentFrame.controls.Add(CREATE_FRAME_ID)
    With pInnerFrame
        .borderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

Public Sub insertBorder(index As XlBordersIndex, color As Long)
    Dim dict As Scripting.Dictionary
    Dim line As xxPropertyLine
    '----------------------------------------------------------------------------------------------------------
    
    If index = xlEdgeTop Then
        Call updateBorderVariable(pTopBorder, color)
    ElseIf index = xlEdgeBottom Then
        Call updateBorderVariable(pBottomBorder, color)
    End If
    
End Sub

Private Sub updateBorderVariable(ByRef borderLabel As MSForms.label, color As Long)
    If borderLabel Is Nothing Then
        Set borderLabel = pInnerFrame.controls.Add(CREATE_LABEL_ID)
    End If
    With borderLabel
        .borderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleOpaque
        .backColor = color
        .height = PIXEL_SIZE
    End With
End Sub

Private Sub updateBordersSizeAndPosition()
    If UI.Forms.isControlValid(pTopBorder) Then
        With pTopBorder
            .left = 0
            .top = PIXEL_SIZE
            .width = pInnerFrame.InsideWidth
            .height = PIXEL_SIZE
        End With
    End If
    
    If UI.Forms.isControlValid(pBottomBorder) Then
        With pBottomBorder
            .left = 0
            .top = pParentFrame.InsideHeight - PIXEL_SIZE
            .width = pParentFrame.InsideWidth
            .height = PIXEL_SIZE
        End With
    End If
    
End Sub

Private Sub updateBackgroundColor()
    On Error Resume Next
    pInnerFrame.backColor = pBackgroundColor
End Sub

Private Sub updateSize()
    Dim lastPoint As Coordinate
    '----------------------------------------------------------------------------------------------------------
    
    lastPoint = UI.Forms.calculateRightBottomPoint(pInnerFrame, True)
    
    With pInnerFrame
        '[Width]
        If pWidth >= 0 Then
            .width = pWidth
        ElseIf pMinWidth > 0 And lastPoint.x < pMinWidth Then
            .width = pMinWidth
        ElseIf pMaxWidth > 0 And lastPoint.x > pMaxWidth Then
            .width = pMaxWidth
        Else
            .width = lastPoint.x + pHorizontalPadding
        End If
        
        '[Height]
        If pHeight >= 0 Then
            .height = pHeight
        ElseIf pMinHeight > 0 And lastPoint.y < pMinHeight Then
            .height = pMinHeight
        ElseIf pMaxHeight > 0 And lastPoint.y > pMaxHeight Then
            .height = pMaxHeight
        Else
            .height = lastPoint.y
        End If
        
    End With
    
    RaiseEvent AfterResize
    
    Call updateScrollBars
    Call updateBordersSizeAndPosition
    
End Sub

Private Sub updateScrollBars()
    Dim lastPoint As Coordinate
    '----------------------------------------------------------------------------------------------------------
    
    lastPoint = UI.Forms.calculateRightBottomPoint(pInnerFrame, True)
    With pInnerFrame
        .ScrollBars = fmScrollBarsNone
        
        '[Vertical scroll bar]
        .ScrollHeight = lastPoint.y
        If lastPoint.y > .InsideHeight Then
            .ScrollBars = .ScrollBars + fmScrollBarsVertical
            .scrollTop = 0
        End If
        
        '[Horizontal scroll bar]
        .ScrollWidth = lastPoint.x
        If lastPoint.x > .InsideWidth Then
            .ScrollBars = .ScrollBars + fmScrollBarsHorizontal
            .ScrollLeft = 0
        End If
        
    End With
     
End Sub

Private Sub updatePosition()
    If UI.Forms.isControlValid(pInnerFrame) Then
        With pInnerFrame
            .left = pLeft
            .top = pTop
        End With
    End If
End Sub


Public Sub updateVisibility()
    If UI.Forms.isControlValid(pInnerFrame) Then
        If pInnerFrame.visible <> pIsVisible Then
            pInnerFrame.visible = pIsVisible
            RaiseEvent AfterVisibilityChange
        End If
    End If
End Sub


Public Sub autoResize(Optional resizeWindowAfterward As Boolean = True)
    Dim coord As Coordinate
    '----------------------------------------------------------------------------------------------------------
    
    If Not pDisableEvents Then
        pDisableEvents = True
        coord = UI.Forms.calculateRightBottomPoint(pInnerFrame)
        With pInnerFrame
            .height = coord.y
            .width = coord.x + pHorizontalPadding
        End With
        
        If resizeWindowAfterward Then RaiseEvent AfterResize
        pDisableEvents = False
    End If
    
End Sub


'[Implementation of IPanel interface]

Private Function IPanel_setName(value As String) As IPanel
    '
End Function

Private Function IPanel_getName() As String
    IPanel_getName = vbNullString
End Function


Private Function IPanel_setIndex(value As Long) As IPanel
    Set IPanel_setIndex = setIndex(value)
End Function

Private Function IPanel_getIndex() As Long
    IPanel_getIndex = getIndex
End Function


Private Sub IPanel_render()
    Call render
End Sub

Private Function IPanel_setParentFrame(value As MSForms.frame) As IPanel
    Set IPanel_setParentFrame = setParentFrame(value)
End Function

Private Function IPanel_setUiProperty(prop As MielkCore.UIProp, value As Variant) As IPanel
    Set IPanel_setUiProperty = setUiProperty(prop, value)
End Function

Private Function IPanel_getUiProperty(prop As MielkCore.UIProp) As Variant
    Call F.Variables.assign(IPanel_getUiProperty, getUiProperty(prop))
End Function

Private Sub IPanel_insertBorder(index As XlBordersIndex, color As Long)
    Call insertBorder(index, color)
End Sub

Private Sub IPanel_addValidationLabel(value As UIValidationLabel)
    Stop
End Sub

Private Function IPanel_setListener(value As IListener) As IPanel
    Set IPanel_setListener = setListener(value)
End Function

Private Function IPanel_getValidation() As MielkCore.ValidationObject
    'Set IPanel_getValidation = getValidation
End Function

Private Function IPanel_getValue() As Variant
    ' stop
End Function

Private Function IPanel_setWindow(value As WTemplate) As IPanel
    Set IPanel_setWindow = setWindow(value)
End Function

Private Function IPanel_isVisible() As Boolean
    If UI.Forms.isControlValid(pInnerFrame) Then
        IPanel_isVisible = pInnerFrame.visible
    End If
End Function
