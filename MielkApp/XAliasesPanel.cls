VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XAliasesPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPanel
Implements IListener

Private Const CLASS_NAME As String = "XAliasesPanel"
'----------------------------------------------------------------------------------------------------------
Private Const DEFAULT_LINE_HEIGHT As Single = 21
Private Const DEFAULT_VERTICAL_PADDING As Single = 3
Private Const DEFAULT_HORIZONTAL_PADDING As Single = 6
Private Const DEFAULT_TITLE As String = "AliasesPanel.Title"
Private Const DEFAULT_BACK_COLOR As Long = 15790320
Private Const DEFAULT_BORDER_COLOR As Long = 13487565
'[Add alias button properties] ----------------------------------------------------------------------------
Private Const ADD_ALIAS_BUTTON_TAG As String = "AddAlias"
Private Const ADD_ALIAS_BUTTON_HEIGHT As Single = 24
Private Const ADD_ALIAS_BUTTON_WIDTH As Single = 90
Private Const ADD_ALIAS_BUTTON_FONT_SIZE As Single = 10
Private Const ADD_ALIAS_BUTTON_FONT_COLOR As Long = 1381653
Private Const ADD_ALIAS_BUTTON_CAPTION As String = "AliasesPanel.AddAliasButton"
Private Const ADD_ALIAS_BUTTON_BACK_COLOR As Long = NEUTRAL_BACK_COLOR
Private Const ADD_ALIAS_BUTTON_BORDER_COLOR As Long = NEUTRAL_BORDER_COLOR
'----------------------------------------------------------------------------------------------------------
Private pWindow As WTemplate
Private pListener As IListener
Private pParentFrame As MSForms.frame
Private pInnerFrame As MSForms.frame
Private pTitleLabel As MSForms.label
Private pLines As Scripting.Dictionary
Private pAddAliasButton As UIButton
'----------------------------------------------------------------------------------------------------------
Private pItem As IAliasable
'[UI properties] ------------------------------------------------------------------------------------------
Private pTitle As String
Private pLineHeight As Single
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
Private pHorizontalAlignment As XlHAlign
Private pHorizontalPadding As Single
Private pVerticalPadding As Single
Private pBackColor As Long
Private pBorderColor As Long
'[State] --------------------------------------------------------------------------------------------------
Private pName As String
Private pIndex As Long
Private pIsRendered As Boolean
Private pIsVisible As Boolean
'----------------------------------------------------------------------------------------------------------
'Event AfterChange()
'Event AfterValidationChange(validation As ValidationObject)
'----------------------------------------------------------------------------------------------------------





Private Sub Class_Initialize()
    Set pLines = F.dictionaries.Create(False)
    pLineHeight = DEFAULT_LINE_HEIGHT
    pIsVisible = True
    pTitle = MsgService.getText(DEFAULT_TITLE)
    
    '[Size]
    pWidth = -1
    pMinWidth = -1
    pMaxWidth = -1
    pHeight = -1
    pMinHeight = -1
    pMaxHeight = -1
    pVerticalPadding = DEFAULT_VERTICAL_PADDING
    pHorizontalPadding = DEFAULT_HORIZONTAL_PADDING
    pBackColor = DEFAULT_BACK_COLOR
    pBorderColor = DEFAULT_BORDER_COLOR
    
End Sub



'[SETTERS]
Public Function setName(value As String) As XAliasesPanel
    Set setName = Me
    pName = value
End Function

Public Function setListener(value As IListener) As XAliasesPanel
    Set setListener = Me
    Set pListener = value
End Function

Public Function setWindow(value As WTemplate) As XAliasesPanel
    Set setWindow = Me
    Set pWindow = value
End Function

Public Function setParentFrame(value As MSForms.frame) As XAliasesPanel
    Set setParentFrame = Me
    Set pParentFrame = value
End Function

Public Function setItem(value As IAliasable) As XAliasesPanel
    Set setItem = Me
    Set pItem = value
End Function

Public Function setTop(value As Single) As XAliasesPanel
    Set setTop = Me
    pTop = value
End Function

Public Function setLeft(value As Single) As XAliasesPanel
    Set setLeft = Me
    pLeft = value
End Function

Public Function setIndex(value As Long) As XAliasesPanel
    Set setIndex = Me
    pIndex = value
End Function

Public Function setVerticalPadding(value As Single) As XAliasesPanel
    Set setVerticalPadding = Me
    pVerticalPadding = value
End Function

Public Function setHorizontalAlignment(value As XlHAlign) As XAliasesPanel
    Set setHorizontalAlignment = Me
    pHorizontalAlignment = value
End Function

Public Function setBackColor(value As Long) As XAliasesPanel
    Set setBackColor = Me
    pBackColor = value
End Function

Public Function setBorderColor(value As Long) As XAliasesPanel
    Set setBorderColor = Me
    pBorderColor = value
End Function

Public Function setTitle(value As String) As XAliasesPanel
    Set setTitle = Me
    pTitle = value
End Function

Public Function setUiProperty(prop As UIProp, value As Variant) As XAliasesPanel
    Set setUiProperty = Me
    Select Case prop.getId
        Case UI.Props.top.getId:                    Call setTop(VBA.CSng(value))
        Case UI.Props.left.getId:                   Call setLeft(VBA.CSng(value))
        Case UI.Props.width.getId:                  pWidth = VBA.CSng(value)
        Case UI.Props.MinWidth.getId:               pMinWidth = VBA.CSng(value)
        Case UI.Props.maxWidth.getId:               pMaxWidth = VBA.CSng(value)
        Case UI.Props.height.getId:                 pHeight = VBA.CSng(value)
        Case UI.Props.minHeight.getId:              pMinHeight = VBA.CSng(value)
        Case UI.Props.maxHeight.getId:              pMaxHeight = VBA.CSng(value)
        Case UI.Props.padding.getId:
                                                    pHorizontalPadding = VBA.CSng(value)
                                                    pVerticalPadding = VBA.CSng(value)
        Case UI.Props.paddingHorizontal.getId:      pHorizontalPadding = VBA.CSng(value)
        Case UI.Props.paddingVertical.getId:        pVerticalPadding = VBA.CSng(value)
        Case UI.Props.BackgroundColor.getId:        pBackColor = VBA.CLng(value)
        Case UI.Props.BorderColor.getId:            pBorderColor = VBA.CLng(value)
    End Select
    
    If pIsRendered Then
        If prop.isLayoutProperty Then Call arrangeControls
        If prop.isPositionProperty Then Call updatePosition
        If prop.isSizeProperty Then Call updateSize
        If prop.isLookProperty Then Call updateLook
    End If
    
End Function

Public Function setVisibility(value As Boolean) As XAliasesPanel
    Dim key As Variant
    Dim line As xxPropertyLine
    '------------------------------------------------------------------------------------------------------
    
    Set setVisibility = Me
    pIsVisible = value
    
    If UI.Forms.isControlValid(pInnerFrame) Then
        pInnerFrame.Visible = pIsVisible
    End If
    
End Function


'[GETTERS]
Public Function getName() As String
    getName = pName
End Function

Public Function getWindow() As WTemplate
    Set getWindow = pWindow
End Function

Public Function getInnerFrame() As MSForms.frame
    Set getInnerFrame = pInnerFrame
End Function

Public Function getItem() As IAliasable
    Set getItem = pItem
End Function

Public Function getBottom() As Single
    getBottom = pBottom
End Function

Public Function getWidth() As Single
    getWidth = pInnerFrame.width
End Function

Public Function getLineHeight() As Single
    getLineHeight = pLineHeight
End Function

Public Function getIndex() As Long
    getIndex = pIndex
End Function

Public Function getUiProperty(prop As UIProp) As Variant
    If UI.Forms.isControlValid(pInnerFrame) Then
        With pInnerFrame
            Select Case prop.getId
                Case UI.Props.top.getId:                getUiProperty = .top
                Case UI.Props.left.getId:               getUiProperty = .left
                Case UI.Props.right.getId:              getUiProperty = .left + .width
                Case UI.Props.bottom.getId:             getUiProperty = .top + .height
            End Select
        End With
    End If
End Function

Public Function getVisibility() As Boolean
    getVisibility = pIsVisible
End Function

Public Function getHorizontalAlignment() As XlHAlign
    getHorizontalAlignment = pHorizontalAlignment
End Function

Public Function getHorizontalPadding() As Single
    getHorizontalPadding = pHorizontalPadding
End Function


'[RENDERING]
Public Sub render()
    If Not pIsRendered Then
        Call insertInnerFrame
        Call insertTitleLabel
        Call insertLines
        Call insertAddButton
        Call arrangeControls
        Call updatePosition
        'Call updateValidation
        pIsRendered = True
    End If
End Sub

Private Sub insertInnerFrame()
    Set pInnerFrame = pParentFrame.Controls.Add(CREATE_FRAME_ID)
    With pInnerFrame
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = pBorderColor
        .SpecialEffect = fmSpecialEffectFlat
        .BackColor = pBackColor
    End With
End Sub

Private Sub insertTitleLabel()
    Set pTitleLabel = pInnerFrame.Controls.Add(CREATE_LABEL_ID)
    With pTitleLabel
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .Font.name = APP_FONT_FAMILY
        .Font.size = 9
        .Font.Bold = 0
        .caption = VBA.UCase$(pTitle)
        .AutoSize = True
    End With
End Sub

Private Sub insertLines()
    Dim aliases As VBA.Collection
    Dim alias As Variant
    Dim line As xxAliasLine
    '------------------------------------------------------------------------------------------------------
    Set pLines = F.dictionaries.Create(False)
    Set aliases = pItem.getAliases
    For Each alias In aliases
        Set line = createLine(VBA.CStr(alias))
        Call F.dictionaries.addItem(pLines, alias, line)
    Next alias
    Call sortLines
End Sub

Private Function createLine(alias As String) As xxAliasLine
    Set createLine = New xxAliasLine
    With createLine
        Call .setParent(Me)
        Call .setAlias(VBA.CStr(alias))
        Call .render
    End With
End Function

Private Sub sortLines()
    Set pLines = F.dictionaries.sortByKey(pLines, True)
End Sub

Private Sub insertAddButton()
    Set pAddAliasButton = UI.Create.UIButton(pInnerFrame)
    With pAddAliasButton
        Call .setTag(ADD_ALIAS_BUTTON_TAG)
        Call .setHeight(ADD_ALIAS_BUTTON_HEIGHT)
        Call .setWidth(ADD_ALIAS_BUTTON_WIDTH)
        Call .setFontSize(ADD_ALIAS_BUTTON_FONT_SIZE)
        Call .setFontName(APP_FONT_FAMILY)
        Call .setFontBold(False)
        Call .setFontColor(ADD_ALIAS_BUTTON_FONT_COLOR)
        Call .setCaption(MsgService.getText(ADD_ALIAS_BUTTON_CAPTION))
        Call .setBackgroundColor(ADD_ALIAS_BUTTON_BACK_COLOR)
        Call .setBorderColor(ADD_ALIAS_BUTTON_BORDER_COLOR)
        Call .setVisibility(True)
        Call .setListener(Me)
        Call .render
    End With
End Sub

Private Sub updateBorderVariable(ByRef borderLabel As MSForms.label, color As Long)
    If borderLabel Is Nothing Then
        Set borderLabel = pInnerFrame.Controls.Add(CREATE_LABEL_ID)
    End If
    With borderLabel
        .BorderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleOpaque
        .BackColor = color
        .height = PIXEL_SIZE
    End With
End Sub



Friend Sub arrangeControls()
    Dim key As Variant
    Dim line As xxAliasLine
    '------------------------------------------------------------------------------------------------------
    Dim top As Single:              top = pVerticalPadding
    Dim left As Single:             left = pHorizontalPadding
    Dim right As Single
    '------------------------------------------------------------------------------------------------------

    Call UI.Forms.locateControl(pTitleLabel, left, top)
    top = UI.Forms.getBottom(pTitleLabel) + pVerticalPadding
    
    If F.dictionaries.isNonEmptyDictionary(pLines) Then
        For Each key In pLines.Keys
            Set line = pLines.item(key)
            Call line.move(left, top)
            top = top + pLineHeight
        Next key
        top = top + 2 * pVerticalPadding
    End If
    
    Call pAddAliasButton.move(, top)

End Sub

Private Sub updateSize()
    Dim lastPoint As Coordinate
    '------------------------------------------------------------------------------------------------------
    
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
        ElseIf pMinHeight > 0 And lastPoint.Y < pMinHeight Then
            .height = pMinHeight
        ElseIf pMaxHeight > 0 And lastPoint.Y > pMaxHeight Then
            .height = pMaxHeight
        Else
            .height = lastPoint.Y + 2 * pVerticalPadding
        End If
        
    End With
    
    Call updateAddButtonPosition
    Call updateAliasLinesWidth
    Call updateScrollBars
    
End Sub

Private Sub updateAddButtonPosition()
    Call pAddAliasButton.move(pInnerFrame.InsideWidth - ADD_ALIAS_BUTTON_WIDTH - pHorizontalPadding)
End Sub

Private Sub updateAliasLinesWidth()
    Dim key As Variant
    Dim line As xxAliasLine
    '------------------------------------------------------------------------------------------------------
    For Each key In pLines.Keys
        Set line = pLines.item(key)
        Call line.arrangeControls
    Next key
End Sub

Private Sub updateScrollBars()
    Dim lastPoint As Coordinate
    '------------------------------------------------------------------------------------------------------
    
    lastPoint = UI.Forms.calculateRightBottomPoint(pInnerFrame, True)
    With pInnerFrame
        .ScrollBars = fmScrollBarsNone
        
        '[Vertical scroll bar]
        .ScrollHeight = lastPoint.Y
        If lastPoint.Y > .InsideHeight Then
            .ScrollBars = .ScrollBars + fmScrollBarsVertical
            .ScrollTop = 0
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
            If pTop <> .top Then .top = pTop
        End With
    End If
End Sub

Private Sub updateLook()
    With pInnerFrame
        .BackColor = pBackColor
        .BorderColor = pBorderColor
    End With
End Sub



'[VALIDATION]

Public Function getValidation() As ValidationObject
    Set getValidation = F.Create.ValidationObject(True)
End Function








Public Sub setFocus()
    Dim key As Variant
    Dim line As xxPropertyLine
    '------------------------------------------------------------------------------------------------------
    
    For Each key In pLines.Keys
        Set line = pLines.items(0)
        If Not line Is Nothing Then
            If Not line.isReadOnly Then
                Call line.setFocus
                Exit For
            End If
        End If
    Next key
    
ExitPoint:
    Exit Sub

ErrHandler:
    Stop
    
End Sub



Private Sub handleAddingAlias()
    Dim alias As EAlias
    Dim validation As ValidationObject
    Dim line As xxAliasLine
    Dim text As String
    '------------------------------------------------------------------------------------------------------
    
    Set alias = New EAlias
    With alias
        Call .setParent(pItem)
        Call .setConfirmWhenOverriding(True)
        Set validation = .edit
        If validation.getStatus <> False Then
            text = .getAlias
            Call pItem.addAlias(text)
            Set line = createLine(text)
            Call F.dictionaries.addItem(pLines, text, line)
            Call sortLines
            
            '[Update view] ------------------------------|
            Call arrangeControls                        '|
            Call updateSize                             '|
            Call pWindow.arrangeComponents              '|
            '--------------------------------------------|
        End If
    End With
    
End Sub


Friend Sub deleteAlias(alias As String)
    Call pItem.removeAlias(alias)
    Call F.dictionaries.removeItem(pLines, alias)
    Call arrangeControls
    Call updateSize
    Call pWindow.arrangeComponents
End Sub




'[Implementation of IPanel interface]

Private Function IPanel_setName(value As String) As IPanel
    Set IPanel_setName = setName(value)
End Function

Private Function IPanel_getName() As String
    IPanel_getName = getName
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
    'Call insertBorder(index, color)
End Sub

Private Sub IPanel_addValidationLabel(value As UIValidationLabel)
    'N/A
End Sub

Private Function IPanel_setListener(value As IListener) As IPanel
    Set IPanel_setListener = setListener(value)
End Function

Private Function IPanel_getValidation() As MielkCore.ValidationObject
    Set IPanel_getValidation = getValidation
End Function

Private Function IPanel_getValue() As Variant
    Set IPanel_getValue = pItem
End Function

Private Function IPanel_setWindow(value As WTemplate) As IPanel
    Set IPanel_setWindow = setWindow(value)
End Function

Private Function IPanel_isVisible() As Boolean
    IPanel_isVisible = getVisibility
End Function


'[Implementation of IListener interface]
Private Sub IListener_catchEvent(source As Variant, eventName As String, Optional properties As Scripting.IDictionary)
    If source Is pAddAliasButton Then
        Call handleAddingAlias
    End If
End Sub

