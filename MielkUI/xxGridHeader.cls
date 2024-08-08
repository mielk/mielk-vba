VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xxGridHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IControl
Implements IListener

Private Const CLASS_NAME As String = "xxGridHeader"
'----------------------------------------------------------------------------------------------------------
Private Const DEFAULT_FONT_COLOR As Long = vbBlack
Private Const DEFAULT_FONT_SIZE As Single = 10
Private Const DEFAULT_FONT_BOLD As Boolean = True
Private Const DEFAULT_FONT_NAME As String = APP_FONT_FAMILY
Private Const CAPTION_MARGIN As Single = 3
Private Const FILTER_BORDER_SIZE As Single = 12
Private Const FILTER_ICON_SIZE As Single = 12
Private Const FILTER_LEFT_MARGIN As Single = 1.5
Private Const FILTER_ACTIVE_BACK_COLOR As Long = 7987710
Private Const FILTER_ACTIVE_BORDER_COLOR As Long = 4041417
Private Const SORT_ICON_LEFT_MARGIN As Single = 3
'----------------------------------------------------------------------------------------------------------
Private pGridManager As XGridManager
Private pListener As IListener
Private pFrame As MSForms.frame
Private pColumn As DColumn
'----------------------------------------------------------------------------------------------------------
Private pBackgroundLabel As MSForms.label
Attribute pBackgroundLabel.VB_VarHelpID = -1
Private WithEvents pCaptionLabel As MSForms.label
Attribute pCaptionLabel.VB_VarHelpID = -1
Private pFilterIcon As UIIcon
Attribute pFilterIcon.VB_VarHelpID = -1
'----------------------------------------------------------------------------------------------------------
Private pCaption As String
Private pTop As Single
Private pLeft As Single
Private pWidth As Single
Private pVisibility As Boolean
Private pIsRendered As Boolean
'----------------------------------------------------------------------------------------------------------
Private pFontColor As Long
Private pFontName As String
Private pFontBold As Boolean
Private pFontSize As Single
'----------------------------------------------------------------------------------------------------------
Private pFilterActive As Boolean
'----------------------------------------------------------------------------------------------------------





Private Sub Class_Initialize()
    pFontColor = DEFAULT_FONT_COLOR
    pFontSize = DEFAULT_FONT_SIZE
    pFontName = DEFAULT_FONT_NAME
    pFontBold = DEFAULT_FONT_BOLD
End Sub



'[SETTERS]
Public Function setFrame(value As MSForms.frame) As xxGridHeader
    Set setFrame = Me
    Set pFrame = value
End Function

Public Function setGridManager(value As XGridManager) As xxGridHeader
    Set setGridManager = Me
    Set pGridManager = value
End Function

Public Function setListener(value As IListener) As xxGridHeader
    Set setListener = Me
    Set pListener = value
End Function

Public Function setColumn(value As DColumn) As xxGridHeader
    Set setColumn = Me
    Set pColumn = value
End Function

Public Function setFontName(value As String) As xxGridHeader
    Set setFontName = Me
    pFontName = value
    Call updateCaption
End Function

Public Function setFontSize(value As Single) As xxGridHeader
    Set setFontSize = Me
    pFontSize = value
    Call updateCaption
End Function

Public Function setFontBold(value As Boolean) As xxGridHeader
    Set setFontBold = Me
    pFontBold = value
    Call updateCaption
End Function

Public Function setFontColor(value As Long) As xxGridHeader
    Set setFontColor = Me
    pFontColor = value
    Call updateCaption
End Function

Public Function setVisibility(value As Boolean) As xxGridHeader
    Set setVisibility = Me
    pVisibility = value
    Call updateVisibility
End Function

Public Function setWidth(value As Single) As xxGridHeader
    Set setWidth = Me
    pWidth = value
    Call updateSize
End Function

Public Function setLeft(value As Single) As xxGridHeader
    Set setLeft = Me
    pLeft = value
    If pIsRendered Then Call arrangeLayout
End Function



'[GETTERS]
Public Function getColumn() As DColumn
    Set getColumn = pColumn
End Function

Public Function getRight() As Single
    getRight = UI.controls.getRight(pBackgroundLabel)
End Function

Private Function getProperFilterPicture() As Object
    If pFilterActive Then
        Set getProperFilterPicture = ufImages.icoFilterYellow.picture
    Else
        Set getProperFilterPicture = ufImages.icoFilterWhite.picture
    End If
End Function

Private Function getCaption() As String
    If VBA.Len(pCaption) Then
        getCaption = pCaption
    Else
        getCaption = pColumn.getLabel
    End If
End Function

Public Function getFilterIcon() As UIIcon
    Set getFilterIcon = pFilterIcon
End Function

Public Function getHeight() As Single
    getHeight = pBackgroundLabel.height
End Function




'[RENDERING]
Public Sub render()
    If Not pIsRendered Then
        Call insertBackgroundLabel
        Call insertCaptionLabel
        Call insertFilterIcon
        Call arrangeLayout
        pIsRendered = True
    End If
End Sub

Private Sub insertBackgroundLabel()
    Set pBackgroundLabel = pFrame.controls.add(CREATE_LABEL_ID)
    With pBackgroundLabel
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .height = 20
        .width = pColumn.getWidth
    End With
End Sub

Private Sub insertCaptionLabel()
    Set pCaptionLabel = pFrame.controls.add(CREATE_LABEL_ID)
    With pCaptionLabel
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignLeft
        .borderStyle = fmBorderStyleNone
        If pColumn.isSortable Then
            .MousePointer = fmMousePointerCustom
           Set .MouseIcon = Icons.getMousePointer
        End If
        Call updateCaption
    End With
End Sub

Private Sub insertFilterIcon()
    If pColumn.hasFilter Then
        Set pFilterIcon = UI.Create.UIIcon(pFrame)
        With pFilterIcon
            Call .setBorderColor(FILTER_ACTIVE_BORDER_COLOR)
            Call .setHasBorder(False)
            Call .setBackColor(FILTER_ACTIVE_BACK_COLOR)
            Call .setIsTransparent(True)
            Call .setListener(Me)
            Call .setHeight(FILTER_ICON_SIZE)
            Call .setWidth(FILTER_ICON_SIZE)
            Call .setVisibility(True)
            Call .render
        End With
        Call updateFilterIconView
    End If
End Sub


Private Sub arrangeLayout()
    '[Background]
    With pBackgroundLabel
        .width = pColumn.getWidth
        pTop = UI.Normalize.position((pFrame.InsideHeight - .height) / 2)
        .top = pTop
        .left = UI.Normalize.position(VBA.CDate(pLeft))
    End With
    
    '[Caption]
    With pCaptionLabel
        .top = UI.Normalize.position((pFrame.InsideHeight - .height) / 2)
        .left = pBackgroundLabel.left
    End With
    
    '[Filter]
    If pColumn.hasFilter Then
        If Not pFilterIcon Is Nothing Then
            With pFilterIcon
                Call .Move(UI.controls.getRight(pCaptionLabel) + FILTER_LEFT_MARGIN, (pFrame.InsideHeight - .getControl.height) / 2)
                Call .alignVertically(pBackgroundLabel)
            End With
        End If
    End If
    
End Sub


Public Function Move(Optional x As Variant, Optional Y As Variant) As xxGridHeader
    Dim offsetX As Single
    Dim offsetY As Single
    '------------------------------------------------------------------------------------------------------
    
    Set Move = Me
    
    With pBackgroundLabel
        If VBA.IsNumeric(x) Then
            offsetX = (x - pLeft)
            pLeft = x
        End If
        If VBA.IsNumeric(Y) Then
            offsetY = (Y - pTop)
            pTop = Y
        End If
    End With
    
    Call offset(offsetX, offsetY)
    
End Function


Public Function offset(Optional x As Single, Optional Y As Single) As xxGridHeader
    Set offset = Me
    Call UI.Forms.offsetControls(x, Y, pBackgroundLabel, pCaptionLabel, pFilterIcon)
End Function

Private Sub updateSize()
    If UI.controls.isValid(pBackgroundLabel) Then
        pBackgroundLabel.width = pWidth
    End If
    If UI.controls.isValid(pCaptionLabel) Then
        pCaptionLabel.width = pWidth
    End If
End Sub

Private Sub updateCaption()
    If UI.controls.isValid(pCaptionLabel) Then
        With pCaptionLabel
            .caption = getCaption
            With .font
                .name = pFontName
                .Bold = VBA.IIf(pFontBold, 1, False)
                .size = pFontSize
            End With
            .ForeColor = pFontColor
            .width = pBackgroundLabel.width
            .AutoSize = True
            .AutoSize = False
            .AutoSize = True
            .width = .width + 2 * CAPTION_MARGIN
            Call arrangeLayout
        End With
    End If
End Sub

Private Sub updateVisibility()
    Stop
End Sub

Public Sub activateFilter(value As Boolean)
    pFilterActive = value
    Call updateFilterIconView
End Sub

Public Sub refreshFilterIcon()
    Dim filter As IFilter
    '------------------------------------------------------------------------------------------------------
    
    Set filter = pColumn.getFilter
    If Not filter Is Nothing Then
        pFilterActive = filter.isActive
    End If
    Call updateFilterIconView
    
End Sub

Private Sub updateFilterIconView()
    With pFilterIcon
        Call .setHasBorder(pFilterActive)
        Call .setIsTransparent(Not pFilterActive)
        Call .setPicture(getProperFilterPicture)
    End With
End Sub

Public Sub appendSortIcon(ctrl As MSForms.label)
    Dim left As Single
    Dim top As Single
    '------------------------------------------------------------------------------------------------------
    
    With ctrl
        If pFilterIcon Is Nothing Then
            left = UI.controls.getRight(pCaptionLabel)
        Else
            left = pFilterIcon.getRight + SORT_ICON_LEFT_MARGIN
        End If
        top = pCaptionLabel.top + (pCaptionLabel.height - .height) / 2
        .left = left
        .top = top
        .visible = True
    End With
    
End Sub





'[Controls' events]

Private Sub pCaptionLabel_Click()
    Call ErrorManager.clear
    
    If Not pListener Is Nothing Then
        Call pListener.catchEvent(Me, ACTION_SORT)
    End If
    
End Sub




'[Implementation of IControl interface]

Private Sub IControl_setVisibility(value As Boolean)
    Call setVisibility(value)
End Sub

Private Sub IControl_setEnability(value As Boolean)
    Stop
End Sub

Private Sub IControl_setWidth(value As Single)
    Call setWidth(value)
End Sub

Private Function IControl_getHeight() As Single
    IControl_getHeight = getHeight
End Function

Private Function IControl_getWidth() As Single
    Stop
End Function

Private Function IControl_getLeft() As Single
    Stop
End Function

Private Function IControl_getTop() As Single
    Stop
End Function

Private Function IControl_getRectangle() As MielkCore.RECT
    Stop
End Function

Private Sub IControl_resize(Optional width As Variant, Optional height As Variant)
    Stop
End Sub

Private Sub IControl_alignVertically(baseCtrl As MSForms.IControl)
    'Call alignVertically(baseCtrl)
End Sub

Private Function IControl_move(Optional x As Variant, Optional Y As Variant) As IControl
    Set IControl_move = Move(x, Y)
End Function

Private Function IControl_offset(Optional x As Single, Optional Y As Single) As IControl
    Set IControl_offset = offset(x, Y)
End Function

Private Sub IControl_triggerEvent(eventId As Long)
    Stop
End Sub





'[Implementation of IListener]
Private Sub IListener_catchEvent(src As Variant, ev As Variant, Optional properties As Scripting.IDictionary)
    If src Is pFilterIcon Then
        If ev = UI.Events.MouseUp.getName Then
            Call handleFilterIconMouseUpEvent(src, props)
        End If
    End If
End Sub

Private Sub handleFilterIconMouseUpEvent(Source As Variant, properties As Scripting.Dictionary)
    Dim shift As Long
    '------------------------------------------------------------------------------------------------------
    
    If F.dictionaries.isNonEmptyDictionary(properties) Then
        shift = F.dictionaries.getItem(properties, props.shift.getName)
        Select Case shift
            Case 0:             Call pListener.catchEvent(Me, ACTION_FILTER_ON)
            Case CTRL_MASK:     Call pListener.catchEvent(Me, ACTION_FILTER_OFF)
        End Select
    End If
    
End Sub
