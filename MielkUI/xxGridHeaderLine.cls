VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xxGridHeaderLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IListener

Private Const CLASS_NAME As String = "xxGridHeaderLine"
'----------------------------------------------------------------------------------------------------------
Private Const BACKGROUND_LABEL_NAME As String = "lblBackground"
Private Const BOTTOM_BORDER_LABEL_NAME As String = "lblBottomBorder"
Private Const SORT_ICON_LABEL_NAME As String = "lblSortIcon"
'----------------------------------------------------------------------------------------------------------
Private Const LINE_HEIGHT As Single = 18
Private Const SORT_ICON_SIZE As Single = 9
Private Const FONT_FAMILY As String = APP_FONT_FAMILY
Private Const FONT_SIZE As Single = 11
Private Const FONT_COLOR As Long = vbBlack
Private Const BORDER_COLOR As Long = 5526612
'----------------------------------------------------------------------------------------------------------
Private pManager As XGridManager
Private pFrame As MSForms.frame
'----------------------------------------------------------------------------------------------------------
Private pHeaders As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------
Private pUuid As String
Private pTop As Single
Private pLeft As Single
Private pIsRendered As Boolean
Private pSortAscending As Boolean
Private pEventsDisabled As Boolean
'[UI controls] --------------------------------------------------------------------------------------------
Private pBackground As MSForms.label
Private pBottomBorder As MSForms.label
Private pSortIcon As MSForms.label
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    '[Set initial values]
    pUuid = F.Crypto.createUUID
    
    '[Initialize containers]
    Set pHeaders = F.dictionaries.Create(False)
End Sub





'[SETTERS]
Public Function setManager(value As XGridManager) As xxGridHeaderLine
    Set setManager = Me
    Set pManager = value
End Function

Public Function setFrame(value As MSForms.frame) As xxGridHeaderLine
    Set setFrame = Me
    Set pFrame = value
End Function

Public Function setTop(value As Single) As xxGridHeaderLine
    Dim offsetY As Single
    '------------------------------------------------------------------------------------------------------
    Set setTop = Me
    offsetY = value - pTop
    pTop = value
    If pIsRendered Then Call offsetControls(Y:=offsetY)
End Function

Public Function setLeft(value As Single) As xxGridHeaderLine
    Dim offsetX As Single
    '------------------------------------------------------------------------------------------------------
    Set setLeft = Me
    offsetX = value - pLeft
    pLeft = value
    If pIsRendered Then Call offsetControls(x:=offsetX)
End Function



'[GETTERS]
Public Function getBottom() As Single
    If UI.controls.isValid(pBackground) Then
        getBottom = pBackground.top + pBackground.height
    End If
End Function

Public Function getHeader(key As String) As xxGridHeader
    Set getHeader = F.dictionaries.getObject(pHeaders, key)
End Function





'[RENDERING]
Public Sub render()
    Dim columns As VBA.Collection
    Dim column As DColumn
    Dim header As xxGridHeader
    Dim left As Single:             left = pManager.getHorizontalPadding
    '------------------------------------------------------------------------------------------------------
    
    If Not pIsRendered Then
        pEventsDisabled = True
        
        Call insertBackground
        Call insertBottomBorder
        Call insertSortLabel
        
        Set pHeaders = F.dictionaries.Create(False)
        Set columns = pManager.getColumnsCollection
        For Each column In columns
            Set header = insertHeader(column, left)
            Call pHeaders.add(column.getName, header)
        Next column
        
        pIsRendered = True
        Call relocateControls
        pEventsDisabled = False
    End If
    
End Sub


Private Sub insertBackground()
    Set pBackground = pFrame.controls.add(CREATE_LABEL_ID, BACKGROUND_LABEL_NAME)
    With pBackground
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .left = 0
        .width = pFrame.InsideWidth
        .height = pFrame.InsideHeight
        .top = pTop
    End With
End Sub

Private Sub insertBottomBorder()
    Set pBottomBorder = pFrame.controls.add(CREATE_LABEL_ID, BOTTOM_BORDER_LABEL_NAME, True)
    With pBottomBorder
        .BackStyle = fmBackStyleOpaque
        .backColor = BORDER_COLOR
        .borderStyle = fmBorderStyleNone
        .height = PIXEL_SIZE
        .width = pFrame.width
    End With
End Sub

Private Sub insertSortLabel()
    Set pSortIcon = pFrame.controls.add(CREATE_LABEL_ID, SORT_ICON_LABEL_NAME, False)
    With pSortIcon
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .height = SORT_ICON_SIZE
        .width = SORT_ICON_SIZE
        Set .picture = getProperSortPicture
    End With
End Sub

Private Function insertHeader(column As DColumn, ByRef left As Single) As xxGridHeader
    Set insertHeader = UI.Create.xxGridHeader(pFrame)
    With insertHeader
        Call .setGridManager(pManager)
        Call .setListener(Me)
        Call .setColumn(column)
        Call .render
        Call .Move(left)
        left = .getRight + pManager.getColumnsSpace
    End With
End Function

Private Sub relocateControls()
    Dim key As Variant
    Dim header As xxGridHeader
    '------------------------------------------------------------------------------------------------------
    
    If pIsRendered Then
        With pBackground
            .top = pTop
            .left = pLeft
            .width = pFrame.width
        End With
        With pBottomBorder
            .top = pFrame.InsideHeight - .height
            .left = pLeft - PIXEL_SIZE
            .width = pFrame.width
        End With
        For Each key In pHeaders.keys
            Set header = pHeaders.item(key)
            If Not header Is Nothing Then
                'Call header.alignVertically(pbackground)
            End If
        Next key
    End If
    
End Sub

Private Sub offsetControls(Optional x As Single, Optional Y As Single)
    Dim key As Variant
    Dim header As xxGridHeader
    '------------------------------------------------------------------------------------------------------
    
    If pIsRendered Then
        If Y Then
            pBackground.top = pBackground.top + Y
            pBottomBorder.top = pBottomBorder.top + Y
            For Each key In pHeaders.keys
                Set header = pHeaders.item(key)
                If Not header Is Nothing Then
                    Call header.offset(Y:=Y)
                End If
            Next key
        End If
    End If
    
End Sub


Public Sub arrangeControls()
    Dim columns As VBA.Collection
    Dim column As DColumn
    Dim header As xxGridHeader
    Dim left As Single:                     left = pManager.getHorizontalPadding
    '------------------------------------------------------------------------------------------------------
    
    pBackground.width = pFrame.InsideWidth
    pBottomBorder.width = pFrame.InsideWidth
    
    Set columns = pManager.getColumnsCollection
    For Each column In columns
        Set header = F.dictionaries.getObject(pHeaders, column.getName)
        If Not header Is Nothing Then
            Call header.setLeft(left)
            left = left + column.getWidth + pManager.getColumnsSpace
        End If
    Next column
    
End Sub







'[SORTING]
Private Sub updateSortIcon(header As xxGridHeader)
    Set pSortIcon.picture = getProperSortPicture
    Call header.appendSortIcon(pSortIcon)
End Sub

Private Function getProperSortPicture() As Object
    If pSortAscending Then
        Set getProperSortPicture = ufImages.ico_SortUp.picture
    Else
        Set getProperSortPicture = ufImages.ico_SortDown.picture
    End If
End Function




'[Implementation of IListener interface]
Private Sub IListener_catchEvent(src As Variant, ev As Variant, Optional properties As Scripting.IDictionary)
    If ev = ACTION_SORT Then
        Call handleSortTrigger(src)
    ElseIf ev = ACTION_FILTER_ON Then
        Call handleFilterOnTrigger(src)
    ElseIf ev = ACTION_FILTER_OFF Then
        Call handleFilterOffTrigger(src)
    ElseIf ev = "exportToExcel" Then
        Call pManager.exportToExcel
    End If
End Sub

Private Sub handleSortTrigger(Source As Variant)
    Dim header As xxGridHeader:         Set header = Source
    '------------------------------------------------------------------------------------------------------
    
    Call pManager.sort(header.getColumn)
    pSortAscending = pManager.isAscendingSortOrder
    Call updateSortIcon(header)
    
End Sub

Private Sub handleFilterOnTrigger(Source As Variant)
    Dim header As xxGridHeader:         Set header = Source
    '------------------------------------------------------------------------------------------------------
    Call pManager.editFilter(header)
End Sub

Private Sub handleFilterOffTrigger(Source As Variant)
    Dim header As xxGridHeader:         Set header = Source
    '------------------------------------------------------------------------------------------------------
    Call pManager.removeFilter(header)
End Sub

