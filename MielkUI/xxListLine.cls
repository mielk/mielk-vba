VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xxListLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IListener

Private Const CLASS_NAME As String = "xxListLine"
'----------------------------------------------------------------------------------------------------------
Private Const ACTION_NOT_ALLOWED_TAG As String = "Warning_ActionNotAllowed"
'----------------------------------------------------------------------------------------------------------
Private Const LINE_HEIGHT As Single = 18
Private Const SORT_ICON_SIZE As Single = 9
Private Const FONT_FAMILY As String = "Segoe UI"
Private Const FONT_SIZE As Single = 9
Private Const FONT_COLOR As Long = vbBlack
Private Const BORDER_COLOR As Long = 14079702
'----------------------------------------------------------------------------------------------------------
Private pManager As XListManager
Private pFrame As MSForms.frame
Private pItem As IListItem
'----------------------------------------------------------------------------------------------------------
Private pUuid As String
Private pTop As Single
Private pIsRendered As Boolean
Private pEventsDisabled As Boolean
Private pVisibility As Boolean
Private pHasBorder As Boolean
Private pLineHeight As Boolean
'[UI controls] --------------------------------------------------------------------------------------------
Private pBackground As MSForms.label
Private pBorderLine As MSForms.label
Private pControls As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------





Private Sub Class_Initialize()
    '[Set initial values]
    pUuid = F.Crypto.createUuid
    pLineHeight = LINE_HEIGHT
    
    '[Initialize containers]
    Set pControls = F.Dictionaries.Create(False)
End Sub



'[SETTERS]
Public Function setManager(value As XListManager) As xxListLine
    Set setManager = Me
    Set pManager = value
End Function

Public Function setFrame(value As MSForms.frame) As xxListLine
    Set setFrame = Me
    Set pFrame = value
End Function

Public Function setItem(value As IListItem) As xxListLine
    Set setItem = Me
    Set pItem = value
End Function

Public Function setTop(value As Single) As xxListLine
    Dim offsetY As Single
    '------------------------------------------------------------------------------------------------------
    Set setTop = Me
    offsetY = value = pTop
    pTop = value
    If pIsRendered Then Call offsetControls(y:=offsetY)
End Function

Public Function setLineHeight(value As Single) As xxListLine
    Set setLineHeight = Me
    pLineHeight = value
End Function

Public Function setVisibility(value As Boolean) As xxListLine
    Set setVisibility = Me
    pVisibility = value
    Call updateVisibility
End Function

Public Function setHasBorder(value As Boolean) As xxListLine
    Set setHasBorder = Me
    pHasBorder = value
    Call updateBorderVisibility
End Function



'[GETTERS]
Public Function getBottom() As Single
    If UI.Forms.isControlValid(pBackground) Then
        getBottom = pBackground.top + pBackground.height
    End If
End Function





'[RENDERING]
Public Sub render()
    Dim columns As VBA.Collection
    Dim column As DColumn
    Dim ctrl As IControl
    Dim left As Single:             left = pManager.getHorizontalPadding
    '------------------------------------------------------------------------------------------------------
    
    If Not pIsRendered Then
        pEventsDisabled = True
        
        Call insertBackground
        Call insertBorderLine
        
        Set pControls = F.Dictionaries.Create(False)
        Set columns = pManager.getColumnsCollection
        For Each column In columns
            Set ctrl = insertControl(column, left)
            Call pControls.Add(column, ctrl)
        Next column
        
        Call relocateControls
        Call adjustLineHeight
        
        pIsRendered = True
        pEventsDisabled = True
        
    End If
    
End Sub

Private Sub insertBackground()
    Set pBorderLine = pFrame.controls.Add(CREATE_LABEL_ID)
    With pBorderLine
        .BackStyle = fmBackStyleOpaque
        .BackColor = BORDER_COLOR
        .borderStyle = fmBorderStyleNone
        .left = 0
        .width = pFrame.InsideWidth + SCROLL_BAR_WIDTH + PIXEL_SIZE
        .height = PIXEL_SIZE
        .top = pTop
    End With
End Sub


Private Sub insertBorderLine()
    Set pBorderLine = pFrame.controls.Add(CREATE_LABEL_ID)
    With pBorderLine
        .BackStyle = fmBackStyleOpaque
        .BackColor = BORDER_COLOR
        .borderStyle = fmBorderStyleNone
        .left = 0
        .width = pFrame.InsideWidth + SCROLL_BAR_WIDTH + PIXEL_SIZE
        .height = PIXEL_SIZE
        .top = pTop
    End With
End Sub

Private Function insertControl(column As DColumn, ByRef left As Single) As IControl
    If column.getControlType = ControlType_Label Then
        Set insertControl = insertLabel(column, left)
    ElseIf column.getControlType = ControlType_Icon Then
        Set insertControl = insertIcon(column, left)
    End If
End Function

Private Function insertLabel(column As DColumn, ByRef left As Single) As UILabel
    Set insertLabel = UI.Create.UILabel(pFrame)
    With insertLabel
        Call .setTag(column.getName)
        Call .setListener(Me)
        Call .setWidth(column.getWidth)
        Call .setHeight(LINE_HEIGHT)
        Call .setCaption(getCaption(column))
        Call .setFontName(FONT_FAMILY)
        Call .setFontSize(FONT_SIZE)
        Call .setFontColor(column.getFontColor)
        Call .setTextAlignment(column.getTextAlignment)
        Call .setIsUnderlined(column.isUnderlined)
        Call .setMousePointer(column.getMousePointer)
        Call .render
        Call .Move(left)
        Call .alignVertically(pBackground)
        left = .getRight + pManager.getColumnsSpace
    End With
End Function

Private Function insertIcon(column As DColumn, ByRef left As Single) As UIIcon
    Set insertIcon = UI.Create.UIIcon(pFrame)
    With insertIcon
        If pItem.isActionAllowed(column.getName) Then
            Call .setPicture(column.getPicture)
        Else
            Call .setIsInactive(True)
            Call .setPicture(column.getInactivePicture)
        End If
        Call .setHeight(column.getWidth)
        Call .setWidth(column.getWidth)
        Call .setActionName(column.getName)
        Call .setListener(Me)
        Call .setEnability(getControlEnability(column))
        Call .render
        Call .alignVertically(pBackground)
        Call .Move(left)
        Call .setTooltip(column.getTooltip)
        left = .getRight + pManager.getColumnsSpace
    End With
End Function

Private Sub relocateControls()
    Dim key As Variant
    Dim ctrl As IControl
    Dim top As Single
    '------------------------------------------------------------------------------------------------------
    
    If pIsRendered Then
        pBorderLine.top = UI.Normalize.position(pTop)
        pBackground.top = UI.Normalize.position(pTop + PIXEL_SIZE)
        For Each key In pControls.Keys
            Set ctrl = pControls.item(key)
            If Not ctrl Is Nothing Then
                top = pTop + PIXEL_SIZE + (pLineHeight - ctrl.getHeight) / 2
                Call ctrl.Move(y:=top)
                'Call ctrl.alignVertically(pBackground)
            End If
        Next key
    End If
    
End Sub

Private Sub offsetControls(Optional x As Single, Optional y As Single)
    Dim key As Variant
    Dim ctrl As IControl
    '------------------------------------------------------------------------------------------------------
    
    If pIsRendered Then
        If y Then
            pBorderLine.top = UI.Normalize.position(pBorderLine.top)
            pBackground.top = UI.Normalize.position(pBackground.top)
            For Each key In pControls.Keys
                Set ctrl = pControls.item(key)
                If Not ctrl Is Nothing Then
                    Call ctrl.offset(y:=y)
                End If
            Next key
        End If
    End If
    
End Sub

Private Sub adjustLineHeight()
    Dim key As Variant
    Dim ctrl As IControl
    Dim height As Single
    Dim maxHeight As Single
    '------------------------------------------------------------------------------------------------------
    
    For Each key In pControls.Keys
        Set ctrl = pControls.item(key)
        height = ctrl.getHeight
        If height > maxHeight Then
            maxHeight = height
        End If
    Next key
    
    If maxHeight > LINE_HEIGHT Then
        pBackground.height = maxHeight
    Else
        pBackground.height = LINE_HEIGHT
    End If
    
End Sub


Private Function getControlEnability(column As DColumn) As Boolean
    Dim enableFunction As String
    '------------------------------------------------------------------------------------------------------
    
    enableFunction = column.getEnableFunction
    If VBA.Len(enableFunction) Then
        getControlEnability = pItem.getValue(enableFunction)
    Else
        getControlEnability = True
    End If
    
End Function

Private Sub updateVisibility()
    Dim key As Variant
    Dim ctrl As IControl
    '------------------------------------------------------------------------------------------------------
    
    pBackground.visible = pVisibility
    pBorderLine.visible = (pVisibility And pHasBorder)
    
    For Each key In pControls.Keys
        Set ctrl = pControls.item(key)
        Call ctrl.setVisibility(pVisibility)
    Next key
    
End Sub

Private Sub updateBorderVisibility()
    If UI.Forms.isControlValid(pBorderLine) Then
        pBorderLine.visible = (pVisibility And pHasBorder)
    End If
End Sub

Private Function getCaption(column As DColumn) As Variant
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Call F.Variables.assign(value, pItem.getValue(column.getValueFunction))
    If F.Variables.isPrimitive(value) Then
        getCaption = value
    ElseIf VBA.IsObject(value) Then
        If Not value Is Nothing Then
            If TypeOf value Is VBA.Collection Then
                getCaption = F.Collections.toString(F.Cast.toCollection(value), vbCrLf)
            End If
        End If
    Else
        Stop
    End If
    
End Function

Public Sub updateValues()
    Dim columns As VBA.Collection
    Dim column As DColumn
    Dim ctrl As IControl
    Dim label As UILabel
    '------------------------------------------------------------------------------------------------------
    
    Set columns = pManager.getColumnsCollection
    For Each column In columns
        Set ctrl = F.Dictionaries.getObject(pControls, column)
        If Not ctrl Is Nothing Then
            If TypeOf ctrl Is UILabel Then
                Set label = ctrl
                Call label.setCaption(getCaption(column))
            ElseIf TypeOf ctrl Is UIIcon Then
                Call ctrl.setEnability(getControlEnability(column))
            End If
        End If
    Next column
    Call adjustLineHeight
    
End Sub

Public Sub arrangeColumnsLayout()
    Dim columns As VBA.Collection
    Dim column As DColumn
    Dim ctrl As IControl
    Dim left As Single:                 left = pManager.getHorizontalPadding
    '------------------------------------------------------------------------------------------------------
    
    Set columns = pManager.getColumnsCollection
    
    For Each column In columns
        Set ctrl = F.Dictionaries.getObject(pControls, column)
        If Not ctrl Is Nothing Then
            Call ctrl.setWidth(column.getWidth)
            Call ctrl.Move(x:=left)
            left = left + column.getWidth + pManager.getColumnsSpace
        End If
    Next column
    
    pBackground.width = left + SCROLL_BAR_WIDTH + PIXEL_SIZE
    pBorderLine.width = left + SCROLL_BAR_WIDTH + PIXEL_SIZE
    
    Call adjustLineHeight
    
End Sub


Public Sub hideOverflow()
    Dim key As Variant
    Dim control As IControl
    Dim label As UILabel
    '------------------------------------------------------------------------------------------------------
    
    For Each key In pControls.Keys
        Set control = pControls.item(key)
        If TypeOf control Is UILabel Then
            Set label = control
            Call label.hideOverflow
        End If
    Next key
    
End Sub




'[Implementation of IListener interface]

Private Sub IListener_catchEvent(source As Variant, eventName As String, Optional properties As Scripting.IDictionary)
    If TypeOf source Is UIIcon Then
        Call handleButtonClick(UI.Cast.toUiIcon(source))
    ElseIf TypeOf source Is UILabel Then
        Call handleLabelClick(UI.Cast.toUiLabel(source))
    End If
End Sub

Private Sub handleButtonClick(icon As UIIcon)
    Dim isActionAllowed As Boolean
    '------------------------------------------------------------------------------------------------------
    
    If pItem.isActionAllowed(icon.getActionName) Then
        isActionAllowed = True
    End If
    
    If isActionAllowed Then
        Call pManager.runItemAction(pItem, icon.getActionName)
    Else
        Call VBA.MsgBox(MsgService.getText(ACTION_NOT_ALLOWED_TAG), vbExclamation)
    End If
End Sub

Private Sub handleLabelClick(lbl As UILabel)
    Dim tag As String
    Dim column As DColumn
    Dim actionName As String
    '------------------------------------------------------------------------------------------------------
    
    tag = lbl.getTag
    Set column = pManager.getColumn(tag)
    If Not column Is Nothing Then
        actionName = column.getActionName(EVENT_CLICK)
        Call pManager.runItemAction(pItem, actionName)
    End If
    
End Sub
