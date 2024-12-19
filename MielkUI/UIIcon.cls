VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IControl

Private Const CLASS_NAME As String = "UIIcon"
'----------------------------------------------------------------------------------------------------------
Private pFrame As MSForms.frame
Private pListener As IListener
Private WithEvents pLabel As MSForms.label
Attribute pLabel.VB_VarHelpID = -1
'----------------------------------------------------------------------------------------------------------
Private pUuid As String
Private pTag As String
Private pIsRendered As Boolean
'----------------------------------------------------------------------------------------------------------
Private pWidth As Single
Private pHeight As Single
Private pLeft As Single
Private pTop As Single
Private pPicture As Variant
Private pVisible As Boolean
Private pEnabled As Boolean
Private pActionName As String
Private pTooltip As String
Private pHasBorder As Boolean
Private pBorderColor As Long
Private pBackColor As Long
Private pIsTransparent As Boolean
Private pIsInactive As Boolean
'[Events] -------------------------------------------------------------------------------------------------
Event AfterClick()
'----------------------------------------------------------------------------------------------------------




Private Sub Class_Initialize()
    Call setInitialValues
End Sub

Private Sub setInitialValues()
    pUuid = F.Crypto.createUUID
    pVisible = True
    pEnabled = True
End Sub





'[SETTERS]
Public Function setFrame(value As MSForms.frame) As UIIcon
    Set setFrame = Me
    Set pFrame = value
End Function

Public Function setListener(value As IListener) As UIIcon
    Set setListener = Me
    Set pListener = value
End Function

Public Function setActionName(value As String) As UIIcon
    Set setActionName = Me
    pActionName = value
End Function

Public Function setTag(value As String) As UIIcon
    Set setTag = Me
    pTag = value
End Function

Public Function setHeight(value As Single) As UIIcon
    Set setHeight = Me
    pHeight = value
    If UI.controls.isValid(pLabel) Then
        pLabel.height = value
    End If
End Function

Public Function setWidth(value As Single) As UIIcon
    Set setWidth = Me
    pWidth = value
    If UI.controls.isValid(pLabel) Then
        pLabel.width = value
    End If
End Function

Public Function setPicture(value As Variant) As UIIcon
    Set setPicture = Me
    Set pPicture = value
    If UI.controls.isValid(pLabel) Then
        On Error Resume Next
        Set pLabel.picture = pPicture
    End If
End Function

Public Function setTooltip(value As String) As UIIcon
    Set setTooltip = Me
    pTooltip = value
    If UI.controls.isValid(pLabel) Then
        pLabel.ControlTipText = pTooltip
    End If
End Function

Public Function setVisibility(value As Boolean) As UIIcon
    Set setVisibility = Me
    pVisible = value
    Call updateVisibility
End Function

Public Function setEnability(value As Boolean) As UIIcon
    Set setEnability = Me
    pEnabled = value
    Call updateVisibility
End Function

Public Function setHasBorder(value As Boolean) As UIIcon
    Set setHasBorder = Me
    pHasBorder = value
    Call updateBorder
End Function

Public Function setBorderColor(value As Long) As UIIcon
    Set setBorderColor = Me
    pBorderColor = value
    Call updateBorder
End Function

Public Function setBackColor(value As Long) As UIIcon
    Set setBackColor = Me
    pBackColor = value
    Call updateBackground
End Function

Public Function setIsTransparent(value As Boolean) As UIIcon
    Set setIsTransparent = Me
    pIsTransparent = value
    Call updateBackground
End Function

Public Function setIsInactive(value As Boolean) As UIIcon
    Set setIsInactive = Me
    pIsInactive = value
    If pIsRendered Then Call updateMousePointer
End Function




'[GETTERS]

Public Function getControl() As MSForms.label
    Set getControl = pLabel
End Function

Public Function getActionName() As String
    getActionName = pActionName
End Function

Public Function getTag() As String
    getTag = pTag
End Function

Public Function getLeft() As Single
    If UI.controls.isValid(pLabel) Then
        getLeft = pLabel.left
    End If
End Function

Public Function getTop() As Single
    If UI.controls.isValid(pLabel) Then
        getTop = pLabel.top
    End If
End Function

Public Function getRight() As Single
    If UI.controls.isValid(pLabel) Then
        getRight = pLabel.left + pLabel.width
    End If
End Function

Public Function getBottom() As Single
    If UI.controls.isValid(pLabel) Then
        getBottom = pLabel.top + pLabel.height
    End If
End Function

Public Function getHeight() As Single
    If UI.controls.isValid(pLabel) Then
        getHeight = pLabel.height
    End If
End Function

Public Function getWidth() As Single
    If UI.controls.isValid(pLabel) Then
        getWidth = pLabel.width
    End If
End Function

Public Function getVisibility() As Boolean
    getVisibility = pVisible
End Function








'[RENDERING]
Public Function render() As UIIcon
    Set render = Me
    If Not pIsRendered Then
        Set pLabel = pFrame.controls.add(CREATE_LABEL_ID)
        With pLabel
            .height = pHeight
            .width = pWidth
            .BackStyle = fmBackStyleTransparent
            .borderColor = pBorderColor
            .borderStyle = VBA.IIf(pHasBorder, fmBorderStyleSingle, fmBorderStyleNone)
            If Not VBA.IsEmpty(pPicture) Then Set .picture = pPicture
            Call updateMousePointer
            .ControlTipText = pTooltip
        End With
        Call updateVisibility
        pIsRendered = True
    End If
End Function


Private Sub updateMousePointer()
    With pLabel
        If Not pIsInactive Then
            .MousePointer = fmMousePointerCustom
            Set .MouseIcon = ufImages.MouseIcon
        Else
            .MousePointer = fmMousePointerDefault
        End If
    End With
End Sub



Public Function Move(Optional x As Variant, Optional y As Variant) As UIIcon
    Dim offsetX As Single
    Dim offsetY As Single
    '------------------------------------------------------------------------------------------------------
    
    Set Move = Me
    
    With pLabel
        If VBA.IsNumeric(x) Then
            offsetX = (x - .left)
        End If
        If VBA.IsNumeric(y) Then
            offsetY = (y - .top)
        End If
    End With
    
    Call offset(offsetX, offsetY)
    
End Function

Public Function offset(Optional x As Single, Optional y As Single) As UIIcon
    Set offset = Me
    Call UI.Forms.offsetControls(x, y, pLabel)
    pTop = pLabel.top
End Function

Public Function alignVertically(ctrl As MSForms.control) As UIIcon
    Dim offset As Single
    '------------------------------------------------------------------------------------------------------
    
    Set alignVertically = Me
    
    If UI.controls.isValid(ctrl) Then
        offset = (ctrl.top + (ctrl.height - pHeight) / 2) - pTop
    End If
    
    If offset Then
        If UI.controls.isValid(pLabel) Then
            pLabel.top = pLabel.top + offset
            pTop = pLabel.top
        End If
    End If
    
End Function






Private Sub updateVisibility()
    If UI.controls.isValid(pLabel) Then
        pLabel.visible = (pVisible And pEnabled)
    End If
End Sub

Private Sub updateBorder()
    If UI.controls.isValid(pLabel) Then
        With pLabel
            .borderColor = pBorderColor
            .borderStyle = VBA.IIf(pHasBorder, fmBorderStyleSingle, fmBorderStyleNone)
        End With
    End If
End Sub

Private Sub updateBackground()
    If UI.controls.isValid(pLabel) Then
        With pLabel
            .BackStyle = VBA.IIf(pIsTransparent, fmBackStyleTransparent, fmBackStyleOpaque)
            .backColor = pBackColor
        End With
    End If
End Sub

Public Sub delete()
    Call pFrame.controls.remove(pLabel.name)
End Sub













'[Controls' events]

Private Sub pLabel_MouseUp(ByVal Button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
    Dim properties As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    Call ErrorManager.clear
    
    If pIsInactive Then
        ' skip
    ElseIf pListener Is Nothing Then
        ' skip
    Else
        Set properties = F.dictionaries.createWithItems(False, _
                                                        KeyValue(TAG_BUTTON, Button), _
                                                        KeyValue(props.shift.getName, shift), _
                                                        KeyValue(TAG_X, x), _
                                                        KeyValue(TAG_Y, y))
        Call pListener.catchEvent(Me, Events.MouseUp.getName, properties)
    End If
    
End Sub

Private Sub pLabel_Click()
    Call ErrorManager.clear
    
    RaiseEvent AfterClick
    
End Sub







'[Implementation of IControl interface]
Private Sub IControl_setVisibility(value As Boolean)
    Call setVisibility(value)
End Sub

Private Sub IControl_setEnability(value As Boolean)
    Call setEnability(value)
End Sub

Private Sub IControl_setWidth(value As Single)
    Call setWidth(value)
End Sub

Private Function IControl_getHeight() As Single
    IControl_getHeight = getHeight
End Function

Private Function IControl_getWidth() As Single
    IControl_getWidth = getWidth
End Function

Private Function IControl_getLeft() As Single
    IControl_getLeft = getLeft
End Function

Private Function IControl_getTop() As Single
    IControl_getTop = getTop
End Function

Private Function IControl_getRectangle() As RECT
    Stop
End Function

Private Sub IControl_resize(Optional width As Variant, Optional height As Variant)
    Stop
End Sub

Private Sub IControl_alignVertically(baseCtrl As MSForms.IControl)
    Call alignVertically(baseCtrl)
End Sub

Private Function IControl_move(Optional x As Variant, Optional y As Variant) As IControl
    Set IControl_move = Move(x, y)
End Function

Private Function IControl_offset(Optional x As Single, Optional y As Single) As IControl
    Set IControl_offset = offset(x, y)
End Function

Private Sub IControl_triggerEvent(eventId As Long)
    Stop
End Sub
