VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xxSubitemLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "xxSubitemLine"
'----------------------------------------------------------------------------------------------------------
Private Const DEFAULT_FONT_FAMILY As String = "Segoe UI"
Private Const DEFAULT_FONT_SIZE As Single = 10
Private Const DEFAULT_FONT_COLOR As Long = 6579300
Private Const DEFAULT_FONT_BOLD As Boolean = False
Private Const DEFAULT_ALIGNMENT As Variant = fmTextAlignCenter
'[Textbox]
Private Const TEXTBOX_BACK_COLOR As Long = 15132390
Private Const TEXTBOX_BORDER_COLOR As Long = 0
Private Const DEFAULT_TEXTBOX_WIDTH As Long = 100
'[Action button]
Private Const DEFAULT_BUTTON_WIDTH As Single = 10.5
Private Const DEFAULT_BUTTON_HEIGHT As Single = 10.5
Private Const DEFAULT_BUTTON_RIGHT_MARGIN As Single = 3
'----------------------------------------------------------------------------------------------------------
Private pParent As XSubitemsPanel
Private pUuid As String
Private pSubitem As String
'----------------------------------------------------------------------------------------------------------
Private WithEvents pDeleteIcon As MSForms.label
Attribute pDeleteIcon.VB_VarHelpID = -1
Private pNameBox As MSForms.textBox
'[UI properties] ------------------------------------------------------------------------------------------
Private pTop As Single
Private pLeft As Single
Private pNameBoxWidth As Single
Private pFontFamily As String
Private pFontSize As Single
Private pFontColor As Long
Private pFontBold As Boolean
'----------------------------------------------------------------------------------------------------------
Private pIsRendered As Boolean
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    pUuid = F.Crypto.createUUID
    pFontFamily = DEFAULT_FONT_FAMILY
    pFontSize = DEFAULT_FONT_SIZE
    pFontColor = DEFAULT_FONT_COLOR
    pFontBold = DEFAULT_FONT_BOLD
    pNameBoxWidth = DEFAULT_BUTTON_WIDTH
End Sub



'[SETTERS]
Friend Function setParent(value As XSubitemsPanel) As xxSubitemLine
    Set setParent = Me
    Set pParent = value
End Function

Friend Function setSubitem(value As String) As xxSubitemLine
    Set setSubitem = Me
    pSubitem = value
End Function

Friend Function setTop(value As Single) As xxSubitemLine
    Set setTop = Me
    pTop = value
    Call arrangeControls
End Function

Friend Function setLeft(value As Single) As xxSubitemLine
    Set setLeft = Me
    pLeft = value
    Call arrangeControls
End Function

Friend Function setFontColor(value As Long) As xxSubitemLine
    Set setFontColor = Me
    pFontColor = value
End Function



'[GETTERS]
Private Function getFrame() As MSForms.frame
    Set getFrame = pParent.getInnerFrame
End Function

Public Function getRight() As Single
    getRight = UI.Forms.getRight(pNameBox)
End Function

Public Function getParent() As XSubitemsPanel
    Set getParent = pParent
End Function

Public Function getTop() As Single
    getTop = pNameBox.top
End Function

Public Function getSubitem() As String
    getSubitem = pSubitem
End Function



'[RENDERING]
Public Sub render()
    If Not pIsRendered Then
        Call insertDeleteIcon
        Call insertNameBox
        pIsRendered = True
        Call arrangeControls
    End If
End Sub
    
Private Sub insertDeleteIcon()
    Set pDeleteIcon = getFrame.Controls.Add(CREATE_LABEL_ID)
    With pDeleteIcon
        .BackStyle = fmBackStyleTransparent
        .width = DEFAULT_BUTTON_WIDTH
        .height = DEFAULT_BUTTON_HEIGHT
        .MousePointer = fmMousePointerCustom
        Set .MouseIcon = UI.Icons.getMousePointer
        Set .Picture = UI.Icons.getDeleteIcon
    End With
End Sub

Private Sub insertNameBox()
    Set pNameBox = getFrame.Controls.Add(CREATE_TEXTBOX_ID)
    With pNameBox
        .BorderStyle = fmBorderStyleSingle
        .width = pNameBoxWidth
        With .Font
            .name = pFontFamily
            .size = pFontSize
            .Bold = 0
        End With
        .ForeColor = pFontColor
        .TextAlign = fmTextAlignLeft
        .value = pSubitem
        .Locked = 1
    End With
End Sub

Public Sub move(left As Single, top As Single)
    pLeft = left
    pTop = top
    Call arrangeControls
End Sub

Public Sub arrangeControls()
    Dim lineHeight As Single:               lineHeight = pParent.getLineHeight
    '------------------------------------------------------------------------------------------------------
    
    If pIsRendered Then
        Call UI.Forms.locateControl(pDeleteIcon, pLeft, pTop + (lineHeight - pDeleteIcon.height) / 2)
        Call UI.Forms.locateControl(pNameBox, _
                            UI.Forms.getRight(pDeleteIcon) + DEFAULT_BUTTON_RIGHT_MARGIN, _
                            pTop + (lineHeight - pNameBox.height) / 2)
        Call UI.Forms.resizeControl(pNameBox, pParent.getWidth - pNameBox.left - pParent.getHorizontalPadding)
    End If
    
End Sub


Private Sub pDeleteIcon_Click()
    Call destroyControls
    Call pParent.deleteSubitem(pSubitem)
End Sub


Private Sub destroyControls()
    Dim frame As MSForms.frame
    '------------------------------------------------------------------------------------------------------
    
    Set frame = getFrame
    If UI.Forms.isControlValid(frame) Then
        With frame.Controls
            On Error Resume Next
            Call .Remove(pDeleteIcon.name)
            Call .Remove(pNameBox.name)
        End With
    End If
End Sub


