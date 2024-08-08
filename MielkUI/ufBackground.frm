VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufBackground 
   Caption         =   "UserForm1"
   ClientHeight    =   3012
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4560
   OleObjectBlob   =   "ufBackground.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "ufBackground"
'----------------------------------------------------------------------------------------------------------
Private pUuid As String
Private pWindow As WTemplate
Private pIsActivated As Boolean
'----------------------------------------------------------------------------------------------------------



Private Sub UserForm_Initialize()
    pUuid = F.Crypto.createUUID
    Me.caption = pUuid
    Me.backColor = TRANSPARENCY_LAYER_COLOR
    Me.StartUpPosition = 0
End Sub

Private Sub UserForm_Activate()
    Dim frm As ufTemplate
    '----------------------------------------------------------------------------------------------------------
    
    If Not pIsActivated Then
        pIsActivated = True
        
        Call UI.Forms.HideTitleBarAndBorder(Me)
        Call UI.Forms.makeUserFormTransparent(Me)
        
        Set frm = pWindow.getForm
        With frm
            Me.left = .left
            Me.top = .top
            Me.width = .InsideWidth - 2 * PIXEL_SIZE
            Me.height = .InsideHeight - 2 * PIXEL_SIZE
            Call .show(vbModal)
        End With
    End If
    
End Sub



'[SETTERS]
Public Sub setWindow(value As WTemplate)
    Set pWindow = value
End Sub

Public Sub setWidth(value As Single)
    Me.width = value
End Sub

Public Sub setHeight(value As Single)
    Me.height = value
End Sub

Public Sub setTop(value As Single)
    Me.top = value
End Sub

Public Sub setLeft(value As Single)
    Me.left = value
End Sub

Friend Sub makeOpaque()
    Me.backColor = vbBlack
End Sub

Friend Sub makeTransparent()
    Me.backColor = TRANSPARENCY_LAYER_COLOR
End Sub
