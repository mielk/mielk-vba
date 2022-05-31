VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufTemplate 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ufTemplate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "ufTemplate"
'----------------------------------------------------------------------------------------------------------
Private pMethodInvoker As Object
Private pMethodName As String
Private pMethodParams As Scripting.Dictionary
Private pMethodErrorMessage As String
Private pCloseWindowAfterward As Boolean
'----------------------------------------------------------------------------------------------------------



Private Sub UserForm_Initialize()
    Me.BackColor = TRANSPARENCY_LAYER_COLOR
    Me.StartUpPosition = 0
End Sub

Private Sub UserForm_Activate()
    Call ErrorManager.clear
    
    Call UI.Forms.hideTitleBarAndBorder(Me)
    Call UI.Forms.makeUserFormTransparent(Me, TRANSPARENCY_LAYER_COLOR)
    
    '[Run underlying method if it is specified]
    If VBA.Len(pMethodName) Then
        If Not pMethodInvoker Is Nothing Then
            Call VBA.CallByName(pMethodInvoker, pMethodName, VbMethod, pMethodParams)
        Else
            Call Excel.Application.run(pMethodName, pMethodParams)
        End If
    End If
    
    If pCloseWindowAfterward Then
        On Error Resume Next
        Call Me.hide
    End If
    
    
ExitPoint:
    'Call errorManager.save
    'Call me.Hide
    
End Sub




'[SETTERS]
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



Public Sub setUnderlyingMethod(methodName As String, Optional invoker As Object, _
                                    Optional params As Scripting.Dictionary, _
                                    Optional closeWindowAfterward As Boolean = False, _
                                    Optional methodErrorMessage As String)
    pMethodName = methodName
    Set pMethodInvoker = invoker
    Set pMethodParams = params
    pCloseWindowAfterward = closeWindowAfterward
    pMethodErrorMessage = methodErrorMessage
End Sub





'[GETTERS]
Public Function getWidth() As Single
    getWidth = Me.width
End Function

Public Function getHeight() As Single
    getHeight = Me.height
End Function

Public Function getTop() As Single
    getTop = Me.top
End Function

Public Function getLeft() As Single
    getLeft = Me.left
End Function

Public Function isVisible() As Boolean
    isVisible = Me.visible
End Function






Public Sub toFront()
    Call BringWindowToTop(UI.Forms.getWindowHandle(Me))
End Sub
