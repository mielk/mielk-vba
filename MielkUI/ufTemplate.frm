VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufTemplate 
   Caption         =   "UserForm1"
   ClientHeight    =   3012
   ClientLeft      =   120
   ClientTop       =   468
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
Event AfterDisplayed()
'----------------------------------------------------------------------------------------------------------


Private Sub UserForm_Initialize()
    Me.backColor = TRANSPARENCY_LAYER_COLOR
    Me.StartUpPosition = 0
End Sub

Private Sub UserForm_Activate()
    Dim errNumber As Long, errDescription As String
    '------------------------------------------------------------------------------------------------------
    
    Call ErrorManager.clear
    
    Call UI.Forms.hideTitleBarAndBorder(Me)
    Call UI.Forms.makeUserFormTransparent(Me, TRANSPARENCY_LAYER_COLOR)
    
    '[Run underlying method if it is specified]
    If VBA.Len(pMethodName) Then
        Call VBA.Err.clear
        
        On Error Resume Next
        If Not pMethodInvoker Is Nothing Then
            Call VBA.CallByName(pMethodInvoker, pMethodName, VbMethod, pMethodParams)
        Else
            Call Excel.Application.run(pMethodName, pMethodParams)
        End If
        
        errNumber = VBA.Err.Number
        
        If Not DEV_MODE Then On Error GoTo ErrHandler
        
        If errNumber = Exceptions.WrongNumberOfArguments.getNumber Then
            errNumber = 0
            If Not pMethodInvoker Is Nothing Then
                Call VBA.CallByName(pMethodInvoker, pMethodName, VbMethod)
            Else
                Call Excel.Application.run(pMethodName)
            End If
        End If
        
    End If
    
    If pCloseWindowAfterward Then
        On Error Resume Next
        Call Me.hide
    End If
    
    RaiseEvent AfterDisplayed
    
ExitPoint:
    'Call errorManager.save
    'Call me.Hide
ErrHandler:

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
