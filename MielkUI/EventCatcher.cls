VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventCatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "EventCatcher"
'----------------------------------------------------------------------------------------------------------
Private pParent As IControl
Private pBasicControl As Object
Private pEventGuide As GUID
Private pCookie As Long
'----------------------------------------------------------------------------------------------------------


Public Sub setParentAndControl(parent As IControl, ctrl As Object)
    Set pParent = parent
    Set pBasicControl = ctrl
    Call connectAllEvents(True)
End Sub

Public Sub clear()
    If (pCookie <> 0) Then
        Call connectAllEvents(False)
    End If
    Set pBasicControl = Nothing
End Sub

Public Sub connectAllEvents(ByVal connect As Boolean)
    With pEventGuide
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    Call ConnectToConnectionPoint(Me, pEventGuide, connect, pBasicControl, pCookie, 0&)
End Sub





'All Other Control-Events also possible

Public Sub handleEnter()
Attribute handleEnter.VB_UserMemId = -2147384830
    If Not pParent Is Nothing Then
        Call pParent.triggerEvent(UI.Events.Enter)
    End If
End Sub

Public Sub handleExit(ByVal cancel As MSForms.ReturnBoolean)
Attribute handleExit.VB_UserMemId = -2147384829
    If Not pParent Is Nothing Then
        Call pParent.triggerEvent(UI.Events.Leave)
    End If
End Sub

Public Sub handleAfterUpdate()
Attribute handleAfterUpdate.VB_UserMemId = -2147384832
    If Not pParent Is Nothing Then
        Call pParent.triggerEvent(UI.Events.AfterUpdate)
    End If
End Sub

Public Sub handleBeforeUpdate(ByVal cancel As MSForms.ReturnBoolean)
Attribute handleBeforeUpdate.VB_UserMemId = -2147384831
    If Not pParent Is Nothing Then
        Call pParent.triggerEvent(UI.Events.BeforeUpdate)
    End If
End Sub
