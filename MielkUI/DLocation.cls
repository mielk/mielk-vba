VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "DLocation"
'----------------------------------------------------------------------------------------------------------
Private pAnchorParent As Object
Private pAnchorPoint As AnchorPointEnum
Private pAnchorLeft As Single
Private pAnchorTop As Single
Private pAnchorWidth As Single
Private pAnchorHeight As Single
Private pOffsetLeft As Single
Private pOffsetTop As Single
Private pIsDefined As Boolean
'----------------------------------------------------------------------------------------------------------




'[SETTERS]

'<simple setters>
Public Function setAnchorParent(value As Object) As DLocation
    Set setAnchorParent = Me
    Set pAnchorParent = value
End Function

Public Function setAnchorPoint(value As AnchorPointEnum) As DLocation
    Set setAnchorPoint = Me
    pAnchorPoint = value
End Function

Public Function setAnchorLeft(value As Single) As DLocation
    Set setAnchorLeft = Me
    pAnchorLeft = value
End Function

Public Function setAnchorTop(value As Single) As DLocation
    Set setAnchorTop = Me
    pAnchorTop = value
End Function

Public Function setAnchorWidth(value As Single) As DLocation
    Set setAnchorWidth = Me
    pAnchorWidth = value
End Function

Public Function setAnchorHeight(value As Single) As DLocation
    Set setAnchorHeight = Me
    pAnchorHeight = value
End Function


Public Function setOffsetLeft(value As Single) As DLocation
    Set setOffsetLeft = Me
    pOffsetLeft = value
End Function

Public Function setOffsetTop(value As Single) As DLocation
    Set setOffsetTop = Me
    pOffsetTop = value
End Function

'</simple setters>





'[GETTERS]
Public Function getTopLeftPointForUserForm(uf As WTemplate) As Coordinate
    Dim myArea As area
    Dim ufArea As area
    Dim left As Single
    Dim top As Single
    '------------------------------------------------------------------------------------------------------
    Dim finalArea As area
    '------------------------------------------------------------------------------------------------------
    
    myArea = getArea
    ufArea = uf.getSizeAndLocation(True)
    
    With myArea
        If pAnchorPoint = AnchorPoint_Middle Then
            left = .left + (.width - ufArea.width) / 2 + pOffsetLeft
            top = .top + (.height - ufArea.height) / 2 + pOffsetTop
        ElseIf pAnchorPoint = AnchorPoint_TopLeft Then
            left = .left + pOffsetLeft
            top = .top + pOffsetTop
        ElseIf pAnchorPoint = AnchorPoint_TopMiddle Then
            top = .top + pOffsetTop
            left = .left + (.width - ufArea.width) / 2 + pOffsetLeft
        End If
    End With
    
    getTopLeftPointForUserForm = adjustToPreventOverflow(left, top, ufArea.width, ufArea.height)
    
End Function

Private Function adjustToPreventOverflow(ByVal left As Single, ByVal top As Single, _
                                         ByVal width As Single, ByVal height As Single) As Coordinate
    Dim dpiScaling As Coordinate
    Dim screenWorkingArea As area
    Dim cursorPotision As Coordinate
    '------------------------------------------------------------------------------------------------------
    
    dpiScaling = F.Screen.getDpiScaling
    screenWorkingArea = F.Screen.getActiveMonitorWorkingArea
    cursorPotision = F.Screen.getCursorPosition
    
    With screenWorkingArea
        
        '[Adjust screen working area to DPI]
        .width = .width / dpiScaling.x
        .height = .height / dpiScaling.y
        .left = .left / dpiScaling.x
        .top = .top / dpiScaling.y
        
        '[Left]
        If left < .left Then
            left = .left
        ElseIf left + width > .left + .width Then
            left = .left + .width - width
        End If
        
        '[Top]
        If top < .top Then
            top = .top
        ElseIf top + height > .top + .height Then
            top = .top + .height - height
        End If
        
    End With
    
    With adjustToPreventOverflow
        .x = left
        .y = top
    End With
    
End Function


Private Function getArea() As area
    Dim uf As WTemplate
    '------------------------------------------------------------------------------------------------------
    
    With getArea
        .left = pAnchorLeft
        .top = pAnchorTop
        .width = pAnchorWidth
        .height = pAnchorHeight
    End With
    
    If Not pAnchorParent Is Nothing Then
        If TypeOf pAnchorParent Is WTemplate Then
            Set uf = pAnchorParent
            getArea = uf.getSizeAndLocation
        End If
    End If
    
End Function


Public Function isDefined() As Boolean
    isDefined = (pAnchorPoint <> AnchorPoint_None)
End Function
