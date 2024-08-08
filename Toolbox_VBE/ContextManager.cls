VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContextManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "ContextManager"
'----------------------------------------------------------------------------------------------------------
Private pActiveMenuItems As VBA.Collection
'----------------------------------------------------------------------------------------------------------

Private Sub Class_Initialize()
    Set pActiveMenuItems = New VBA.Collection
End Sub

Public Sub addItem(submenuCaption As String, itemCaption As String, _
                    OnAction As String, Optional faceId As Long)
    Dim menu As Office.CommandBar
    Dim subMenu As Office.commandBarPopup
    Dim barItem As Office.CommandBarButton
    Dim item As ControlWithEvent
    '------------------------------------------------------------------------------------------------------
    
    Set item = New ControlWithEvent
    
    Set menu = Excel.Application.VBE.CommandBars(CODE_WINDOW_NAME)
    Set subMenu = Fn.Create.commandBarPopup(menu, CUSTOM_MENU_CAPTION)
    Set barItem = subMenu.Controls.Add(msoControlButton)
    
    With subMenu
        .caption = submenuCaption
        .Tag = CUSTOM_MENU_ITEM_TAG
    End With
    
    With barItem
        .caption = itemCaption
        .faceId = faceId
        .OnAction = F.Reflection.getFunctionFullName(OnAction, Excel.ThisWorkbook)
        .Tag = CUSTOM_MENU_ITEM_TAG
    End With
    
    Set item.pControl = Application.VBE.Events.CommandBarEvents(barItem)
    Call pActiveMenuItems.Add(item)
    
End Sub


Public Sub removeCustomMenuItems()
    Dim Controls As Office.CommandBarControls
    Dim control As Office.CommandBarControl
    '------------------------------------------------------------------------------------------------------
    
    Set Controls = Excel.Application.VBE.CommandBars.FindControls(Tag:=CUSTOM_MENU_ITEM_TAG)
    
    If Not Controls Is Nothing Then
        For Each control In Controls
           On Error Resume Next
           Call control.Delete
        Next control
    End If
    
    Call removeCommandControlWithoutCaption
    
End Sub


Private Sub removeCommandControlWithoutCaption()
    Dim menu As Office.CommandBar
    Dim control As Office.CommandBarControl
    '------------------------------------------------------------------------------------------------------
    
    Set menu = Excel.Application.VBE.CommandBars(CODE_WINDOW_NAME)
    For Each control In menu.Controls
        If control.caption = vbNullString Then
            On Error Resume Next
            Call control.Delete
        End If
    Next control
    
End Sub


Public Sub removeMenuItemById(id As Long)
    Dim Controls As Office.CommandBarControls
    Dim control As Office.CommandBarControl
    '------------------------------------------------------------------------------------------------------
    
    Set Controls = Excel.Application.VBE.CommandBars.FindControls(id:=id)
    
    For Each control In Controls
        On Error Resume Next
        Call control.Delete
    Next control
    
End Sub
