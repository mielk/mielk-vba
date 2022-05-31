VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FContextMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FContextMenu"
'----------------------------------------------------------------------------------------------------------
Private Const MENU_ITEM_PREFIX As String = "custom_"
Private Const CONTEXT_MENU_COMMAND_BAR_NAME As String = "Cell"
'[Tags]
Private Const TAG_METHOD_NAME As String = "methodName"
Private Const TAG_CAPTION As String = "caption"
Private Const TAG_TAG As String = "tag"
Private Const TAG_FACE_ID As String = "faceId"
Private Const TAG_SEPARATOR As String = "separator"
Private Const TAG_COLUMNS As String = "columns"
'----------------------------------------------------------------------------------------------------------


Public Sub addContextMenuItemFromDto(ByRef dto As Scripting.Dictionary)
    Dim methodName As String
    Dim buttonCaption As String
    Dim buttonTag As String
    Dim faceId As Long
    Dim separateFromNext As Boolean
    '------------------------------------------------------------------------------------------------------
    
    methodName = F.Dictionaries.getItem(dto, TAG_METHOD_NAME)
    buttonCaption = MsgService.getText(F.Dictionaries.getItem(dto, TAG_CAPTION))
    buttonTag = F.Dictionaries.getItem(dto, TAG_TAG)
    faceId = F.Dictionaries.getItem(dto, TAG_FACE_ID)
    separateFromNext = F.Dictionaries.getItem(dto, TAG_SEPARATOR)
    
    Call addContextMenuItem(methodName, buttonCaption, buttonTag, faceId, 1, separateFromNext)

End Sub


Public Sub addContextMenuItem(methodName As String, buttonCaption As String, buttonTag As String, _
                            Optional faceId As Long, Optional before As Long = 1, _
                            Optional separateFromNext As Boolean = False)
    Dim contextMenu As Object:                      Set contextMenu = getContextMenu
    Dim contextMenuItem As Object
    '------------------------------------------------------------------------------------------------------
    
    Set contextMenuItem = getContextMenuItem(buttonTag)
    If Not contextMenuItem Is Nothing Then
        contextMenuItem.OnAction = methodName
    Else
        With contextMenu.Controls.Add(Type:=msoControlButton, before:=before)
            .OnAction = methodName
            .faceId = faceId
            .Caption = buttonCaption
            .tag = buttonTag
        End With
        
        If separateFromNext Then
            On Error Resume Next
            contextMenu.Controls(before + 1).BeginGroup = True
        End If
    End If
    
End Sub

Private Function getContextMenu() As Object
    Set getContextMenu = Excel.Application.CommandBars(CONTEXT_MENU_COMMAND_BAR_NAME)
End Function

Private Function getContextMenuItem(tag As String) As Object
    Dim contextMenu As Object:                      Set contextMenu = getContextMenu
    Dim ctrl As Object
    '------------------------------------------------------------------------------------------------------
    For Each ctrl In contextMenu.Controls
        If ctrl.tag = tag Then
            Set getContextMenuItem = ctrl
            Exit For
        End If
    Next ctrl
End Function

Public Sub clearContextMenu()
    Dim contextMenu As Object:                      Set contextMenu = getContextMenu
    Dim item As Object
    '------------------------------------------------------------------------------------------------------

    If Not contextMenu Is Nothing Then
        For Each item In contextMenu.Controls
            Call item.delete
        Next item
    End If
    
    Application.ShowMenuFloaties = True

End Sub

Public Sub resetContextMenu()
    Dim contextMenu As Object:                      Set contextMenu = getContextMenu
    '------------------------------------------------------------------------------------------------------
    
    If Not contextMenu Is Nothing Then
        Call contextMenu.Reset
    End If
    
    Application.ShowMenuFloaties = False
    
End Sub

Public Sub clearContextMenuFromCustomItems(Optional projectPrefix As String)
    Call deleteContextMenuByRegex("^" & MENU_ITEM_PREFIX, "^" & projectPrefix)
End Sub

Private Sub deleteContextMenuByRegex(ParamArray items() As Variant)
    Dim contextMenu As Object:                      Set contextMenu = getContextMenu
    Dim objRegex As Object
    Dim ctrl As Object
    Dim varItem As Variant
    '------------------------------------------------------------------------------------------------------

    Set objRegex = F.regex.Create
    
    For Each varItem In items
        objRegex.pattern = VBA.CStr(varItem)
        For Each ctrl In contextMenu.Controls
            If objRegex.test(ctrl.tag) Then
                Call ctrl.delete
            End If
        Next ctrl
    Next varItem

    If contextMenu.Controls.count Then
        contextMenu.Controls(1).BeginGroup = False
    End If

End Sub
