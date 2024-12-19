VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FConstructors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FConstructors"
'----------------------------------------------------------------------------------------------------------


Public Function commandBarPopup(parent As Office.CommandBar, caption As String) As Office.commandBarPopup
    Const MAX_ATTEMPT As Long = 10
    '------------------------------------------------------------------------------------------------------
    Dim attempt As Long
    '------------------------------------------------------------------------------------------------------
    
    Do
        Set commandBarPopup = getCommandBarControlByCaption(parent, CUSTOM_MENU_CAPTION)
        If commandBarPopup Is Nothing Then
            Set commandBarPopup = parent.Controls.Add(Office.msoControlPopup, , , 1)
        End If
        
        attempt = attempt + 1
        If attempt > MAX_ATTEMPT Then Exit Do
        
    Loop Until isValidCommandBarPopup(commandBarPopup)
    
End Function

Private Function getCommandBarControlByCaption(parent As Office.CommandBar, caption As String) As Office.CommandBarControl
    Dim ctrl As Office.CommandBarControl
    '------------------------------------------------------------------------------------------------------
    
    For Each ctrl In parent.Controls
        If ctrl.caption = caption Then
            Set getCommandBarControlByCaption = ctrl
            Exit For
        End If
    Next ctrl
    
End Function

Private Function isValidCommandBarPopup(popup As Office.commandBarPopup) As Boolean
    Dim controlsCounter As Long:                controlsCounter = -1
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    controlsCounter = popup.Controls.count
    
    isValidCommandBarPopup = (controlsCounter >= 0)

End Function




Public Function codeSelection() As ECodeSelection
    Set codeSelection = New ECodeSelection
End Function

Public Function currentCodeSelection() As ECodeSelection
    Set currentCodeSelection = New ECodeSelection
    Call currentCodeSelection.readFromActivePane
End Function
