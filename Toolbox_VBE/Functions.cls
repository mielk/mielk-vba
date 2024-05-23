VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Functions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "Functions"
'----------------------------------------------------------------------------------------------------------





Public Function Create() As FConstructors
    Static instance As FConstructors
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then
        Set instance = New FConstructors
    End If
    Set Create = instance
End Function




Public Function getActiveProject() As VBIDE.VBProject
    Set getActiveProject = ThisWorkbook.VBProject.VBE.ActiveVBProject
End Function

Public Function getActivePane() As VBIDE.CodePane
    Set getActivePane = getActiveProject.VBE.ActiveCodePane
End Function
