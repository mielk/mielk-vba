VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Toolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "Toolbox"
'----------------------------------------------------------------------------------------------------------







Public Sub createNewProject(sheets As Scripting.Dictionary)
    Dim project As EProject
    '------------------------------------------------------------------------------------------------------
    
    Set project = New EProject
    With project
        If .isValid Then
            Call .create(sheets)
        End If
    End With
    
End Sub
