VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ESession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "ESession"
'----------------------------------------------------------------------------------------------------------
Private Const LOADING_APP_SETTINGS As String = "App.Open.LoadSettings"
'----------------------------------------------------------------------------------------------------------
Private pStartTime As Date
'----------------------------------------------------------------------------------------------------------


Private Sub Class_Initialize()
    pStartTime = VBA.Now
End Sub





Public Sub setup(Optional params As Scripting.Dictionary)
    'Implement your own logic
End Sub
