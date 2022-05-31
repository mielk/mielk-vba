VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "DbEngine"
'----------------------------------------------------------------------------------------------------------
Private pTimeout As Long
Private pMaxAttempts As Long
'[Services] -----------------------------------------------------------------------------------------------
Private pConstructors As FConstructors
Private pConnections As FConnections
Private pFunctions As FFunctions
'Private pRecordsets As FRecordsets
'----------------------------------------------------------------------------------------------------------


'[Initializer]
Private Sub Class_Initialize()
    Set pConstructors = New FConstructors
    Set pConnections = New FConnections
    Set pFunctions = New FFunctions
    'Set pRecordsets = New FRecordsets
End Sub


'[Submodules]
Public Function Create() As FConstructors
    Set Create = pConstructors
End Function

Public Function Connections() As FConnections
    Set Connections = pConnections
End Function

Public Function Functions() As FFunctions
    Set Functions = pFunctions
End Function

'Public Function Recordsets() As FRecordsets
'    Set Recordsets = pRecordsets
'End Function



'[TIMEOUT]

'<simple setter & getter>

Public Sub setTimeout(value As Long)
    pTimeout = value
End Sub

Public Function getTimeout() As Long
    getTimeout = pTimeout
End Function

'</simple setter & getter>





'[ATTEMPTS]

'<simple setter & getter>

Public Sub setMaxAttempts(value As Long)
    pMaxAttempts = value
End Sub

Public Function getMaxAttempts() As Long
    getMaxAttempts = pMaxAttempts
End Function

'</simple setter & getter>
