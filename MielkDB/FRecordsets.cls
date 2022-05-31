VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FRecordsets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FRecordsets"
'----------------------------------------------------------------------------------------------------------

Public Sub closeAndDestroy(ByRef rs As ADODB.recordset)
    If Not rs Is Nothing Then
        If rs.State = 1 Then
            Call rs.Close
        End If
        Set rs = Nothing
    End If
End Sub
