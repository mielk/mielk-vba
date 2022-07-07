VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "ActionListener"
'----------------------------------------------------------------------------------------------------------
Event AfterClick(callerName As String, sheet As Excel.Worksheet, row As Long, column As Long)
'----------------------------------------------------------------------------------------------------------

Public Sub click(callerName As String)
    Dim wks As Excel.Worksheet
    Dim shape As Excel.shape
    Dim cell As Excel.Range
    '----------------------------------------------------------------------------------------------------------
    
    Set wks = Excel.ActiveSheet
    Set shape = F.Sheets.getShape(wks, callerName)

    If Not shape Is Nothing Then
        Set cell = shape.TopLeftCell
        RaiseEvent AfterClick(callerName, wks, cell.row, cell.column)
    End If

End Sub
