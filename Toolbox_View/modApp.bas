Attribute VB_Name = "modApp"
Option Explicit

Private Const CLASS_NAME As String = "modApp"
'----------------------------------------------------------------------------------------------------------


'[FEATURES]
Public Sub quitApp(Optional message As String)
    If VBA.Len(message) Then
        Call VBA.MsgBox(message, vbCritical, APPLICATION_NAME)
    End If
    
    If F.System.isDeveloper Then Stop
    Application.EnableEvents = False
    Call ThisWorkbook.Close(False)
    
End Sub

Public Sub convertToAddIn()
    Call F.Utils.convertToAddIn(ThisWorkbook.FullName, , True)
End Sub







'[FUNCTIONS]
Public Function getSheetsDictionary() As Scripting.Dictionary
    Set getSheetsDictionary = F.dictionaries.Create(False)
    With getSheetsDictionary
        Call .Add(DataTypes.ribbon.getName, wksRibbon)
    End With
End Function
