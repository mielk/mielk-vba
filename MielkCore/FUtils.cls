VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FUtils"
'----------------------------------------------------------------------------------------------------------
Private Const BACKUP_FOLDER_PATH As String = "E:\dane\VBA\backups"
'----------------------------------------------------------------------------------------------------------

Public Sub printfilepaths()
    Dim addIn As Excel.addIn
    '------------------------------------------------------------------------------------------------------
    
    For Each addIn In Application.addIns2
        If addIn.isOpen Then
            Debug.Print addIn.fullName
        End If
    Next addIn
    
End Sub

Public Sub saveAllAddIns(Optional ByVal saveBackups As Boolean = False, Optional ByVal saveFolder As String = BACKUP_FOLDER_PATH)
    Const REGEX_MACRO_ENABLED As String = "\.xlsm$"
    '------------------------------------------------------------------------------------------------------
    Dim addIn As Excel.addIn
    Dim wkb As Excel.Workbook
    Dim Timestamp As Date:                          Timestamp = VBA.Now
    '------------------------------------------------------------------------------------------------------
    
    'Save all add-ins
    For Each addIn In Application.addIns2
        If Not isAddInSkippedFromSaving(addIn.name) Then
            If addIn.isOpen Then
                On Error Resume Next
                Set wkb = Application.Workbooks(addIn.name)
                On Error GoTo 0
                
                If wkb.name = addIn.name Then
                    Call wkb.save
                    Debug.Print wkb.fullName & " saved"
                    If saveBackups Then
                        Call saveBackup(wkb, Timestamp)
                    End If
                End If
            End If
        End If
    Next addIn
    
    'Save all xlam files.
    For Each wkb In Application.Workbooks
        If f.regex.checkIfMatch(wkb.name, REGEX_MACRO_ENABLED) Then
            Call wkb.save
            Debug.Print wkb.fullName & " saved"
            If saveBackups Then
                Call saveBackup(wkb, Timestamp)
            End If
        End If
    Next wkb
    
    Debug.Print "Saving completed [" & VBA.format(VBA.Now, "YYYY-MM-DD HH:MM:SS") & "]"
    
End Sub

Private Function isAddInSkippedFromSaving(ByVal addinName As String) As Boolean
    If VBA.InStr(1, addinName, "solver", vbTextCompare) Then
        isAddInSkippedFromSaving = True
    ElseIf VBA.InStr(1, addinName, ".xll", vbTextCompare) Then
        isAddInSkippedFromSaving = True
    Else
        isAddInSkippedFromSaving = False
    End If
End Function

Private Sub saveBackup(wkb As Excel.Workbook, Timestamp As Date)
    Static fso As Scripting.FileSystemObject
    '------------------------------------------------------------------------------------------------------
    Dim folderName As String
    Dim folderPath As String
    Dim folder As Scripting.folder
    Dim path As String
    '------------------------------------------------------------------------------------------------------
    
    If fso Is Nothing Then
        Set fso = New Scripting.FileSystemObject
    End If
    
    folderName = VBA.format$(Timestamp, "YYYY-MM-DD HHMMSS")
    folderPath = f.files.concatenatePath(BACKUP_FOLDER_PATH, folderName)
    If fso.folderExists(folderPath) Then
        Set folder = fso.getFolder(folderPath)
    Else
        Set folder = fso.createFolder(folderPath)
    End If
    path = f.files.concatenatePath(folder.path, wkb.name)
    
    Call fso.CopyFile(wkb.fullName, path, True)
    
End Sub


Public Sub unfreezeExcel()
    With Excel.Application
        .Interactive = True
        .EnableEvents = True
        .ScreenUpdating = True
        .cursor = xlDefault
    End With
End Sub




Public Function convertToAddIn(originalfilepath As String, _
                                Optional ByVal destinationPath As String, _
                                Optional overwriteIfExists As Boolean = False, _
                                Optional ByVal xls As Excel.Application) As String
    Dim wkb As Excel.Workbook
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(destinationPath) = 0 Then
        destinationPath = VBA.replace(originalfilepath, ".xlsm", ".xlam")
    End If
    
    If Not overwriteIfExists Then
        destinationPath = f.files.uniquefilepath(f.files.getUncPath(destinationPath))
    End If
    
    If xls Is Nothing Then Set xls = f.Excel.getBackgroundInstance
    
    Set wkb = f.Books.open_(originalfilepath, False, xls)
    With wkb
        .IsAddin = True
        .CheckCompatibility = False
        Call f.files.deleteFile(destinationPath)
        Call .SaveAs(destinationPath, xlAddIn)
        Call .Close(False)
    End With
    
End Function
