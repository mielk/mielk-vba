VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeComparisonManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "CodeComparisonManager"
'----------------------------------------------------------------------------------------------------------
Private pBaseFilePath As String
Private pCompareFilePath As String
Private pBaseProject As EVbaProject
Private pCompareProject As EVbaProject
'----------------------------------------------------------------------------------------------------------
Private Const PROGRESS_BAR_COMPARING_FILES As String = "ComparingCode.ProgressBar.ComparingFiles"
Private Const PROGRESS_BAR_LOADING_BASE_FILE As String = "ComparingCode.ProgressBar.LoadingBaseFile"
Private Const PROGRESS_BAR_LOADING_COMPARE_FILE As String = "ComparingCode.ProgressBar.LoadingCompareFile"
Private Const PROGRESS_BAR_PRINTING_RESULT As String = "ComparingCode.ProgressBar.PrintingResult"
Private Const SELECT_BASE_FILE_CAPTION As String = "ComparingCode.WindowSelectBaseFile.Title"
Private Const SELECT_COMPARE_FILE_CAPTION As String = "ComparingCode.WindowSelectCompareFile.Title"
'----------------------------------------------------------------------------------------------------------

Public Sub run()
    Dim baseFilePath As String
    Dim comparedFilePath As String
    '------------------------------------------------------------------------------------------------------
    
    Call selectFiles
    
    If checkIfValidFilePaths Then
        'Call ProgressBar.reset.setWidth(300). _
                            setCaption(Msg.getText(PROGRESS_BAR_COMPARING_FILES)). _
                            setAction("compareProjects", Me). _
                            run
        Call compareProjects
    End If
    
End Sub



Private Sub selectFiles()
    pBaseFilePath = f.files.selectFile(pBaseFilePath, Msg.getText(SELECT_BASE_FILE_CAPTION), FILES_PATTERN_EXCEL_MACRO)
    If VBA.Len(pBaseFilePath) Then
        pCompareFilePath = f.files.selectFile(pCompareFilePath, Msg.getText(SELECT_COMPARE_FILE_CAPTION), FILES_PATTERN_EXCEL_MACRO)
    End If
End Sub

Private Function checkIfValidFilePaths() As Boolean
    If Not f.files.FileExists(pBaseFilePath) Then
        checkIfValidFilePaths = False
    ElseIf Not f.files.FileExists(pCompareFilePath) Then
        checkIfValidFilePaths = False
    Else
        checkIfValidFilePaths = True
    End If
End Function



Public Sub compareProjects(Optional params As Scripting.Dictionary)
    Dim comparison As EProjectsComparison
    '------------------------------------------------------------------------------------------------------
        
Call ProgressBar.openSubProcess(0, 0.3)
    Call loadProjects
Call ProgressBar.closeCurrentSubProcess

Call ProgressBar.setCaption(Msg.getText(PROGRESS_BAR_COMPARING_FILES)).openSubProcess(0.3, 0.6)
    Set comparison = newProjectsComparison.run(pBaseProject, pCompareProject)
Call ProgressBar.closeCurrentSubProcess

Call ProgressBar.setCaption(Msg.getText(PROGRESS_BAR_PRINTING_RESULT)).openSubProcess(0.6, 1)
    Call CodeComparisonPrinter.run(comparison)
Call ProgressBar.closeCurrentSubProcess

End Sub


Private Sub loadProjects()

Call ProgressBar.setCaption(f.Strings.Format(Msg.getText(PROGRESS_BAR_LOADING_BASE_FILE), pBaseFilePath)).setCaptionLocked(True)
Call ProgressBar.openSubProcess(0, 0.5)
    Set pBaseProject = createVbaProjectFromPath(pBaseFilePath)
Call ProgressBar.closeCurrentSubProcess


Call ProgressBar.setCaption(f.Strings.Format(Msg.getText(PROGRESS_BAR_LOADING_COMPARE_FILE), pCompareFilePath)).setCaptionLocked(True)
Call ProgressBar.openSubProcess(0.5, 1)
    Set pCompareProject = createVbaProjectFromPath(pCompareFilePath)
Call ProgressBar.closeCurrentSubProcess

End Sub
