VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "ProjectCreator"
'[Config] -------------------------------------------------------------------------------------------------
Private Const CONFIG_NAME As String = CONFIG_NEW_PROJECT
'[Captions] -----------------------------------------------------------------------------------------------
Private Const WINDOW_TITLE As String = "CreatingNewProject.Window.Title"
'[Errors] -------------------------------------------------------------------------------------------------
'[Default values] -----------------------------------------------------------------------------------------
'[Other messages] -----------------------------------------------------------------------------------------
Private Const SUCCESS_MESSAGE As String = "CreatingNewProject.Messages.FinalSuccessMessage"
'----------------------------------------------------------------------------------------------------------
Private pConfig As Scripting.Dictionary
Private pProject As EProject
Private pProcessUnits As Scripting.Dictionary
Private WithEvents pProcessor As ItemsProcessor
Attribute pProcessor.VB_VarHelpID = -1
Private pNoMacrosExcel As Excel.Application
'[Settings] -----------------------------------------------------------------------------------------------
Private pIsAborted As Boolean
'[References] ---------------------------------------------------------------------------------------------
Private pSheets As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------


Private Sub Class_Initialize()
    Call loadConfig
End Sub

Private Sub loadConfig()
    Set pConfig = Config.getConfig(CONFIG_NAME)
End Sub


Public Function getNoMacrosExcel() As Excel.Application
    If pNoMacrosExcel Is Nothing Then
        Set pNoMacrosExcel = VBA.CreateObject(EXCEL_APPLICATION)
        With pNoMacrosExcel
            .Visible = False
            .AutomationSecurity = msoAutomationSecurityForceDisable
        End With
    End If
    Set getNoMacrosExcel = pNoMacrosExcel
End Function

Private Sub Class_Terminate()
    If Not pNoMacrosExcel Is Nothing Then
        On Error Resume Next
        Call pNoMacrosExcel.Quit
    End If
End Sub

Private Sub resetState()
    pIsAborted = False
End Sub

Friend Sub abort()
    pIsAborted = True
End Sub



'[SETTERS]
Public Sub setSheets(sheets As Scripting.Dictionary)
    Set pSheets = sheets
End Sub



'[GETTERS]
Public Function getProject() As EProject
    Set getProject = pProject
End Function

Public Function getProcessor() As ItemsProcessor
    Set getProcessor = pProcessor
End Function

Friend Function isAborted() As Boolean
    isAborted = pIsAborted
End Function

Friend Function getReferencesDictionary() As Scripting.Dictionary
    Set getReferencesDictionary = f.dictionaries.GetObject(pConfig, props.references.getName)
End Function

Friend Function getSheet(dataType As enumDataType) As Excel.Worksheet
    Set getSheet = f.dictionaries.GetObject(pSheets, dataType.getName)
End Function







Public Sub run(project As EProject)

    Call resetState

    Set pProject = project
    Call createProcessUnits
    
    If f.dictionaries.isNonEmptyDictionary(pProcessUnits) Then
        Set pProcessor = setupItemsProcessor
        Call pProcessor.run
    End If
    
End Sub

Private Sub createProcessUnits()
    Dim steps As Variant
    Dim varStep As Variant
    Dim processUnit As IProjectCreatorProcessUnit
    '------------------------------------------------------------------------------------------------------
    
    steps = VBA.Array( _
                        CreatingProjectStep_CreatingProjectFolder, _
                        CreatingProjectStep_ApplyingChangesToTextFiles, _
                        CreatingProjectStep_ApplyingChangesToCode, _
                        CreatingProjectStep_FixingReferencesBetweenFiles, _
                        CreatingProjectStep_CreatingRibbonComponents _
                     )
                        
    Set pProcessUnits = f.dictionaries.create(False)
    For Each varStep In steps
        Set processUnit = createProcessUnit(VBA.CLng(varStep))
        If Not processUnit Is Nothing Then
            Call f.dictionaries.AddItem(pProcessUnits, varStep, processUnit)
        End If
    Next varStep
    
End Sub


Private Function createProcessUnit(step As CreatingProjectStepEnum) As IProjectCreatorProcessUnit
    Select Case step
        Case CreatingProjectStep_CreatingProjectFolder:             Set createProcessUnit = New PC_FolderCreator
        Case CreatingProjectStep_ApplyingChangesToCode:             Set createProcessUnit = New PC_CodeReplacer
        Case CreatingProjectStep_ApplyingChangesToTextFiles:        Set createProcessUnit = New PC_TextReplacer
        Case CreatingProjectStep_FixingReferencesBetweenFiles:      Set createProcessUnit = New PC_ReferencesFixer
        Case CreatingProjectStep_CreatingRibbonComponents:          Set createProcessUnit = New PC_RibbonGenerator
    End Select
    
    If Not createProcessUnit Is Nothing Then
        With createProcessUnit
            Call .setParent(Me)
        End With
    End If
End Function

Private Function setupItemsProcessor() As ItemsProcessor
    Dim viewer As IItemsProcessorViewer
    '------------------------------------------------------------------------------------------------------
    Set viewer = getProcessViewer
    Set setupItemsProcessor = MApp.create.ItemsProcessor(pProcessUnits, viewer)
End Function


Private Function getProcessViewer() As WDefaultItemsProcessorViewer
    Const USER_FORM_WIDTH As Single = 500
    '----------------------------------------------------------------------------------------------------------
    
    Set getProcessViewer = MApp.create.DefaultItemsProcessorViewer
    With getProcessViewer
        Call .setHasSummaryPanel(False)
        Call .setHasDetailsProgressBar(True)
        Call .setWidth(USER_FORM_WIDTH)
        Call .setButtonsCaptions(okButtonCaption:=Msg.getText(OK))
        Call .setHasCancelButton(False)
        Call .setWindowTitle(f.Strings.Format(Msg.getText(WINDOW_TITLE), pProject.getName, pProject.getCodeName))
        Call .setStatusDictionary(getStatusViewDictionary)
    End With
    
End Function

Private Function getStatusViewDictionary() As Scripting.Dictionary
    Dim categories As Variant
    '------------------------------------------------------------------------------------------------------
    
    categories = VBA.Array( _
                        ItemProcessStatus_Unknown, _
                        ItemProcessStatus_Correct, _
                        ItemProcessStatus_Warning, _
                        ItemProcessStatus_Error)
                        
    Set getStatusViewDictionary = ItemsProcessor.getDefaultCategories(categories)
    
End Function

Private Sub pProcessor_AfterConfirm()
    Call f.files.openFolderInWindowsExplorer(pProject.getFolderPath)
End Sub
