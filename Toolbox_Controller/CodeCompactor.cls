VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeCompactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "CodeCompactor"
'----------------------------------------------------------------------------------------------------------
Private WithEvents pProcessor As ItemsProcessor
Attribute pProcessor.VB_VarHelpID = -1
Private pItems As Scripting.Dictionary
Private pPassword As String
'----------------------------------------------------------------------------------------------------------
Private pExcel As Excel.Application
'----------------------------------------------------------------------------------------------------------


Public Sub run()
    '[Select file]
    Call createProcessUnits
    
    If f.dictionaries.isNonEmptyDictionary(pItems) Then
        '[Enter password]
        pPassword = VBA.InputBox("Password for VBA project")
        Call setup
        
        Set pProcessor = setupItemsProcessor
        Call pProcessor.run
    End If
    
End Sub


Private Sub createProcessUnits()
    Dim files As Variant
    Dim unit As CodeCompactorProcessUnit
    Dim file As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set pItems = f.dictionaries.create(False)
    Set files = f.files.selectFiles(FileType:="Excel files, *.xlsm; *.xla; *.xlam; *.xlsb")
    
    If f.Collections.isNonEmptyCollection(files) Then
        For Each file In files
            Set unit = New CodeCompactorProcessUnit
            With unit
                Call .setParent(Me)
                Call .setFilePath(VBA.CStr(file))
            End With
            Call f.dictionaries.AddItem(pItems, VBA.CStr(file), unit)
        Next file
    End If
    
End Sub


Private Sub setup()
    Set pExcel = New Excel.Application
    With pExcel
        .Visible = True
        .AutomationSecurity = msoAutomationSecurityForceDisable
    End With
End Sub



'[Settings]
Friend Function getVbaCodePassword() As String
    getVbaCodePassword = pPassword
End Function

Friend Function getVbaProjects() As VBA.Collection
    Stop
End Function

Friend Function getExcel() As Excel.Application
    Set getExcel = pExcel
End Function




Private Function setupItemsProcessor() As ItemsProcessor
    Dim viewer As IItemsProcessorViewer
    '------------------------------------------------------------------------------------------------------
    Set viewer = getProcessViewer
    Set setupItemsProcessor = MApp.create.ItemsProcessor(pItems, viewer)
End Function


Private Function getProcessViewer() As WDefaultItemsProcessorViewer
    Const USER_FORM_WIDTH As Single = 500
    '----------------------------------------------------------------------------------------------------------
    
    Set getProcessViewer = MApp.create.DefaultItemsProcessorViewer
    With getProcessViewer
        Call .setHasSummaryPanel(False)
        Call .setHasDetailsProgressBar(True)
        Call .setWidth(USER_FORM_WIDTH)
        Call .setButtonsCaptions(okButtonCaption:="OK")
        Call .setHasCancelButton(False)
        Call .setWindowTitle("Compacting files")
        Call .setStatusDictionary(getStatusViewDictionary)
    End With
    
End Function

Private Function getStatusViewDictionary() As Scripting.Dictionary
    Dim categories As Variant
    '------------------------------------------------------------------------------------------------------
    
    categories = VBA.Array( _
                        ItemProcessStatus_Unknown, _
                        ItemProcessStatus_Correct, _
                        ItemProcessStatus_Error)
                        
    Set getStatusViewDictionary = ItemsProcessor.getDefaultCategories(categories)
    
End Function



Private Sub pProcessor_AfterConfirm()
    Call pExcel.Quit
    Set pExcel = Nothing
End Sub
