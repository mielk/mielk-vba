VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WDefaultItemsProcessorViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IItemsProcessorViewer
Implements MielkUI.IProgressControlController

Private Const CLASS_NAME As String = "WDefaultItemsProcessorViewer"
'[Control label IDs & tags] -------------------------------------------------------------------------------
Private Const PROGRESS_FRAME_ID As String = "frmProgress"
Private Const ITEMS_LIST_FRAME_ID As String = "frmList"
Private Const OK_BUTTON_CAPTION_TAG As String = "Save"
Private Const CANCEL_BUTTON_CAPTION_TAG As String = "Cancel"
Private Const DEFAULT_WINDOW_TITLE_TAG As String = "ProcessingItems"
Private Const DEFAULT_MAIN_PROGRESS_BAR_CAPTION_TAG As String = "ProcessingItems.Controls.MainProgressBarHeader"
Private Const DEFAULT_DETAIL_PROGRESS_BAR_CAPTION_TAG As String = "ProcessingItems.Controls.DetailProgressBarHeader"
'[Errors & warnings messages] -----------------------------------------------------------------------------
Private Const NO_ITEMS_FOUND_WARNING As String = "ProcessingItems.Error.NoItemFound"
Private Const ERROR_CANCELLING As String = "ProcessingItems.Error.CancelWindow"
Private Const ERROR_CONFIRMING As String = "ProcessingItems.Error.ConfirmWindow"
'[UI constants] -------------------------------------------------------------------------------------------
Private Const DEFAULT_HAS_DETAILS_PROGRESS_BAR As Boolean = True
Private Const DEFAULT_HAS_SUMMARY_PANEL As Boolean = True
Private Const DEFAULT_CONTROLS_VERTICAL_SPACE As Single = 6
Private Const DEFAULT_LEFT_OFFSET As Single = 6
Private Const DEFAULT_FILES_SELECTOR_WIDTH As Single = 402
Private Const DEFAULT_USER_FORM_WIDTH As Single = 439.5
Private Const DEFAULT_USER_FORM_HEIGHT As Single = 300
'[References] ---------------------------------------------------------------------------------------------
Private WithEvents pProcessor As ItemsProcessor
Attribute pProcessor.VB_VarHelpID = -1
'[UI components] ------------------------------------------------------------------------------------------
Private WithEvents pWindow As MielkUI.WTemplate
Attribute pWindow.VB_VarHelpID = -1
Private pProgressPanel As MielkUI.XProgressPanel
Private WithEvents pSummaryPanel As XProcessSummaryPanel
Attribute pSummaryPanel.VB_VarHelpID = -1
Private WithEvents pItemsListControl As XProcessItemsList
Attribute pItemsListControl.VB_VarHelpID = -1
'[UI properties] ------------------------------------------------------------------------------------------
Private pHasDetailsProgressBar As Boolean
Private pHasSummaryPanel As Boolean
Private pControlsVerticalSpace As Single
Private pLeftOffset As Single
Private pFilesSelectorWidth As Single
Private pWidth As Single
'[Captions] -----------------------------------------------------------------------------------------------
Private pWindowTitle As String
Private pMainProgressBarCaption As String
Private pDetailProgressBarCaption As String
'[Buttons]
Private pOkButtonCaption As String
Private pCancelButtonCaption As String
Private pHasCancelButton As Boolean
'[Statuses]
Private pStatusesDict As Scripting.Dictionary
'[State] ----------------------------------------------------------------------------------------------
Private pCancelledByUser As Boolean
'[Events] ---------------------------------------------------------------------------------------------
Event Confirmed()
Event Cancelled()
'------------------------------------------------------------------------------------------------------




'[INITIALIZATION]
Private Sub Class_Initialize()
    Call setDefaultValues
End Sub

Private Sub setDefaultValues()
    '[UI properties]
    pWidth = DEFAULT_USER_FORM_WIDTH
    pLeftOffset = DEFAULT_LEFT_OFFSET
    pFilesSelectorWidth = DEFAULT_FILES_SELECTOR_WIDTH
    pControlsVerticalSpace = DEFAULT_CONTROLS_VERTICAL_SPACE
    pHasDetailsProgressBar = DEFAULT_HAS_DETAILS_PROGRESS_BAR
    pHasSummaryPanel = DEFAULT_HAS_SUMMARY_PANEL
    '[Captions]
    pWindowTitle = DEFAULT_WINDOW_TITLE_TAG
    pMainProgressBarCaption = DEFAULT_MAIN_PROGRESS_BAR_CAPTION_TAG
    pDetailProgressBarCaption = DEFAULT_DETAIL_PROGRESS_BAR_CAPTION_TAG
    '[Statuses]
    Set pStatusesDict = createDefaultStatusDictionary
    '[Buttons]
    pHasCancelButton = True
End Sub

Private Sub Class_Terminate()
    If Not pWindow Is Nothing Then Call pWindow.destroy
    Set pWindow = Nothing
End Sub





'[SETTERS]
Public Function setProcessor(value As ItemsProcessor) As WDefaultItemsProcessorViewer
    Set setProcessor = Me
    Set pProcessor = value
End Function

Public Function setWindowTitle(value As String) As WDefaultItemsProcessorViewer
    Set setWindowTitle = Me
    pWindowTitle = value
End Function

Public Function setHasDetailsProgressBar(value As Boolean) As WDefaultItemsProcessorViewer
    Set setHasDetailsProgressBar = Me
    pHasDetailsProgressBar = value
End Function

Public Function setHasSummaryPanel(value As Boolean) As WDefaultItemsProcessorViewer
    Set setHasSummaryPanel = Me
    pHasSummaryPanel = value
End Function

Public Function setWidth(value As Single) As WDefaultItemsProcessorViewer
    Set setWidth = Me
    pWidth = value
End Function

Public Function setHasCancelButton(value As Boolean) As WDefaultItemsProcessorViewer
    Set setHasCancelButton = Me
    pHasCancelButton = value
End Function

Public Function setButtonsCaptions(Optional okButtonCaption As String, Optional cancelButtonCaption As String) As WDefaultItemsProcessorViewer
    Set setButtonsCaptions = Me
    pOkButtonCaption = okButtonCaption
    pCancelButtonCaption = cancelButtonCaption
End Function

Public Function setMainProgressPanelHeader(value As String) As WDefaultItemsProcessorViewer
    Set setMainProgressPanelHeader = Me
    pMainProgressBarCaption = value
End Function




'[GETTERS]
Public Function getProcessor() As ItemsProcessor
    Set getProcessor = pProcessor
End Function

Public Function isCancelledByUser() As Boolean
    isCancelledByUser = pCancelledByUser
End Function





'[DEFINING STATUSES]
Private Function createDefaultStatusDictionary() As Scripting.Dictionary
    Dim categories As Variant
    '------------------------------------------------------------------------------------------------------
    
    categories = VBA.Array( _
                        ItemProcessStatus_Unknown, _
                        ItemProcessStatus_Correct, _
                        ItemProcessStatus_Warning, _
                        ItemProcessStatus_Error)
                        
    Set createDefaultStatusDictionary = ItemsProcessor.getDefaultCategories(categories)
    
End Function

Public Function setStatusDictionary(value As Scripting.Dictionary) As WDefaultItemsProcessorViewer
    Set setStatusDictionary = Me
    Set pStatusesDict = value
End Function

Public Function addStatus(value As Long, settings As Scripting.Dictionary) As WDefaultItemsProcessorViewer
    Set addStatus = Me
    If pStatusesDict Is Nothing Then Set pStatusesDict = F.dictionaries.Create(False)
    Call F.dictionaries.addItem(pStatusesDict, value, settings)
End Function


Public Function getStatusSettings(value As Long) As Scripting.Dictionary
    Set getStatusSettings = F.dictionaries.GetObject(pStatusesDict, value)
End Function

Public Function getStatusesDictionary() As Scripting.Dictionary
    Set getStatusesDictionary = pStatusesDict
End Function


'[API]
Public Sub display()
    Call createUserForm
    Call insertProgressPanel
    Call insertSummaryPanel
    Call insertItemsListControl
    Call arrangeLayout
    Call showUserForm
End Sub



'[RENDERING]
Private Sub createUserForm()
    Set pWindow = UI.Create.WindowTemplate
    With pWindow
        Call .setCaption(MsgService.getText(pWindowTitle))
        Call .setContentFrameTopMargin(pControlsVerticalSpace)
        Call .setHasCancelButton(pHasCancelButton)
        Call .setOkCaption(pOkButtonCaption)
        Call .setCancelCaption(pCancelButtonCaption)
        Call .setWarningsForUnsavedData(False)
        Call .render
        Call .setValidation(F.Create.ValidationObject(False))
    End With
End Sub

Private Sub insertProgressPanel()
    Dim frame As MSForms.frame
    '------------------------------------------------------------------------------------------------------
    
    Set frame = pWindow.getContentFrame.controls.Add(CREATE_FRAME_ID, PROGRESS_FRAME_ID)
    Set pProgressPanel = UI.Create.ProgressPanel(pWindow, frame)
    With pProgressPanel
        Call .setController(Me)
        Call .setWidth(pWidth)
        Call .setHasDetailsProgressBar(pHasDetailsProgressBar)
        Call .setMainHeaderCaption(MsgService.getText(pMainProgressBarCaption))
        Call .setDetailHeaderCaption(MsgService.getText(pDetailProgressBarCaption))
        Call .setHasBottomBorderOnly(False)
        Call .render
    End With
End Sub


Private Sub insertSummaryPanel()
    Dim frame As MSForms.frame
    '------------------------------------------------------------------------------------------------------
    If pHasSummaryPanel Then
        Set frame = pWindow.getContentFrame.controls.Add(CREATE_FRAME_ID, PROGRESS_FRAME_ID)
        Set pSummaryPanel = MApp.Create.processSummaryPanel
        With pSummaryPanel
            Call .setController(Me)
            Call .setWindow(pWindow)
            Call .setWidth(pWidth)
            Call .setFrame(frame)
            Call .render
        End With
    End If
End Sub


Private Sub insertItemsListControl()
    Dim frame As MSForms.frame
    '------------------------------------------------------------------------------------------------------
    Set frame = pWindow.getContentFrame.controls.Add(CREATE_FRAME_ID, ITEMS_LIST_FRAME_ID)
    Set pItemsListControl = MApp.Create.processItemsList
    With pItemsListControl
        Call .setProcessor(pProcessor)
        Call .setView(Me)
        Call .setWindow(pWindow)
        Call .setWidth(pWidth)
        Call .setFrame(frame)
        Call .render
    End With
End Sub

Private Sub showUserForm()
    With pWindow
        Call .locateRelativeTo(UI.Create.LocationPropsForScreenCenter)
        Call .display(vbModeless, False)
    End With
End Sub

Private Sub arrangeLayout()
    Dim top As Single:              top = 0
    Dim left As Single:             left = pWindow.getContentFrameHorizontalMargin
    '------------------------------------------------------------------------------------------------------

    '[Progress panel]
    Call pProgressPanel.move(left, top)
    top = pProgressPanel.getBottom + pControlsVerticalSpace

    '[Summary panel]
    If pHasSummaryPanel Then
        Call pSummaryPanel.move(left, top)
        top = pSummaryPanel.getBottom + pControlsVerticalSpace
    End If
    
    '[Items list]
    With pItemsListControl
        Call .move(left, top)
        Call .updateScrollBarVisibility
        top = .getBottom
    End With

    Call pWindow.arrangeComponents

End Sub




'[PROCESSING ITEMS]
Private Sub pProcessor_BeforeProcessingItem(item As IProcessItem)
    Call pProgressPanel.startNewMainItem(item.getCaption)
End Sub

Private Sub pProcessor_AfterFinish()
    Call ErrorManager.clear
    Call pWindow.setValidation(F.Create.ValidationObject(True))
End Sub

Private Sub pProcessor_AfterProcessingItem(item As IProcessItem)
    Call ErrorManager.clear
    Call pProgressPanel.endCurrentMainItem
    With item
        Call pItemsListControl.updateItemLine(.getName, .getStatus, .getMessage)
    End With
End Sub

Private Sub pProcessor_DetailProgressChange(currentOperation As String, currentProgress As Single)
    Call ErrorManager.clear
    Call pProgressPanel.updateDetailProgress(currentOperation, currentProgress)
End Sub

Private Sub pProcessor_AfterItemRemoved(item As IProcessItem)
    Call ErrorManager.clear
    Call pItemsListControl.triggerAfterRemoveItemEvent(item)
End Sub








Private Sub pItemsListControl_AfterPreviewFileClicked(item As IProcessItem)
    Call ErrorManager.clear
    If Not item Is Nothing Then
        Call item.preview
    End If
End Sub

Private Sub pItemsListControl_AfterReprocessClicked(item As IProcessItem)
    Call ErrorManager.clear
    If Not item Is Nothing Then
        Call pProcessor.triggerAfterReprocessItemEvent(item)
        With item
            Call pItemsListControl.updateItemLine(.getName, .getStatus, .getMessage)
        End With
    End If
End Sub

Private Sub pItemsListControl_AfterRemoveClicked(item As IProcessItem)
    Call ErrorManager.clear
    Stop
    Call pProcessor.removeItem(item.getName)
    Stop            'SHOULD UPDATE COUNTNERS IN SUMMARY PANEL
    'Call arrangeLayout
    'Call updateButtonsVisibility
End Sub




'[Filtering]
Private Sub pSummaryPanel_AfterClick(status As Variant)
    Call ErrorManager.clear
    Call pItemsListControl.filter(status)
End Sub




'[Events]
Private Sub pWindow_cancel()
    Call ErrorManager.clear
    Call cancelProcessing
End Sub

Private Sub pWindow_Confirm()
    Call ErrorManager.clear
    Call confirmProcessing
End Sub


Private Sub cancelProcessing()
    pCancelledByUser = True
    Call pWindow.hide
    RaiseEvent Cancelled
    Call pProcessor.Cancel
End Sub

Private Sub confirmProcessing()
    pCancelledByUser = False
    Call pWindow.hide
    RaiseEvent Confirmed
    Call pProcessor.confirm
End Sub



'[Triggers]
Friend Sub triggerAfterExpandCollapseEvent(item As IProcessItem)
    Call pProcessor.triggerAfterExpandCollapseEvent(item)
End Sub

Friend Sub triggerAfterPreviewItemEvent(item As IProcessItem)
    Call pProcessor.triggerAfterPreviewItemEvent(item)
End Sub

Friend Sub triggerAfterReprocessItemEvent(item As IProcessItem)
    Call pProcessor.triggerAfterReprocessItemEvent(item)
End Sub



'[Implementation of IFileProcessorViewer]
Private Function IItemsProcessorViewer_getProcessor() As ItemsProcessor
    Set IItemsProcessorViewer_getProcessor = pProcessor
End Function

Private Function IItemsProcessorViewer_getItems() As VBA.Collection
    Set IItemsProcessorViewer_getItems = pProcessor.getItemsCollection
End Function

Private Function IItemsProcessorViewer_getCanvas() As Variant
    Stop
    'Set IItemsProcessorViewer_getCanvas = Me
End Function

Private Sub IItemsProcessorViewer_arrangeLayout()
    Stop
    'Call arrangeLayout
End Sub

Private Sub IItemsProcessorViewer_display()
    Call display
End Sub

Private Sub IItemsProcessorViewer_hide()
    Call pWindow.hide
End Sub

Private Sub IItemsProcessorViewer_afterCancel()
    Stop
End Sub

Private Sub IItemsProcessorViewer_afterConfirm()
    Stop
End Sub

Private Function IItemsProcessorViewer_getStatusesDictionary() As Scripting.IDictionary
    Set IItemsProcessorViewer_getStatusesDictionary = getStatusesDictionary
End Function

Private Function IItemsProcessorViewer_getStatusSettings(value As Long) As Scripting.IDictionary
    Set IItemsProcessorViewer_getStatusSettings = getStatusSettings(value)
End Function

Private Function IItemsProcessorViewer_setProcessor(value As ItemsProcessor) As IItemsProcessorViewer
    Set IItemsProcessorViewer_setProcessor = setProcessor(value)
End Function








'[Implementation of IProgressControlController]
Private Function IProgressControlController_getTotalItemsCounter() As Long
    IProgressControlController_getTotalItemsCounter = pProcessor.countItems
End Function

Private Function IProgressControlController_getCanvas() As Variant
    Stop
    'Set IProgressControlController_getCanvas = Me
End Function


