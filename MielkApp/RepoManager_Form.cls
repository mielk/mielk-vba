VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepoManager_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Private Const CLASS_NAME As String = "RepoManager"
''----------------------------------------------------------------------------------------------------------
'Private pRepo As Repository2
'Private pRepoParent As IRepoParent
'Private WithEvents pListWindow As WTemplate
'Private WithEvents pListManager As XListManager
''----------------------------------------------------------------------------------------------------------
'
'
''[Setters]
'Public Function setRepo(value As Repository2) As RepoManager
'    Set setRepo = Me
'    Set pRepo = value
'    Set pRepoParent = pRepo.getParent
'End Function
'
'
'Public Sub show()
'    Const DEFAULT_WINDOW_TITLE_TAG As String = "ListManager.WindowTitle"
'    '----------------------------------------------------------------------------------------------------------
'    Dim frame As MSForms.frame
'    '----------------------------------------------------------------------------------------------------------
'
'    If pRepoParent Is Nothing Then Call VBA.Err.Raise(1, Description:="pRepoParent not set")
'
'Call ProgressBar.openSubProcess(0, 1)
'Call ProgressBar.setProgress(0)
'
'    Set pListWindow = UI.Create.WindowTemplate.setHasOkButton(False).setHasCancelButton(False)
'    With pListWindow
'        Call .render
'        Call .locateRelativeTo(UI.Create.LocationPropsForScreenCenter)
'        Call .setCaption(F.Strings.coalesce(pRepoParent.getWindowTitle, MsgService.getText(DEFAULT_WINDOW_TITLE_TAG)))
'        If pRepoParent.isAddingAllowed Then
'            Call .addExtraButton(ACTION_ADD, MsgService.getText("Add"), MIELK_COLOR_GRAY)
'        Else
'            Call pListWindow.setHasButtonsPanel(False)
'        End If
'    End With
'
'Call ProgressBar.setProgress(0.025)
'
'    Set pListManager = UI.Create.ListManager(createFrame)
'    With pListManager
'        Call .setWindow(pListWindow)
'        Call .setColumns(pRepoParent.getColumnsDictionary)
'        Call .setFilters(getFiltersDictionary(pListManager))
'        Call .setItems(F.Collections.Clone(pRepo.getItemsCollection))
'
'        Call .setSortParameters(pRepoParent.getSortColumnKey, True)
'        Call .sethasRecordsCounter(pRepoParent.hasRecordsCounter)
'        '[Captions]
'        Call .setDeleteItemWindowTitle(pRepoParent.getDeleteItemWindowTitle)
'        Call .setDeleteItemQuestion(pRepoParent.getDeleteItemQuestion)
'        Call .setDeleteSuccessInfo(pRepoParent.getDeleteSuccessInfo)
'Call ProgressBar.setProgress(0.05)
'Call ProgressBar.openSubProcess(0.05, 0.95)
'        Call .Load
'        Call pListWindow.arrangeComponents
'Call ProgressBar.closeCurrentSubProcess
'    End With
'
'Call ProgressBar.setProgress(1)
'Call ProgressBar.hide
'
'    Call pListWindow.locateRelativeTo(UI.Create.LocationPropsForScreenCenter)
'    Call pListWindow.display(vbModeless, True)
'
'
'End Sub
'
'
'
'Private Function createFrame() As MSForms.frame
'    Const DEFAULT_FRAME_WIDTH As Single = 350
'    Const DEFAULT_FRAME_HEIGHT As Single = 350
'    Const DEFAULT_FRAME_TOP As Single = 6
'    Const DEFAULT_FRAME_LEFT As Single = 6
'    '------------------------------------------------------------------------------------------------------
'    Dim width As Single
'    '------------------------------------------------------------------------------------------------------
'
'    width = calculateFrameWidth(pRepoParent.getColumnsDictionary) + 2 * pListWindow.getContentFrameHorizontalMargin + SCROLL_BAR_WIDTH
'
'    With pListWindow
'        Set createFrame = .getContentFrame.Controls.Add(CREATE_FRAME_ID, "frmList", True)
'        With createFrame
'            .height = F.Numerics.coalesce(pRepoParent.getFrameHeight, DEFAULT_FRAME_HEIGHT)
'            .width = F.Numerics.coalesce(width, DEFAULT_FRAME_WIDTH)
'            .left = DEFAULT_FRAME_LEFT
'            .top = DEFAULT_FRAME_TOP
'        End With
'        Call .arrangeComponents
'    End With
'
'End Function
'
'
'Private Function calculateFrameWidth(columns As Scripting.Dictionary) As Single
'    Dim key As Variant
'    Dim col As DColumn
'    '------------------------------------------------------------------------------------------------------
'
'    If F.Dictionaries.isNonEmptyDictionary(columns) Then
'        For Each key In columns.Keys
'            Set col = columns.item(key)
'            calculateFrameWidth = calculateFrameWidth + col.getWidth
'        Next key
'    End If
'
'End Function
'
'
'
'Private Function getFiltersDictionary(manager As XListManager) As Scripting.Dictionary
'    Dim fields As VBA.Collection
'    Dim varField As Variant
'    Dim key As String
'    '------------------------------------------------------------------------------------------------------
'
'    Set getFiltersDictionary = F.Dictionaries.Create(False)
'    Set fields = pRepoParent.getFilterColumns
'
'    For Each varField In fields
'        key = VBA.CStr(varField)
'        Call getFiltersDictionary.Add(key, UI.Create.filterList(manager, key).setHasSearchField(True))
'    Next varField
'
'End Function
'
'
'Private Sub handleAddingItem()
'    Dim newItem As IEdititem
'    Dim listItem As IListItem
'    Dim validation As ValidationObject
'    Dim frm As IEditForm
'    '------------------------------------------------------------------------------------------------------
'
'    Set newItem = pRepoParent.createNewItem
'
'    Call pListWindow.sendToBackground
'    Set validation = newItem.edit
'    Call pListWindow.restoreFromBackground
'
'    If validation.getStatus Then
'        Call pRepo.addItem(newItem)
'
'        '[Convert IEditItem to IListItem] -------------------------------------------|
'        On Error Resume Next                                                        '|
'        Set listItem = newItem                                                      '|
'        If Not DEV_MODE Then On Error GoTo ErrHandler Else On Error GoTo 0          '|
'        '----------------------------------------------------------------------------|
'
'        If Not listItem Is Nothing Then Call pListManager.addItem(listItem)
'
'    ElseIf VBA.Len(validation.getMessage) = 0 Then
'        'cancelled by user
'    Else
'        Call VBA.Err.Raise(1, Description:=validation.getMessage)
'    End If
'
'ErrHandler:
'
'End Sub
'
'
'
'Private Sub pListManager_AfterItemEdit(item As MielkUI.IListItem, ByRef updateAfterward As Boolean)
'    Call pRepo.getParent.runAfterEditAction(updateAfterward)
'    Call pRepo.updateItemKeys(item)
'End Sub
'
'
'Private Sub pListWindow_CustomAction(actionTag As String, params As Scripting.Dictionary)
'    Select Case actionTag
'        Case ACTION_ADD:                    Call handleAddingItem
'    End Select
'End Sub
