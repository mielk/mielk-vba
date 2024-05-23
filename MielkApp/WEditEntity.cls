VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WEditEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IEditForm
Implements IListener

Private Const CLASS_NAME As String = "WEditEntity"
'----------------------------------------------------------------------------------------------------------
Private Const PROCESS_TAG As String = "edit"
Private Const EDIT_WINDOW_TITLE_TAG As String = "Edit"
Private Const ADD_WINDOW_TITLE_TAG As String = "Add"
Private Const SAVE_SUCCESS_INFO As String = "Success"
'[UI constants] -------------------------------------------------------------------------------------------
Private Const PANELS_VERTICAL_SPACE As Single = 6
Private Const DEFAULT_VERTICAL_PADDING As Single = 3
Private Const DEFAULT_HORIZONTAL_PADDING As Single = 6
Private Const DEFAULT_LABEL_WIDTH As Single = 84
Private Const DEFAULT_VALUE_WIDTH As Single = 204
Private Const VALIDATION_PANEL_WIDTH As Single = 250
Private Const SEPARATOR_BACK_COLOR As Long = 13487565
Private Const DEFAULT_BUTTON_HEIGHT As Single = 24
Private Const DEFAULT_BUTTON_WIDTH As Single = 100
'----------------------------------------------------------------------------------------------------------
Private WithEvents pForm As WTemplate
Attribute pForm.VB_VarHelpID = -1
Private pFrame As MSForms.frame
'----------------------------------------------------------------------------------------------------------
Private pEditItem As IEditItem
Private pItem As IEditItem
Private pCancelledByUser As Boolean
Private pEventsDisabled As Boolean
'[UI components] ------------------------------------------------------------------------------------------
Private WithEvents pPropertiesPanel As XPropertiesPanel
Attribute pPropertiesPanel.VB_VarHelpID = -1
Private pLinesDefinitions As VBA.Collection
Private pCustomPanels As VBA.Collection
'[UI properties] ------------------------------------------------------------------------------------------
Private pVerticalPadding As Single
Private pHorizontalPadding As Single
Private pLabelWidth As Single
Private pValueWidth As Single
Private pHasWarningsForUnsavedData As Boolean
Private pWindowTitle As String
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    pVerticalPadding = DEFAULT_VERTICAL_PADDING
    pHorizontalPadding = DEFAULT_HORIZONTAL_PADDING
    pLabelWidth = DEFAULT_LABEL_WIDTH
    pValueWidth = DEFAULT_VALUE_WIDTH
    pHasWarningsForUnsavedData = False
End Sub


'[Setters]
Public Function setItem(value As IEditItem) As WEditEntity
    Set setItem = Me
    Set pItem = value
    Set pEditItem = value.clone
End Function

Public Function setWindowTitle(value As String) As WEditEntity
    Set setWindowTitle = Me
    pWindowTitle = value
End Function

Public Function setLabelWidth(value As Single) As WEditEntity
    Set setLabelWidth = Me
    pLabelWidth = value
End Function

Public Function setValueWidth(value As Single) As WEditEntity
    Set setValueWidth = Me
    pValueWidth = value
End Function

Public Function setHasWarningsForUnsavedData(value As Boolean) As WEditEntity
    Set setHasWarningsForUnsavedData = Me
    pHasWarningsForUnsavedData = value
End Function

Public Function setLinesDefinitions(value As VBA.Collection) As WEditEntity
    Set setLinesDefinitions = Me
    Set pLinesDefinitions = value
End Function

Public Function addCustomPanel(value As IPanel) As WEditEntity
    Set addCustomPanel = Me
    If pCustomPanels Is Nothing Then Set pCustomPanels = New VBA.Collection
    Call pCustomPanels.Add(value)
End Function



'[Getters]
Public Function getItem() As IEditItem
    Set getItem = pItem
End Function

Public Function getEditItem() As IEditItem
    Set getEditItem = pEditItem
End Function

Public Function isCancelledByUser() As Boolean
    isCancelledByUser = pCancelledByUser
End Function





'[Rendering]
Public Sub run(Optional mode As FormShowConstants = vbModal, Optional locationProps As DLocation)
    Call createForm
    Call render
    Call pForm.locateRelativeTo(locationProps)
    Call pForm.display(mode, True)
End Sub


Private Sub createForm()
    If pForm Is Nothing Then
        Set pForm = UI.Create.WindowTemplate
        With pForm
            Call .setCaption(getWindowTitle)
            Call .setMinHeight(1)
            Call .setValidationPanelWidth(VALIDATION_PANEL_WIDTH)
            Call .setWarningsForUnsavedData(pHasWarningsForUnsavedData)
            Call .render
            Set pFrame = .getContentFrame
        End With
    End If
End Sub


Private Function getWindowTitle() As String
    If VBA.Len(pWindowTitle) Then
        getWindowTitle = pWindowTitle
    ElseIf pEditItem.getId > 0 Then
        getWindowTitle = F.Strings.format(MsgService.getText(EDIT_WINDOW_TITLE_TAG), pEditItem.getName)
    Else
        getWindowTitle = MsgService.getText(ADD_WINDOW_TITLE_TAG)
    End If
End Function


Private Sub render()
    If Not pEventsDisabled Then
        pEventsDisabled = True
        Call insertPropertiesPanel
        Call renderCustomPanels
        Call arrangeComponents
        Call pPropertiesPanel.setFocus
        pEventsDisabled = False
    End If
End Sub

Private Sub insertPropertiesPanel()
    Set pPropertiesPanel = UI.Create.PropertiesPanel(pEditItem, pForm, pFrame). _
                                    setValidationFrame(pForm.getValidationFrame). _
                                    setLabelWidth(pLabelWidth). _
                                    setValueWidth(pValueWidth). _
                                    setTop(pVerticalPadding). _
                                    setLeft(pHorizontalPadding). _
                                    setLinesDefinitions(pLinesDefinitions)
    With pPropertiesPanel
        Call .render
    End With
End Sub

Private Sub renderCustomPanels()
    Dim panel As IPanel
    '------------------------------------------------------------------------------------------------------
    
    If F.Collections.isNonEmptyCollection(pCustomPanels) Then
        For Each panel In pCustomPanels
            With panel
                Call .setParentFrame(pFrame)
                Call .setWindow(pForm)
                Call .render
            End With
        Next panel
    End If
    
End Sub

Friend Sub arrangeComponents()
    Dim panels As VBA.Collection
    Dim panel As IPanel
    '------------------------------------------------------------------------------------------------------
    Dim horizontalMargin As Single:         horizontalMargin = pForm.getContentFrameHorizontalMargin
    Dim top As Single:                      top = pVerticalPadding
    Dim isNotFirstPanel As Boolean
    '------------------------------------------------------------------------------------------------------
    
    Set panels = getPanelsInDisplayOrder
    For Each panel In panels
        If panel.isVisible Then
            If isNotFirstPanel Then Call panel.insertBorder(xlEdgeTop, SEPARATOR_BACK_COLOR)
            Call panel.setUiProperty(UI.props.left, VBA.CSng(horizontalMargin)). _
                        setUiProperty(UI.props.top, top)
            '[Adjust width of all panels to the width of properties panel]
            If Not panel Is pPropertiesPanel Then
                Call panel.setUiProperty(UI.props.width, pPropertiesPanel.getWidth)
            End If
            '[Update variables]
            isNotFirstPanel = True
            top = panel.getUiProperty(UI.props.bottom) + PANELS_VERTICAL_SPACE
        End If
    Next panel
    
    Call pForm.arrangeComponents
    
End Sub

Private Function getPanelsInDisplayOrder() As VBA.Collection
    Dim panels As VBA.Collection
    Dim panel As IPanel
    Dim arr As Variant
    Dim counter As Long
    '------------------------------------------------------------------------------------------------------
    
    Set panels = F.Collections.Create(pPropertiesPanel)
    Call F.Collections.appendCollection(panels, pCustomPanels)
    
    ReDim arr(1 To 2, 1 To panels.Count)
    For Each panel In panels
        If Not panel Is Nothing Then
            counter = counter + 1
            Set arr(1, counter) = panel
            arr(2, counter) = panel.getIndex
        End If
    Next panel
    
    ReDim Preserve arr(1 To 2, 1 To counter)
    Call F.arrays.Sort(arr, True, 2)
    
    Set getPanelsInDisplayOrder = F.Collections.fromArray(F.arrays.getRow(arr, 1, False))
    
End Function







'[Events]

Private Sub pPropertiesPanel_AfterValidationChange(validation As MielkCore.ValidationObject)
    Call pForm.setValidation(validation)
End Sub

Private Sub pPropertiesPanel_AfterChange()
    If Not pEventsDisabled Then
        Call pForm.setHasUnsavedData(True)
    End If
End Sub

Private Sub pPropertiesPanel_AfterLineChange(line As MielkUI.xxPropertyLine)
    Call pEditItem.runAction(line.getChangeCallbackMethod, F.dictionaries.createSingleEntryDict(TAG_ITEM, line))
End Sub



Private Sub pForm_Confirm()
    Dim validation As ValidationObject
    Dim continue As Boolean
    '------------------------------------------------------------------------------------------------------
    
    Set validation = pPropertiesPanel.getValidation
    If validation.getStatus = WARNING_CODE Then
        continue = pEditItem.getValue(TAG_CONFIRM_IF_WARNINGS)
    Else
        continue = True
    End If
    
    pCancelledByUser = Not continue
    
End Sub

Private Sub pForm_Cancel()
    pCancelledByUser = True
End Sub





'[Implementation of IEditForm interface]
Private Function IEditForm_isCancelledByUser() As Boolean
    IEditForm_isCancelledByUser = isCancelledByUser
End Function

Private Sub IEditForm_run(Optional mode As FormShowConstants = 1&, Optional locationProps As MielkUI.DLocation)
    Call run(mode, locationProps)
End Sub



'[Implementation of IListener interface]
Private Sub IListener_catchEvent(source As Variant, eventName As String, Optional properties As Scripting.IDictionary)
    Stop
End Sub
