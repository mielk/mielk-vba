VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WStandarizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "WStandarizer"
'[Default values] -----------------------------------------------------------------------------------------
Private Const DEFAULT_WINDOW_TITLE As String = "NamesStandarizer.DefaultWindowTitle"
Private Const DEFAULT_BORDER_COLOR As Long = 13487565
Private Const DEFAULT_LIST_WIDTH As Single = 300
Private Const DEFAULT_LIST_HEIGHT As Single = 300
Private Const DEFAULT_BUTTON_HEIGHT As Single = 24
Private Const DEFAULT_PADDING As Single = 6
Private Const DEFAULT_CONTROLS_SPACE As Single = 6
'[Warning panel] ------------------------------------------------------------------------------------------
Private Const WARNING_PANEL_BACK_COLOR As Long = 4837114    'rgb(250, 206, 73)
Private Const WARNING_PANEL_BORDER_COLOR As Long = 4029906    'rgb(210, 125, 61)
Private Const WARNING_PANEL_FONT_SIZE As Single = 10
Private Const WARNING_PANEL_FONT_COLOR As Long = 3955870    'rgb(158, 92, 60)
Private Const WARNING_PANEL_HEIGHT As Single = 36
Private Const WARNING_ICON_SIZE As Single = 24
'[Captions] -----------------------------------------------------------------------------------------------
Private Const CREATE_NEW_ITEM_BUTTON_CAPTION As String = "NamesStandarizer.Labels.CreateNewItem"
Private Const ASSIGN_ALIAS_BUTTON_CAPTION As String = "NamesStandarizer.Labels.AssignAlias"
'----------------------------------------------------------------------------------------------------------
Private WithEvents pParent As Standarizer
Attribute pParent.VB_VarHelpID = -1
'[UI references] ------------------------------------------------------------------------------------------
Private WithEvents pForm As WTemplate
Attribute pForm.VB_VarHelpID = -1
Private pFrame As MSForms.frame
'[UI components] ------------------------------------------------------------------------------------------
Private pWarningBackground As MSForms.label
Private pWarningIcon As MSForms.label
Private pWarningMessage As MSForms.label
Private WithEvents pAliasesList As UIList
Attribute pAliasesList.VB_VarHelpID = -1
Private pVerticalSeparator As MSForms.label
Private WithEvents pStandardNamesList As UIList
Attribute pStandardNamesList.VB_VarHelpID = -1
Private WithEvents pNewItemButton As UIButton
Attribute pNewItemButton.VB_VarHelpID = -1
Private WithEvents pAssignButton As UIButton
Attribute pAssignButton.VB_VarHelpID = -1
'[UI properties] ------------------------------------------------------------------------------------------
Private pWindowTitle As String
Private pBorderColor As Long
Private pAliasesListWidth As Single
Private pStandardNamesListWidth As Single
Private pListHeight As Single
Private pButtonHeight As Single
Private pPadding As Single
Private pControlsSpace As Single
'[Captions] -----------------------------------------------------------------------------------------------
Private pCreateNewItemButtonCaption As String
Private pAssignAliasButtonCaption As String
'[State] --------------------------------------------------------------------------------------------------
Private pIsRendered As Boolean
Private pWarningStatus As Boolean
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    pWindowTitle = DEFAULT_WINDOW_TITLE
    pBorderColor = DEFAULT_BORDER_COLOR
    pAliasesListWidth = DEFAULT_LIST_WIDTH
    pStandardNamesListWidth = DEFAULT_LIST_WIDTH
    pListHeight = DEFAULT_LIST_HEIGHT
    pButtonHeight = DEFAULT_BUTTON_HEIGHT
    pPadding = DEFAULT_PADDING
    pControlsSpace = DEFAULT_CONTROLS_SPACE
    pCreateNewItemButtonCaption = MsgService.getText(CREATE_NEW_ITEM_BUTTON_CAPTION)
    pAssignAliasButtonCaption = MsgService.getText(ASSIGN_ALIAS_BUTTON_CAPTION)
End Sub



'[Setters]
Public Function setParent(value As Standarizer) As WStandarizer
    Set setParent = Me
    Set pParent = value
End Function

Public Function setWindowTitle(value As String) As WStandarizer
    Set setWindowTitle = Me
    pWindowTitle = value
End Function

Public Function setCreateNewItemButtonCaption(value As String) As WStandarizer
    Set setCreateNewItemButtonCaption = Me
    pCreateNewItemButtonCaption = value
End Function

Public Function setAliasesListWidth(value As Single) As WStandarizer
    Set setAliasesListWidth = Me
    pAliasesListWidth = value
End Function




'[Getters]
Public Function getParent() As Standarizer
    Set getParent = pParent
End Function



'[Rendering]
Public Sub run(Optional mode As FormShowConstants = vbModal, Optional locationProps As DLocation)
    Call createForm
    Call render
    Call setupValues
    Call pForm.locateRelativeTo(locationProps)
    Call pForm.display(mode, True)
End Sub


Private Sub createForm()
    If pForm Is Nothing Then
        Set pForm = UI.Create.WindowTemplate
        With pForm
            Call .setCaption(MsgService.getText(pWindowTitle))
            Call .setMinHeight(1)
            Call .setWarningsForUnsavedData(False)
            Call .render
            Set pFrame = .getContentFrame
        End With
    End If
End Sub

Private Sub render()
    If Not pIsRendered Then
        Call insertWarningsControls
        Call insertAliasesList
        Call insertVerticalSeparator
        Call insertStandardNamesList
        Call insertNewNameButton
        Call insertAssignButton
        Call arrangeComponents
        pIsRendered = True
    End If
End Sub

Private Sub insertWarningsControls()
    Set pWarningBackground = pFrame.Controls.Add(CREATE_LABEL_ID, False)
    With pWarningBackground
        .BackStyle = fmBackStyleOpaque
        .BackColor = WARNING_PANEL_BACK_COLOR
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = WARNING_PANEL_BORDER_COLOR
    End With
    
    Set pWarningIcon = pFrame.Controls.Add(CREATE_LABEL_ID, False)
    With pWarningIcon
        .BackStyle = fmBackStyleTransparent
        .height = WARNING_ICON_SIZE
        .width = WARNING_ICON_SIZE
        Set .Picture = UI.Icons.getWarningOrangeRedIcon
    End With
    
    Set pWarningMessage = pFrame.Controls.Add(CREATE_LABEL_ID, False)
    With pWarningMessage
        .BackStyle = fmBackStyleTransparent
        .Font.size = WARNING_PANEL_FONT_SIZE
        .ForeColor = WARNING_PANEL_FONT_COLOR
        .Font.name = APP_FONT_FAMILY
        .Font.Bold = 1
    End With
    
    pWarningStatus = pParent.hasAnyUnassigned
    
End Sub

Private Sub insertAliasesList()
    Set pAliasesList = UI.Create.List(pFrame). _
                            setMultiSelectMode(fmMultiSelectMulti). _
                            setBorderColor(pBorderColor). _
                            setHasHeaderPanel(True). _
                            setHeaderDefinitions(pParent.getAliasesListHeaders). _
                            setWidth(pAliasesListWidth). _
                            setHeight(pListHeight)
    Call pAliasesList.render
End Sub

Private Sub insertVerticalSeparator()
    Set pVerticalSeparator = pFrame.Controls.Add(CREATE_LABEL_ID)
    With pVerticalSeparator
        .BackStyle = fmBackStyleOpaque
        .BackColor = pBorderColor
    End With
End Sub

Private Sub insertStandardNamesList()
    Set pStandardNamesList = UI.Create.List(pFrame). _
                            setMultiSelectMode(fmMultiSelectSingle). _
                            setBorderColor(pBorderColor). _
                            setHasHeaderPanel(True). _
                            setItemsPreviewable(True). _
                            setHeaderDefinitions(pParent.getStandardNamesListHeaders). _
                            setWidth(pStandardNamesListWidth). _
                            setHeight(pListHeight)
    Call pStandardNamesList.render
End Sub

Private Sub insertNewNameButton()
    Set pNewItemButton = UI.Create.UIButton(pFrame). _
                            setBackgroundColor(NEUTRAL_BACK_COLOR). _
                            setBorderColor(NEUTRAL_BORDER_COLOR). _
                            setFontColor(NEUTRAL_FONT_COLOR). _
                            setHeight(pButtonHeight). _
                            setCaption(MsgService.getText(CREATE_NEW_ITEM_BUTTON_CAPTION))
    Call pNewItemButton.render
End Sub

Private Sub insertAssignButton()
    Set pAssignButton = UI.Create.UIButton(pFrame). _
                            setBackgroundColor(NEUTRAL_BACK_COLOR). _
                            setBorderColor(NEUTRAL_BORDER_COLOR). _
                            setFontColor(NEUTRAL_FONT_COLOR). _
                            setHeight(pButtonHeight). _
                            setCaption(MsgService.getText(ASSIGN_ALIAS_BUTTON_CAPTION))
    Call pAssignButton.render
End Sub

Private Sub arrangeComponents()
'    Dim validator As MielkUI.UIControl
    Dim top As Single:                      top = pPadding
    Dim left As Single:                     left = pPadding
    '------------------------------------------------------------------------------------------------------

    Call arrangeWarningControls(top, left)

    With pAliasesList
        Call .move(left, top)
        Call .resize(pAliasesListWidth, pListHeight)
        left = .getRight + pControlsSpace
    End With
    
    Call UI.Forms.locateControl(pVerticalSeparator, left, top)
    Call UI.Forms.resizeControl(pVerticalSeparator, PIXEL_SIZE, pListHeight)
    left = UI.Forms.getRight(pVerticalSeparator) + pControlsSpace
    
    With pStandardNamesList
        Call .move(left, top)
        Call .resize(pStandardNamesList, calculateStandardNamesListHeight)
        top = .getBottom + pControlsSpace
    End With

    With pAssignButton
        Call .move(left, top)
        Call .setWidth(calculateButtonWidth)
        left = .getRight + pControlsSpace
    End With

    With pNewItemButton
        Call .move(left, top)
        Call .setWidth(calculateButtonWidth)
    End With

    If pWarningStatus Then
        '[Resize warning status again to adjust width]
        Call UI.Forms.resizeControl(pWarningBackground, calculateWarningPanelWidth)
    End If

    Call pForm.arrangeComponents

End Sub

Private Sub arrangeWarningControls(ByRef top As Single, ByRef left As Single)
    pWarningBackground.Visible = pWarningStatus
    pWarningIcon.Visible = pWarningStatus
    pWarningMessage.Visible = pWarningStatus
    If pWarningStatus Then
        Call UI.Forms.locateControl(pWarningBackground, left, top)
        Call UI.Forms.resizeControl(pWarningBackground, pFrame.InsideWidth, WARNING_PANEL_HEIGHT)
        Call UI.Forms.locateControl(pWarningIcon, left + pControlsSpace, calculateWarningIconTop)
                
        With pWarningMessage
            .caption = pParent.getWarningMessage
            .width = pAliasesListWidth + pStandardNamesListWidth
            .AutoSize = True
            .AutoSize = False
        End With
        Call UI.Forms.locateControl(pWarningMessage, calculateWarningMessageLeft, calculateWarningMessageTop)
        
        top = UI.Forms.getBottom(pWarningBackground) + pControlsSpace
    End If

End Sub

Private Sub updateWarningMessage()
    If UI.Forms.isControlValid(pWarningMessage) Then
        pWarningMessage.caption = pParent.getWarningMessage
    End If
End Sub

Private Function calculateWarningPanelWidth() As Single
    calculateWarningPanelWidth = pStandardNamesList.getRight - pAliasesList.getLeft
End Function

Private Function calculateWarningIconTop() As Single
    calculateWarningIconTop = pWarningBackground.top + (WARNING_PANEL_HEIGHT - WARNING_ICON_SIZE) / 2
End Function

Private Function calculateWarningMessageLeft() As Single
    calculateWarningMessageLeft = UI.Forms.getRight(pWarningIcon) + 2 * pControlsSpace
End Function

Private Function calculateWarningMessageTop() As Single
    calculateWarningMessageTop = pWarningBackground.top + (WARNING_PANEL_HEIGHT - pWarningMessage.height) / 2
End Function

Private Function calculateStandardNamesListHeight() As Single
    calculateStandardNamesListHeight = pListHeight - pControlsSpace - pButtonHeight
End Function

Private Function calculateButtonWidth() As Single
    calculateButtonWidth = (pStandardNamesListWidth - pControlsSpace) / 2
End Function



'[Values]
Private Sub setupValues()
    Dim items As Scripting.Dictionary
    Dim captions As Scripting.Dictionary
    Dim specialFormats As Scripting.Dictionary
    Dim constKeys As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    With pParent
        Set items = .getStandardNamesListItems
        Set captions = .getStandardNamesDisplayeValues
        Set constKeys = .getAttachedStandardNames
        Set specialFormats = getSpecialItemsFormatsDict
    End With
    
    Call pAssignButton.setIsEnabled(False)
    With pStandardNamesList
        Call .setValues(items, , captions, constKeys, specialFormats)
        Call .setFocus(False, False, True)
    End With
End Sub

Private Function getSpecialItemsFormatsDict() As Scripting.Dictionary
    Dim formats As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If formats Is Nothing Then
        Set formats = F.dictionaries.Create(False)
        With formats
            Call .Add(UI.Props.FontColor.getName, VBA.RGB(100, 100, 100))
            Call .Add(UI.Props.FontItalic.getName, True)
        End With
    End If
    
    Set getSpecialItemsFormatsDict = F.dictionaries.Create
    With getSpecialItemsFormatsDict
        Call .Add(StandarizerContainerType_Unassigned, formats)
        Call .Add(StandarizerContainerType_Junk, formats)
    End With
    
End Function





'[Actions]
Private Sub handleStandardNameSelectionChange()
    Call pStandardNamesList.resetSelection(True)
    With pAliasesList
        Call .changeHeaderCaption(1, pParent.getAliasesListTitle)
        Call .setValues(pParent.getPreviewItemAliases)
        Call .setFocus(False, True)
    End With
End Sub

Private Sub updateAssignButtonEnability()
    Dim enabled As Boolean
    '------------------------------------------------------------------------------------------------------
    
    If Not pAliasesList.isAnyItemSelected Then
    ElseIf Not pStandardNamesList.isAnyItemSelected Then
    ElseIf pStandardNamesList.getSelectedKey = pParent.getPreviewKey Then
    Else
        enabled = True
    End If
    
    Call pAssignButton.setIsEnabled(enabled)
    
End Sub

Private Sub refreshWarningStatus()
    Dim status As Boolean
    '------------------------------------------------------------------------------------------------------
    status = pParent.hasAnyUnassigned
    Call updateWarningMessage
    If status <> pWarningStatus Then
        pWarningStatus = status
        Call arrangeComponents
    End If
End Sub



Private Sub changeAssignment()
    Dim aliases As VBA.Collection
    Dim standardNameObject As Variant
    Dim standardNameListKey As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set aliases = pAliasesList.getSelectedKey
    Call F.Variables.assign(standardNameObject, pStandardNamesList.getSelectedValue)
    standardNameListKey = pStandardNamesList.getSelectedKey
    
    Call pParent.changeAssignment(aliases, standardNameObject, standardNameListKey)
    
End Sub

Private Sub updateStandardNameListLine(key As Variant)
    Dim value As Variant
    Dim caption As String
    '------------------------------------------------------------------------------------------------------
    Call F.Variables.assign(value, pStandardNamesList.getValueForKey(key))
    caption = pParent.getCaptionForStandardName(value)
    Call pStandardNamesList.changeItemCaption(key, caption)
End Sub


Private Sub handleNewItemClick()
    Dim aliases As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    Set aliases = pAliasesList.getSelectedKey
    If F.Collections.isNonEmptyCollection(aliases) Then
        Call pParent.assignToNewItem(aliases)
    Else
        Call pParent.addNewItem
    End If
End Sub




'[Events]
Private Sub pStandardNamesList_AfterPreviewChanged(key As Variant, item As Variant, index As Long)
    Call pParent.changePreviewStandardName(item, key)
End Sub


Private Sub pAliasesList_AfterChange()
    Call updateAssignButtonEnability
End Sub

Private Sub pStandardNamesList_AfterChange()
    Call updateAssignButtonEnability
End Sub

Private Sub pAssignButton_Click()
    Call changeAssignment
End Sub

Private Sub pNewItemButton_Click()
    Call handleNewItemClick
End Sub



'[Parent's events]

Private Sub pParent_PreviewChange()
    Call handleStandardNameSelectionChange
End Sub

Private Sub pParent_AssignmentChange(previewKey As Variant, newKey As Variant)
    Call updateStandardNameListLine(previewKey)
    Call updateStandardNameListLine(newKey)
    Call runAfterAssignmentActions
End Sub

Private Sub pParent_AssignedToNewItem(item As IAliasable, previewKey As Variant)
    Call updateStandardNameListLine(previewKey)
    Call addNewItemToStandardNamesList(item)
    Call runAfterAssignmentActions
End Sub

Private Sub pParent_ItemAdded(item As IAliasable)
    Call addNewItemToStandardNamesList(item)
End Sub

Private Sub runAfterAssignmentActions()
    Call pAliasesList.resetActiveLine
    Call pAliasesList.setValues(pParent.getPreviewItemAliases)
    Call pStandardNamesList.resetSelection(True)
    Call refreshWarningStatus
    Call pAliasesList.setFocus(False, True, False)
End Sub

Private Sub addNewItemToStandardNamesList(item As IAliasable)
    Dim label As String
    '------------------------------------------------------------------------------------------------------
    label = pParent.getCaptionForStandardName(item)
    Call pStandardNamesList.addItem(item.getId, item, label)
End Sub



'[Form's events]

Private Sub pForm_Confirm()
    Call pParent.confirmChanges
End Sub
