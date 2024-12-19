VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FConstructors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FConstructors"
'----------------------------------------------------------------------------------------------------------


Public Function filterList(parent As IFilterable, key As String) As WFilterList
    Set filterList = New WFilterList
    With filterList
        Call .setParent(parent)
        Call .setKey(key)
    End With
End Function


Public Function UIIcon(frame As MSForms.frame) As UIIcon
    Set UIIcon = New UIIcon
    With UIIcon
        Call .setFrame(frame)
    End With
End Function

Public Function UILabel(frame As MSForms.frame) As UILabel
    Set UILabel = New UILabel
    With UILabel
        Call .setFrame(frame)
    End With
End Function

Public Function UITextBox(frame As MSForms.frame) As UITextBox
    Set UITextBox = New UITextBox
    With UITextBox
        Call .setFrame(frame)
    End With
End Function

Public Function UIComboBox(frame As MSForms.frame) As UIComboBox
    Set UIComboBox = New UIComboBox
    With UIComboBox
        Call .setFrame(frame)
    End With
End Function

Public Function UIButton(frame As MSForms.frame) As UIButton
    Set UIButton = New UIButton
    With UIButton
        Call .setFrame(frame)
    End With
End Function

Public Function xxGridHeader(frame As MSForms.frame) As xxGridHeader
    Set xxGridHeader = New xxGridHeader
    With xxGridHeader
        Call .setFrame(frame)
    End With
End Function

Public Function GridManager(frame As MSForms.frame) As XGridManager
    Set GridManager = New XGridManager
    With GridManager
        Call .setFrame(frame)
    End With
End Function

Public Function GridLine(manager As XGridManager, frame As MSForms.frame, item As IGridItem) As xxGridLine
    Set GridLine = New xxGridLine
    With GridLine
        Call .setManager(manager)
        Call .setFrame(frame)
        Call .setItem(item)
    End With
End Function

Public Function GridHeaderLine(manager As XGridManager, frame As MSForms.frame) As xxGridHeaderLine
    Set GridHeaderLine = New xxGridHeaderLine
    With GridHeaderLine
        Call .setManager(manager)
        Call .setFrame(frame)
    End With
End Function

Public Function column() As DColumn
    Set column = New DColumn
End Function


Public Function PropertiesPanel(item As IEditItem, window As WTemplate, frame As MSForms.frame) As XPropertiesPanel
    Set PropertiesPanel = New XPropertiesPanel
    With PropertiesPanel
        Call .setEditItem(item)
        Call .setWindow(window)
        Call .setParentFrame(frame)
    End With
End Function


Public Function PropertyLine(parent As XPropertiesPanel, dictDefinition As Scripting.Dictionary) As xxPropertyLine
    Set PropertyLine = New xxPropertyLine
    With PropertyLine
        Call .setParent(parent)
        Call .setDefinition(dictDefinition)
    End With
End Function

Public Function ProgressPanel(window As WTemplate, frame As MSForms.frame) As XProgressPanel
    Set ProgressPanel = New XProgressPanel
    With ProgressPanel
        Call .setWindow(window)
        Call .setFrame(frame)
    End With
End Function


Public Function WindowTemplate() As WTemplate
    Set WindowTemplate = New WTemplate
End Function

Public Function SelectItemWindow(items As Scripting.Dictionary, Optional ByVal windowTitle As String) As WSelectItem
    Const DEFAULT_WINDOW_TITLE_TAG As String = "SelectItem.WindowTitle"
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(windowTitle) = 0 Then windowTitle = MsgService.getText(DEFAULT_WINDOW_TITLE_TAG)
    
    Set SelectItemWindow = New WSelectItem
    With SelectItemWindow
        Call .setTitle(windowTitle)
        Call .setItems(items)
    End With
    
End Function

Public Function datePicker(Optional defaultDate As Variant, Optional ByVal windowTitle As String) As WDatePicker
    Dim initialDate As Date
    '------------------------------------------------------------------------------------------------------
    If VBA.IsDate(defaultDate) Then initialDate = VBA.CDate(defaultDate) Else initialDate = VBA.Date
    Set datePicker = New WDatePicker
    With datePicker
        Call .setDefaultDate(initialDate)
        If VBA.Len(windowTitle) Then Call .setTitle(windowTitle)
    End With
End Function

Public Function CustomPanel() As UICustomPanel
    Set CustomPanel = New UICustomPanel
End Function

Public Function OptionsPanel() As XOptionsPanel
    Set OptionsPanel = New XOptionsPanel
End Function

Public Function OptionLine(parent As XOptionsPanel, key As Variant, value As Variant) As xxOptionLine
    Set OptionLine = New xxOptionLine
    With OptionLine
        Call .setParent(parent)
        Call .setKey(VBA.CStr(key))
        Call .setItem(value)
    End With
End Function


Public Function validationLabel(frame As MSForms.frame) As UIValidationLabel
    Set validationLabel = New UIValidationLabel
    With validationLabel
        Call .setFrame(frame)
    End With
End Function





Public Function locationProps() As DLocation
    Set locationProps = New DLocation
End Function

Public Function LocationPropsForCurrentCursor() As DLocation
    Dim position As Coordinate
    Dim dpiScaling As Coordinate
    '------------------------------------------------------------------------------------------------------
    
    position = F.Screen.getCursorPosition
    dpiScaling = F.Screen.getDpiScaling
    
    Set LocationPropsForCurrentCursor = UI.Create.locationProps. _
                                setAnchorLeft(position.x * PIXEL_SIZE / dpiScaling.x). _
                                setAnchorTop(position.y * PIXEL_SIZE / dpiScaling.y). _
                                setAnchorPoint(AnchorPoint_TopMiddle)
    
End Function

Public Function LocationPropsForScreenCenter() As DLocation
    Dim screenArea As area
    Dim dpiScaling As Coordinate
    Dim x As Single
    Dim y As Single
    '------------------------------------------------------------------------------------------------------
    
    screenArea = F.Screen.getActiveMonitorWorkingArea
    dpiScaling = F.Screen.getDpiScaling
    
    With screenArea
        x = .left + .width / 2
        y = .top + .height / 2
    End With
    
    Set LocationPropsForScreenCenter = UI.Create.locationProps. _
                                setAnchorLeft(x / dpiScaling.x). _
                                setAnchorTop(y / dpiScaling.y). _
                                setAnchorPoint(AnchorPoint_Middle)
    
End Function

Public Function TreeView() As clsTreeView
    Set TreeView = New clsTreeView
End Function

Public Function TreeNode() As clsNode
    TreeNode = New clsNode
End Function

Public Function SelectFromTreeWindow(items As Scripting.Dictionary, _
                                            Optional initialSelection As Scripting.Dictionary, _
                                            Optional ByVal windowTitle As String) As WSelectFromTree
    Const DEFAULT_WINDOW_TITLE As String = "SelectItem.WindowTitle"
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(windowTitle) = 0 Then windowTitle = MsgService.getText(DEFAULT_WINDOW_TITLE)
    Set SelectFromTreeWindow = New WSelectFromTree
    With SelectFromTreeWindow
        Call .setTitle(windowTitle)
        Call .setDataStructure(items)
        Call .setInitialSelection(initialSelection)
    End With
    
End Function



'[UIList]
Public Function list(value As MSForms.UserForm) As UIList
    Set list = New UIList
    With list
        Call .setParent(value)
    End With
End Function

Friend Function ListHeader(list As UIList, definition As DListHeader) As UIListHeader
    Set ListHeader = New UIListHeader
    With ListHeader
        Call .setParent(list)
        Call .setDefinition(definition)
    End With
End Function

Public Function ListHeaderDefinition(caption As String, prop As Variant) As DListHeader
    Set ListHeaderDefinition = New DListHeader
    With ListHeaderDefinition
        Call .setCaption(caption)
        Call .setProperty(prop)
    End With
End Function

Friend Function ListLine() As UIListLine
    Stop
End Function

