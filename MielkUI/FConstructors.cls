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

Public Function UIButton(frame As MSForms.frame) As UIButton
    Set UIButton = New UIButton
    With UIButton
        Call .setFrame(frame)
    End With
End Function

Public Function xxListHeader(frame As MSForms.frame) As xxListHeader
    Set xxListHeader = New xxListHeader
    With xxListHeader
        Call .setFrame(frame)
    End With
End Function

Public Function ListManager(frame As MSForms.frame) As XListManager
    Set ListManager = New XListManager
    With ListManager
        Call .setFrame(frame)
    End With
End Function

Public Function ListLine(manager As XListManager, frame As MSForms.frame, item As IListItem) As xxListLine
    Set ListLine = New xxListLine
    With ListLine
        Call .setManager(manager)
        Call .setFrame(frame)
        Call .setItem(item)
    End With
End Function

Public Function ListHeaderLine(manager As XListManager, frame As MSForms.frame) As xxListHeaderLine
    Set ListHeaderLine = New xxListHeaderLine
    With ListHeaderLine
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

Public Function DatePicker(Optional defaultDate As Variant, Optional ByVal windowTitle As String) As WDatePicker
    Dim initialDate As Date
    '------------------------------------------------------------------------------------------------------
    If VBA.IsDate(defaultDate) Then initialDate = VBA.CDate(defaultDate) Else initialDate = VBA.Date
    Set DatePicker = New WDatePicker
    With DatePicker
        Call .setDefaultDate(initialDate)
        If VBA.Len(windowTitle) Then Call .setTitle(windowTitle)
    End With
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


Public Function ValidationLabel(frame As MSForms.frame) As UIValidationLabel
    Set ValidationLabel = New UIValidationLabel
    With ValidationLabel
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
    
    Set LocationPropsForScreenCenter = UI.Create.locationProps. _
                                setAnchorLeft(x / dpiScaling.x). _
                                setAnchorTop(y / dpiScaling.y). _
                                setAnchorPoint(AnchorPoint_Middle)
    
End Function
