VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FDialogs"
'----------------------------------------------------------------------------------------------------------




Public Function selectItems(items As Scripting.Dictionary, Optional windowTitle As String, _
                            Optional defaultItem As Variant, _
                            Optional hasSearchPanel As Boolean = False, _
                            Optional hasTitleBar As Boolean = True, _
                            Optional multiSelectAllowed As Boolean = False, _
                            Optional addingAllowed As Boolean = False, _
                            Optional locationProps As DLocation) As VBA.Collection
    Const METHOD_NAME As String = "selectItems"
    '------------------------------------------------------------------------------------------------------
    Dim window As WSelectItem
    '------------------------------------------------------------------------------------------------------
    
    Set window = UI.Create.SelectItemWindow(items, windowTitle). _
                    setAddingAllowed(addingAllowed). _
                    setMultiSelectAllowed(multiSelectAllowed). _
                    setHasTitlePanel(hasTitleBar). _
                    setHasSearchPanel(hasSearchPanel)

    With window
        Call .display(locationProps)
        Set selectItems = .getSelectedItems
    End With
    
    Set window = Nothing

End Function





'----------------------------------------------------------------------------------------------------------
'Description:       Create new Select window.
'                   This method was introduced due to more and more optional properties in WSelectItem class
'                   and therefore in [selectItem] function that is based on that class.
'
'                   This method is another way to invoke [selectItem] method but it allows to add any
'                   number of optional properties (or skip them all):
'                     Call UI.Dialogs.SelectItems(items, hasSearchPanel:=True, multiSelectAllowed:=True)
'                       =
'                     Call UI.Dialogs.SelectWindow(items).setHasSearchPanel(True).setMutliSelectAllowed(True).execute
'----------------------------------------------------------------------------------------------------------
Public Function selectWindow(items As Scripting.Dictionary, Optional captions As Scripting.Dictionary) As WSelectItem
    Set selectWindow = UI.Create.SelectItemWindow(items)
    Call selectWindow.setCaptions(captions)
End Function



Public Function selectDateFromCalendar(Optional defaultDate As Variant, _
                                       Optional windowTitle As String = "Select date", _
                                       Optional invokingWindow As WTemplate) As Variant
    Dim datePicker As WDatePicker
    Dim d As Date
    Dim location As DLocation
    '------------------------------------------------------------------------------------------------------
    
    If Not F.Variables.isUndefined(defaultDate, True, True) Then
        d = VBA.CDate(defaultDate)
    Else
        d = VBA.Date
    End If
    
    Set datePicker = UI.Create.datePicker(d, MsgService.getText(windowTitle))
    With datePicker
        If Not invokingWindow Is Nothing Then
            Call invokingWindow.sendToBackground
            Set location = UI.Create.LocationPropsForCurrentCursor
        Else
            Set location = UI.Create.LocationPropsForScreenCenter
        End If
        
        Call .display(vbModal, location)
        
        If Not invokingWindow Is Nothing Then
            Call invokingWindow.restoreFromBackground
        End If
        
        If Not .isCancelledByUser Then
            selectDateFromCalendar = VBA.format(.getDate, "YYYY-MM-DD")
        ElseIf defaultDate Then
            selectDateFromCalendar = VBA.format(defaultDate, "YYYY-MM-DD")
        Else
            selectDateFromCalendar = Empty
        End If
        
    End With
    
End Function

