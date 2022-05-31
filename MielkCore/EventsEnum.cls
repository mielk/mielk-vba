VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventsEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "EventsEnum"
'----------------------------------------------------------------------------------------------------------
Private Const EVENT_AFTER_UPDATE As String = "after-update"
Private Const EVENT_BEFORE_UPDATE As String = "before-update"
Private Const EVENT_CHANGE As String = "change"
Private Const EVENT_CLICK As String = "click"
Private Const EVENT_DOUBLE_CLICK As String = "double-click"
Private Const EVENT_ENTER As String = "enter"
Private Const EVENT_LEAVE As String = "leave"
Private Const EVENT_KEY_DOWN As String = "key-down"
Private Const EVENT_KEY_UP As String = "key-up"
Private Const EVENT_MOUSE_DOWN As String = "mouse-down"
Private Const EVENT_MOUSE_UP As String = "mouse-up"
Private Const EVENT_MOVE As String = "move"
Private Const EVENT_RESIZE As String = "resize"
Private Const EVENT_VALIDATION_CHANGE As String = "validation-change"
'----------------------------------------------------------------------------------------------------------
Private pEventsById As Scripting.Dictionary
Private pEventsByName As Scripting.Dictionary
Private pEventsByAliases As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------





Private Sub Class_Initialize()
    Call initializeEvents
End Sub


Private Sub initializeEvents()
    '[Initialize containers]
    Set pEventsById = F.Dictionaries.Create(False)
    Set pEventsByName = F.Dictionaries.Create(False)
    Set pEventsByAliases = F.Dictionaries.Create(False)
    
    '[Add event objects]
    Call addEvent(1, EVENT_CLICK, "click")
    Call addEvent(2, EVENT_DOUBLE_CLICK, "doubleClick", "double-click", "double_click", "dblClick", "dbl-click", "dbl_click")
    Call addEvent(3, EVENT_AFTER_UPDATE, "afterUpdate", "after-update", "after_update")
    Call addEvent(4, EVENT_BEFORE_UPDATE, "beforeUpdate", "before-update", "before_update")
    Call addEvent(5, EVENT_ENTER, "enter")
    Call addEvent(6, EVENT_LEAVE, "exit", "leave")
    Call addEvent(7, EVENT_KEY_DOWN, "key-down", "keyDown", "key_down")
    Call addEvent(8, EVENT_KEY_UP, "key-up", "keyUp", "key_up")
    Call addEvent(9, EVENT_MOUSE_DOWN, "mouse-down", "mouseDown", "mouse_down")
    Call addEvent(10, EVENT_MOUSE_UP, "mouse-up", "mouseUp", "mouse_up")
    Call addEvent(11, EVENT_CHANGE, "change")
    Call addEvent(12, EVENT_VALIDATION_CHANGE, "validationChange", "validation-change", "validation_change")
    'Size & position
    Call addEvent(13, EVENT_RESIZE, "resize")
    Call addEvent(14, EVENT_MOVE, "move")
    
End Sub


Private Sub addEvent(id As Long, eventName As String, ParamArray aliases() As Variant)
    Dim item As Event_
    Dim varAlias As Variant
    '----------------------------------------------------------------------------------------------------------
    
    Set item = New Event_
    With item
        Call .setId(id)
        Call .setName(eventName)
        '[Add to repositories]
        Call pEventsById.Add(id, item)
        Call pEventsByName.Add(eventName, item)
        '[Add aliases]
        For Each varAlias In aliases
            Call item.addAlias(VBA.CSng(varAlias))
            Call pEventsByAliases.Add(varAlias, item)
        Next varAlias
    End With
    
End Sub




Public Function ById(id As Long) As Event_
    Set ById = F.Dictionaries.getObject(pEventsById, id)
End Function

Public Function ByName(eventName As String) As Event_
    Set ByName = F.Dictionaries.getObject(pEventsByName, eventName)
End Function

Public Function ByAlias(alias As String) As Event_
    Set ByAlias = F.Dictionaries.getObject(pEventsByAliases, alias)
End Function


'[Quick access]
Public Function AfterUpdate() As Event_: Set AfterUpdate = F.Dictionaries.getObject(pEventsByName, EVENT_AFTER_UPDATE): End Function
Public Function BeforeUpdate() As Event_: Set BeforeUpdate = F.Dictionaries.getObject(pEventsByName, EVENT_BEFORE_UPDATE): End Function
Public Function Change() As Event_: Set Change = F.Dictionaries.getObject(pEventsByName, EVENT_CHANGE): End Function
Public Function Click() As Event_: Set Click = F.Dictionaries.getObject(pEventsByName, EVENT_CLICK): End Function
Public Function DoubleClick() As Event_: Set DoubleClick = F.Dictionaries.getObject(pEventsByName, EVENT_DOUBLE_CLICK): End Function
Public Function Enter() As Event_: Set Enter = F.Dictionaries.getObject(pEventsByName, EVENT_ENTER): End Function
Public Function Leave() As Event_: Set Leave = F.Dictionaries.getObject(pEventsByName, EVENT_LEAVE): End Function
Public Function KeyDown() As Event_: Set KeyDown = F.Dictionaries.getObject(pEventsByName, EVENT_KEY_DOWN): End Function
Public Function KeyUp() As Event_: Set KeyUp = F.Dictionaries.getObject(pEventsByName, EVENT_KEY_UP): End Function
Public Function MouseDown() As Event_: Set MouseDown = F.Dictionaries.getObject(pEventsByName, EVENT_MOUSE_DOWN): End Function
Public Function MouseUp() As Event_: Set MouseUp = F.Dictionaries.getObject(pEventsByName, EVENT_MOUSE_UP): End Function
Public Function Move() As Event_: Set Move = F.Dictionaries.getObject(pEventsByName, EVENT_MOVE): End Function
Public Function Resize() As Event_: Set Resize = F.Dictionaries.getObject(pEventsByName, EVENT_RESIZE): End Function
Public Function ValidationChange() As Event_: Set ValidationChange = F.Dictionaries.getObject(pEventsByName, EVENT_VALIDATION_CHANGE): End Function

