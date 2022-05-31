VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "WDatePicker"
'[UI default values] --------------------------------------------------------------------------------------
Private Const DEFAULT_TITLE_TAG As String = "SelectDate"
Private Const DEFAULT_DAY_NAME_LENGTH As Long = 3
Private Const DEFAULT_HEIGHT As Single = 180
Private Const DEFAULT_WIDTH As Single = 220
Private Const DEFAULT_GRID_FONT_SIZE As Single = 10
Private Const DEFAULT_DAY_FONT_SIZE As Single = 10
Private Const DEFAULT_HORIZONTAL_MARGIN As Single = 12
Private Const DEFAULT_TOP_MARGIN As Single = 3
'----------------------------------------------------------------------------------------------------------
Private WithEvents pForm As WTemplate
Attribute pForm.VB_VarHelpID = -1
Private pCalendar As cCalendar
Private pTitle As String
'[State] --------------------------------------------------------------------------------------------------
Private pDate As Date
Private pCancelledByUser As Boolean
'----------------------------------------------------------------------------------------------------------







Private Sub Class_Initialize()
    pTitle = MsgService.getText(DEFAULT_TITLE_TAG)
End Sub



'[SETTERS]
Public Function setTitle(value As String) As WDatePicker
    Set setTitle = Me
    pTitle = value
End Function

Public Function setDefaultDate(value As Date) As WDatePicker
    Set setDefaultDate = Me
    pDate = value
    Call updateDate
End Function



'[GETTERS]
Public Function getDate() As Date
    getDate = VBA.DateSerial(pCalendar.Year, pCalendar.Month, pCalendar.Day)
End Function

Public Function isCancelledByUser() As Boolean
    isCancelledByUser = pCancelledByUser
End Function



'[RENDERING]
Public Sub display(Optional mode As FormShowConstants = vbModal, Optional locationProps As DLocation)
    Call createForm
    Call createCalendar
    Call adjustContentFramePosition
    With pForm
        Call .arrangeComponents
        Call .locateRelativeTo(locationProps)
        Call .display(mode, False)
    End With
End Sub


Private Sub createForm()
    Set pForm = UI.Create.WindowTemplate
    With pForm
        Call .setCaption(pTitle)
        Call .setContentFrameHorizontalMargin(DEFAULT_HORIZONTAL_MARGIN)
        Call .setContentFrameTopMargin(DEFAULT_TOP_MARGIN)
        Call .render
        .getContentFrame.width = DEFAULT_WIDTH
    End With
End Sub

Private Sub createCalendar()
    Set pCalendar = New cCalendar
    With pCalendar
        Call .Add_Calendar_into_Frame(pForm.getContentFrame)
        .UseDefaultBackColors = False
        .DayLength = DEFAULT_DAY_NAME_LENGTH
        .MonthLength = mlENShort
        .height = DEFAULT_HEIGHT
        .width = DEFAULT_WIDTH
        .GridFont.size = DEFAULT_GRID_FONT_SIZE
        .DayFont.size = DEFAULT_DAY_FONT_SIZE
        Call .Refresh
    End With
    Call updateDate
End Sub

Private Sub adjustContentFramePosition()
    Dim frame As MSForms.frame
    '----------------------------------------------------------------------------------------------------------
    With pForm
        Set frame = .getContentFrame
        frame.left = DEFAULT_HORIZONTAL_MARGIN / 2
    End With
End Sub

Private Sub updateDate()
    If Not pCalendar Is Nothing Then
        With pCalendar
            .Year = VBA.Year(pDate)
            .Month = VBA.Month(pDate)
            .Day = VBA.Day(pDate)
        End With
    End If
End Sub







'[EVENTS]
Private Sub pForm_Cancel()
    Call ErrorManager.clear
    pCancelledByUser = True
End Sub
