Attribute VB_Name = "modEnums"
Option Explicit

Private Const CLASS_NAME As String = "modEnums"
'----------------------------------------------------------------------------------------------------------

Public Enum StandarizerContainerTypeEnum
    StandarizerContainerType_Unassigned = -2
    StandarizerContainerType_Junk = -1
    StandarizerContainerType_AliasableObject = 1
End Enum

Public Enum ItemProcessStatusEnum
    ItemProcessStatus_Before = 0
    ItemProcessStatus_Right = 1
    ItemProcessStatus_Warning = 2
    ItemProcessStatus_Error = 3
    ItemProcessStatus_Rejected = 4
End Enum



Sub test()
    Dim xls As Excel.Application
    Set xls = VBA.CreateObject("Excel.Application")
    xls.Visible = True
    Stop
End Sub
