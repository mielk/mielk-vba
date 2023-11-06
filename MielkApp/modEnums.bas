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
    ItemProcessStatus_Unknown = 0
    ItemProcessStatus_Correct = 1
    ItemProcessStatus_Warning = 2
    ItemProcessStatus_Error = 3
    ItemProcessStatus_Rejected = 4
    ItemProcessStatus_Skipped = 5
End Enum
'----------------------------------------------------------------------------------------------------------




Public Function getItemProcessStatusName(status As ItemProcessStatusEnum) As String
    Select Case status
        Case ItemProcessStatus_Unknown:                 getItemProcessStatusName = MsgService.getText("Status.Unknown")
        Case ItemProcessStatus_Correct:                 getItemProcessStatusName = MsgService.getText("Status.Success")
        Case ItemProcessStatus_Error:                   getItemProcessStatusName = MsgService.getText("Status.Error")
        Case ItemProcessStatus_Warning:                 getItemProcessStatusName = MsgService.getText("Status.Warning")
        Case ItemProcessStatus_Rejected:                getItemProcessStatusName = MsgService.getText("Status.Rejected")
        Case ItemProcessStatus_Skipped:                 getItemProcessStatusName = MsgService.getText("Status.Skipped")
    End Select
End Function

