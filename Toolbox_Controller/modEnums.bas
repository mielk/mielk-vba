Attribute VB_Name = "modEnums"
Option Explicit

Private Const CLASS_NAME As String = "modEnums"
'----------------------------------------------------------------------------------------------------------

Public Enum CreatingProjectStepEnum
    CreatingProjectStep_Unknown = 0
    CreatingProjectStep_CreatingProjectFolder = 1
    CreatingProjectStep_ApplyingChangesToTextFiles = 2
    CreatingProjectStep_ApplyingChangesToCode = 3
    CreatingProjectStep_CreatingRibbonComponents = 4
    CreatingProjectStep_FixingReferencesBetweenFiles = 5
End Enum

Public Enum RibbonControlTypeEnum
    RibbonControlType_Unknown = 0
    RibbonControlType_Tab = 1
    RibbonControlType_Group = 2
    RibbonControlType_Menu = 3
    RibbonControlType_Label = 4
    RibbonControlType_Button = 5
    RibbonControlType_Separator = 6
End Enum



Public Function getCreatingProjectStepCaption(step As CreatingProjectStepEnum) As String
    Select Case step
        Case CreatingProjectStep_CreatingProjectFolder
                getCreatingProjectStepCaption = Msg.getText("CreatingNewProject.Steps.CreatingProjectFolder")
        Case CreatingProjectStep_ApplyingChangesToTextFiles
                getCreatingProjectStepCaption = Msg.getText("CreatingNewProject.Steps.ApplyingChangesToTextFiles")
        Case CreatingProjectStep_ApplyingChangesToCode
                getCreatingProjectStepCaption = Msg.getText("CreatingNewProject.Steps.ApplyingChangesToCode")
        Case CreatingProjectStep_CreatingRibbonComponents
                getCreatingProjectStepCaption = Msg.getText("CreatingNewProject.Steps.CreatingRibbonComponents")
        Case CreatingProjectStep_FixingReferencesBetweenFiles
                getCreatingProjectStepCaption = Msg.getText("CreatingNewProject.Steps.FixingReferencesBetweenFiles")
    End Select
End Function



Public Function getRibbonControlTypeFromString(name As String) As RibbonControlTypeEnum
    Select Case VBA.LCase(VBA.Trim(name))
        Case "tab":               getRibbonControlTypeFromString = RibbonControlType_Tab
        Case "group":             getRibbonControlTypeFromString = RibbonControlType_Group
        Case "menu":              getRibbonControlTypeFromString = RibbonControlType_Menu
        Case "label":             getRibbonControlTypeFromString = RibbonControlType_Label
        Case "button":            getRibbonControlTypeFromString = RibbonControlType_Button
        Case "separator":         getRibbonControlTypeFromString = RibbonControlType_Separator
    End Select
End Function
