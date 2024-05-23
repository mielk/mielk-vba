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

Public Enum CodeComparisonStatusEnum
    CodeComparisonStatus_Unknown = 0
    CodeComparisonStatus_Equal = 1
    CodeComparisonStatus_Different = 2
    CodeComparisonStatus_BaseOnly = 3
    CodeComparisonStatus_CompareOnly = 4
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




'[RIBBONS]

Public Function getRibbonControlSizeName(size As RibbonControlSize) As String
    Select Case size
        Case RibbonControlSizeRegular:      getRibbonControlSizeName = "normal"
        Case RibbonControlSizeLarge:        getRibbonControlSizeName = "large"
    End Select
End Function

Public Function getRibbonControlSizeFromName(value As String) As RibbonControlSize
    Select Case VBA.LCase(value)
        Case "regular", "normal":           getRibbonControlSizeFromName = RibbonControlSizeRegular
        Case "large":                       getRibbonControlSizeFromName = RibbonControlSizeLarge
    End Select
End Function

Public Function isCallbackProperty(prop As enumProperty) As Boolean
    If prop Is props.id Then
        isCallbackProperty = False
    ElseIf prop Is Props_Project.OnAction Then
        isCallbackProperty = False
    ElseIf prop Is Props_Project.size Then
        isCallbackProperty = False
    Else
        isCallbackProperty = True
    End If
End Function

Public Function getRibbonPropertyXmlTag(prop As enumProperty) As String
    If isCallbackProperty(prop) Then
        getRibbonPropertyXmlTag = f.Strings.Format("get" & f.Strings.toSentenceCase(prop.getName))
    Else
        getRibbonPropertyXmlTag = f.Strings.convertLetterCasing(prop.getName, LetterCasing_StartWithLower)
    End If
End Function

Public Function isStringProperty(prop As enumProperty) As Boolean
    If prop Is Props_Project.Label Then
        isStringProperty = True
    ElseIf prop Is Props_Project.ScreenTip Then
        isStringProperty = True
    Else
        isStringProperty = False
    End If
End Function

Public Function getRibbonPropertyDefaultValue(prop As enumProperty, _
                                    Optional controlType As enumRibbonControlType) As Variant
    If prop Is Props_Project.Visible Then
        '[Visible] ----------------------------------------------------------|
        If controlType Is Nothing Then                                      '|
            getRibbonPropertyDefaultValue = "{checkUserPermission}"         '|
        ElseIf controlType.isContainer Then                                 '|
            getRibbonPropertyDefaultValue = True                            '|
        Else                                                                '|
            getRibbonPropertyDefaultValue = "{checkUserPermission}"         '|
        End If                                                              '|
        '--------------------------------------------------------------------|
        
    ElseIf prop Is Props_Project.Enabled Then
        getRibbonPropertyDefaultValue = True
    ElseIf prop Is Props_Project.size Or prop Is Props_Project.ItemsSize Then
        getRibbonPropertyDefaultValue = getRibbonControlSizeName(RibbonControlSizeLarge)
    End If
End Function

Public Function getAdjustedRibbonProperty(prop As enumProperty, value As Variant) As Variant
    If prop Is Props_Project.size Or prop Is Props_Project.ItemsSize Then
        getAdjustedRibbonProperty = getRibbonControlSizeFromName(VBA.CStr(value))
    Else
        getAdjustedRibbonProperty = value
    End If
End Function
