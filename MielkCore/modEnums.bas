Attribute VB_Name = "modEnums"
'#FORCHECK


'
''[Containers]
'Public Enum ContainerTypeEnum
'    ContainerType_Array1D = 1
'    ContainerType_Collection = 2
'    ContainerType_Dictionary = 3
'    ContainerType_Array2D = 4
'End Enum
'
'
'
''[Application]
'Public Enum AppModeEnum
'    AppMode_PROD = 1
'    AppMode_DEV = 2
'    AppMode_TEST = 3
'End Enum





Option Explicit

Private Const CLASS_NAME As String = "modEnums"
'----------------------------------------------------------------------------------------------------------


'[Errors]
Public Enum ErrorHandlingStatusEnum
    errorHandling_AllErrors = 1
    errorHandling_ClassErrors = 2
    errorHandling_UnhandledErrors = 3
End Enum


'[Dictionaries]
Public Enum DuplicateBehaviourEnum
    duplicateBehaviour_ThrowError = 0
    duplicateBehaviour_Override = 1
    duplicateBehaviour_Skip = 2
    duplicateBehaviour_WarningInImmediateWindow = 3
End Enum

Public Enum DictPartEnum
    DictPart_KeyAndValue = 0
    DictPart_KeyOnly = 1
    DictPart_ValueOnly = 2
End Enum

Public Enum DictCompareModeEnum
    DictCompareMode_Binary = 0
    DictCompareMode_Text = 1
    DictCompareMode_Database = 2
End Enum


'[Strings]
Public Enum StringifyModeEnum
    StringifyMode_Normal = 1
    StringifyMode_Db = 2
    StringifyMode_Xml = 3
End Enum


Public Enum LetterCasingEnum
    LetterCasing_Unknown = 0
    LetterCasing_Lower = 1
    LetterCasing_Upper = 2
    LetterCasing_Sentence = 3
    LetterCasing_Camel = 4
    LetterCasing_StartWithLower = 5
End Enum


'[Sql]
Public Enum SqlWhereEnum
    SqlWhere_Equal = 0
    SqlWhere_LessThan = 1
    SqlWhere_LessEqualThan = 2
    SqlWhere_GreaterThan = 3
    SqlWhere_GreaterEqualThan = 4
    SqlWhere_Like = 5
    SqlWhere_In = 6
End Enum


'[Access]
Public Enum ReadWriteModeEnum
    ReadWriteMode_ReadOnly = 0
    ReadWriteMode_ReadWrite = 1
End Enum

'[Dates]
Public Enum DateVisibilityEnum
    DateVisibility_DateOnly = 0
    DateVisibility_DateTime = 1
    DateVisibility_TimeOnly = 2
End Enum

Public Enum MonthNameTypeEnum
    MonthName_ByRegionalSettings = 0
    MonthName_English = 0
End Enum


'[Developers]
Public Enum ReferenceTypeEnum
    ReferenceType_Unknown = 0
    ReferenceType_Project = 1
    ReferenceType_InternalLibrary = 2
    ReferenceType_ExternalLibrary = 3
End Enum


'[Other]
Public Enum ComparisonResultEnum
    ComparisonResult_Common = 1
    ComparisonResult_OnlyInBase = 2
    ComparisonResult_OnlyInCompared = 3
End Enum


'[System]
Public Enum SystemSettingEnum
    SystemSetting_DecimalSeparator = 1
    SystemSetting_ListSeparator = 2
    SystemSetting_DateFormat = 3
End Enum



Public Function convertErrorHandlingStatusToString(Status As ErrorHandlingStatusEnum) As String
    Select Case Status
        Case errorHandling_AllErrors:           convertErrorHandlingStatusToString = "Break on All Errors"
        Case errorHandling_ClassErrors:         convertErrorHandlingStatusToString = "Break in Class Module"
        Case errorHandling_UnhandledErrors:     convertErrorHandlingStatusToString = "Break on Unhandled Errors"
    End Select
End Function
