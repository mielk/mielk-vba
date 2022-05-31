Attribute VB_Name = "modEnums"
Option Explicit

Private Const CLASS_NAME As String = "modEnums"
'---------------------------------------------------------------------------------------------------------

Public Enum FilterTypeEnum
    FilterType_Text = 1
    FilterType_List = 2
    FilterType_Numbers = 3
End Enum

Public Enum ControlTypeEnum
    ControlType_Label = 1
    ControlType_Icon = 2
    ControlType_TextBox = 3
    ControlType_CheckBox = 4
    ControlType_ComboBox = 5
    ControlType_LabelWithActionButton = 6
End Enum

Public Enum AnchorPointEnum
    AnchorPoint_None = 0
    AnchorPoint_Middle = 1
    AnchorPoint_TopLeft = 2
    AnchorPoint_BottomLeft = 3
    AnchorPoint_TopRight = 4
    AnchorPoint_BottomRight = 5
    AnchorPoint_TopMiddle = 6
End Enum

Public Enum ControlAlignmentEnum
    ControlAlignment_Center = 0
    ControlAlignment_Left = 1
    ControlAlignment_Right = 2
End Enum
