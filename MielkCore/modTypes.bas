Attribute VB_Name = "modTypes"
Option Explicit

Private Const CLASS_NAME As String = "modTypes"
    '----------------------------------------------------------------------------------------------------------

Public Type Coordinate
    X As Single
    Y As Single
End Type

Public Type area
    Left As Single
    Top As Single
    width As Single
    height As Single
End Type

Public Type ExcelState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    Interactive As Boolean
    '[Sheet protection]
    Sheet As Excel.Worksheet
    SheetProtection As Boolean
    ProtectionLevels As Excel.Protection
    protectionPassword As String
    ProtectDrawingObjects As Boolean
    ProtectContents As Boolean
    ProtectScenarios As Boolean
End Type

Public Type RgbArray
    red As Byte
    green As Byte
    blue As Byte
End Type



Public Function stateToString(state As ExcelState) As String
    With state
        stateToString = "ScreenUpdating: " & .ScreenUpdating & "; " & _
                        "EnableEvents: " & .EnableEvents & "; " & _
                        "Interactive: " & .Interactive & "; " & _
                        "Sheet: " & (Not .Sheet Is Nothing) & "; " & _
                        "SheetProtection: " & .SheetProtection & "; " & _
                        "ProtectionPassword: " & .protectionPassword
    End With
End Function

Public Function areaToString(area As area) As String
    With area
        areaToString = .width & "x" & .height & _
                        " | x: " & .Left & ", y: " & .Top
    End With
End Function
