Attribute VB_Name = "modTypes"
Option Explicit

Private Const CLASS_NAME As String = "modTypes"

'[Types] --------------------------------------------------------------------------------------------------
Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
'----------------------------------------------------------------------------------------------------------

