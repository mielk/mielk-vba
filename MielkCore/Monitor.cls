VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "Monitor"
'----------------------------------------------------------------------------------------------------------
#If VBA7 Then
    Private pHandle As LongPtr
#Else
    Private pHandle As Long
#End If
Private pLeft As Long
Private pTop As Long
Private pRight As Long
Private pBottom As Long
'----------------------------------------------------------------------------------------------------------



'[Setters]
#If VBA7 Then
    Public Function setHandle(value As LongPtr) As Monitor
#Else
    Public Function setHandle(value As Long) As Monitor
#End If
    Set setHandle = Me
    pHandle = value
End Function

Public Function setBounds(Top As Long, Right As Long, Bottom As Long, Left As Long)
    pTop = Top
    pRight = Right
    pBottom = Bottom
    pLeft = Left
End Function





'[Getters]
#If VBA7 Then
    Public Function getHandle() As LongPtr
#Else
    Public Function getHandle() As Long
#End If
    getHandle = pHandle
End Function

Public Function isPointInMonitor(pt As POINTAPI) As Boolean
    If pt.X < pLeft Then
        isPointInMonitor = False
    ElseIf pt.X > pRight Then
        isPointInMonitor = False
'    ElseIf pt.y < pTop Then
'        isPointInMonitor = False
'    ElseIf pt.y > pBottom Then
'        isPointInMonitor = False
    Else
        isPointInMonitor = True
    End If
End Function

Public Function getWorkingArea() As area
    With getWorkingArea
        .Top = pTop
        .Left = pLeft
        .width = pRight - pLeft
        .height = pBottom - pTop
    End With
End Function
