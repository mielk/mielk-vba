VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ECodeSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "ECodeSelection"
'----------------------------------------------------------------------------------------------------------
Private pComponent As VBIDE.VBComponent
Private pStartRow As Long
Private pStartCol As Long
Private pEndRow As Long
Private pEndCol As Long
'----------------------------------------------------------------------------------------------------------


'[SETTERS]
Public Function setComponent(value As VBIDE.VBComponent) As ECodeSelection
    Set setComponent = Me
    Set pComponent = value
End Function

Public Function setStartRow(value As Long) As ECodeSelection
    Set setStartRow = Me
    pStartRow = value
End Function

Public Function setStartCol(value As Long) As ECodeSelection
    Set setStartCol = Me
    pStartCol = value
End Function

Public Function setEndRow(value As Long) As ECodeSelection
    Set setEndRow = Me
    pEndRow = value
End Function

Public Function setEndCol(value As Long) As ECodeSelection
    Set setEndCol = Me
    pEndCol = value
End Function

Public Sub readFromActivePane()
    Dim activePane As VBIDE.CodePane
    '------------------------------------------------------------------------------------------------------
    Set activePane = Fn.getActivePane
    Set pComponent = activePane.CodeModule.parent
    Call activePane.GetSelection(pStartRow, pStartCol, pEndRow, pEndCol)
End Sub





'[GETTERS]
Public Function getComponent() As VBIDE.VBComponent
    Set getComponent = pComponent
End Function

Public Function getStartRow() As Long
    getStartRow = pStartRow
End Function

Public Function getStartCol() As Long
    getStartCol = pStartCol
End Function

Public Function getEndRow() As Long
    getEndRow = pEndRow
End Function

Public Function getEndCol() As Long
    getEndCol = pEndCol
End Function

Public Function countLines() As Long
    countLines = pEndRow - pStartRow + 1
End Function

Public Function readContent() As String
    Dim cm As VBIDE.CodeModule
    '------------------------------------------------------------------------------------------------------
    Set cm = pComponent.CodeModule
    readContent = F.Dev.getCodeModuleContent(cm, pStartRow, pEndRow - pStartRow + 1)
End Function




'[API]
Public Function toString() As String
    toString = "Current selection || " & _
                "Module: " & pComponent.name & " ||| " & _
                "Start: " & pStartRow & " / " & pStartCol & " | " & _
                "End: " & pEndRow & " / " & pEndCol
End Function

Public Sub replaceWithGivenCode(code As String)
    Dim cm As VBIDE.CodeModule
    Dim linesCounter As Long
    '------------------------------------------------------------------------------------------------------
    
    Set cm = pComponent.CodeModule
    linesCounter = countLines
    
    With cm
        If linesCounter > 1 Then
            Call .DeleteLines(pStartRow + 1, linesCounter - 1)
        End If
        Call .ReplaceLine(pStartRow, code)
    End With
End Sub

Public Sub insertCode(code As String)
    Dim cm As VBIDE.CodeModule
    Dim linesCounter As Long
    '------------------------------------------------------------------------------------------------------
    
    Set cm = pComponent.CodeModule
    linesCounter = countLines
    
    With cm
        If linesCounter > 1 Then
            Call .DeleteLines(pStartRow + 1, linesCounter - 1)
        End If
        
        If pStartRow > .CountOfLines Then
            Call .AddFromString(code)
        Else
            Call .ReplaceLine(pStartRow, code)
        End If
    End With
End Sub
