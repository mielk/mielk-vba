VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeCompactorProcessUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IProcessItem

Private Const CLASS_NAME As String = "CodeCompactorProcessUnit"
'----------------------------------------------------------------------------------------------------------
Private pParent As CodeCompactor
Private pProcessor As ItemsProcessor
Private pfilepath As String
Private pStatus As ItemProcessStatusEnum
Private pMessage As String
'----------------------------------------------------------------------------------------------------------



'[Setters]
Public Function setParent(value As CodeCompactor) As CodeCompactorProcessUnit
    Set setParent = Me
    Set pParent = value
End Function

Public Function setItemsProcessor(value As ItemsProcessor) As CodeCompactorProcessUnit
    Set setItemsProcessor = Me
    Set pProcessor = value
End Function

Public Function setfilepath(value As String) As CodeCompactorProcessUnit
    Set setfilepath = Me
    pfilepath = value
End Function


'[Getters]
Public Function getCaption() As String
    getCaption = pfilepath
End Function

Public Function getMessage() As String
    getMessage = pMessage
End Function

Public Function getStatus() As Long
    getStatus = pStatus
End Function



'[API]
Public Sub process()
    Const COMPILING_TIMEOUT As Long = 60
    '------------------------------------------------------------------------------------------------------
    Dim xls As Excel.Application
    Dim wkb As Excel.Workbook
    Dim project As Object       'VBIDE.VBProject
    Dim password As String
    '------------------------------------------------------------------------------------------------------
    Dim initialTimestamp As Date
    Dim finalTimestamp As Date
    Dim startTime As Double
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    Set xls = pParent.getExcel
    Set wkb = xls.Workbooks.Open(pfilepath, , False)
    Set project = wkb.VBProject
    
    Call project.VBE.Windows.Item(1).SetFocus
    password = pParent.getVbaCodePassword
    If VBA.Len(password) Then Call setPasswordOnProject(project, password)
    
    startTime = VBA.Timer
    initialTimestamp = getFileTimestamp(pfilepath)
    Debug.Print pfilepath & " | Initial timestamp: " & initialTimestamp
    
    Call xls.SendKeys("^s")
    Call Sleep(1000)
    
    Do
        finalTimestamp = getFileTimestamp(pfilepath)
        If VBA.Timer - startTime > COMPILING_TIMEOUT Then
            Call VBA.Err.Raise(1, Description:="Timeout exceeded")
            Exit Do
        ElseIf finalTimestamp > initialTimestamp Then
            Exit Do
        End If
    Loop
    
    Call closeBook(wkb)
    
'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then
        pStatus = ItemProcessStatus_Error
        pMessage = errDescription
    Else
        pStatus = ItemProcessStatus_Correct
        pMessage = "OK"
    End If
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = Err.Number: errDescription = Err.Description
    GoTo ExitPoint
    
End Sub



Private Function getFileTimestamp(filepath As String) As Date
    Const TIMEOUT As Long = 10
    '------------------------------------------------------------------------------------------------------
    Dim f As Scripting.file
    Dim startTime As Double
    '------------------------------------------------------------------------------------------------------
    
    startTime = VBA.Timer
    Do Until getFileTimestamp > 0
        On Error Resume Next
        Set f = fso.GetFile(filepath)
        getFileTimestamp = f.DateLastModified
        On Error GoTo 0
        
        If VBA.Timer - startTime > TIMEOUT Then
            Call VBA.Err.Raise(1, Description:="Error while trying to refer to file " & filepath)
        End If
    Loop
    
End Function


Private Sub closeBook(wkb As Excel.Workbook)
    Const TIMEOUT As Long = 10
    '------------------------------------------------------------------------------------------------------
    Dim startTime As Double
    '------------------------------------------------------------------------------------------------------
    
    startTime = VBA.Timer
    
    Do
        Call VBA.Err.Clear
        On Error Resume Next
        Call wkb.Close(False)
        
        If VBA.Timer - startTime > TIMEOUT Then
            Call VBA.Err.Raise(1, Description:="Error while trying to close workbook " & wkb.FullName)
        End If
        
    Loop While Err.Number
    
End Sub


'https://stackoverflow.com/questions/53323655/locking-a-project-programmatically-in-vba
Private Sub setPasswordOnProject(project As Object, password As String)

    Call project.VBE.CommandBars(1).FindControl(ID:=2578, recursive:=True).Execute
    
    With pParent.getExcel
        Call Sleep(250)
        .SendKeys "^{PGDN}"
        Call Sleep(250)
        
        'Check [Lock project for viewing] checkbox
        .SendKeys "%v"
        Call Sleep(250)
        
        'Enter password
        .SendKeys "%p"
        Call Sleep(250)
        .SendKeys password
        Call Sleep(250)
        
        'Confirm password
        .SendKeys "%c"
        Call Sleep(250)
        .SendKeys password
        Call Sleep(250)
        
        .SendKeys "{ENTER}"
        Call Sleep(250)
        
    End With

End Sub



'[Implementation of IProcessItem]

Private Function IProcessItem_getCaption() As String
    IProcessItem_getCaption = getCaption
End Function

Private Function IProcessItem_getMessage() As Variant
    IProcessItem_getMessage = getMessage
End Function

Private Function IProcessItem_getName() As String
    IProcessItem_getName = getCaption
End Function

Private Function IProcessItem_getStatus() As Long
    IProcessItem_getStatus = getStatus
End Function

Private Sub IProcessItem_preview()
    Stop
End Sub

Private Sub IProcessItem_process()
    Call process
End Sub

Private Function IProcessItem_setItemsProcessor(value As MielkApp.ItemsProcessor) As MielkApp.IProcessItem
    Set IProcessItem_setItemsProcessor = setItemsProcessor(value)
End Function

Private Function IProcessItem_setParams(params As Scripting.IDictionary) As MielkApp.IProcessItem
    Stop
End Function

