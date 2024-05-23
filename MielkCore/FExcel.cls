VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FExcel"
'----------------------------------------------------------------------------------------------------------
Private pBackgroundExcel As Excel.Application
'----------------------------------------------------------------------------------------------------------


'[Metainfo]
Public Function getVbaVersion() As Long
    #If VBA7 Then
        getVbaVersion = 7
    #Else
        getVbaVersion = 6
    #End If
End Function

Public Function getExcelVersion() As Long
    getExcelVersion = f.Numerics.toNumeric(Excel.Application.Version)
End Function



'***** WITHOUT UNIT TESTS *****


Public Function isApplicationValid(app As Excel.Application) As Boolean
    Dim name As String
    '------------------------------------------------------------------------------------------------------
    
    If Not app Is Nothing Then
        On Error Resume Next
        name = app.name
        On Error GoTo 0
        isApplicationValid = (VBA.Len(name) > 0)
    End If
    
End Function

Public Function getColumnHeader(columnIndex As Long) As String
    Const METHOD_NAME As String = "columnHeader"
    '------------------------------------------------------------------------------------------------------
    Const ALPHABET_LENGTH As Byte = 26
    Const ASCII_BEFORE_A As Byte = 64
    '------------------------------------------------------------------------------------------------------
    Dim lTemp As Long
    Dim iLetter As Integer
    '------------------------------------------------------------------------------------------------------

    If columnIndex <= 0 Then GoTo NegativeColumnIndexException

    lTemp = VBA.Int(columnIndex)
    Do
        iLetter = lTemp Mod ALPHABET_LENGTH
        If iLetter = 0 Then iLetter = ALPHABET_LENGTH
        getColumnHeader = VBA.Chr$(iLetter + ASCII_BEFORE_A) & getColumnHeader
        lTemp = Excel.WorksheetFunction.RoundUp((lTemp - ALPHABET_LENGTH) / ALPHABET_LENGTH, 0)
    Loop While lTemp > 0

'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
NegativeColumnIndexException:
    GoTo ExitPoint

End Function

Public Function getColumnHeadersDictionary(Optional maxIndex As Long = 26) As Scripting.Dictionary
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    Set getColumnHeadersDictionary = f.dictionaries.Create(False)
    For i = 1 To maxIndex
        Call getColumnHeadersDictionary.Add(i, getColumnHeader(i))
    Next i
End Function

Public Function getLastAction() As String
    Const UNDO_CTRL_ID As Long = 128
    '------------------------------------------------------------------------------------------------------
    Dim cb As CommandBar
    Dim ctrl As CommandBarControl
    '------------------------------------------------------------------------------------------------------
    Set cb = Application.CommandBars("Standard")
    Set ctrl = cb.FindControl(id:=UNDO_CTRL_ID)
    If Not ctrl Is Nothing Then
        On Error Resume Next
        getLastAction = ctrl.List(1)
    End If
End Function

Public Function isLastActionPaste(Optional includeFilling As Boolean = True, Optional includePasting As Boolean = True) As Boolean
    Const PASTE_REGEX_PATTERN As String = "^(Paste|Wklej)$"
    Const AUTOFILL_REGEX_PATTERN As String = "^(Auto|)(Wype.ni(e|a)nie|Filling)$"
    '------------------------------------------------------------------------------------------------------
    Static objPasteRegex As Object
    Static objAutofillRegex As Object
    '------------------------------------------------------------------------------------------------------
    Dim lastAction As String
    '------------------------------------------------------------------------------------------------------
    
    lastAction = getLastAction
    
    If objPasteRegex Is Nothing Then Set objPasteRegex = f.regex.Create(PASTE_REGEX_PATTERN)
    
    If includePasting Then isLastActionPaste = objPasteRegex.test(lastAction)
    
    If Not isLastActionPaste Then
        If includeFilling Then
            If objAutofillRegex Is Nothing Then Set objAutofillRegex = f.regex.Create(AUTOFILL_REGEX_PATTERN)
            isLastActionPaste = objAutofillRegex.test(lastAction)
        End If
    End If
End Function

Public Sub showHideBook(wkb As Excel.Workbook, show As Boolean)
    Dim ScreenUpdating As Boolean
    Dim wnd As Excel.Window
    '------------------------------------------------------------------------------------------------------
    
    With Application
        ScreenUpdating = .ScreenUpdating: .ScreenUpdating = False
    End With
    
    For Each wnd In Application.Windows
        If wnd.caption = wkb.name Then
            wnd.visible = show
            Exit For
        End If
    Next wnd
    
    Application.ScreenUpdating = ScreenUpdating
    
End Sub




Public Sub freeze(ByRef state As ExcelState)
    Dim app As Excel.Application:                       Set app = Excel.Application
    '------------------------------------------------------------------------------------------------------
    
    With state
        If .ScreenUpdating Then app.ScreenUpdating = False
        If .EnableEvents Then app.EnableEvents = False
        If .Interactive Then
            app.Interactive = False
            app.EnableCancelKey = xlDisabled
            app.cursor = xlWait
        End If
        If Not .Sheet Is Nothing Then
            If .SheetProtection Then
                Call .Sheet.Unprotect(.protectionPassword)
            End If
        End If
    End With
    
End Sub


Public Sub unfreeze(ByRef state As ExcelState)
    Dim app As Excel.Application:                       Set app = Excel.Application
    '------------------------------------------------------------------------------------------------------
    
    With state
        If Not .Sheet Is Nothing Then
            If .SheetProtection Then
                Call .Sheet.Protect(.protectionPassword, _
                                    .ProtectDrawingObjects, _
                                    .ProtectContents, _
                                    .ProtectScenarios, _
                                    True, _
                                    .ProtectionLevels.AllowFormattingCells, _
                                    .ProtectionLevels.AllowFormattingColumns, _
                                    .ProtectionLevels.AllowFormattingRows, _
                                    .ProtectionLevels.AllowInsertingColumns, _
                                    .ProtectionLevels.AllowInsertingRows, _
                                    .ProtectionLevels.AllowInsertingHyperlinks, _
                                    .ProtectionLevels.AllowDeletingColumns, _
                                    .ProtectionLevels.AllowDeletingRows, _
                                    .ProtectionLevels.AllowSorting, _
                                    .ProtectionLevels.AllowFiltering, _
                                    .ProtectionLevels.AllowUsingPivotTables)
            End If
        End If
        
        If .EnableEvents Then app.EnableEvents = True
        If .Interactive Then
            app.Interactive = True
            app.EnableCancelKey = xlInterrupt
            app.cursor = xlDefault
        End If
        
        If .ScreenUpdating Then app.ScreenUpdating = True
        
    End With
    
End Sub


Public Function getDefaultExcelState(Optional Sheet As Excel.Worksheet, Optional protectionPassword As String) As ExcelState
    Dim app As Excel.Application:               Set app = Excel.Application
    '------------------------------------------------------------------------------------------------------
    
    With getDefaultExcelState
        .ScreenUpdating = app.ScreenUpdating
        .EnableEvents = app.EnableEvents
        .Interactive = app.Interactive
        If Not VBA.IsMissing(Sheet) Then
            If f.sheets.IsValid(Sheet) Then
                Set .Sheet = Sheet
                .SheetProtection = Sheet.ProtectContents
                Set .ProtectionLevels = Sheet.Protection
                .ProtectDrawingObjects = Sheet.ProtectDrawingObjects
                .ProtectContents = Sheet.ProtectContents
                .ProtectScenarios = Sheet.ProtectScenarios
            End If
        End If
        
        If Not VBA.IsMissing(protectionPassword) Then
            .protectionPassword = protectionPassword
        End If
        
    End With
    
End Function


Public Sub quitApp(Optional ByVal xls As Excel.Application)
    If xls Is Nothing Then Set xls = Excel.Application
    With xls
        .EnableEvents = False
        .displayAlerts = False
        Call .Quit
    End With
End Sub



Public Function isValidAppReference(value As Variant) As Boolean
    Dim checkValue As Long:             checkValue = -1
    '------------------------------------------------------------------------------------------------------
    
    If Not VBA.IsObject(value) Then
    ElseIf value Is Nothing Then
    ElseIf Not TypeOf value Is Excel.Application Then
    Else
        On Error Resume Next
        checkValue = VBA.Len(value.name)
        On Error GoTo ErrHandler
        
        isValidAppReference = (checkValue > 0)
        
    End If
    

ErrHandler:

End Function





Public Function getBackgroundInstance(Optional visible As Boolean = False, _
                                      Optional macroEnabled As Boolean = True) As Excel.Application
    If Not f.Excel.isValidAppReference(pBackgroundExcel) Then
        Set pBackgroundExcel = New Excel.Application
        With pBackgroundExcel
            .visible = visible
            .AutomationSecurity = VBA.IIf(macroEnabled, msoAutomationSecurityByUI, msoAutomationSecurityForceDisable)
            .displayAlerts = False
        End With
    End If
    Set getBackgroundInstance = pBackgroundExcel
End Function


Public Function closeBackgroundInstance() As Excel.Application
    If f.Excel.isValidAppReference(pBackgroundExcel) Then
        Call pBackgroundExcel.Quit
        Set pBackgroundExcel = Nothing
    End If
End Function


Public Function getMacroDisabledInstance() As Excel.Application
    If Not f.Excel.isValidAppReference(pBackgroundExcel) Then
        Set pBackgroundExcel = New Excel.Application
        With pBackgroundExcel
            .visible = False
            .displayAlerts = False
        End With
    End If
    Set getMacroDisabledInstance = pBackgroundExcel
End Function


Public Sub adjustExcelSettings(Optional ByVal xls As Excel.Application)
    If xls Is Nothing Then Set xls = Excel.Application
    With xls.ErrorCheckingOptions
        .UnlockedFormulaCells = False
        .OmittedCells = False
        .InconsistentTableFormula = False
        .InconsistentFormula = False
        .EmptyCellReferences = False
        .NumberAsText = False
    End With
    
    '[Application settings]
    With xls
        .displayAlerts = False
        .WindowState = xlMaximized
    End With
    
    Excel.ActiveWindow.WindowState = xlMaximized
    
End Sub



Public Function createNewInstance(Optional visible As Boolean = True) As Excel.Application
    Set createNewInstance = New Excel.Application
    With createNewInstance
        .visible = visible
        .displayAlerts = VBA.IIf(visible, .displayAlerts, False)
    End With
End Function
