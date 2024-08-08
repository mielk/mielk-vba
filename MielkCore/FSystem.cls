VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FSystem"
'----------------------------------------------------------------------------------------------------------

'[Windows API] --------------------------------------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" _
                        (ByVal flags As Long, ByVal name As String, ByVal Level As Long, _
                        pPrinterEnum As Long, ByVal cdBuf As Long, pcbNeeded As Long, _
                        pcReturned As Long) As LongPtr
    Private Declare PtrSafe Function PtrToStr Lib "kernel32" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As LongPtr) As Long
    Private Declare PtrSafe Function StrLen Lib "kernel32" Alias "lstrlenA" (ByVal Ptr As Long) As Long
#Else
    Private Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" _
                        (ByVal flags As Long, ByVal name As String, ByVal Level As Long, _
                        pPrinterEnum As Long, ByVal cdBuf As Long, pcbNeeded As Long, _
                        pcReturned As Long) As Long
    Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As Long) As Long
    Private Declare Function StrLen Lib "kernel32" Alias "lstrlenA" (ByVal Ptr As Long) As Long
#End If
'----------------------------------------------------------------------------------------------------------
'[Printers]
Private Const PRINTER_ENUM_CONNECTIONS = &H4
Private Const PRINTER_ENUM_LOCAL = &H2
'----------------------------------------------------------------------------------------------------------



'[System settings]

Public Function getSystemBits() As Long
    #If Win64 Then
        getSystemBits = 64
    #Else
        getSystemBits = 32
    #End If
End Function


Public Function getVbaVersion() As Long
    #If VBA7 Then
        getVbaVersion = 7
    #Else
        getVbaVersion = 6
    #End If
End Function

Public Function getDateFormat() As String
    getDateFormat = VBA.CreateObject("WScript.Shell").RegRead("HKCU\Control Panel\International\sShortDate")
End Function

Public Function getDecimalSeparator() As String
    getDecimalSeparator = VBA.CreateObject("WScript.Shell").RegRead("HKCU\Control Panel\International\sDecimal")
End Function

Public Function getThousandSeparator() As String
    getThousandSeparator = VBA.CreateObject("WScript.Shell").RegRead("HKCU\Control Panel\International\sThousand")
End Function







'[ERROR HANDLING]
Public Function getErrorHandlingSetting() As ErrorHandlingStatusEnum
    Const MS_COMMON_REG_KEY As String = "HKEY_CURRENT_USER\Software\Microsoft\VBA\7.1\Common\"
    Const ALL_ERRORS_TAG As String = "BreakOnAllErrors"
    Const CLASS_ERRORS_TAG As String = "BreakOnServerErrors"
    '------------------------------------------------------------------------------------------------------
    Dim objShell As Object:                 Set objShell = VBA.CreateObject("WScript.Shell")
    '------------------------------------------------------------------------------------------------------

    If objShell.RegRead(MS_COMMON_REG_KEY & ALL_ERRORS_TAG) = 1 Then
        getErrorHandlingSetting = errorHandling_AllErrors
    ElseIf objShell.RegRead(MS_COMMON_REG_KEY & CLASS_ERRORS_TAG) = 1 Then
        getErrorHandlingSetting = errorHandling_ClassErrors
    Else
        getErrorHandlingSetting = errorHandling_UnhandledErrors
    End If

    Set objShell = Nothing

End Function

Public Sub setErrorHandling(errorHandlingState As ErrorHandlingStatusEnum)
    Const MS_COMMON_REG_KEY As String = "HKEY_CURRENT_USER\Software\Microsoft\VBA\7.1\Common"
    Const ALL_ERRORS_TAG As String = "BreakOnAllErrors"
    Const CLASS_ERRORS_TAG As String = "BreakOnServerErrors"
    '------------------------------------------------------------------------------------------------------

    Select Case errorHandlingState
        Case errorHandling_AllErrors:
            Call f.Registry.updateValue(HKEY_CURRENT_USER, MS_COMMON_REG_KEY, CLASS_ERRORS_TAG, 0)
            Call f.Registry.updateValue(HKEY_CURRENT_USER, MS_COMMON_REG_KEY, ALL_ERRORS_TAG, 1)
        Case errorHandling_ClassErrors:
            Call f.Registry.updateValue(HKEY_CURRENT_USER, MS_COMMON_REG_KEY, CLASS_ERRORS_TAG, 1)
            Call f.Registry.updateValue(HKEY_CURRENT_USER, MS_COMMON_REG_KEY, ALL_ERRORS_TAG, 0)
        Case errorHandling_UnhandledErrors:
            Call f.Registry.updateValue(HKEY_CURRENT_USER, MS_COMMON_REG_KEY, CLASS_ERRORS_TAG, 0)
            Call f.Registry.updateValue(HKEY_CURRENT_USER, MS_COMMON_REG_KEY, ALL_ERRORS_TAG, 0)
    End Select

End Sub








'#todo - zmieniæ na nowy sposób, korzystaj¹cy z WinApi
Public Function getWindowsUserName() As String
    getWindowsUserName = VBA.Environ("Username")
End Function

Public Function isDeveloper() As Boolean
    Const DEVELOPERS_COMPUTER_NAMES As String = ";DESKTOP-PE9TCEP;"
    Const DEVELOPERS_UIDS As String = ";______;"
    '------------------------------------------------------------------------------------------------------
    Dim compName As String
    Dim userName As String
    '------------------------------------------------------------------------------------------------------
    
    compName = f.System.GetComputerName
    userName = f.System.getUserUid
    
    If VBA.InStr(1, DEVELOPERS_COMPUTER_NAMES, ";" & compName & ";", vbTextCompare) > 0 Then
        isDeveloper = True
    ElseIf VBA.InStr(1, DEVELOPERS_UIDS, ";" & userName & ";", vbTextCompare) > 0 Then
        isDeveloper = True
    End If
    
End Function


'https://www.access-programmers.co.uk/forums/threads/ms-access-2003-environ-username-returning-empty-string.200115/
Public Function getUserUid() As String
    Dim cn As String
    Dim ls As Long
    Dim res As Long
    '------------------------------------------------------------------------------------------------------

    cn = String(1024, 0)
    ls = 1024
    res = GetUserName(cn, ls)
    If res <> 0 Then
        getUserUid = mid(cn, 1, InStr(cn, Chr(0)) - 1)
    Else
        getUserUid = ""
    End If
    
End Function


Public Function GetComputerName() As String
    Dim sBuffer As String
    Dim nLen As Long
    Dim nSize As Long
    '------------------------------------------------------------------------------------------------------
    
    nLen = 16
    sBuffer = VBA.String$(nLen, 0)
    nSize = apiGetComputerName(sBuffer, nLen)
    If nSize <> 0 Then
       GetComputerName = VBA.Left$(sBuffer, nLen)
    Else
       GetComputerName = vbNullString
    End If
    
End Function



'[FOLDERS]
Public Function getDownloadFolderPath() As String
    getDownloadFolderPath = VBA.Environ("USERPROFILE") & "\Downloads"
End Function

Public Function getDesktopFolderPath() As String
    getDesktopFolderPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
End Function











'[Printers]
'#todo - przenieœæ do modu³u printers
Public Function listPrinters() As Variant
    Dim bSuccess As Boolean
    Dim iBufferRequired As Long
    Dim iBufferSize As Long
    Dim iBuffer() As Long
    Dim iEntries As Long
    Dim iIndex As Long
    Dim strPrinterName As String
    Dim iDummy As Long
    Dim iDriverBuffer() As Long
    Dim StrPrinters() As String
    '------------------------------------------------------------------------------------------------------
    
    iBufferSize = 3072
    ReDim iBuffer((iBufferSize \ 4) - 1) As Long
    
    'EnumPrinters will return a value False if the buffer is not big enough
    bSuccess = EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
                            1, iBuffer(0), iBufferSize, iBufferRequired, iEntries)
    
    If Not bSuccess Then
        If iBufferRequired > iBufferSize Then
            iBufferSize = iBufferRequired
            ReDim iBuffer(iBufferSize \ 4) As Long
        End If
        'Try again with new buffer
        bSuccess = EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
                            1, iBuffer(0), iBufferSize, iBufferRequired, iEntries)
    End If
    
    If Not bSuccess Then
        MsgBox "Error enumerating printers."
        Exit Function
    Else
        ReDim StrPrinters(iEntries - 1)
        For iIndex = 0 To iEntries - 1
            strPrinterName = VBA.Space$(StrLen(iBuffer(iIndex * 4 + 2)))
            iDummy = PtrToStr(strPrinterName, iBuffer(iIndex * 4 + 2))
            StrPrinters(iIndex) = strPrinterName
        Next iIndex
    End If
    
    listPrinters = StrPrinters

End Function



'[Others]
'#todo - przenieœæ do oddzielnego modu³u FClipboard
Public Sub saveToClipboard(value As Variant)
    Dim dataObj As MSForms.DataObject
    '------------------------------------------------------------------------------------------------------
    Set dataObj = New MSForms.DataObject
    With dataObj
        Call .SetText(VBA.CStr(value))
        Call .PutInClipboard
    End With
End Sub

Public Function getClipboardContentAsString() As Variant
    Dim dataObj As MSForms.DataObject
    '------------------------------------------------------------------------------------------------------
    Set dataObj = New MSForms.DataObject
    With dataObj
        Call .GetFromClipboard
        On Error Resume Next
        getClipboardContentAsString = .getText
        On Error GoTo 0
    End With
End Function

Public Function getClipboardContentAsArray() As Variant
    Dim clipboardContent As String
    Dim arr As Variant
    Dim values As Variant
    Dim i As Long
    Dim j As Long
    Dim rowColumns As Long
    Dim maxColumns As Long
    Dim result As Variant
    '------------------------------------------------------------------------------------------------------
    
    clipboardContent = getClipboardContentAsString
    If f.strings.endsWith(clipboardContent, VBA.vbCrLf) Then clipboardContent = VBA.Left$(clipboardContent, VBA.Len(clipboardContent) - 2)
    arr = VBA.Split(clipboardContent, vbCrLf)
    
    If f.arrays.isDefined(arr) Then
        ReDim values(LBound(arr) To UBound(arr))
        For i = LBound(arr) To UBound(arr)
            values(i) = VBA.Split(arr(i), vbTab)
            rowColumns = UBound(values(i)) - LBound(values(i)) + 1
            If rowColumns > maxColumns Then maxColumns = rowColumns
        Next i
        
        If maxColumns >= 1 Then
            ReDim result(LBound(values) To UBound(values), 1 To maxColumns)
            For i = LBound(values) To UBound(values)
                For j = LBound(values(i)) To UBound(values(i))
                    result(i, j + 1 - LBound(values(i))) = values(i)(j)
                Next j
            Next i
        End If
        
        getClipboardContentAsArray = result
    End If
    
End Function



Public Function getWindowsUpdates() As Scripting.Dictionary
    Const REGEX_PATTERN As String = "(?:\()(KB\d{7})(?:\))"
    '------------------------------------------------------------------------------------------------------
    Dim session As Object
    Dim searcher As Object
    Dim updates As Variant
    Dim update As Object
    Dim updateName As String
    '------------------------------------------------------------------------------------------------------
    
    Set getWindowsUpdates = f.dictionaries.Create(False)
    Set session = VBA.CreateObject("Microsoft.Update.Session")
    Set searcher = session.createUpdateSearcher
    With searcher
        Set updates = .QueryHistory(0, .getTotalHistoryCount)
        For Each update In updates
            updateName = f.regex.getFirstGroupMatch(update.title, REGEX_PATTERN)
            Call f.dictionaries.addItem(getWindowsUpdates, updateName, vbNullString)
        Next update
    End With
    
End Function
