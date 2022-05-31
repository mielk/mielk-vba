VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FFiles"
'----------------------------------------------------------------------------------------------------------
#If VBA7 And Win64 Then
    Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As LongPtr)
    Private Declare PtrSafe Function lStrCat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    'Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As LongPtr
    Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As LongPtr
    Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidList As LongPtr, ByVal lpBuffer As String) As Long
#ElseIf VBA7 Then
    Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
    Private Declare PtrSafe Function lStrCat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
    Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
#Else
    Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
    Private Declare Function lStrCat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
    Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
#End If
'----------------------------------------------------------------------------------------------------------
#If VBA7 Then
    Private Type BROWSEINFO
        hwndOwner As LongPtr
        pidlRoot As Long
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfnCallback As LongPtr
        lParam As LongPtr
        iImage As Long
    End Type
#Else
    Private Type BROWSEINFO
        hwndOwner As Long
        pidlRoot As Long
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfnCallback As Long
        lParam As Long
        iImage As Long
    End Type
#End If
'----------------------------------------------------------------------------------------------------------
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
'----------------------------------------------------------------------------------------------------------



Public Function fileExists(filePath As String) As Boolean
    fileExists = fso.fileExists(filePath)
End Function

Public Function folderExists(path As String) As Boolean
    folderExists = fso.folderExists(path)
End Function

Public Function getFileName(filePath As String, Optional includeExtension As Boolean = True) As String
    Const REGEX_WITH_EXTENSION As String = "([^\\]+)$"
    Const REGEX_NO_EXTENSION As String = "([^\\]+)\..*$"
    '------------------------------------------------------------------------------------------------------
    If includeExtension Then
        getFileName = F.regex.getFirstGroupMatch(filePath, REGEX_WITH_EXTENSION)
    Else
        getFileName = F.regex.getFirstGroupMatch(filePath, REGEX_NO_EXTENSION)
    End If
End Function

Public Function createFolder(folderPath As String) As Object
    Dim strDriveName As String
    Dim strParentFolder As String
    '------------------------------------------------------------------------------------------------------

    With fso
        If .folderExists(folderPath) Then
            Set createFolder = .getFolder(folderPath)
        Else
            strDriveName = .GetDriveName(folderPath)
            If .DriveExists(strDriveName) Then
                                                                                                      
                strParentFolder = .GetParentFolderName(folderPath)
                If Not .folderExists(strParentFolder) Then
                    Call createFolder(strParentFolder)
                End If
                
                On Error Resume Next
                Call .createFolder(folderPath)
                Set createFolder = .getFolder(folderPath)
                On Error GoTo 0
                                                                                                      
            End If
        End If
    End With

End Function

Public Sub deleteFile(filePath As String)
    With fso
        If .fileExists(filePath) Then
            Call .deleteFile(filePath)
        End If
    End With
End Sub

Public Function getFileFolderPath(filePath As String) As String
    Dim slashPosition As Long
    '------------------------------------------------------------------------------------------------------
    slashPosition = VBA.InStr(1, VBA.StrReverse(filePath), "\")
    If slashPosition Then
        getFileFolderPath = VBA.left$(filePath, VBA.Len(filePath) - slashPosition + 1)
    End If
End Function

Public Function hasWritePermission(folderPath As String) As Boolean
    Dim textFileUrl As String
    '------------------------------------------------------------------------------------------------------
    
    textFileUrl = F.files.uniqueFilePath(folderPath & VBA.IIf(VBA.right$(folderPath, 1) = "\", vbNullString, "\") & _
                        F.System.getWindowsUserName & "_test.text")
    With fso
        On Error Resume Next
        Call .CreateTextFile(textFileUrl, True)
        On Error GoTo ErrHandler
        
        If .fileExists(textFileUrl) Then
            hasWritePermission = True
            Call .deleteFile(textFileUrl)
        End If
        
    End With
    
ErrHandler:
    
End Function


Public Function isFileLocked(filePath As String) As Boolean
    Dim fileIndex As Integer
    '------------------------------------------------------------------------------------------------------
    
    fileIndex = VBA.FreeFile
    
    Call VBA.err.clear
    On Error Resume Next
    
    Open filePath For Binary Access Read Write Lock Read Write As #fileIndex
    Close #fileIndex
    
    If VBA.err.number <> 0 Then
        isFileLocked = True
        Call VBA.err.clear
    Else
        isFileLocked = False
    End If
    
End Function



Public Function uniqueFilePath(filePath As String) As String
    Dim strFileExtension As String
    Dim strFileName As String
    Dim strParentFolder As String
    Dim strTempFilePath As String
    Dim intCounter As Integer
    '------------------------------------------------------------------------------------------------------

    With fso
        If .fileExists(filePath) Then
                                                                                                        
            strParentFolder = .GetParentFolderName(filePath)
            If Not VBA.right$(strParentFolder, 1) = "\" Then strParentFolder = strParentFolder & "\"
            strFileName = .GetBaseName(filePath)
            strFileExtension = "." & .GetExtensionName(filePath)
                                                                                                        
            Do
                intCounter = intCounter + 1
                strTempFilePath = strParentFolder & strFileName & _
                                                " (" & intCounter & ")" & strFileExtension
            Loop While .fileExists(strTempFilePath)
                                                                                                        
            uniqueFilePath = strTempFilePath
                                                                                                        
        Else
            uniqueFilePath = filePath
        End If
    End With

End Function




'----------------------------------------------------------------------------------------------------------
' Description:          Function converts the given path to UNC path.
'
' Parameters:
'   * path              Original path.
    '----------------------------------------------------------------------------------------------------------
Public Function getUncPath(path As String) As String
    Dim driveName As String
    Dim uncName As String
    Dim uncDrive As String
    '------------------------------------------------------------------------------------------------------
    
    getUncPath = path
    
    If VBA.mid$(path, 2, 1) = ":" Then
        driveName = fso.GetDriveName(path)
        
        On Error Resume Next
        uncDrive = fso.Drives(driveName & "\").ShareName
        On Error GoTo ErrHandler
        
        If VBA.Len(uncDrive) > 0 Then
            getUncPath = VBA.Replace(path, driveName, uncDrive)
        End If
        
    End If
    
ErrHandler:
    
End Function



Public Function concatenatePath(ParamArray pathParts() As Variant) As String
    Const NETWORK_PATH_PREFIX As String = "\\"
    '------------------------------------------------------------------------------------------------------
    Dim varPathPart As Variant
    Dim strPathPart As String
    Dim isNetworkPath As String
    '------------------------------------------------------------------------------------------------------
    
    For Each varPathPart In pathParts
        If Not F.Variables.isString(varPathPart) Then
        ElseIf VBA.Len(varPathPart) = 0 Then
        Else
            strPathPart = varPathPart
            If VBA.Len(concatenatePath) = 0 Then
                isNetworkPath = F.strings.startsWith(strPathPart, NETWORK_PATH_PREFIX)
            End If
            
            Do While VBA.right$(strPathPart, 1) = "\"
                strPathPart = VBA.left$(strPathPart, VBA.Len(strPathPart) - 1)
            Loop
            
            Do While VBA.left$(strPathPart, 1) = "\"
                strPathPart = VBA.mid$(strPathPart, 2)
            Loop
            
            Do While VBA.InStr(1, strPathPart, "\\")
                strPathPart = VBA.Replace(strPathPart, "\\", "\")
            Loop
            
            If VBA.Len(concatenatePath) Then
                concatenatePath = concatenatePath & "\" & strPathPart
            Else
                concatenatePath = strPathPart
            End If
            
        End If
    Next varPathPart
    
    If isNetworkPath Then
        concatenatePath = NETWORK_PATH_PREFIX & concatenatePath
    End If

End Function


Public Sub openFolderInWindowsExplorer(folderPath As String)
    Const COMMAND_PATTERN As String = "C:\WINDOWS\explorer.exe ""{0}"""
    '------------------------------------------------------------------------------------------------------
    If fso.folderExists(folderPath) Then
        Call VBA.Shell(F.strings.format(COMMAND_PATTERN, folderPath), vbNormalFocus)
    End If
End Sub


Public Function isValidPath(value As String) As Boolean
    Const REGEX_PATTERN As String = "^([A-Z]:|\\)(\\[a-zA-Z0-9\.\-_ ]{1,})*(?:\\|)$"
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(value) Then
        isValidPath = F.regex.checkIfMatch(value, REGEX_PATTERN)
    End If
    
End Function

Public Function getNameFromUncPath(path As String) As String
    Const REGEX_PATTERN As String = "([^\\]+)$"
    '------------------------------------------------------------------------------------------------------
    getNameFromUncPath = F.regex.getFirstGroupMatch(path, REGEX_PATTERN)
End Function

Public Function getFolderFiles(folderPath As String, _
                                    Optional includeSubfolder As Boolean = False, _
                                    Optional filePathPattern As String) As VBA.Collection
    Dim folder As Scripting.folder
    Dim subfolder As Scripting.folder
    Dim file As Scripting.file
    '------------------------------------------------------------------------------------------------------
    
    Set getFolderFiles = New VBA.Collection
    
    'Initial check
    With fso
        If Not .folderExists(folderPath) Then Call Exceptions.InvalidDirectory.raise(folderPath)
        Set folder = .getFolder(folderPath)
    End With
    
    For Each file In folder.files
        If VBA.Len(filePathPattern) Then
            If F.regex.checkIfMatch(file.path, filePathPattern) Then Call getFolderFiles.Add(file)
        Else
            Call getFolderFiles.Add(file)
        End If
    Next file
    
    If includeSubfolder Then
        For Each subfolder In folder.SubFolders
            Call F.Collections.appendCollection(getFolderFiles, getFolderFiles(subfolder.path, True, filePathPattern))
        Next subfolder
    End If
    
End Function

Public Function getFileExtension(filePath As String) As String
    Const REGEX_PATTERN As String = "\.([a-zA-Z0-9_]*)$"
    '------------------------------------------------------------------------------------------------------
    getFileExtension = F.regex.getFirstGroupMatch(filePath, REGEX_PATTERN)
End Function






'[SELECTING FILES & FOLDER]
Public Function selectFile(Optional defaultPath As String, Optional title As String = "Select file(s)", _
                           Optional fileType As Variant = "Excel files, *.xls; *.xlsm; *.xlsx; *.xlsb", _
                           Optional multiSelect As Boolean = True) As Variant()
    Dim drive As String * 1:                drive = VBA.left(defaultPath, 1)
    Dim selection As Variant
    Dim emptyArray() As Variant
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    Call VBA.ChDrive(drive)
    Call VBA.ChDir(defaultPath)
    On Error GoTo ErrHandler
    
    selection = Excel.Application.GetOpenFilename(filefilter:=fileType, title:=title, multiSelect:=multiSelect)
    
    If F.Variables.isBoolean(selection) Then
        selectFile = emptyArray
    ElseIf F.Variables.isString(selection) Then
        selectFile = VBA.Array(selection)
    Else
        selectFile = selection
    End If


ErrHandler:

End Function











Public Function getParentFolder(path As String) As String
    getParentFolder = getFileFolderPath(path)
End Function

Public Function getDependentFilePath(pattern As String, sourceFilePath As String) As String
    Dim countAsterisks As Integer
    Dim topLevelFolder As folder
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    countAsterisks = F.strings.countOccurrences(pattern, "*", False)
    Set topLevelFolder = fso.getFolder(sourceFilePath)
    For i = 1 To countAsterisks
        Set topLevelFolder = topLevelFolder.ParentFolder
    Next i
    getDependentFilePath = topLevelFolder.path & VBA.Replace(pattern, "*", vbNullString)
    
End Function


'[PATHS]
Public Function isProperPath(path As String) As Boolean
    Const REGEX_PATTERN As String = "^(\\\\|[\\|a-zA-Z]:\\)((?!\\\\)[^?*|"":<>])*[^ \.]$"
    Const SUBFOLDER_NAME_REGEX_PATTERN As String = "^(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])$"
    '------------------------------------------------------------------------------------------------------
    Static objRegex As Object
    Static objSubfolderNameRegex As Object
    '------------------------------------------------------------------------------------------------------
    Dim tempPath As String:                         tempPath = VBA.Replace(path, "/", "\")
    Dim i As Integer
    Dim pathParts() As String
    Dim pathPartName As String
    '------------------------------------------------------------------------------------------------------
    
    If objRegex Is Nothing Then Set objRegex = F.regex.Create(REGEX_PATTERN)
    If objSubfolderNameRegex Is Nothing Then Set objSubfolderNameRegex = F.regex.Create(SUBFOLDER_NAME_REGEX_PATTERN)
    
    isProperPath = objRegex.test(tempPath)
    
    If isProperPath Then
        For i = 0 To 30
            If VBA.InStr(1, path, VBA.Chr(i)) Then
                isProperPath = False
                Exit For
            End If
        Next i
    End If
    
    If isProperPath Then
        tempPath = VBA.mid$(tempPath, VBA.IIf(F.strings.startsWith(tempPath, "\\"), 3, 4))
        If F.strings.endsWith(tempPath, "\") Then tempPath = VBA.left$(tempPath, VBA.Len(tempPath) - 1)
        
        pathParts = VBA.Split(tempPath, "\")
        If F.arrays.isDefined(pathParts) Then
            For i = LBound(pathParts) To UBound(pathParts)
                pathPartName = pathParts(i)
                If VBA.Len(pathPartName) = 0 Then
                    isProperPath = False
                    Exit For
                ElseIf objSubfolderNameRegex.test(pathPartName) Then
                    isProperPath = False
                    Exit For
                End If
            Next i
        End If
    End If
    
End Function







Public Function getShortenedFilePath(filePath As String) As String
    Dim fileName As String
    Dim folderName As String
    '------------------------------------------------------------------------------------------------------
    fileName = getFileName(filePath)
    folderName = getFileName(VBA.Replace(filePath, "\" & fileName, vbNullString))
    getShortenedFilePath = "...\" & folderName & "\" & fileName
End Function

Public Function isFolderWriteable(folderPath As String) As Boolean
    Const TEMP_FILE_NAME As String = "TestFile.txt"
    '------------------------------------------------------------------------------------------------------
    Static objFSO As Object
    Dim strTempFile As String
    '------------------------------------------------------------------------------------------------------

    If objFSO Is Nothing Then
        Set objFSO = VBA.CreateObject("Scripting.FileSystemObject")
    End If


    If objFSO.folderExists(folderPath) Then
                                                                                                        
        strTempFile = folderPath
        If Not VBA.right$(folderPath, 1) = "\" Then strTempFile = strTempFile & "\"
        strTempFile = uniqueFilePath(strTempFile & TEMP_FILE_NAME)
                                                                                                        
        On Error Resume Next
        Call objFSO.CreateTextFile(strTempFile)
        On Error GoTo 0
        
        If objFSO.fileExists(strTempFile) Then
            isFolderWriteable = True
            Call objFSO.deleteFile(strTempFile)
        End If
    
    Else
        GoTo FolderNotExistException
    End If

'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
FolderNotExistException:
    GoTo ExitPoint

End Function


Public Function isSubfoldersAddingAllowed(folderPath As String) As Boolean
    Const SUBFOLDER_NAME As String = "TestFolder"
    '------------------------------------------------------------------------------------------------------
    Static objFSO As Object
    Dim strTempFolder As String
    '------------------------------------------------------------------------------------------------------

    If objFSO Is Nothing Then
        Set objFSO = VBA.CreateObject("Scripting.FileSystemObject")
    End If

    If objFSO.folderExists(folderPath) Then
                                                                                                        
        strTempFolder = folderPath
        If Not VBA.right$(folderPath, 1) = "\" Then strTempFolder = strTempFolder & "\"
        strTempFolder = uniqueFolderName(strTempFolder & SUBFOLDER_NAME)
                                                                                                        
        On Error Resume Next
        Call objFSO.createFolder(strTempFolder)
        On Error GoTo 0
                                                                                                        
        If objFSO.folderExists(strTempFolder) Then
            isSubfoldersAddingAllowed = True
            Call objFSO.DeleteFolder(strTempFolder)
        End If
                                                                                                        
    Else
        GoTo FolderNotExistException
    End If

'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
FolderNotExistException:
    GoTo ExitPoint
                                        
End Function


Public Function uniqueFolderName(folderName As String) As String
    Static objFSO As Object
    Dim strFolderName As String
    Dim strParentFolder As String
    Dim strTempName As String
    Dim intCounter As Integer
    '------------------------------------------------------------------------------------------------------

    If objFSO Is Nothing Then
        Set objFSO = VBA.CreateObject("Scripting.FileSystemObject")
    End If

    With objFSO
        If .folderExists(folderName) Then
                                                                                                        
            strParentFolder = .GetParentFolderName(folderName)
            If Not VBA.right$(strParentFolder, 1) = "\" Then strParentFolder = strParentFolder & "\"
            strFolderName = .GetBaseName(folderName)
                                                                                                        
            Do
                intCounter = intCounter + 1
                strTempName = strParentFolder & strFolderName & " (" & intCounter & ")"
            Loop While .folderExists(strTempName)
            uniqueFolderName = strTempName
                                                                                                        
        Else
            uniqueFolderName = folderName
        End If
    End With

End Function



Public Function browseForFolder(hwndOwner As Long, sPrompt As String) As String
    Dim iNull As Integer
#If VBA7 Then
    Dim lpIDList As LongPtr
#Else
    Dim lpIDList As Long
#End If
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BROWSEINFO
    
    With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lStrCat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    lpIDList = SHBrowseForFolder(udtBI)
    
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        
        If iNull Then sPath = left$(sPath, iNull - 1)
    
    End If
    
    browseForFolder = sPath
    
End Function


Public Sub clearFolder(folderPath As String)
    Dim folder As Scripting.folder
    Dim subfolder As Scripting.folder
    Dim file As Scripting.file
    '------------------------------------------------------------------------------------------------------
    Set folder = fso.getFolder(folderPath)
    If Not folder Is Nothing Then
        For Each subfolder In folder.SubFolders
            Call subfolder.delete
        Next subfolder
        For Each file In folder.files
            Call file.delete
        Next file
    End If
End Sub

