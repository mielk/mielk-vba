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



Public Function fileExists(filepath As String) As Boolean
    fileExists = fso.fileExists(filepath)
End Function

Public Function folderExists(path As String) As Boolean
    folderExists = fso.folderExists(path)
End Function

Public Function getFileName(filepath As String, Optional includeExtension As Boolean = True) As String
    Const REGEX_WITH_EXTENSION As String = "([^\\]+)$"
    Const REGEX_NO_EXTENSION As String = "([^\\]+)\..*$"
    '------------------------------------------------------------------------------------------------------
    If includeExtension Then
        getFileName = f.regex.getFirstGroupMatch(filepath, REGEX_WITH_EXTENSION)
    Else
        getFileName = f.regex.getFirstGroupMatch(filepath, REGEX_NO_EXTENSION)
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

Public Sub deleteFile(filepath As String)
    With fso
        If .fileExists(filepath) Then
            Call .deleteFile(filepath)
        End If
    End With
End Sub

Public Function getFileFolderPath(filepath As String, Optional removeEndSlash As Boolean = False) As String
    Const REGEX_PATTERN As String = "(.*)(?:\\[^\\]+)(?:\\*)$"
    '------------------------------------------------------------------------------------------------------
    Dim slashPosition As Long
    '------------------------------------------------------------------------------------------------------
    If removeEndSlash Then
        getFileFolderPath = f.regex.getFirstGroupMatch(filepath, REGEX_PATTERN)
    Else
        slashPosition = VBA.InStr(1, VBA.StrReverse(filepath), "\")
        If slashPosition Then
            getFileFolderPath = VBA.left$(filepath, VBA.Len(filepath) - slashPosition + 1)
        End If
    End If
End Function

Public Function hasWritePermission(folderPath As String) As Boolean
    Dim textFileUrl As String
    '------------------------------------------------------------------------------------------------------
    
    textFileUrl = f.files.uniqueFilePath(folderPath & VBA.IIf(VBA.right$(folderPath, 1) = "\", vbNullString, "\") & _
                        f.System.getWindowsUserName & "_test_____fY7JqjsOg7.text")
    With fso
        On Error Resume Next
        Call .deleteFile(textFileUrl)
        Call .CreateTextFile(textFileUrl, True)
        On Error GoTo ErrHandler
        
        If .fileExists(textFileUrl) Then
            hasWritePermission = True
            Call .deleteFile(textFileUrl)
        End If
        
    End With
    
ErrHandler:
    
End Function


Public Function isFileLocked(filepath As String) As Boolean
    Dim fileIndex As Integer
    '------------------------------------------------------------------------------------------------------
    
    fileIndex = VBA.FreeFile
    
    Call VBA.err.clear
    On Error Resume Next
    
    Open filepath For Binary Access Read Write Lock Read Write As #fileIndex
    Close #fileIndex
    
    If VBA.err.number <> 0 Then
        isFileLocked = True
        Call VBA.err.clear
    Else
        isFileLocked = False
    End If
    
End Function



Public Function uniqueFilePath(filepath As String) As String
    Dim strFileExtension As String
    Dim strFileName As String
    Dim strParentFolder As String
    Dim strTempFilePath As String
    Dim intCounter As Integer
    '------------------------------------------------------------------------------------------------------

    With fso
        If .fileExists(filepath) Then
                                                                                                        
            strParentFolder = .GetParentFolderName(filepath)
            If Not VBA.right$(strParentFolder, 1) = "\" Then strParentFolder = strParentFolder & "\"
            strFileName = .GetBaseName(filepath)
            strFileExtension = "." & .GetExtensionName(filepath)
                                                                                                        
            Do
                intCounter = intCounter + 1
                strTempFilePath = strParentFolder & strFileName & _
                                                " (" & intCounter & ")" & strFileExtension
            Loop While .fileExists(strTempFilePath)
                                                                                                        
            uniqueFilePath = strTempFilePath
                                                                                                        
        Else
            uniqueFilePath = filepath
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
            getUncPath = VBA.replace(path, driveName, uncDrive)
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
        If Not f.Variables.isString(varPathPart) Then
        ElseIf VBA.Len(varPathPart) = 0 Then
        Else
            strPathPart = varPathPart
            If VBA.Len(concatenatePath) = 0 Then
                isNetworkPath = f.strings.startsWith(strPathPart, NETWORK_PATH_PREFIX)
            End If
            
            Do While VBA.right$(strPathPart, 1) = "\"
                strPathPart = VBA.left$(strPathPart, VBA.Len(strPathPart) - 1)
            Loop
            
            Do While VBA.left$(strPathPart, 1) = "\"
                strPathPart = VBA.mid$(strPathPart, 2)
            Loop
            
            Do While VBA.InStr(1, strPathPart, "\\")
                strPathPart = VBA.replace(strPathPart, "\\", "\")
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
        Call VBA.Shell(f.strings.format(COMMAND_PATTERN, folderPath), vbNormalFocus)
    End If
End Sub


Public Function isValidPath(value As String) As Boolean
    Const REGEX_PATTERN As String = "^([A-Z]:|\\)(\\[a-zA-Z0-9\.\-_ ]{1,})*(?:\\|)$"
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(value) Then
        isValidPath = f.regex.checkIfMatch(value, REGEX_PATTERN)
    End If
    
End Function

Public Function getNameFromUncPath(path As String) As String
    Const REGEX_PATTERN As String = "([^\\]+)$"
    '------------------------------------------------------------------------------------------------------
    getNameFromUncPath = f.regex.getFirstGroupMatch(path, REGEX_PATTERN)
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
            If f.regex.checkIfMatch(file.path, filePathPattern) Then Call getFolderFiles.add(file)
        Else
            Call getFolderFiles.add(file)
        End If
    Next file
    
    If includeSubfolder Then
        For Each subfolder In folder.SubFolders
            Call f.Collections.appendCollection(getFolderFiles, getFolderFiles(subfolder.path, True, filePathPattern))
        Next subfolder
    End If
    
End Function

Public Function getFolderNewestFile(folderPath As String, _
                                    Optional includeSubfolders As Boolean = False, _
                                    Optional filePathPattern As String) As Scripting.file
    Dim files As VBA.Collection
    Dim file As Scripting.file
    Dim modifiedDate As Date
    '------------------------------------------------------------------------------------------------------
    
    Set files = getFolderFiles(folderPath, includeSubfolders, filePathPattern)
    For Each file In files
        If file.DateLastModified > modifiedDate Then
            Set getFolderNewestFile = file
            modifiedDate = file.DateLastModified
        End If
    Next file
    
End Function

Public Function getFileExtension(filepath As String) As String
    Const REGEX_PATTERN As String = "\.([a-zA-Z0-9_]*)$"
    '------------------------------------------------------------------------------------------------------
    getFileExtension = f.regex.getFirstGroupMatch(filepath, REGEX_PATTERN)
End Function






'[SELECTING FILES & FOLDER]
Public Function selectFiles(Optional defaultPath As String, Optional title As String = "Select files", _
                            Optional FileType As Variant = "Excel files, *.xls; *.xlsm; *.xlsx; *.xlsb") As VBA.Collection
    Dim selection As Variant
    '------------------------------------------------------------------------------------------------------
    selection = selectFile___(defaultPath, title, FileType, True)
    If f.arrays.isNonEmptyArray(selection) Then
        Set selectFiles = f.Collections.fromArray(selection)
    End If
End Function

Public Function selectFile(Optional defaultPath As String, Optional title As String = "Select file", _
                           Optional FileType As Variant = "Excel files, *.xls; *.xlsm; *.xlsx; *.xlsb") As String
    Dim selection As Variant
    '------------------------------------------------------------------------------------------------------
    selection = selectFile___(defaultPath, title, FileType, False)
    If f.arrays.isNonEmptyArray(selection) Then
        selectFile = f.arrays.getFirstItem(selection)
    End If
End Function

Private Function selectFile___(Optional defaultPath As String, Optional title As String = "Select file(s)", _
                           Optional FileType As Variant = "Excel files, *.xls; *.xlsm; *.xlsx; *.xlsb", _
                           Optional multiSelect As Boolean = True) As Variant()
    Dim drive As String * 1:                drive = VBA.left(defaultPath, 1)
    Dim selection As Variant
    Dim emptyArray() As Variant
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    Call VBA.ChDrive(drive)
    Call VBA.ChDir(defaultPath)
    On Error GoTo ErrHandler
    
    selection = Excel.Application.GetOpenFilename(filefilter:=FileType, title:=title, multiSelect:=multiSelect)
    
    If f.Variables.isBoolean(selection) Then
        selectFile___ = emptyArray
    ElseIf f.Variables.isString(selection) Then
        selectFile___ = VBA.Array(selection)
    Else
        selectFile___ = selection
    End If


ErrHandler:

End Function











Public Function getParentFolder(path As String) As String
    getParentFolder = getFileFolderPath(path)
End Function

Public Function getDependentFilePath(pattern As String, SourceFilePath As String) As String
    Dim countAsterisks As Integer
    Dim topLevelFolder As folder
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    countAsterisks = f.strings.countOccurrences(pattern, "*", False)
    Set topLevelFolder = fso.getFolder(SourceFilePath)
    For i = 1 To countAsterisks
        Set topLevelFolder = topLevelFolder.ParentFolder
    Next i
    getDependentFilePath = topLevelFolder.path & VBA.replace(pattern, "*", vbNullString)
    
End Function


'[PATHS]
Public Function isProperPath(path As String) As Boolean
    Const REGEX_PATTERN As String = "^(\\\\|[\\|a-zA-Z]:\\)((?!\\\\)[^?*|"":<>])*[^ \.]$"
    Const SUBFOLDER_NAME_REGEX_PATTERN As String = "^(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])$"
    '------------------------------------------------------------------------------------------------------
    Static objRegex As Object
    Static objSubfolderNameRegex As Object
    '------------------------------------------------------------------------------------------------------
    Dim tempPath As String:                         tempPath = VBA.replace(path, "/", "\")
    Dim i As Integer
    Dim pathParts() As String
    Dim pathPartName As String
    '------------------------------------------------------------------------------------------------------
    
    If objRegex Is Nothing Then Set objRegex = f.regex.Create(REGEX_PATTERN)
    If objSubfolderNameRegex Is Nothing Then Set objSubfolderNameRegex = f.regex.Create(SUBFOLDER_NAME_REGEX_PATTERN)
    
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
        tempPath = VBA.mid$(tempPath, VBA.IIf(f.strings.startsWith(tempPath, "\\"), 3, 4))
        If f.strings.endsWith(tempPath, "\") Then tempPath = VBA.left$(tempPath, VBA.Len(tempPath) - 1)
        
        pathParts = VBA.Split(tempPath, "\")
        If f.arrays.isDefined(pathParts) Then
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







Public Function getShortenedFilePath(filepath As String) As String
    Dim FileName As String
    Dim folderName As String
    '------------------------------------------------------------------------------------------------------
    FileName = getFileName(filepath)
    folderName = getFileName(VBA.replace(filepath, "\" & FileName, vbNullString))
    getShortenedFilePath = "...\" & folderName & "\" & FileName
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



Public Sub clearFolder(folderPath As String)
    Dim folder As Scripting.folder
    Dim subfolder As Scripting.folder
    Dim file As Scripting.file
    '------------------------------------------------------------------------------------------------------
    If fso.folderExists(folderPath) Then
        Set folder = fso.getFolder(folderPath)
        
        '[Delete folders]
        For Each file In folder.files
            Call file.delete(True)
        Next file
        
        '[Delete subfolders]
        For Each subfolder In folder.SubFolders
            Call subfolder.delete(True)
        Next subfolder
        
    End If
End Sub




Public Function getFileSize(SourceFilePath As String) As Long
    Dim f As file
    '------------------------------------------------------------------------------------------------------
    If fileExists(SourceFilePath) Then
        Set f = fso.GetFile(SourceFilePath)
        getFileSize = f.size
    End If
End Function

Public Function getFileTimestamp(SourceFilePath As String) As Date
    Dim f As file
    '------------------------------------------------------------------------------------------------------
    If fileExists(SourceFilePath) Then
        Set f = fso.GetFile(SourceFilePath)
        getFileTimestamp = f.DateLastModified
    End If
End Function

Public Function getSubfolder(baseFolder As Variant, subfolderName As String) As Scripting.folder
    Dim folder As Scripting.folder
    '------------------------------------------------------------------------------------------------------
    
    If f.Variables.isUndefined(baseFolder) Then
    ElseIf VBA.IsObject(baseFolder) Then
        If TypeOf baseFolder Is Scripting.folder Then
            Set folder = baseFolder
        End If
    ElseIf f.Variables.isString(baseFolder) Then
        On Error Resume Next
        Set folder = fso.getFolder(baseFolder)
        On Error GoTo 0
    End If
    
    If Not folder Is Nothing Then
        On Error Resume Next
        Set getSubfolder = folder.SubFolders(subfolderName)
    End If
    
End Function



Public Function browseForFolder(Optional initialPath As String, _
                                Optional windowTitle As String = "Select folder") As String
    Dim dialog As Office.FileDialog
    '------------------------------------------------------------------------------------------------------
    
    Set dialog = Excel.Application.FileDialog(msoFileDialogFolderPicker)
    With dialog
        .InitialFileName = initialPath
        .title = windowTitle
        Call .show
        If .SelectedItems.count Then
            browseForFolder = .SelectedItems.Item(1)
        End If
    End With
End Function

