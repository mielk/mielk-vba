Option Explicit

'[Paths]
Const LIB_FOLDER = "D:\Dropbox\tm\mielk\mielk-vba\code\"
Const TOOL_FOLDER = "D:\Dropbox\tm\mielk\mielk-vba\Toolbox\"

'[Open mode]
Const READ_ONLY = False
Const COPY_TO_LOCAL_DRIVE = True
Const APP_NAME = "Toolbox"
Const LOCAL_DRIVE_PATH = "D:\vba-tests\apps\"

'[Registry]
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const REGISTRY_PATH = "Software\Microsoft\VBA\7.1\Common"
Const BREAK_ON_ALL_ERRORS = "BreakOnAllErrors"
Const BREAK_ON_SERVER_ERRORS = "BreakOnServerErrors"



Call runApp



Sub runApp()
    Call updateVbaErrorHandlingRegistry
	If copyFiles Then
		Call addProperLocationsToTrusted
		Call openFiles
	End If
End Sub

Sub updateVbaErrorHandlingRegistry()
    Dim strComputer
    Dim objRegistry
    
    strComputer = "."
    Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
    
    objRegistry.setDWORDValue HKEY_CURRENT_USER, REGISTRY_PATH, BREAK_ON_ALL_ERRORS, 0
    objRegistry.setDWORDValue HKEY_CURRENT_USER, REGISTRY_PATH, BREAK_ON_SERVER_ERRORS, 0
    
End Sub


Sub openFiles()
    Dim excelApp
	Dim libFolder
	Dim toolFolder
    
	If COPY_TO_LOCAL_DRIVE Then
		libFolder = getDestinationFolderPath
		toolFolder = libFolder
	Else
		libFolder = LIB_FOLDER
		toolFolder = TOOL_FOLDER
	End If
	
    Set excelApp = CreateObject("Excel.Application")
    With excelApp
        .Visible = True
        .WindowState = -4137
    End With
    
    Call excelApp.Workbooks.Open(libFolder & "mielk-core.xlam", , READ_ONLY)
    Call excelApp.Workbooks.Open(libFolder & "mielk-ui.xlam", , READ_ONLY)
    Call excelApp.Workbooks.Open(libFolder & "mielk-db.xlam", , READ_ONLY)
	Call excelApp.Workbooks.Open(libFolder & "mielk-app.xlam", , READ_ONLY)
    Call excelApp.Workbooks.Open(toolFolder & "toolbox-core.xlam", , READ_ONLY)
	Call excelApp.Workbooks.Open(toolFolder & "toolbox-controller.xlam", , READ_ONLY)
    Call excelApp.Workbooks.Open(toolFolder & "toolbox.xlsm", , READ_ONLY)
    
End Sub


Function getDestinationFolderPath()
	getDestinationFolderPath = Replace(LOCAL_DRIVE_PATH & "\" & APP_NAME & "\", "\\", "\")
End Function


Function copyFiles()
	Dim fso
	Dim path
	Dim destinationFolder
	
	If COPY_TO_LOCAL_DRIVE Then	
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		path = getDestinationFolderPath		
		Set destinationFolder = createFolder(path, True)
		
		If Not destinationFolder Is Nothing Then

			Err.clear
			On Error Resume Next
			Call fso.CopyFolder(clearPath(TOOL_FOLDER), clearPath(destinationFolder), True)
			If TOOL_FOLDER <> LIB_FOLDER Then
				Call fso.CopyFile(LIB_FOLDER & "mielk-app.xlam", destinationFolder & "\mielk-app.xlam")
				Call fso.CopyFile(LIB_FOLDER & "mielk-core.xlam", destinationFolder & "\mielk-core.xlam")
				Call fso.CopyFile(LIB_FOLDER & "mielk-db.xlam", destinationFolder & "\mielk-db.xlam")
				Call fso.CopyFile(LIB_FOLDER & "mielk-ui.xlam", destinationFolder & "\mielk-ui.xlam")
			End If
			
			copyFiles = (Err.number = 0)
			If Err.number Then
				Call MsgBox(getErrorMessage_CopyingFolder(destinationFolder.path, Err.number, Err.description), vbCritical, APP_NAME)
			End If
		Else
			copyFiles = False
		End if
	Else
		copyFiles = True
	End If
End Function




Function createFolder(path, removeIfExists)
    Dim fso
    Dim parentFolderPath
    Dim parentFolder
    '------------------------------------------------------------------------------------------------------
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    
    If fso.folderExists(path) And (Not removeIfExists) Then
        Set createFolder = fso.getFolder(path)
    Else
	
        Set createFolder = Nothing
		
        '[Removing previous version of this folder] ---------------------------------------------|
        If fso.folderExists(path) Then                                                          '|
            On Error Resume Next                                                                '|
            Call fso.DeleteFolder(clearPath(path))                                              '|
            If fso.folderExists(path) Then                                                      '|
                Call MsgBox(getErrorMessage_DeletingFolder(path), vbCritical, APP_NAME)         '|
                Exit Function                                                                   '|
            End If                                                                              '|
        End If                                                                                  '|
        '----------------------------------------------------------------------------------------|
    
        parentFolderPath = fso.GetParentFolderName(path)
        If Len(parentFolderPath) Then
            If Not fso.folderExists(parentFolderPath) Then
                Set parentFolder = createFolder(parentFolderPath, False)
            Else
                Set parentFolder = fso.getFolder(parentFolderPath)
            End If
            
            If Not parentFolder Is Nothing Then
                On Error Resume Next
                Set createFolder = fso.createFolder(path)
                
                If createFolder Is Nothing Then
                    Call MsgBox(getErrorMessage_CreatingFolder(parentFolderPath), vbCritical, APP_NAME)
                End If
            End If
        Else
            Call MsgBox(getErrorMessage_InvalidPath(path), vbCritical, APP_NAME)
        End If
    End If
    
End Function

Function clearPath(path)
    If right(path, 1) = "\" Then
        clearPath = left(path, Len(path) - 1)
    Else
        clearPath = path
    End If
End Function

Function getErrorMessage_CreatingFolder(parentFolderPath)
    getErrorMessage_CreatingFolder = "Error while trying to create application folder [" & APP_NAME & "] in location [" & _
                                        parentFolderPath & "] " & vbCrLf & vbCrLf & _
                                        "The most probable reasons are: " & vbCrLf & _
                                        "   * No write permission for folder [" & parentFolderPath & "]," & vbCrLf & _
                                        "   * Invalid characters in project name or destination path, " & vbCrLf & _
                                        "   * No space on the disk"
End Function

Function getErrorMessage_DeletingFolder(path)
    getErrorMessage_DeletingFolder = "Error while trying to delete the previous version of application located in [" & _
                                        path & "] " & vbCrLf & vbCrLf & _
                                        "The most probable reason is that the application is opened in another Excel instance" & _
                                        vbCrLf & vbCrLf & _
                                        "Close all open instances of " & APP_NAME & " application, including hidden Excel instances, and try again"
End Function

Function getErrorMessage_InvalidPath(path)
    getErrorMessage_InvalidPath = "It seems that the folder path set for the application in configuration file has some errors, because drive [" & _
                                        path & "] does not exist"
End Function

Function getErrorMessage_CopyingFolder(path, errNumber, errDescription)
    getErrorMessage_CopyingFolder = "Error while copying application files to [" & path & "]." & vbCrLf & _
                                        "   Error: " & errDescription & " (" & errNumber & ")"
End Function







Sub addProperLocationsToTrusted()
	If COPY_TO_LOCAL_DRIVE Then
		Call addTrustedLocation(LOCAL_DRIVE_PATH)
	Else
		Call addTrustedLocation(LIB_FOLDER)
		If TOOL_FOLDER <> LIB_FOLDER Then Call addTrustedLocation(TOOL_FOLDER)
	End If
End Sub

Public Sub addTrustedLocation(folderPath)
    Const REGISTRY_KEY = "Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations"
    Const ENTRY_DESCRIPTION = "Toolbox"
    '------------------------------------------------------------------------------------------------------
    Dim objRegistry
    Dim iLocCounter
    Dim arrChildKeys
    Dim sChildKey
    Dim sNewKey
    Dim sPath
    Dim sDescription
    Dim bAlreadyExists
    Dim bAllowNetworkLocations
    Dim bAllowSubFolders
	Dim value
    '------------------------------------------------------------------------------------------------------
    
    Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")
    bAllowSubFolders = True
    bAlreadyExists = False
    
    iLocCounter = 0
    Call objRegistry.enumKey(HKEY_CURRENT_USER, REGISTRY_KEY, arrChildKeys)
    For Each sChildKey In arrChildKeys
        Call objRegistry.getstringvalue(HKEY_CURRENT_USER, REGISTRY_KEY & "\" & sChildKey, "Description", sDescription)
        Call objRegistry.getstringvalue(HKEY_CURRENT_USER, REGISTRY_KEY & "\" & sChildKey, "Path", sPath)
        
        If sPath = folderPath Then bAlreadyExists = True
		value = Mid(sChildKey, 9)
		If IsNumeric(value) Then
			If CInt(value) > iLocCounter Then
				iLocCounter = CInt(Mid(sChildKey, 9))
			End If
		End If
    Next
    
    'Uncomment the following 4 lines if you wish to enable network locations as Trusted locations
    bAllowNetworkLocations = True
    If bAllowNetworkLocations Then
        objRegistry.setDWORDValue HKEY_CURRENT_USER, REGISTRY_KEY, "AllowNetworkLocations", 1
    End If
    
    If Not bAlreadyExists Then
        sNewKey = REGISTRY_KEY & "\Location" & CStr(iLocCounter + 1)
        objRegistry.createKey HKEY_CURRENT_USER, sNewKey
        objRegistry.setStringValue HKEY_CURRENT_USER, sNewKey, "Path", folderPath
        objRegistry.setStringValue HKEY_CURRENT_USER, sNewKey, "Description", ENTRY_DESCRIPTION
        
        If bAllowSubFolders Then
            objRegistry.setDWORDValue HKEY_CURRENT_USER, sNewKey, "AllowSubFolders", 1
        End If
    End If
    
End Sub