Attribute VB_Name = "modTest"
Option Explicit

Private Const CLASS_NAME As String = "modTest"
'----------------------------------------------------------------------------------------------------------


Const LIB_FOLDER = "D:\Dropbox\tm\mielk\mielk-vba\code\"
Const TOOL_FOLDER = "D:\Dropbox\tm\mielk\mielk-vba\Toolbox\"

Const APP_NAME = "Toolbox"
Const COPY_TO_LOCAL_DRIVE = False
Const LOCAL_DRIVE_PATH = "D:\vba-tests\apps"



Public Sub test()
    Dim fso As Scripting.FileSystemObject
    Dim path As String
    Dim destinationFolder As Scripting.folder
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = LOCAL_DRIVE_PATH & "\" & APP_NAME & "\"
    path = replace(path, "\\", "\")
    Set destinationFolder = createFolder(path, True)
    
    If Not destinationFolder Is Nothing Then
    
        err.clear
        On Error Resume Next
        Call fso.copyFolder(clearPath(TOOL_FOLDER), clearPath(destinationFolder), True)
        
        If err.number Then
            Call MsgBox(getErrorMessage_CopyingFolder(destinationFolder.path, err.number, err.description), vbCritical, APP_NAME)
        End If
        
    End If
    
    Debug.Print path
        
End Sub


Private Function createFolder(path As String, removeIfExists)
    Dim fso As Scripting.FileSystemObject
    Dim parentFolderPath As String
    Dim parentFolder As Scripting.folder
    '------------------------------------------------------------------------------------------------------
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    
    If fso.folderExists(path) And (Not removeIfExists) Then
        Set createFolder = fso.getFolder(path)
    Else
    
        Set createFolder = Nothing
    
        '[Removing previous version of this folder] ---------------------------------------------|
        If fso.folderExists(path) Then                                                          '|
            On Error Resume Next                                                                '|
            Call fso.deleteFolder(clearPath(path))                                         '|
                                                                                                '|
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
                                        parentFolderPath & "]." & vbCrLf & vbCrLf & _
                                        "The most probable reasons are: " & vbCrLf & _
                                        "   * No write permission for folder [" & parentFolderPath & "]," & vbCrLf & _
                                        "   * Invalid characters in project name or destination path, " & vbCrLf & _
                                        "   * No space on the disk"
End Function

Function getErrorMessage_DeletingFolder(path)
    getErrorMessage_DeletingFolder = "Error while trying to delete the previous version of application located in [" & _
                                        path & "]." & vbCrLf & vbCrLf & _
                                        "The most probable reason is that the application is opened in another Excel instance." & _
                                        vbCrLf & vbCrLf & _
                                        "Close all open instances of " & APP_NAME & " application, including hidden Excel instances, and try again."
End Function

Function getErrorMessage_InvalidPath(path)
    getErrorMessage_InvalidPath = "It seems that the folder path set for the application in configuration file has some errors, because drive [" & _
                                        path & "] does not exist."
End Function

Function getErrorMessage_CopyingFolder(path, errNumber, errDescription)
    getErrorMessage_CopyingFolder = "Error while copying application files to [" & path & "]." & vbCrLf & _
                                        "   Error: " & errDescription & " (" & errNumber & ")"
End Function



Public Sub dictTest()
    Dim dictA As Scripting.Dictionary
    Dim dictB As Scripting.Dictionary
    Dim col As VBA.Collection
    
    Set dictA = f.dictionaries.Create(False)
    With dictA
        Call .Add("a", 1)
        Call .Add("b", 2)
        Call .Add("c", 3)
    End With
    
    Set dictB = f.dictionaries.Create(False)
    With dictB
        Call .Add("a", 1)
        Call .Add("b", 1)
    End With
    
    Set col = f.Collections.Create(1, 3)
    
    
    Dim result As Scripting.Dictionary
    'Set result = f.dictionaries.removeDuplicates(dictA, dictB)
    Set result = f.dictionaries.removeDuplicates(dictA, dictB, True)
    'Set result = f.dictionaries.removeDuplicates(dictA, col, True)
    
    Stop
    
End Sub
