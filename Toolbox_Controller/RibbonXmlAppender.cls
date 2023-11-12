VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonXmlAppender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "RibbonXmlAppender"
'----------------------------------------------------------------------------------------------------------
Private Const CUSTOM_UI_FILE_NAME As String = "customUI14.xml"
'----------------------------------------------------------------------------------------------------------
'--2007 attributes and namespace
Private Const ATT_TARGET_2007 As String = "customUI/customUI.xml"
Private Const ATT_TYPE_2007 As String = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
Private Const NS_2007 As String = "http://schemas.microsoft.com/office/2006/01/customui"

'--2010 attributes and namespace
Private Const ATT_TARGET_2010 As String = "customUI/customUI14.xml"
Private Const ATT_TYPE_2010 As String = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
Private Const NS_2010 As String = "http://schemas.microsoft.com/office/2009/07/customui"
'----------------------------------------------------------------------------------------------------------
Private Const ERR_APPENDING_CUSTOM_RIBBON As String = "CreatingNewProject.Errors.AppendingCustomRibbon"
'----------------------------------------------------------------------------------------------------------


Public Sub append(filePath As String, xmlPath As String)
    Dim xml As String
    Dim zipFile As Scripting.file
    Dim unzippedFiles As Scripting.folder
    '------------------------------------------------------------------------------------------------------
    Dim rezippedFilePath As String
    Dim rezippedFile As Scripting.file
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    xml = f.TextFiles.readTextFile(xmlPath)
    
    Set zipFile = createZipFile(filePath)
    Set unzippedFiles = unzipFiles(zipFile.path)
    
    Call writeCustomUI_XML_ToFile(xml, _
                    f.files.concatenatePath(unzippedFiles.path, "customUI"), _
                    CUSTOM_UI_FILE_NAME)
    'Call updateRels(unzippedFiles.path, CUSTOM_UI_FILE_NAME)
    Set rezippedFile = rezipFiles(unzippedFiles.path, zipFile.path)
    
    Call fso.CopyFile(rezippedFile.path, filePath, True)
    
    Call f.files.DeleteFolder(f.files.getParentFolder(unzippedFiles.path))
    
    
'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = Err.Number: errDescription = Err.Description
    Call VBA.Err.Raise(errNumber, Description:=f.Strings.Format(Msg.getText(ERR_APPENDING_CUSTOM_RIBBON), _
                                                                                errNumber, errDescription))
       
End Sub

Private Function createZipFile(filePath As String) As Scripting.file
    Dim fileName As String
    Dim fileExtension As String
    Dim folder As Scripting.folder
    Dim zipFilePath As String
    '------------------------------------------------------------------------------------------------------
    
    fileName = f.files.GetFileName(filePath, True)
    fileExtension = f.files.getFileExtension(filePath, True)
    
    Set folder = createFolderWithFileCopy(filePath)
    zipFilePath = f.files.concatenatePath(folder.path, fileName)
    
    Set createZipFile = f.files.GetFile(zipFilePath)
    createZipFile.name = VBA.Replace(createZipFile.name, fileExtension, EXTENSION_ZIP)
    
End Function

Private Function createFolderWithFileCopy(filePath As String) As Scripting.folder
    Dim folderPath As String
    '------------------------------------------------------------------------------------------------------
    
    folderPath = f.files.concatenatePath(f.files.getParentFolder(filePath), f.files.GetFileName(filePath, False))
    
    Call f.files.DeleteFolder(folderPath)
    
    Set createFolderWithFileCopy = f.files.CreateFolder(folderPath)
    Call fso.CopyFile(filePath, f.files.concatenatePath(folderPath, f.files.GetFileName(filePath)), True)
    
End Function

Private Function unzipFiles(filePath As String) As Scripting.folder
    Dim sourceFilePath As String
    '------------------------------------------------------------------------------------------------------
    
    With f.files
        Set unzipFiles = .CreateFolder(.concatenatePath(.getParentFolder(filePath), "Items"))
        Call .unzip(filePath, unzipFiles.path)
    End With
    
End Function

Private Function rezipFiles(sourceFolderPath As String, destinationFilePath As String) As Scripting.file
    Call f.files.DeleteFile(destinationFilePath)
    Set rezipFiles = f.files.zip(sourceFolderPath, destinationFilePath)
End Function


''--rezip
' sTargetZipFilePath = sTempFolderPath & "\RibbonRestored.zip"
'
' Call Zip(sSourceFolderPath:=sTempFolderPath & "\Items", _
'   sTargetFilePath:=sTargetZipFilePath)
'
''--copy file unique name and Excel extension
' oFSO.CopyFile sTargetZipFilePath, sNewWorkbookFilePath
'
' MsgBox "A copy of this workbook with its custom ribbon restored was saved to: " _
'   & vbCr & vbCr & sNewWorkbookFilePath
'
'ExitProc:
' On Error Resume Next
'
' '--delete temp files and folder
' If oFSO.FolderExists(sTempFolderPath) Then
'   oFSO.DeleteFolder (sTempFolderPath)
' End If
'
' If Len(msErrMsg) Then
'   MsgBox msErrMsg, vbCritical
'   msErrMsg = vbNullString
' End If
'
' Exit Sub
'
'ErrProc:
' msErrMsg = Err.Number & "-" & Err.Description
' Resume ExitProc
'End Sub
'





Private Sub updateRels(topFolderOfItems As String, customUI_Filename As String)
    Const METHOD_NAME = "updateRels"
    '------------------------------------------------------------------------------------------------------
    '--handle no relationships node?
     Dim xmlDoc As Object
     Dim xmlNode As Object, xmlNewNode As Object
     Dim xmlNodes As Object
     Dim relsFilePath As String, namespace As String
     '-----------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo errHandler
    
    relsFilePath = topFolderOfItems & "\_rels\.rels"
    
    Set xmlDoc = CreateObject("Microsoft.XMLDOM")
    xmlDoc.Load relsFilePath
    
    With xmlDoc.SelectSingleNode("/Relationships")
        namespace = .NamespaceURI
    
        '--remove any existing nodes that would conflict with new relationship
        Set xmlNodes = xmlDoc.SelectNodes( _
            "//Relationship[@Id='customUIRelID' or  Target='" _
            & ATT_TARGET_2007 & "' or  Target='" & ATT_TARGET_2010 & "']")
    
        For Each xmlNode In xmlNodes
            Debug.Print "Deleting.." & xmlNode.Attributes.getNamedItem("Target").text
            xmlNode.ParentNode.RemoveChild xmlNode
        Next xmlNode
    
        '--add new node by cloning existing
        Set xmlNewNode = .ChildNodes(0).CloneNode(True)
        xmlNewNode.Attributes.getNamedItem("Id").text = "customUIRelID"
    
        Select Case customUI_Filename
             Case "customUI.xml" '2007
                    xmlNewNode.Attributes.getNamedItem("Type").text = ATT_TYPE_2007
                    xmlNewNode.Attributes.getNamedItem("Target").text = ATT_TARGET_2007
    
            Case "customUI14.xml" '2010
                    xmlNewNode.Attributes.getNamedItem("Type").text = ATT_TYPE_2010
                    xmlNewNode.Attributes.getNamedItem("Target").text = ATT_TARGET_2010
    
            Case Else
                    Call VBA.Err.Raise(1, , "XML filename for Custom Ribbon is unrecognized version.")
                    
      End Select
    
      Call .appendChild(xmlNewNode)
    
    End With
    
    Call xmlDoc.Save(relsFilePath)

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then
        Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = Err.Number: errDescription = Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
    
End Sub

Private Sub writeCustomUI_XML_ToFile(ribbonXml As String, customUiFolderPath As String, customUiFilename As String)
    Const METHOD_NAME = "writeCustomUI_XML_ToFile"
    '------------------------------------------------------------------------------------------------------
    Dim file As Object
    Dim customUiFilePath As String
    Dim folder As Scripting.folder
    '-----------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo errHandler
       
    customUiFilePath = f.files.concatenatePath(customUiFolderPath, customUiFilename)
    Set folder = f.files.CreateFolder(customUiFolderPath)
    Set file = fso.CreateTextFile(customUiFilePath, True)
    With file
        .WriteLine (ribbonXml)
        .Close
    End With

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then
        Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = Err.Number: errDescription = Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
    
End Sub

    
    
