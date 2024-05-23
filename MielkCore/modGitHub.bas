Attribute VB_Name = "modGitHub"
Option Explicit

Private Const URL_ADDRESS As String = "https://github.com/mielk/mielk-vba/tree/master/{0}"
Private Const DESTINATION_FOLDER As String = "D:\vba-tests\github"
Private Const ADDRESS_PREFIX As String = "https://github.com"
Private Const RAW_LINES_TAG As String = "rawLines"
'----------------------------------------------------------------------------------------------------------


Public Sub run()
    Call printContentToTextFiles
End Sub


Public Sub printContentToTextFiles()
    'Const REGEX_PATTERN As String = "<a title=""(\w*\.(?:cls|bas|frm)).*?href=""([\w\/-]*\.(?:cls|bas|frm))"
    Const REGEX_PATTERN As String = "<a title=""(\w*\.(?:clz|baz|frm)).*?href=""([\w\/-]*\.(?:clz|baz|frm))"
    '------------------------------------------------------------------------------------------------------
    Dim files As Variant
    Dim file As Variant
    Dim folderPath As String
    Dim url As String
    Dim response As String
    '------------------------------------------------------------------------------------------------------
    Dim sites As VBA.Collection
    Dim site As Variant
    Dim siteName As String
    Dim siteUrl As String
    Dim done As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    'files = VBA.Array("MielkApp", "MielkCore", "MielkDB", "MielkUI")
    files = VBA.Array("MielkUI")
    For Each file In files
    
        folderPath = f.files.concatenatePath(DESTINATION_FOLDER, file)
        'Call f.files.clearFolder(folderPath)
    
        Set done = f.dictionaries.Create(False)
        
        url = f.strings.format(URL_ADDRESS, file)
        response = f.Http.getResponse(url)
        
        Set sites = f.regex.getMatchArraysCollection(response, REGEX_PATTERN)
        For Each site In sites
            siteName = site(1)
            siteUrl = site(2)
            
            If Not done.exists(siteName) Then
                Call printSite(VBA.CStr(file), site)
                Call f.dictionaries.addItem(done, siteName, siteName)
            End If
        Next site
    Next file
    
End Sub


Private Sub printSite(project As String, arr As Variant)
    Const CODE_PATTERN As String = "(?:""Attribute VB_\w* = (?:True|False)"",)+(.*)(?:],""stylingDirectives"")"
    Const CODE_PREFIX As String = "{""rawLines"":["
    Const CODE_SUFFIX As String = "]}"
    '------------------------------------------------------------------------------------------------------
    Dim url As String
    Dim response As String
    Dim className As String
    Dim filepath As String
    Dim code As String
    Dim codeLines As Variant
    '------------------------------------------------------------------------------------------------------
    
    url = ADDRESS_PREFIX & arr(2)
    className = VBA.CStr(arr(1))
    filepath = f.files.concatenatePath(DESTINATION_FOLDER, project, className & ".txt")
    
    response = f.Http.getResponse(url)
    code = CODE_PREFIX & f.regex.getFirstGroupMatch(response, CODE_PATTERN) & CODE_SUFFIX
    codeLines = getCodeLinesFromRawText(code)
    
    
    Call f.TextFiles.printToTextFile(VBA.join(codeLines, VBA.vbCrLf), filepath, True)
    
End Sub


Private Function getCodeLinesFromRawText(text As String) As Variant
    Dim json As Scripting.Dictionary
    Dim rawLines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    Set json = f.json.ParseJson(text)
    Set rawLines = json.item(RAW_LINES_TAG)
    getCodeLinesFromRawText = f.Collections.toArray(rawLines)
    
End Function





Public Sub createFileFromTextfiles(folderPath As String)
    Const FILE_NAME_PATTERN As String = "([^\\]*)$"
    '------------------------------------------------------------------------------------------------------
    Dim wkb As Excel.Workbook
    Dim project As VBIDE.VBProject
    Dim fileName As String
    '------------------------------------------------------------------------------------------------------
    Dim files As VBA.Collection
    Dim file As Variant
    Dim fileType As VBIDE.vbext_ComponentType
    Dim component As VBIDE.VBComponent
    Dim content As String
    '------------------------------------------------------------------------------------------------------
    
    Set wkb = f.Books.addNew(1, Excel.Application)
    Set project = f.Developer.getVbProject(wkb)
    
    Set files = f.files.getFolderFiles(folderPath, False, "\.txt")
    For Each file In files
        content = f.TextFiles.readTextFile(file.path)
        fileType = getFileType(file.path)
        Set component = project.VBComponents.Add(fileType)
        component.name = f.strings.substring(file.name, vbNullString, ".")
        Call component.codeModule.AddFromString(clearContent(content))
    Next file
    
    fileName = f.regex.getFirstGroupMatch(folderPath, FILE_NAME_PATTERN) & ".xlsm"
    
    Call wkb.SaveAs(f.files.concatenatePath(folderPath, fileName), xlOpenXMLWorkbookMacroEnabled)
    
End Sub



Private Function getFileType(filepath As String) As VBIDE.vbext_ComponentType
    Const CLASS_REGEX_PATTERN As String = "\.cls\."
    Const STD_MOD_REGEX_PATTERN As String = "\.bas\."
    Const FORM_REGEX_PATTERN As String = "\.frm\."
    '------------------------------------------------------------------------------------------------------
    
    If f.regex.checkIfMatch(filepath, CLASS_REGEX_PATTERN) Then
        getFileType = vbext_ct_ClassModule
    ElseIf f.regex.checkIfMatch(filepath, STD_MOD_REGEX_PATTERN) Then
        getFileType = vbext_ct_StdModule
    ElseIf f.regex.checkIfMatch(filepath, FORM_REGEX_PATTERN) Then
        getFileType = vbext_ct_MSForm
    End If
    
End Function



Private Function clearContent(content As String) As String
    Const REGEX_PATTERN As String = "(?:Option Explicit|Attribute ).*$"
    '------------------------------------------------------------------------------------------------------
    clearContent = f.regex.replace(content, REGEX_PATTERN, vbNullString)
End Function
