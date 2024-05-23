VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FTextfiles"
'[Windows API] ----------------------------------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
        ByVal CodePage As LongPtr, ByVal dwFlags As LongPtr, _
        ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As LongPtr, _
        ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As LongPtr) As Long
#Else
    Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, _
        ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, _
        ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
#End If
'------------------------------------------------------------------------------------------------------
Private Const CP_UTF8 As Long = 65001
'------------------------------------------------------------------------------------------------------
Private Const SEPARATOR_TEMP_REPLACEMENT As String = "™"
Private Const TRIM_FILE_REGEX As String = "^[\s\n]*|[\s\n]*$"
Private Const OPEN_CLOSE_QUOTE_REGEX As String = "(^""|""$)"
Private Const FIND_INNER_FIELDS_QUOTES_REGEX As String = """([^,""]*)"""
Private Const REPLACE_INNER_FIELDS_QUOTES_REGEX As String = "$1"
Private Const FIND_SEPARATOR_REGEX_PATTERN As String = "(^|{0})""+([^{0}]*){0}([^{0}]*?)""+(?=$|{0})"
Private Const REPLACE_SEPARATOR_REGEX_PATTERN As String = "$1$2{0}$3"
Private Const DEQUOTATION_CHECK_REGEX_PATTERN As String = "^""(?:[^""]*|""{1}[^""]*""{1}|""{2}[^""]*""{2})(?:{0}""[^""]*""|{0}""{1}[^""]*""{1}|{0}""{2}[^""]*""{2})*""$"
Private Const DEQUOTATION_REPLACED_REGEX_PATTERN As String = "(^|{0})""{0,3}([^""]*)""{0,3}(?={0}|$|\n)"
Private Const DEQUOTATION_REPLACING_REGEX_PATTERN As String = "$1""$2"""
Private Const EMPTY_LINE_REPLACED_REGEX_PATTERN As String = "^""*$(\n|)"
'------------------------------------------------------------------------------------------------------



'[WRITING TO TEXT FILES]

Public Sub printToTextFile(content As Variant, filepath As String, Optional override As Boolean = False)
    Dim intFile As Integer
    '------------------------------------------------------------------------------------------------------
    
    If override Then Call f.files.deleteFile(filepath)
    
    If f.files.createFolder(f.files.getParentFolder(filepath)) Is Nothing Then
        Call VBA.err.raise(1, description:="Error while creating parent folder")
    End If
    
    intFile = VBA.FreeFile
    Open filepath For Append As #intFile
    Print #intFile, content;
    Close intFile

End Sub


Public Sub printToTextFile_Scripting(content As String, filepath As String, Optional override As Boolean = False)
    Const MAX_ATTEMPTS As Long = 1000000
    '----------------------------------------------------------------------------------------------------------------------
    Dim fileStream As TextStream
    Dim attemptsCounter As Long
    '----------------------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    
    Call f.files.createFolder(f.files.getFileFolderPath(filepath))
    
    Do While fileStream Is Nothing
        Set fileStream = fso.OpenTextFile(filepath, ForAppending, True, -1)
        attemptsCounter = attemptsCounter + 1
        If attemptsCounter > MAX_ATTEMPTS Then Exit Do
    Loop
    
    If Not fileStream Is Nothing Then
        With fileStream
            .Write content
            .Close
        End With
    Else
        Call VBA.err.raise(1, description:="File is locked")
    End If
    
End Sub


Public Sub printToTextFileUnicode(content As Variant, filepath As String, Optional override As Boolean)
    Dim objStream As Object
    '------------------------------------------------------------------------------------------------------
    
    If override Then
        Call f.files.deleteFile(filepath)
    End If
    
    Call f.files.createFolder(f.files.getParentFolder(filepath))
    
    If Not f.Variables.isUndefined(content, True, False, True) Then
        Set objStream = VBA.CreateObject("ADODB.Stream")
        With objStream
            .Type = 2                       'Specify stream type - we want To save text/string data.
            .Charset = "unicode"            'Specify charset For the source text data.
            .Open                           'Open the stream And write binary data To the object
            .WriteText VBA.CStr(content)    'M.Strings.stringify(content, StringifyMode_Normal)
            .SaveToFile filepath, 2         'Save binary data To disk
        End With
    End If
End Sub


Public Function isTextfileWriteable(filepath As String) As Boolean
    Dim TEXT_EXTENSIONS As Variant:         TEXT_EXTENSIONS = Array("txt", "csv")
    '------------------------------------------------------------------------------------------------------
    Dim strDriveName As String
    Dim strBaseFolder As String
    Dim strParentFolder As String
    Dim strExtension As String
    '------------------------------------------------------------------------------------------------------
    
    With fso
        If .fileExists(filepath) Then
            isTextfileWriteable = Not (VBA.GetAttr(filepath) And VBA.vbReadOnly)
        Else
            
            strDriveName = .GetDriveName(filepath)
            If .DriveExists(strDriveName) Then
                strExtension = .GetExtensionName(filepath)
                
                If f.arrays.contains(TEXT_EXTENSIONS, strExtension, False) Then
                    strParentFolder = .GetParentFolderName(filepath)
                    strBaseFolder = .GetParentFolderName(filepath)
                    Do Until .folderExists(strBaseFolder)
                        strBaseFolder = .GetParentFolderName(strBaseFolder)
                    Loop
                    
                    If VBA.StrComp(strBaseFolder, strParentFolder, vbTextCompare) = 0 Then
                        isTextfileWriteable = f.files.isFolderWriteable(strParentFolder)
                    Else
                        isTextfileWriteable = f.files.isSubfoldersAddingAllowed(strBaseFolder)
                    End If
                
                Else
                    isTextfileWriteable = False
                End If
                                                                                                      
            Else
                isTextfileWriteable = False
            End If
                                                                                                        
        End If

    End With

End Function



'[READING TEXT FILES]

Public Function readTextFile(filepath As String) As String
    Dim FileSize As Long
    Dim fileIndex As Integer
    Dim byteArray() As Byte
    '----------------------------------------------------------------------------------------------------------------------
    
    FileSize = VBA.FileLen(filepath)
    If FileSize > 0 Then
        ReDim byteArray(0 To FileSize - 1)

        ' Read the existing UTF-8 file
        fileIndex = VBA.FreeFile()
        Open filepath For Binary As #fileIndex
        Get #fileIndex, , byteArray
        Close #fileIndex

        ' Convert all the data to Unicode (all VB Strings are Unicode)
        readTextFile = convertUTF8ToUni(byteArray)
    End If
    
End Function

Private Function convertUTF8ToUni(data() As Byte) As String
    Const METHOD_NAME As String = "readTextFile"
    '----------------------------------------------------------------------------------------------------------------------
    Dim size As Long
    Dim nc As Long
    Dim ret As Long
    '----------------------------------------------------------------------------------------------------------------------

    size = f.arrays.getSize(data, 1)
    nc = size
    convertUTF8ToUni = VBA.String$(nc, VBA.Chr(0))
    ret = MultiByteToWideChar(CP_UTF8, 0, VBA.VarPtr(data(LBound(data))), _
                                    size, VBA.StrPtr(convertUTF8ToUni), nc)
    convertUTF8ToUni = VBA.left$(convertUTF8ToUni, ret)

End Function

Private Function readTextFile2(filepath As String) As String
    Const METHOD_NAME As String = "readTextFile"
    '----------------------------------------------------------------------------------------------------------------------
    Dim fileIndex As Integer
    Dim col As VBA.Collection
    Dim textLine As String
    '----------------------------------------------------------------------------------------------------------------------

    fileIndex = VBA.FreeFile()
    Set col = New VBA.Collection
    
    Open filepath For Input As #fileIndex
        Do Until EOF(fileIndex)
            Line Input #fileIndex, textLine
            Call col.Add(textLine)
        Loop
    Close #fileIndex
    
    readTextFile2 = f.Collections.toString(col, StringifyMode_Normal, VBA.vbCrLf)

End Function

Public Function getCsvContentAsArray(filepath As String, _
                                     Optional delimiter As String = ",", _
                                     Optional customAdjustments As Scripting.Dictionary) As Variant
    Dim content As String
    Dim lines() As String
    Dim text As String
    Dim columns() As Variant
    Dim maxColumns As Long
    Dim arr() As Variant
    Dim i As Long
    Dim j As Long
    '----------------------------------------------------------------------------------------------------------------------
    Dim findRegex As String
    Dim replacementRegex As String
    Dim innerQuotesRegex As String
    Dim dequotationCheckRegex As String
    Dim dequotationReplacedRegex As String
    '[Custom adjustments] -------------------------------------------------------------------------------------------------
    Dim key As Variant
    Dim adjustment As String
    '----------------------------------------------------------------------------------------------------------------------
    
    content = f.strings.cutFromEnd(readTextFile(filepath), VBA.vbCrLf)
    
    If f.dictionaries.isNonEmptyDictionary(customAdjustments) Then
        For Each key In customAdjustments.Keys
            adjustment = customAdjustments.item(key)
            content = f.regex.replace(content, VBA.CStr(key), adjustment)
        Next key
    End If
    
    findRegex = f.strings.format(FIND_SEPARATOR_REGEX_PATTERN, SEPARATOR_TEMP_REPLACEMENT)
    replacementRegex = f.strings.format(REPLACE_SEPARATOR_REGEX_PATTERN, delimiter)
    innerQuotesRegex = f.strings.format(FIND_INNER_FIELDS_QUOTES_REGEX, delimiter)
    dequotationCheckRegex = f.strings.format(DEQUOTATION_CHECK_REGEX_PATTERN, delimiter)
    dequotationReplacedRegex = f.strings.format(DEQUOTATION_REPLACED_REGEX_PATTERN, delimiter)
    
    
    '[Adjust CSV content - replace all separator characters with other character;
    ' leave only separator characters used inside strings in quotes]
    content = readTextFile(filepath)
    content = f.regex.replace(content, TRIM_FILE_REGEX, vbNullString, multiline:=False)
    content = f.regex.replace(content, VBA.Chr(13) & VBA.Chr(10), VBA.Chr(10))
    
    If f.regex.checkIfMatch(content, dequotationCheckRegex) Then
        content = f.regex.replace(content, dequotationReplacedRegex, DEQUOTATION_REPLACING_REGEX_PATTERN)
    End If
    
    content = f.regex.replace(content, innerQuotesRegex, REPLACE_INNER_FIELDS_QUOTES_REGEX)
    content = f.regex.replace(content, delimiter, SEPARATOR_TEMP_REPLACEMENT)
    content = f.regex.replace(content, findRegex, replacementRegex)
    
    
    
    
    If VBA.Len(content) Then
        lines = VBA.Split(content, VBA.vbLf)
        If f.arrays.isDefined(lines) Then
            ReDim columns(LBound(lines) To UBound(lines))
            For i = LBound(lines) To UBound(lines)
                text = lines(i)
                columns(i) = VBA.Split(text, SEPARATOR_TEMP_REPLACEMENT)
                If f.arrays.isDefined(columns(i)) Then
                    maxColumns = f.Math.getMax(maxColumns, f.arrays.getSize(columns(i)))
                End If
            Next i
        End If
    End If
    
    ReDim arr(1 To f.arrays.getSize(columns), 1 To maxColumns)
    
    For i = LBound(columns) To UBound(columns)
        If f.arrays.isDefined(columns(i)) Then
            For j = LBound(columns(i)) To UBound(columns(i))
                arr(i + LBound(arr, 1) - LBound(columns, 1), j + LBound(arr, 2) - LBound(columns(i))) = columns(i)(j)
            Next j
        End If
    Next i
    
    getCsvContentAsArray = arr
    
End Function

Public Function getCsvContentWithModifiedSeparator(filepath As String, originalSeparator As String, designatedSeparator As String) As String
    Dim arr As Variant
    '----------------------------------------------------------------------------------------------------------------------
    arr = f.TextFiles.getCsvContentAsArray(filepath, originalSeparator)
    getCsvContentWithModifiedSeparator = f.arrays.toString(arr, , designatedSeparator, rowStartTag:=vbNullString, rowEndTag:=vbNullString)
End Function

Public Sub printArrayToCsv(arr As Variant, filepath As String, Optional delimiter As String = ",", _
                            Optional override As Boolean = False)
    Dim arrResult As Variant
    Dim i As Long
    Dim subArr As Variant
    Dim j As Long
    '----------------------------------------------------------------------------------------------------------------------
    
    If f.arrays.isDefined(arr) Then
        ReDim arrResult(LBound(arr, 1) To UBound(arr, 1))
        For i = LBound(arr) To UBound(arr)
            ReDim subArr(LBound(arr, 2) To UBound(arr, 2))
            For j = LBound(arr, 2) To UBound(arr, 2)
                If VBA.InStr(1, arr(i, j), delimiter) > 0 Then
                    subArr(j) = VBA.Chr(34) & VBA.replace(arr(i, j), VBA.Chr(34), VBA.Chr(34) & VBA.Chr(34)) & VBA.Chr(34)
                Else
                    subArr(j) = arr(i, j)
                End If
            Next j
            arrResult(i) = VBA.join(subArr, delimiter)
        Next i
    End If
    
    Call f.TextFiles.printToTextFile(VBA.join(arrResult, VBA.vbLf), filepath, override)
    
End Sub
