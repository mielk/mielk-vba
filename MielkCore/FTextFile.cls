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
Private Const CP_UTF8 = 65001
'------------------------------------------------------------------------------------------------------



'[WRITING TO TEXT FILES]

Public Sub printToTextFile(content As Variant, filePath As String, Optional override As Boolean = False)
    Dim intFile As Integer
    '------------------------------------------------------------------------------------------------------
    
    If override Then Call F.files.deleteFile(filePath)
    
    If F.files.createFolder(F.files.getParentFolder(filePath)) Is Nothing Then
        Call VBA.err.raise(1, description:="Error while creating parent folder")
    End If
    
    intFile = VBA.FreeFile
    Open filePath For Append As #intFile
    Print #intFile, content;
    Close intFile

End Sub


Public Sub printToTextFile_Scripting(content As String, filePath As String, Optional override As Boolean = False)
    Const MAX_ATTEMPTS As Long = 1000000
    '----------------------------------------------------------------------------------------------------------------------
    Dim fileStream As TextStream
    Dim attemptsCounter As Long
    '----------------------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    Do While fileStream Is Nothing
        Set fileStream = fso.OpenTextFile(filePath, ForAppending, True, -1)
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


Public Sub printToTextFileUnicode(content As Variant, filePath As String, Optional override As Boolean)
    Dim objStream As Object
    '------------------------------------------------------------------------------------------------------
    
    If override Then
        Call F.files.deleteFile(filePath)
    End If
    
    Call F.files.createFolder(F.files.getParentFolder(filePath))
    
    If Not F.Variables.isUndefined(content, True, False, True) Then
        Set objStream = VBA.CreateObject("ADODB.Stream")
        With objStream
            .Type = 2                       'Specify stream type - we want To save text/string data.
            .Charset = "unicode"            'Specify charset For the source text data.
            .Open                           'Open the stream And write binary data To the object
            .WriteText VBA.CStr(content)    'M.Strings.stringify(content, StringifyMode_Normal)
            .SaveToFile filePath, 2         'Save binary data To disk
        End With
    End If
End Sub


Public Function isTextfileWriteable(filePath As String) As Boolean
    Dim TEXT_EXTENSIONS As Variant:         TEXT_EXTENSIONS = Array("txt", "csv")
    '------------------------------------------------------------------------------------------------------
    Dim strDriveName As String
    Dim strBaseFolder As String
    Dim strParentFolder As String
    Dim strExtension As String
    '------------------------------------------------------------------------------------------------------
    
    With fso
        If .fileExists(filePath) Then
            isTextfileWriteable = Not (VBA.GetAttr(filePath) And VBA.vbReadOnly)
        Else
            
            strDriveName = .GetDriveName(filePath)
            If .DriveExists(strDriveName) Then
                strExtension = .GetExtensionName(filePath)
                
                If F.arrays.contains(TEXT_EXTENSIONS, strExtension, False) Then
                    strParentFolder = .GetParentFolderName(filePath)
                    strBaseFolder = .GetParentFolderName(filePath)
                    Do Until .folderExists(strBaseFolder)
                        strBaseFolder = .GetParentFolderName(strBaseFolder)
                    Loop
                    
                    If VBA.StrComp(strBaseFolder, strParentFolder, vbTextCompare) = 0 Then
                        isTextfileWriteable = F.files.isFolderWriteable(strParentFolder)
                    Else
                        isTextfileWriteable = F.files.isSubfoldersAddingAllowed(strBaseFolder)
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

Public Function readTextFile(filePath As String) As String
    Dim fileSize As Long
    Dim fileIndex As Integer
    Dim byteArray() As Byte
    '----------------------------------------------------------------------------------------------------------------------
    
    fileSize = VBA.FileLen(filePath)
    If fileSize > 0 Then
        ReDim byteArray(0 To fileSize - 1)

        ' Read the existing UTF-8 file
        fileIndex = VBA.FreeFile()
        Open filePath For Binary As #fileIndex
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

    size = F.arrays.getSize(data, 1)
    nc = size
    convertUTF8ToUni = VBA.String$(nc, VBA.Chr(0))
    ret = MultiByteToWideChar(CP_UTF8, 0, VBA.VarPtr(data(LBound(data))), _
                                    size, VBA.StrPtr(convertUTF8ToUni), nc)
    convertUTF8ToUni = VBA.left$(convertUTF8ToUni, ret)

End Function

Public Function readTextFile2(filePath As String) As String
    Const METHOD_NAME As String = "readTextFile"
    '----------------------------------------------------------------------------------------------------------------------
    Dim fileIndex As Integer
    Dim col As VBA.Collection
    Dim textLine As String
    '----------------------------------------------------------------------------------------------------------------------

    fileIndex = VBA.FreeFile()
    Set col = New VBA.Collection
    
    Open filePath For Input As #fileIndex
        Do Until EOF(fileIndex)
            Line Input #fileIndex, textLine
            Call col.Add(textLine)
        Loop
    Close #fileIndex
    
    readTextFile2 = F.Collections.toString(col, StringifyMode_Normal, VBA.vbCrLf)

End Function

Public Function getCsvContentAsArray(filePath As String, Optional delimiter As String = ",") As Variant
    Dim content As String
    Dim lines() As String
    Dim columns() As Variant
    Dim maxColumns As Long
    Dim arr() As Variant
    Dim i As Long
    Dim j As Long
    '----------------------------------------------------------------------------------------------------------------------
    
    content = F.strings.cutFromEnd(readTextFile(filePath), VBA.vbCrLf)
    
    If VBA.Len(content) Then
        lines = VBA.Split(content, VBA.vbLf)
        If F.arrays.isDefined(lines) Then
            ReDim columns(LBound(lines) To UBound(lines))
            For i = LBound(lines) To UBound(lines)
                columns(i) = VBA.Split(lines(i), delimiter)
                If F.arrays.isDefined(columns(i)) Then
                    maxColumns = F.Math.getMax(maxColumns, F.arrays.getSize(columns(i)))
                End If
            Next i
        End If
    End If
    
    ReDim arr(1 To F.arrays.getSize(columns), 1 To maxColumns)
    
    For i = LBound(columns) To UBound(columns)
        If F.arrays.isDefined(columns(i)) Then
            For j = LBound(columns(i)) To UBound(columns(i))
                arr(i + LBound(arr, 1) - LBound(columns, 1), j + LBound(arr, 2) - LBound(columns(i))) = columns(i)(j)
            Next j
        End If
    Next i
    
    getCsvContentAsArray = arr
    
End Function
