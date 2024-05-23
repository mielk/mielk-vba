VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FPdfs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PDF_SAM_JAR_FILE As String = """C:\Program Files (x86)\OpenSource\PDF Split and Merge\lib\pdfsam-console-2.3.1e.jar"""
'----------------------------------------------------------------------------------------------------------



Public Function getPdfObject(filepath As String) As Object
    Const MAX_ATTEMPTS As Long = 10
    '------------------------------------------------------------------------------------------------------
    Dim attempt As Long
    '------------------------------------------------------------------------------------------------------
    
    Set getPdfObject = createNewPdfObject
    
    If getPdfObject Is Nothing Then
        Call VBA.err.raise(Exceptions.CUSTOM_NUMBER, _
                            description:="Error when trying to create instances of PDFPlus.DDDoc")
    Else
        Do While attempt < MAX_ATTEMPTS
            Call getPdfObject.Open(filepath)
            Call Sleep(50)
            If f.files.getFileName(filepath, True) = getPdfObject.getFileName Then Exit Do
            Call Sleep(50)
        Loop
    End If
    
End Function



Public Function createNewPdfObject() As Object
    Const MAX_ATTEMPTS As Long = 10
    '------------------------------------------------------------------------------------------------------
    Dim attempts As Long
    '------------------------------------------------------------------------------------------------------
    
    Do While attempts < MAX_ATTEMPTS
        Set createNewPdfObject = VBA.CreateObject("NuancePDF.DDDoc")
        If Not createNewPdfObject Is Nothing Then Exit Do
        
        Set createNewPdfObject = VBA.CreateObject("PDFPlus.DDDoc")
        If Not createNewPdfObject Is Nothing Then Exit Do
        
        Call Sleep(50)
    Loop
    
End Function



Public Sub mergePdfs(ByVal filePaths As Variant, destinationPath As String)
    Dim colFilePaths As VBA.Collection
    Dim sourcePath As String
    Dim tempPath As String
    Dim finalPdf As Object                  'PDFPlus.DDDoc
    Dim partPdf As Object                   'PDFPlus.DDDoc
    '------------------------------------------------------------------------------------------------------
    Dim path As String
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    Set colFilePaths = f.Collections.createFrom(filePaths)
    
    If Not f.Collections.isNonEmptyCollection(colFilePaths) Then
        ' Nothing to merge
    ElseIf colFilePaths.count = 1 Then              ' One PDF only, no point merging it.
        Call fso.CopyFile(VBA.CStr(colFilePaths(1)), destinationPath, True)
    Else
        
        'Create temporary PDF file for collection pages from all partia PDFs
        sourcePath = colFilePaths(1)
        If Not f.files.fileExists(sourcePath) Then
            Call VBA.err.raise(Exceptions.CUSTOM_NUMBER, description:="File [" & sourcePath & "] doesn't exist")
        End If
        
        tempPath = f.files.concatenatePath(f.files.getFileFolderPath(destinationPath), f.Crypto.createUUID & ".pdf")
        Call f.files.createFolder(f.files.getFileFolderPath(tempPath))
        Call fso.CopyFile(sourcePath, tempPath, True)
        If Not f.files.fileExists(tempPath) Then
            Call VBA.err.raise(Exceptions.CUSTOM_NUMBER, description:="Error while copying file [" & sourcePath & "] to [" & tempPath & "]")
        End If
        Set finalPdf = f.Pdfs.getPdfObject(tempPath)
        
        'Append pages from rest of PDFs.
        If finalPdf Is Nothing Then
            If f.System.isDeveloper Then Stop
            Call VBA.err.raise(Exceptions.CUSTOM_NUMBER, description:="Variable [finalPdf] cannot be Nothing (1)")
        Else
            With finalPdf
                For i = 2 To colFilePaths.count
                    path = colFilePaths(i)
                    If f.files.fileExists(path) Then
                        Set partPdf = f.Pdfs.getPdfObject(path)
                        If partPdf Is Nothing Then
                            If f.System.isDeveloper Then Stop
                            Call VBA.err.raise(Exceptions.CUSTOM_NUMBER, description:="Variable [partPdf] cannot be Nothing (2)")
                        Else
                            If Not finalPdf.insertpages(finalPdf.getNumPages - 1, partPdf, 0, partPdf.getNumPages, True) Then
                                Call VBA.err.raise(Exceptions.CUSTOM_NUMBER, description:="Eror while merging PDF files")
                            End If
                            Call partPdf.Close
                        End If
                    End If
                Next i
            End With
        End If
        
        
        'Save temporary PDF
        If Not finalPdf.save(1, destinationPath) Then
            Call VBA.err.raise(Exceptions.CUSTOM_NUMBER, description:="Eror while saving PDF file")
        End If
        
        Call f.files.deleteFile(tempPath)
        
    End If
    
End Sub



Public Sub mergePdfs_SAM(ByVal filePaths As Variant, destinationPath As String)
    Const Q As String = """"
    '------------------------------------------------------------------------------------------------------
    Dim pdfSam As String
    Dim pdfFiles As String
    Dim pdfOut As String
    Dim pdfSamStr As String
    '------------------------------------------------------------------------------------------------------
    
    pdfSam = "java -jar " & PDF_SAM_JAR_FILE
    pdfFiles = f.arrays.toString(filePaths, StringifyMode_Normal, " ", "-f """, """")
    pdfOut = "-o " & Q & destinationPath & Q
    pdfSamStr = pdfSam & " " & pdfFiles & " " & pdfOut & " -overwrite concat"
    
    Debug.Print pdfSamStr
    
    VBA.Shell pdfSamStr, vbNormalFocus
    
End Sub
