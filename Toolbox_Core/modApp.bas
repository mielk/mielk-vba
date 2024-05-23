Attribute VB_Name = "modApp"
Option Explicit

Private Const CLASS_NAME As String = "modApp"
'----------------------------------------------------------------------------------------------------------

Public Sub setParentApp()
    With App
        Call .setName(APPLICATION_NAME)
        Call .setVersion(APPLICATION_VERSION)
        Call .setPath(F.files.getUncPath(ThisWorkbook.FullName))
    End With
End Sub
