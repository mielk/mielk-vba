Attribute VB_Name = "modExcelFormulas"
Option Explicit

Private Const CLASS_NAME As String = "modExcelFormulas"
'----------------------------------------------------------------------------------------------------------

Public Function APP_NAME() As String
    APP_NAME = APPLICATION_NAME
End Function

Public Function APP_CODE_NAME() As String
    APP_CODE_NAME = APPLICATION_CODE_NAME
End Function

Public Function APP_VERSION(Optional onlyNumber As Boolean = False) As String
    If onlyNumber Then
        APP_VERSION = APPLICATION_VERSION
    Else
        APP_VERSION = F.Strings.Format(Msg.getText("VersionInfo"), APPLICATION_VERSION)
    End If
End Function

