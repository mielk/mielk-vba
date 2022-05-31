VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FConstructors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FConstructors"
'----------------------------------------------------------------------------------------------------------


'[Connections]
Public Function SqlServerConnection_WindowsAuthentication(dataSource As String, dbName As String) As SqlServerConnection
    Set SqlServerConnection_WindowsAuthentication = New SqlServerConnection
    Call SqlServerConnection_WindowsAuthentication.setConnectionDetails(dataSource, dbName)
End Function

Public Function SqlServerConnection_PasswordAuthentication(dataSource As String, dbName As String, _
                                                           userName As String, password As String, _
                                                           Optional provider As String = DEFAULT_OLEDB_PROVIDER) As SqlServerConnection
    Set SqlServerConnection_PasswordAuthentication = New SqlServerConnection
    Call SqlServerConnection_PasswordAuthentication.setConnectionDetails(dataSource, dbName, provider, userName, password)
End Function

Public Function AccessConnection(dataSource As String, readWriteMode As ReadWriteModeEnum, Optional password As String = vbNullString) As AccessConnection
    Set AccessConnection = New AccessConnection
    Call AccessConnection.setConnectionDetails(dataSource, readWriteMode, password)
End Function




'[Other objects]
Public Function Where(fieldName As String, comparisonMode As ComparisonModeEnum, comparedValue As Variant) As EWhere
    Set Where = New EWhere
    With Where
        Call .setFieldName(fieldName)
        Call .setComparisonMode(comparisonMode)
        Call .setComparedValue(comparedValue)
    End With
End Function

