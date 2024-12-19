VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "CDataTypes"
'----------------------------------------------------------------------------------------------------------
Private pByName As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------




'[INITIALIZING]
Private Sub Class_Initialize()
    Call initializeDictionaries
    Call insertEnumOptions
End Sub

Private Sub initializeDictionaries()
    Set pByName = F.dictionaries.Create(False)
End Sub

Private Sub insertEnumOptions()
    Call addEnumOption("Ribbon", "ribbon")
    '[Channels]
    Call addEnumOption("Channels", "[dbo].[Channels]", "[dbo].[Channels]")
    Call addEnumOption("AliasesChannels", "[dbo].[AliasesChannels]", "[dbo].[AliasesChannels]")

End Sub

Private Sub addEnumOption(ByVal name As String, ByVal readDbTable As String, Optional ByVal writeDbTable As String)
    Dim enumDataType As enumDataType
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(writeDbTable) = 0 Then writeDbTable = readDbTable
    Set enumDataType = F.Create.DataType(name, readDbTable, writeDbTable)
    
    Call F.dictionaries.AddItem(pByName, enumDataType.getName, enumDataType, duplicateBehaviour_WarningInImmediateWindow)
    
End Sub




'[ACCESS]
Public Function ByName(tableName As String) As enumDataType
    Set ByName = F.dictionaries.GetObject(pByName, tableName)
End Function





'[INTELLISENSE]
Public Function Ribbon() As enumDataType: Set Ribbon = ByName("Ribbon"): End Function

'[Channels]
Public Function Channels() As enumDataType: Set Channels = ByName("Channels"): End Function
Public Function AliasesChannels() As enumDataType: Set AliasesChannels = ByName("AliasesChannels"): End Function
