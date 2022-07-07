VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommonProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "CommonProperties"
'----------------------------------------------------------------------------------------------------------
Private pByNames As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------



'[INITIALIZING]
Private Sub Class_Initialize()
    Call initializeDictionaries
    Call insertEnumOptions
End Sub

Private Sub initializeDictionaries()
    Set pByNames = f.dictionaries.Create(False)
End Sub

Private Sub insertEnumOptions()
    '[Common]
    Call addEnumOption("Caption")
    Call addEnumOption("ContextMenu")
    Call addEnumOption("CreatedDate", "CreatedDate")
    Call addEnumOption("Date_")
    Call addEnumOption("Dates")
    Call addEnumOption("DbProp")
    Call addEnumOption("Header")
    Call addEnumOption("Id", "Id")
    Call addEnumOption("Index")
    Call addEnumOption("IsActive", "IsActive")
    Call addEnumOption("IsComplete")
    Call addEnumOption("IsConfirmed")
    Call addEnumOption("IsEditable")
    Call addEnumOption("IsMissing")
    Call addEnumOption("IsSelected")
    Call addEnumOption("Item")
    Call addEnumOption("Items")
    Call addEnumOption("Key")
    Call addEnumOption("Layout")
    Call addEnumOption("Method")
    Call addEnumOption("Month")
    Call addEnumOption("Name")
    Call addEnumOption("Parent")
    Call addEnumOption("Prop")
    Call addEnumOption("Separator")
    Call addEnumOption("SessionId")
    Call addEnumOption("SettingType")
    Call addEnumOption("SettingValue")
    Call addEnumOption("SourceFile")
    Call addEnumOption("Status")
    Call addEnumOption("Timestamp")
    Call addEnumOption("Total")
    Call addEnumOption("TypeId")
    Call addEnumOption("Uuid")
    Call addEnumOption("Version")
    Call addEnumOption("Value")
    Call addEnumOption("Year")
    '[Common - worksheet]
    Call addEnumOption("Range")
    Call addEnumOption("RangeName")
    Call addEnumOption("RowHeight")
    Call addEnumOption("Sheet")
    Call addEnumOption("SheetName")
    '[Common - config]
    Call addEnumOption("Rows")
    Call addEnumOption("Columns")
    Call addEnumOption("Styles")
    Call addEnumOption("StylesConditional")
    Call addEnumOption("Events")
    Call addEnumOption("HeaderEvents")
    Call addEnumOption("Template")
    '[Common - files]
    Call addEnumOption("FilePath", "FilePath")
    Call addEnumOption("FileType", "FileType")
    Call addEnumOption("FileName", "FileName")
    Call addEnumOption("FileSize", "FileSize")
    Call addEnumOption("FileTimestamp", "FileTimestamp")
    Call addEnumOption("FileCreator", "FileCreator")
    Call addEnumOption("SourceFilePath", "SourceFilePath")
    Call addEnumOption("DestinationFilePath", "DestinationFilePath")
    '[Common - DB]
    Call addEnumOption("DbUploadRecordsSeparator", "RecordsSeparator")
    Call addEnumOption("DbUploadFieldsSeparator", "FieldsSeparator")
    Call addEnumOption("OutputPrefix", "Prefix")
    
    '[Common - styling]
    Call addEnumOption("IndexFrom", "from")
    Call addEnumOption("IndexTo", "to")
    Call addEnumOption("RowsFirst", "rowsFirst")
    Call addEnumOption("HeaderLast", "headerLast")
    Call addEnumOption("Subranges", "subranges")

    '[Dates range]
    Call addEnumOption("StartDate", "StartDate")
    Call addEnumOption("EndDate", "EndDate")
    Call addEnumOption("StartMonth", "StartMonth")
    Call addEnumOption("EndMonth", "EndMonth")
    
    '[Aliases]
    Call addEnumOption("Alias", "Alias")
    Call addEnumOption("Aliases", "Aliases")
    Call addEnumOption("AliasesString")

End Sub

Private Sub addEnumOption(name As String, Optional ByVal dbField As String)
    Dim enumProperty As enumProperty
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(dbField) = 0 Then dbField = name
    Set enumProperty = f.Create.Prop(name, dbField)
    
    Call f.dictionaries.addItem(pByNames, enumProperty.getName, enumProperty, duplicateBehaviour_WarningInImmediateWindow)
    
End Sub



'[ACCESS]
Public Function ByName(name As String) As enumProperty
    Set ByName = f.dictionaries.getObject(pByNames, name)
End Function

Public Function getProperty(ByVal Prop As Variant) As enumProperty
    If VBA.IsArray(Prop) Then
        'not possible to get property
    ElseIf Not VBA.IsObject(Prop) Then
        Set getProperty = f.dictionaries.getObject(pByNames, Prop)
    ElseIf Prop Is Nothing Then
        'Nothing should be returned
    ElseIf TypeOf Prop Is enumProperty Then
        Set getProperty = Prop
    End If
End Function


'[INTELLISENSE]
'[Common]
Public Function Caption() As enumProperty: Set Caption = ByName("Caption"): End Function
Public Function ContextMenu() As enumProperty: Set ContextMenu = ByName("ContextMenu"): End Function
Public Function CreatedDate() As enumProperty: Set CreatedDate = ByName("CreatedDate"): End Function
Public Function Date_() As enumProperty: Set Date_ = ByName("Date_"): End Function
Public Function Dates() As enumProperty: Set Dates = ByName("Dates"): End Function
Public Function DbProp() As enumProperty: Set DbProp = ByName("DbProp"): End Function
Public Function Header() As enumProperty: Set Header = ByName("Header"): End Function
Public Function Id() As enumProperty: Set Id = ByName("Id"): End Function
Public Function Index() As enumProperty: Set Index = ByName("Index"): End Function
Public Function IsActive() As enumProperty: Set IsActive = ByName("IsActive"): End Function
Public Function IsComplete() As enumProperty: Set IsComplete = ByName("IsComplete"): End Function
Public Function IsConfirmed() As enumProperty: Set IsConfirmed = ByName("IsConfirmed"): End Function
Public Function IsEditable() As enumProperty: Set IsEditable = ByName("IsEditable"): End Function
Public Function IsMissing() As enumProperty: Set IsMissing = ByName("IsMissing"): End Function
Public Function IsSelected() As enumProperty: Set IsSelected = ByName("IsSelected"): End Function
Public Function Item() As enumProperty: Set Item = ByName("Item"): End Function
Public Function Items() As enumProperty: Set Items = ByName("Items"): End Function
Public Function Key() As enumProperty: Set Key = ByName("Key"): End Function
Public Function Layout() As enumProperty: Set Layout = ByName("Layout"): End Function
Public Function Method() As enumProperty: Set Method = ByName("Method"): End Function
Public Function Month() As enumProperty: Set Month = ByName("Month"): End Function
Public Function name() As enumProperty: Set name = ByName("Name"): End Function
Public Function Parent() As enumProperty: Set Parent = ByName("Parent"): End Function
Public Function Prop() As enumProperty: Set Prop = ByName("Prop"): End Function
Public Function Separator() As enumProperty: Set Separator = ByName("Separator"): End Function
Public Function SessionId() As enumProperty: Set SessionId = ByName("SessionId"): End Function
Public Function SettingType() As enumProperty: Set SettingType = ByName("SettingType"): End Function
Public Function SettingValue() As enumProperty: Set SettingValue = ByName("SettingValue"): End Function
Public Function SourceFile() As enumProperty: Set SourceFile = ByName("SourceFile"): End Function
Public Function Status() As enumProperty: Set Status = ByName("Status"): End Function
Public Function Timestamp() As enumProperty: Set Timestamp = ByName("Timestamp"): End Function
Public Function Total() As enumProperty: Set Total = ByName("Total"): End Function
Public Function TypeId() As enumProperty: Set TypeId = ByName("TypeId"): End Function
Public Function Uuid() As enumProperty: Set Uuid = ByName("Uuid"): End Function
Public Function Version() As enumProperty: Set Version = ByName("Version"): End Function
Public Function value() As enumProperty: Set value = ByName("Value"): End Function
Public Function Year() As enumProperty: Set Year = ByName("Year"): End Function

'[Common - worksheet]
Public Function Range() As enumProperty: Set Range = ByName("Range"): End Function
Public Function RangeName() As enumProperty: Set RangeName = ByName("RangeName"): End Function
Public Function RowHeight() As enumProperty: Set RowHeight = ByName("RowHeight"): End Function
Public Function Sheet() As enumProperty: Set Sheet = ByName("Sheet"): End Function
Public Function SheetName() As enumProperty: Set SheetName = ByName("SheetName"): End Function

'[Common - config]
Public Function Rows() As enumProperty: Set Rows = ByName("Rows"): End Function
Public Function Columns() As enumProperty: Set Columns = ByName("Columns"): End Function
Public Function Styles() As enumProperty: Set Styles = ByName("Styles"): End Function
Public Function StylesConditional() As enumProperty: Set StylesConditional = ByName("StylesConditional"): End Function
Public Function Events() As enumProperty: Set Events = ByName("Events"): End Function
Public Function HeaderEvents() As enumProperty: Set HeaderEvents = ByName("HeaderEvents"): End Function
Public Function Template() As enumProperty: Set Template = ByName("Template"): End Function

'[Common - files]
Public Function filepath() As enumProperty: Set filepath = ByName("FilePath"): End Function
Public Function FileType() As enumProperty: Set FileType = ByName("FileType"): End Function
Public Function FileName() As enumProperty: Set FileName = ByName("FileName"): End Function
Public Function FileSize() As enumProperty: Set FileSize = ByName("FileSize"): End Function
Public Function FileTimestamp() As enumProperty: Set FileTimestamp = ByName("FileTimestamp"): End Function
Public Function FileCreator() As enumProperty: Set FileCreator = ByName("FileCreator"): End Function
Public Function SourceFilePath() As enumProperty: Set SourceFilePath = ByName("SourceFilePath"): End Function
Public Function DestinationFilePath() As enumProperty: Set DestinationFilePath = ByName("DestinationFilePath"): End Function

'[Common - DB]
Public Function DbUploadRecordsSeparator() As enumProperty: Set DbUploadRecordsSeparator = ByName("DbUploadRecordsSeparator"): End Function
Public Function DbUploadFieldsSeparator() As enumProperty: Set DbUploadFieldsSeparator = ByName("DbUploadFieldsSeparator"): End Function

'[Common - styling]
Public Function IndexFrom() As enumProperty: Set IndexFrom = ByName("IndexFrom"): End Function
Public Function IndexTo() As enumProperty: Set IndexTo = ByName("IndexTo"): End Function
Public Function RowsFirst() As enumProperty: Set RowsFirst = ByName("RowsFirst"): End Function
Public Function HeaderLast() As enumProperty: Set HeaderLast = ByName("HeaderLast"): End Function
Public Function Subranges() As enumProperty: Set Subranges = ByName("Subranges"): End Function

'[Dates range]
Public Function StartDate() As enumProperty: Set StartDate = ByName("StartDate"): End Function
Public Function EndDate() As enumProperty: Set EndDate = ByName("EndDate"): End Function
Public Function StartMonth() As enumProperty: Set StartMonth = ByName("StartMonth"): End Function
Public Function EndMonth() As enumProperty: Set EndMonth = ByName("EndMonth"): End Function

'[Aliases]
Public Function Alias() As enumProperty: Set Alias = ByName("Alias"): End Function
Public Function Aliases() As enumProperty: Set Aliases = ByName("Aliases"): End Function
Public Function AliasesString() As enumProperty: Set AliasesString = ByName("AliasesString"): End Function

