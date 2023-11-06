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

    '[Common - basic]
    Call addEnumOption("Alias", "Alias")
    Call addEnumOption("Id", "id")
    Call addEnumOption("Key", "key")
    Call addEnumOption("Name", "name")
    Call addEnumOption("Uuid", "uuid")
    Call addEnumOption("Value")

    '[Common - session]
    Call addEnumOption("DevOnly")
    Call addEnumOption("SessionId", "session_id")
    
    '[Common - application]
    Call addEnumOption("SettingType", "type")
    Call addEnumOption("SettingValue", "value")
    Call addEnumOption("Version", "version")

    '[Common - processing]
    Call addEnumOption("DbProp")
    Call addEnumOption("Index", "index")
    Call addEnumOption("Item")
    Call addEnumOption("Items")
    Call addEnumOption("Message")
    Call addEnumOption("Method")
    Call addEnumOption("Parent")
    Call addEnumOption("Prop")

    '[Common - status]
    Call addEnumOption("IsActive", "is_active")
    Call addEnumOption("IsComplete", "is_complete")
    Call addEnumOption("IsConfirmed", "is_confirmed")
    Call addEnumOption("IsEditable")
    Call addEnumOption("IsFixed")
    Call addEnumOption("IsIncluded", "is_included")
    Call addEnumOption("IsMissing", "is_missing")
    Call addEnumOption("IsSelected")
    Call addEnumOption("Status", "status")
    
    '[Common - date & time]
    Call addEnumOption("Date_")
    Call addEnumOption("Dates")
    Call addEnumOption("EndDate", "EndDate")
    Call addEnumOption("EndMonth", "EndMonth")
    Call addEnumOption("Month")
    Call addEnumOption("StartDate", "StartDate")
    Call addEnumOption("StartMonth", "StartMonth")
    Call addEnumOption("Timestamp")
    Call addEnumOption("Year")
    
    '[Common - user forms]
    Call addEnumOption("Caption")
    Call addEnumOption("ContextMenu")
    Call addEnumOption("Header", "header")
    Call addEnumOption("Layout")
    Call addEnumOption("WindowTitle")

    '[Common - sheets & ranges]
    Call addEnumOption("Column")
    Call addEnumOption("Columns")
    Call addEnumOption("Range")
    Call addEnumOption("RangeName")
    Call addEnumOption("Row")
    Call addEnumOption("Rows")
    Call addEnumOption("Shape")
    Call addEnumOption("Sheet")
    Call addEnumOption("SheetName")
    
    '[Common - Styling]
    Call addEnumOption("Styles")
    Call addEnumOption("StylesConditional")
    Call addEnumOption("Events")
    Call addEnumOption("HeaderEvents")
    Call addEnumOption("Template")
    
    '[Common - file types & paths]
    Call addEnumOption("Destinationfilepath", "destination_file_path")
    Call addEnumOption("FileCreator", "file_creator")
    Call addEnumOption("FileName", "file_name")
    Call addEnumOption("filepath", "file_path")
    Call addEnumOption("FileSize", "file_size")
    Call addEnumOption("FileTimestamp", "file_timestamp")
    Call addEnumOption("FileType", "file_type")
    Call addEnumOption("FolderPath", "folder_path")
    Call addEnumOption("Sourcefilepath", "source_file_path")

    '[Common - Developers]
    Call addEnumOption("ProjectName", "project_name")
    Call addEnumOption("ProjectCodeName", "project_code_name")
    Call addEnumOption("References", "references")
    Call addEnumOption("RefName", "ref_name")
    Call addEnumOption("Reffilepath", "ref_file_path")

End Sub

Private Sub addEnumOption(name As String, Optional ByVal dbField As String)
    Dim enumProperty As enumProperty
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(dbField) = 0 Then dbField = name
    Set enumProperty = f.Create.Prop(name, dbField)
    
    Call f.dictionaries.addItem(pByNames, _
                                    enumProperty.getName, _
                                    enumProperty, _
                                    duplicateBehaviour_WarningInImmediateWindow)
    
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

'[Common - basic]
Public Function Alias() As enumProperty: Set Alias = ByName("Alias"): End Function
Public Function Id() As enumProperty: Set Id = ByName("Id"): End Function
Public Function key() As enumProperty: Set key = ByName("Key"): End Function
Public Function name() As enumProperty: Set name = ByName("Name"): End Function
Public Function Uuid() As enumProperty: Set Uuid = ByName("Uuid"): End Function
Public Function value() As enumProperty: Set value = ByName("Value"): End Function

'[Common - session]
Public Function DevOnly() As enumProperty: Set DevOnly = ByName("DevOnly"): End Function
Public Function SessionId() As enumProperty: Set SessionId = ByName("SessionId"): End Function
    
'[Common - application]
Public Function SettingType() As enumProperty: Set SettingType = ByName("SettingType"): End Function
Public Function SettingValue() As enumProperty: Set SettingValue = ByName("SettingValue"): End Function
Public Function Version() As enumProperty: Set Version = ByName("Version"): End Function

'[Common - processing]
Public Function DbProp() As enumProperty: Set DbProp = ByName("DbProp"): End Function
Public Function Index() As enumProperty: Set Index = ByName("Index"): End Function
Public Function Item() As enumProperty: Set Item = ByName("Item"): End Function
Public Function Items() As enumProperty: Set Items = ByName("Items"): End Function
Public Function message() As enumProperty: Set message = ByName("Message"): End Function
Public Function Method() As enumProperty: Set Method = ByName("Method"): End Function
Public Function Parent() As enumProperty: Set Parent = ByName("Parent"): End Function
Public Function Prop() As enumProperty: Set Prop = ByName("Prop"): End Function

'[Common - status]
Public Function IsActive() As enumProperty: Set IsActive = ByName("IsActive"): End Function
Public Function IsComplete() As enumProperty: Set IsComplete = ByName("IsComplete"): End Function
Public Function IsConfirmed() As enumProperty: Set IsConfirmed = ByName("IsConfirmed"): End Function
Public Function IsEditable() As enumProperty: Set IsEditable = ByName("IsEditable"): End Function
Public Function IsFixed() As enumProperty: Set IsFixed = ByName("IsFixed"): End Function
Public Function IsIncluded() As enumProperty: Set IsIncluded = ByName("IsIncluded"): End Function
Public Function IsMissing() As enumProperty: Set IsMissing = ByName("IsMissing"): End Function
Public Function IsSelected() As enumProperty: Set IsSelected = ByName("IsSelected"): End Function
Public Function Status() As enumProperty: Set Status = ByName("Status"): End Function
    
'[Common - date & time]
Public Function Date_() As enumProperty: Set Date_ = ByName("Date_"): End Function
Public Function Dates() As enumProperty: Set Dates = ByName("Dates"): End Function
Public Function EndDate() As enumProperty: Set EndDate = ByName("EndDate"): End Function
Public Function EndMonth() As enumProperty: Set EndMonth = ByName("EndMonth"): End Function
Public Function Month() As enumProperty: Set Month = ByName("Month"): End Function
Public Function StartDate() As enumProperty: Set StartDate = ByName("StartDate"): End Function
Public Function StartMonth() As enumProperty: Set StartMonth = ByName("StartMonth"): End Function
Public Function Timestamp() As enumProperty: Set Timestamp = ByName("Timestamp"): End Function
Public Function Year() As enumProperty: Set Year = ByName("Year"): End Function

'[Common - user forms]
Public Function Caption() As enumProperty: Set Caption = ByName("Caption"): End Function
Public Function ContextMenu() As enumProperty: Set ContextMenu = ByName("ContextMenu"): End Function
Public Function Header() As enumProperty: Set Header = ByName("Header"): End Function
Public Function Layout() As enumProperty: Set Layout = ByName("Layout"): End Function
Public Function WindowTitle() As enumProperty: Set WindowTitle = ByName("WindowTitle"): End Function

'[Common - sheets & ranges]
Public Function Column() As enumProperty: Set Column = ByName("Column"): End Function
Public Function Columns() As enumProperty: Set Columns = ByName("Columns"): End Function
Public Function Range() As enumProperty: Set Range = ByName("Range"): End Function
Public Function RangeName() As enumProperty: Set RangeName = ByName("RangeName"): End Function
Public Function Row() As enumProperty: Set Row = ByName("Row"): End Function
Public Function Rows() As enumProperty: Set Rows = ByName("Rows"): End Function

Public Function Shape() As enumProperty: Set Shape = ByName("Shape"): End Function
Public Function Sheet() As enumProperty: Set Sheet = ByName("Sheet"): End Function
Public Function SheetName() As enumProperty: Set SheetName = ByName("SheetName"): End Function

'[Common - Styling]
Public Function Styles() As enumProperty: Set Styles = ByName("Styles"): End Function
Public Function StylesConditional() As enumProperty: Set StylesConditional = ByName("StylesConditional"): End Function
Public Function Events() As enumProperty: Set Events = ByName("Events"): End Function
Public Function HeaderEvents() As enumProperty: Set HeaderEvents = ByName("HeaderEvents"): End Function
Public Function Template() As enumProperty: Set Template = ByName("Template"): End Function

'[Common - file types & paths]
Public Function Destinationfilepath() As enumProperty: Set Destinationfilepath = ByName("Destinationfilepath"): End Function
Public Function FileCreator() As enumProperty: Set FileCreator = ByName("FileCreator"): End Function
Public Function FileName() As enumProperty: Set FileName = ByName("FileName"): End Function
Public Function filepath() As enumProperty: Set filepath = ByName("filepath"): End Function
Public Function FileSize() As enumProperty: Set FileSize = ByName("FileSize"): End Function
Public Function FileTimestamp() As enumProperty: Set FileTimestamp = ByName("FileTimestamp"): End Function
Public Function FileType() As enumProperty: Set FileType = ByName("FileType"): End Function
Public Function folderPath() As enumProperty: Set folderPath = ByName("FolderPath"): End Function
Public Function Sourcefilepath() As enumProperty: Set Sourcefilepath = ByName("Sourcefilepath"): End Function

'[Common - Developers]
Public Function ProjectName() As enumProperty: Set ProjectName = ByName("ProjectName"): End Function
Public Function ProjectCodeName() As enumProperty: Set ProjectCodeName = ByName("ProjectCodeName"): End Function
Public Function References() As enumProperty: Set References = ByName("References"): End Function
Public Function RefName() As enumProperty: Set RefName = ByName("RefName"): End Function
Public Function Reffilepath() As enumProperty: Set Reffilepath = ByName("Reffilepath"): End Function
