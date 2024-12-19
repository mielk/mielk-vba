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
    Call addEnumOption("Counter", "Counter")
    Call addEnumOption(DEFAULT_TAG)
    Call addEnumOption("Id", "id")
    Call addEnumOption("Key", "key")
    Call addEnumOption("Name", "name")
    Call addEnumOption("Uuid", "uuid")
    Call addEnumOption("Value")

    '[Common - session]
    Call addEnumOption("DevOnly")
    Call addEnumOption("SessionId", "SessionId")
    
    '[Common - application]
    Call addEnumOption("SettingType", "type")
    Call addEnumOption("SettingValue", "value")
    Call addEnumOption("Version")

    '[Common - processing]
    Call addEnumOption("DbProp")
    Call addEnumOption("Index")
    Call addEnumOption("IsDate")
    Call addEnumOption("Item")
    Call addEnumOption("Items")
    Call addEnumOption("Message")
    Call addEnumOption("Messages")
    Call addEnumOption("Method")
    Call addEnumOption("Parent")
    Call addEnumOption("Prop")
    Call addEnumOption("Regex")
    Call addEnumOption("Required")
    Call addEnumOption("Validation")
    Call addEnumOption("Warning")

    '[Common - status]
    Call addEnumOption("IsActive", "IsActive")
    Call addEnumOption("IsComplete", "IsComplete")
    Call addEnumOption("IsConfirmed", "IsConfirmed")
    Call addEnumOption("IsEditable")
    Call addEnumOption("IsFixed")
    Call addEnumOption("IsIncluded", "IsIncluded")
    Call addEnumOption("IsMissing", "IsMissing")
    Call addEnumOption("IsSelected")
    Call addEnumOption("IsValid")
    Call addEnumOption("Status")
    
    '[Common - date & time]
    Call addEnumOption("Date_")
    Call addEnumOption("Dates")
    Call addEnumOption("EndDate", "EndDate")
    Call addEnumOption("EndMonth", "EndMonth")
    Call addEnumOption("EndTime")
    Call addEnumOption("Month")
    Call addEnumOption("StartDate", "StartDate")
    Call addEnumOption("StartMonth", "StartMonth")
    Call addEnumOption("StartTime")
    Call addEnumOption("Timestamp")
    Call addEnumOption("WeekDay")
    Call addEnumOption("Year")
    
    '[Common - user forms]
    Call addEnumOption("Caption")
    Call addEnumOption("ContextMenu")
    Call addEnumOption("Control")
    Call addEnumOption("Header", "header")
    Call addEnumOption("Layout")
    Call addEnumOption("WindowTitle")
    
    '[Common - sheets & ranges]
    Call addEnumOption("Column")
    Call addEnumOption("ColumnIndex")
    Call addEnumOption("ColumnSymbol")          'i.e. D - for column 4, AA - for column 27, etc.
    Call addEnumOption("Columns")
    Call addEnumOption("Formula")
    Call addEnumOption("IsFormulaR1C1")
    Call addEnumOption("Range")
    Call addEnumOption("RangeName")
    Call addEnumOption("Row")
    Call addEnumOption("RowIndex")
    Call addEnumOption("Rows")
    Call addEnumOption("Shape")
    Call addEnumOption("Shapes")
    Call addEnumOption("Sheet")
    Call addEnumOption("SheetName")
    
    '[Common - DB]
    Call addEnumOption("DbUploadRecordsSeparator", "RecordsSeparator")
    Call addEnumOption("DbUploadFieldsSeparator", "FieldsSeparator")
    Call addEnumOption("DbTableName", "TableName")
    Call addEnumOption("DbFieldName", "FieldName")
    Call addEnumOption("OutputPrefix", "Prefix")
    
    '[Common - Styling]
    Call addEnumOption("Styles")
    Call addEnumOption("StylesConditional")
    Call addEnumOption("Events")
    Call addEnumOption("HeaderEvents")
    Call addEnumOption("Template")
    Call addEnumOption("Ribbon")
    
    '[Common - file types & paths]
    Call addEnumOption("DestinationFilePath", "destinationFilePath")
    Call addEnumOption("FileCreator", "fileCreator")
    Call addEnumOption("FileName", "fileName")
    Call addEnumOption("FilePath", "FilePath")
    Call addEnumOption("FileSize", "fileSize")
    Call addEnumOption("FileTimestamp", "fileTimestamp")
    Call addEnumOption("FileType", "fileType")
    Call addEnumOption("FolderPath", "folderPath")
    Call addEnumOption("SourceFilePath", "sourceFilePath")
    
    '[Common - Developers]
    Call addEnumOption("ErrorMessage")
    Call addEnumOption("ProjectName")
    Call addEnumOption("ProjectCodeName")
    Call addEnumOption("References")
    Call addEnumOption("RefName")
    Call addEnumOption("RefFilePath")
    
    '[Common - events handling]
    Call addEnumOption("KeyCode")
    Call addEnumOption("Shift")
    Call addEnumOption("Button")
    Call addEnumOption("XPos")
    Call addEnumOption("YPos")
    
    '[Common - standarization]
    Call addEnumOption("Alias", "Alias")
    Call addEnumOption("Aliases")
    Call addEnumOption("AliasesString")
    
End Sub

Private Sub addEnumOption(name As String, Optional ByVal dbField As String)
    Dim enumProperty As enumProperty
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(dbField) = 0 Then dbField = name
    Set enumProperty = f.Create.prop(name, dbField)
    
    Call f.dictionaries.addItem(pByNames, _
                                    enumProperty.getName, _
                                    enumProperty, _
                                    duplicateBehaviour_WarningInImmediateWindow)
    
End Sub



'[ACCESS]
Public Function ByName(name As String) As enumProperty
    Set ByName = f.dictionaries.getObject(pByNames, name)
End Function

Public Function getProperty(ByVal prop As Variant) As enumProperty
    If VBA.IsArray(prop) Then
        'not possible to get property
    ElseIf Not VBA.IsObject(prop) Then
        Set getProperty = f.dictionaries.getObject(pByNames, prop)
    ElseIf prop Is Nothing Then
        'Nothing should be returned
    ElseIf TypeOf prop Is enumProperty Then
        Set getProperty = prop
    End If
End Function


'[INTELLISENSE]

'[Common - basic]
Public Function Counter() As enumProperty: Set Counter = ByName("Counter"): End Function
Public Function default() As enumProperty: Set default = ByName(DEFAULT_TAG): End Function
Public Function id() As enumProperty: Set id = ByName("Id"): End Function
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
Public Function version() As enumProperty: Set version = ByName("Version"): End Function

'[Common - processing]
Public Function DbProp() As enumProperty: Set DbProp = ByName("DbProp"): End Function
Public Function index() As enumProperty: Set index = ByName("Index"): End Function
Public Function IsDate() As enumProperty: Set IsDate = ByName("IsDate"): End Function
Public Function item() As enumProperty: Set item = ByName("Item"): End Function
Public Function items() As enumProperty: Set items = ByName("Items"): End Function
Public Function message() As enumProperty: Set message = ByName("Message"): End Function
Public Function Messages() As enumProperty: Set Messages = ByName("Messages"): End Function
Public Function Method() As enumProperty: Set Method = ByName("Method"): End Function
Public Function Parent() As enumProperty: Set Parent = ByName("Parent"): End Function
Public Function prop() As enumProperty: Set prop = ByName("Prop"): End Function
Public Function regex() As enumProperty: Set regex = ByName("Regex"): End Function
Public Function Required() As enumProperty: Set Required = ByName("Required"): End Function
Public Function validation() As enumProperty: Set validation = ByName("Validation"): End Function
Public Function Warning() As enumProperty: Set Warning = ByName("Warning"): End Function

'[Common - status]
Public Function IsActive() As enumProperty: Set IsActive = ByName("IsActive"): End Function
Public Function IsComplete() As enumProperty: Set IsComplete = ByName("IsComplete"): End Function
Public Function IsConfirmed() As enumProperty: Set IsConfirmed = ByName("IsConfirmed"): End Function
Public Function IsEditable() As enumProperty: Set IsEditable = ByName("IsEditable"): End Function
Public Function IsFixed() As enumProperty: Set IsFixed = ByName("IsFixed"): End Function
Public Function IsIncluded() As enumProperty: Set IsIncluded = ByName("IsIncluded"): End Function
Public Function IsMissing() As enumProperty: Set IsMissing = ByName("IsMissing"): End Function
Public Function IsSelected() As enumProperty: Set IsSelected = ByName("IsSelected"): End Function
Public Function IsValid() As enumProperty: Set IsValid = ByName("IsValid"): End Function
Public Function Status() As enumProperty: Set Status = ByName("Status"): End Function

'[Common - date & time]
Public Function Date_() As enumProperty: Set Date_ = ByName("Date_"): End Function
Public Function Dates() As enumProperty: Set Dates = ByName("Dates"): End Function
Public Function EndDate() As enumProperty: Set EndDate = ByName("EndDate"): End Function
Public Function EndMonth() As enumProperty: Set EndMonth = ByName("EndMonth"): End Function
Public Function EndTime() As enumProperty: Set EndTime = ByName("EndTime"): End Function
Public Function month() As enumProperty: Set month = ByName("Month"): End Function
Public Function StartDate() As enumProperty: Set StartDate = ByName("StartDate"): End Function
Public Function StartMonth() As enumProperty: Set StartMonth = ByName("StartMonth"): End Function
Public Function StartTime() As enumProperty: Set StartTime = ByName("StartTime"): End Function
Public Function timestamp() As enumProperty: Set timestamp = ByName("Timestamp"): End Function
Public Function WeekDay() As enumProperty: Set WeekDay = ByName("WeekDay"): End Function
Public Function year() As enumProperty: Set year = ByName("Year"): End Function

'[Common - user forms]
Public Function caption() As enumProperty: Set caption = ByName("Caption"): End Function
Public Function ContextMenu() As enumProperty: Set ContextMenu = ByName("ContextMenu"): End Function
Public Function Control() As enumProperty: Set Control = ByName("Control"): End Function
Public Function Header() As enumProperty: Set Header = ByName("Header"): End Function
Public Function Layout() As enumProperty: Set Layout = ByName("Layout"): End Function
Public Function windowTitle() As enumProperty: Set windowTitle = ByName("WindowTitle"): End Function

'[Common - sheets & ranges]
Public Function column() As enumProperty: Set column = ByName("Column"): End Function
Public Function columnIndex() As enumProperty: Set columnIndex = ByName("columnIndex"): End Function
Public Function ColumnSymbol() As enumProperty: Set ColumnSymbol = ByName("ColumnSymbol"): End Function
Public Function columns() As enumProperty: Set columns = ByName("Columns"): End Function
Public Function formula() As enumProperty: Set formula = ByName("Formula"): End Function
Public Function isFormulaR1C1() As enumProperty: Set isFormulaR1C1 = ByName("isFormulaR1C1"): End Function
Public Function Range() As enumProperty: Set Range = ByName("Range"): End Function
Public Function rangeName() As enumProperty: Set rangeName = ByName("RangeName"): End Function
Public Function row() As enumProperty: Set row = ByName("Row"): End Function
Public Function RowIndex() As enumProperty: Set RowIndex = ByName("RowIndex"): End Function
Public Function Rows() As enumProperty: Set Rows = ByName("Rows"): End Function

Public Function shape() As enumProperty: Set shape = ByName("Shape"): End Function
Public Function shapes() As enumProperty: Set shapes = ByName("Shapes"): End Function
Public Function Sheet() As enumProperty: Set Sheet = ByName("Sheet"): End Function
Public Function sheetName() As enumProperty: Set sheetName = ByName("SheetName"): End Function

'[Common - Styling]
Public Function Styles() As enumProperty: Set Styles = ByName("Styles"): End Function
Public Function StylesConditional() As enumProperty: Set StylesConditional = ByName("StylesConditional"): End Function
Public Function Events() As enumProperty: Set Events = ByName("Events"): End Function
Public Function HeaderEvents() As enumProperty: Set HeaderEvents = ByName("HeaderEvents"): End Function
Public Function Template() As enumProperty: Set Template = ByName("Template"): End Function
Public Function Ribbon() As enumProperty: Set Ribbon = ByName("Ribbon"): End Function

'[Common - file types & paths]
Public Function DestinationFilePath() As enumProperty: Set DestinationFilePath = ByName("DestinationFilePath"): End Function
Public Function FileCreator() As enumProperty: Set FileCreator = ByName("FileCreator"): End Function
Public Function fileName() As enumProperty: Set fileName = ByName("FileName"): End Function
Public Function filepath() As enumProperty: Set filepath = ByName("FilePath"): End Function
Public Function FileSize() As enumProperty: Set FileSize = ByName("FileSize"): End Function
Public Function FileTimestamp() As enumProperty: Set FileTimestamp = ByName("FileTimestamp"): End Function
Public Function fileType() As enumProperty: Set fileType = ByName("FileType"): End Function
Public Function folderPath() As enumProperty: Set folderPath = ByName("FolderPath"): End Function
Public Function sourceFilePath() As enumProperty: Set sourceFilePath = ByName("SourceFilePath"): End Function

'[Common - Developers]
Public Function ErrorMessage() As enumProperty: Set ErrorMessage = ByName("ErrorMessage"): End Function
Public Function ProjectName() As enumProperty: Set ProjectName = ByName("ProjectName"): End Function
Public Function ProjectCodeName() As enumProperty: Set ProjectCodeName = ByName("ProjectCodeName"): End Function
Public Function References() As enumProperty: Set References = ByName("References"): End Function
Public Function RefName() As enumProperty: Set RefName = ByName("RefName"): End Function
Public Function RefFilePath() As enumProperty: Set RefFilePath = ByName("RefFilePath"): End Function

'[Common - events handling]
Public Function KeyCode() As enumProperty: Set KeyCode = ByName("KeyCode"): End Function
Public Function Shift() As enumProperty: Set Shift = ByName("Shift"): End Function
Public Function Button() As enumProperty: Set Button = ByName("Button"): End Function
Public Function XPos() As enumProperty: Set XPos = ByName("XPos"): End Function
Public Function YPos() As enumProperty: Set YPos = ByName("YPos"): End Function

'[Common - standarization]
Public Function Alias() As enumProperty: Set Alias = ByName("Alias"): End Function
Public Function Aliases() As enumProperty: Set Aliases = ByName("Aliases"): End Function
Public Function AliasesString() As enumProperty: Set AliasesString = ByName("AliasesString"): End Function
'[Common - DB]
Public Function DbUploadRecordsSeparator() As enumProperty: Set DbUploadRecordsSeparator = ByName("DbUploadRecordsSeparator"): End Function
Public Function DbUploadFieldsSeparator() As enumProperty: Set DbUploadFieldsSeparator = ByName("DbUploadFieldsSeparator"): End Function
Public Function DbTableName() As enumProperty: Set DbTableName = ByName("DbTableName"): End Function
Public Function DbFieldName() As enumProperty: Set DbFieldName = ByName("DbFieldName"): End Function
