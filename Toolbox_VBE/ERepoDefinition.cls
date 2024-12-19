VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ERepoDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "ERepoDefinition"
'----------------------------------------------------------------------------------------------------------
Private Const EN_ADDING_ALLOWED As String = "RANGE_ADDING_ALLOWED"
Private Const EN_DELETING_ALLOWED As String = "RANGE_DELETING_ALLOWED"
Private Const EN_SORTING_ALLOWED As String = "RANGE_SORTING_ALLOWED"
Private Const EN_FILTERING_ALLOWED As String = "RANGE_FILTERING_ALLOWED"
Private Const EN_ENTITY_ALIASABLE_ALLOWED As String = "RANGE_ENTITY_ALIASABLE"
Private Const EN_ENTITY_EDITABLE_ALLOWED As String = "RANGE_ENTITY_EDITABLE"
Private Const EN_ENTITY_NAME As String = "RANGE_ENTITY_NAME"
Private Const EN_REPO_NAME As String = "RANGE_REPO_NAME"
Private Const EN_REPO_ID_FIELD As String = "RANGE_REPO_ID_FIELD"
Private Const EN_PROPERTIES As String = "RANGE_PROPERTIES"
Private Const EN_READ_TABLE As String = "RANGE_READ_TABLE"
Private Const EN_WRITE_TABLE As String = "RANGE_WRITE_TABLE"
Private Const EN_ID_FIELD As String = "RANGE_ID_FIELD"
'---
Private Const ALIASES_PREFIX As String = "Aliases"
'----------------------------------------------------------------------------------------------------------
'[Main]
Private pName As String
Private pAddingAllowed As Boolean
Private pDeletingAllowed As Boolean
Private pSortingAllowed As Boolean
Private pFilteringAllowed As Boolean
'[Db]
Private pReadTable As String
Private pWriteTable As String
Private pIdField As String
'[Entity]
Private pEntityName As String
Private pEntityAliasable As Boolean
Private pEntityEditable As Boolean
Private pProperties As VBA.Collection
'----------------------------------------------------------------------------------------------------------



Public Sub readFromSheet(wks As Excel.Worksheet)
    '[Main]
    pName = readValueFromSheet(wks, EN_REPO_NAME)
    pAddingAllowed = readValueFromSheet(wks, EN_ADDING_ALLOWED)
    pDeletingAllowed = readValueFromSheet(wks, EN_DELETING_ALLOWED)
    pSortingAllowed = readValueFromSheet(wks, EN_SORTING_ALLOWED)
    pFilteringAllowed = readValueFromSheet(wks, EN_FILTERING_ALLOWED)
    '[Db]
    pReadTable = readValueFromSheet(wks, EN_READ_TABLE)
    pWriteTable = readValueFromSheet(wks, EN_WRITE_TABLE)
    pIdField = readValueFromSheet(wks, EN_ID_FIELD)
    '[Entity]
    pEntityName = readValueFromSheet(wks, EN_ENTITY_NAME)
    pEntityAliasable = readValueFromSheet(wks, EN_ENTITY_ALIASABLE_ALLOWED)
    pEntityEditable = readValueFromSheet(wks, EN_ENTITY_EDITABLE_ALLOWED)
    Call loadProperties(wks)
End Sub

Private Function readValueFromSheet(wks As Excel.Worksheet, rangeName As String) As Variant
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    Set rng = F.Sheets.getNamedRange(wks, rangeName)
    If F.Ranges.IsValid(rng) Then
        readValueFromSheet = rng.Value2
    End If
    
End Function

Private Sub loadProperties(wks As Excel.Worksheet)
    Dim rng As Excel.Range
    Dim arr As Variant
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    Dim name As String
    Dim dataType As String
    Dim propName As String
    Dim dbField As String
    Dim linkedRepo As String
    Dim extraProps As String
    Dim prop As EProperty
    '------------------------------------------------------------------------------------------------------
    
    Set pProperties = New VBA.Collection
    
    Set rng = F.Sheets.getNamedRange(wks, EN_PROPERTIES)
    If F.Ranges.IsValid(rng) Then
        arr = F.arrays.fromRange(rng.CurrentRegion, False)
        
        If F.arrays.isDefined(arr) Then
            For i = LBound(arr) + 1 To UBound(arr)
                name = arr(i, LBound(arr, 2))
                dataType = arr(i, LBound(arr, 2) + 1)
                propName = arr(i, LBound(arr, 2) + 2)
                dbField = arr(i, LBound(arr, 2) + 3)
                linkedRepo = arr(i, LBound(arr, 2) + 4)
                extraProps = arr(i, LBound(arr, 2) + 5)
                Set prop = New EProperty
                Call prop. _
                        setRepo(Me). _
                        setName(name). _
                        setType(dataType). _
                        setPropName(propName). _
                        setDbField(dbField). _
                        setLinkedRepo(linkedRepo). _
                        setExtraProps(extraProps)
                Call pProperties.Add(prop)
            Next i
        End If
        
    End If

End Sub



'[GETTERS]
Public Function getName() As String
    getName = pName
End Function

Public Function getAddingAllowed() As Boolean
    getAddingAllowed = pAddingAllowed
End Function

Public Function getDeletingAllowed() As Boolean
    getDeletingAllowed = pDeletingAllowed
End Function

Public Function getSortingAllowed() As Boolean
    getSortingAllowed = pSortingAllowed
End Function

Public Function getFilteringAllowed() As Boolean
    getFilteringAllowed = pFilteringAllowed
End Function

Public Function getEntityName() As String
    getEntityName = pEntityName
End Function

Public Function getEntityClassName() As String
    getEntityClassName = "E" & F.Strings.convertLetterCasing(pEntityName, LetterCasing_Sentence)
End Function

Public Function getRepoClassName() As String
    getRepoClassName = "Repo_" & pName
End Function

Public Function isEntityAliasable() As Boolean
    isEntityAliasable = pEntityAliasable
End Function

Public Function isEntityEditable() As Boolean
    isEntityEditable = pEntityEditable
End Function

Public Function getReadTable() As String
    getReadTable = pReadTable
End Function

Public Function getWriteTable() As String
    getWriteTable = pWriteTable
End Function

Public Function getIdField() As String
    getIdField = pIdField
End Function

Public Function getAliasName() As String
    getAliasName = F.Strings.toSentenceCase(pName) & ALIASES_PREFIX
End Function

Public Function getAliasReadTable() As String
    getAliasReadTable = appendAliasesPrefixToTableName(pWriteTable)
End Function

Public Function getAliasWriteTable() As String
    getAliasWriteTable = appendAliasesPrefixToTableName(pWriteTable)
End Function

Public Function hasNameProperty() As Boolean
    Dim prop As EProperty
    '------------------------------------------------------------------------------------------------------
    For Each prop In pProperties
        If prop.isNameProperty Then
            hasNameProperty = True
            Exit For
        End If
    Next prop
End Function

Public Function hasDateProperty() As Boolean
    Dim prop As EProperty
    '------------------------------------------------------------------------------------------------------
    For Each prop In pProperties
        If prop.isDate Then
            hasDateProperty = True
            Exit For
        End If
    Next prop
End Function

Public Function getNameProperty() As String
    Dim prop As EProperty
    '------------------------------------------------------------------------------------------------------
    For Each prop In pProperties
        If prop.isNameProperty Then
            getNameProperty = prop.getPropName
            Exit For
        End If
    Next prop
End Function

Private Function appendAliasesPrefixToTableName(fullTableName As String) As String
    Dim tableName As String
    '------------------------------------------------------------------------------------------------------
    If F.regex.checkIfMatch(fullTableName, "^\w+$") Then
        appendAliasesPrefixToTableName = ALIASES_PREFIX & F.Strings.toSentenceCase(fullTableName)
    ElseIf F.regex.checkIfMatch(fullTableName, DBO_TABLE_REGEX_PATTERN) Then
        tableName = F.regex.getFirstGroupMatch(pWriteTable, DBO_TABLE_REGEX_PATTERN)
        appendAliasesPrefixToTableName = F.Strings.Format(DBO_TABLE_BUILD_PATTERN, _
                                        ALIASES_PREFIX & F.Strings.toSentenceCase(tableName))
    End If
End Function

Public Function getProps() As VBA.Collection
    Set getProps = pProperties
End Function

Public Function getPropsDictionary() As Scripting.Dictionary
    Dim prop As EProperty
    '------------------------------------------------------------------------------------------------------
    
    Set getPropsDictionary = F.dictionaries.Create(False)
    
    For Each prop In pProperties
        Call F.dictionaries.addItem(getPropsDictionary, prop.getPropName, prop.getDbField, duplicateBehaviour_Skip)
    Next prop
    
End Function
