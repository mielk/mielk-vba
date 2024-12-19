VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTypesGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "DataTypesGenerator"
'----------------------------------------------------------------------------------------------------------
Private Const DATA_TYPES_CLASS_NAME As String = "CDataTypes"
Private Const INSERT_METHOD_NAME As String = "insertEnumOptions"
Private Const ADD_ENUM_METHOD_NAME As String = "addEnumOption"
'[Captions]
Private Const NO_ACCESS_METHOD_WARNING As String = "VBE.Warnings.AddDataType.NoAccessMethods"
Private Const NO_ENUM_OPTION_WARNING As String = "VBE.Warnings.AddDataType.NoEnumOptions"
Private Const DUPLICATE_NAMES_WARNING As String = "VBE.Warnings.AddDataType.DuplicateNames"
Private Const INVALID_NAMES_WARNING As String = "VBE.Warnings.AddDataType.InvalidNames"
'[Settings] -----------------------------------------------------------------------------------------------
Private pDisplayAlerts As Boolean
'----------------------------------------------------------------------------------------------------------
Private pCoreProject As VBIDE.VBProject
Private pDataTypesComponent As VBIDE.VBComponent
'----------------------------------------------------------------------------------------------------------
Private pContent As String
Private pMethodBody As String
Private pExistingDataTypes As Scripting.Dictionary
Private pOneLineMethods As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    pDisplayAlerts = True
End Sub



'[SETTERS]
Public Function setDisplayAlerts(value As Boolean) As DataTypesGenerator
    Set setDisplayAlerts = Me
    pDisplayAlerts = value
End Function


Public Sub addDataTypes(project As VBIDE.VBProject, ByVal newDataTypes As VBA.Collection, _
                                                    Optional category As String = "Other")
    Dim body As String
    '----------------------------------------------------------------------------------------------------------
    
    Call setup(project)
    Call parseCodeModule
    
    If Not clearPropsCollection(newDataTypes).getStatus Then
        ' cancelled by user
    ElseIf Not comparePropsAndMethodsList.getStatus Then
        ' cancelled by user
    ElseIf F.Collections.isNonEmptyCollection(newDataTypes) Then
        Call updateInsertEnumOptionsMethod(newDataTypes, category)
        Call updateAccessMethods(newDataTypes, category)
        Call F.Developer.overrideCodeModuleContent(pDataTypesComponent.CodeModule, pContent)
    End If
    
End Sub


Private Sub setup(project As VBIDE.VBProject)
    Const REGEX_PATTERN As String = "^(\w+?)(?:_|$)"
    Const CORE_FILE_NAME_PATTERN As String = "{0}_Core"
    '------------------------------------------------------------------------------------------------------
    Dim projectPrefix As String
    Dim coreName As String
    '------------------------------------------------------------------------------------------------------
    projectPrefix = F.regex.getFirstGroupMatch(project.name, REGEX_PATTERN)
    coreName = F.Strings.Format(CORE_FILE_NAME_PATTERN, projectPrefix)
    Set pCoreProject = F.Developer.getVbProjectByProjectName(coreName)
    Set pDataTypesComponent = F.Developer.getVbComponent(pCoreProject, DATA_TYPES_CLASS_NAME)
End Sub


Private Sub parseCodeModule()
    Const REGEX_PATTERN As String = "(Private Sub $\(\)(?:.|\n)*?End Sub)"
    '------------------------------------------------------------------------------------------------------
    Dim regexPattern As String
    '------------------------------------------------------------------------------------------------------
    
    pContent = F.Developer.getCodeModuleContent(pDataTypesComponent.CodeModule)
    
    '[Loading existing enum options.] ---------------------------------------------------------------|
    regexPattern = VBA.Replace(REGEX_PATTERN, "$", INSERT_METHOD_NAME)                              '|
    pMethodBody = F.regex.getFirstGroupMatch(pContent, regexPattern)                                '|
    Call extractExistingDataTypes(pMethodBody)                                                      '|
    Call loadExistingMethods(pContent)                                                              '|
    '------------------------------------------------------------------------------------------------|
    
End Sub

Private Sub extractExistingDataTypes(methodBody As String)
    Const REGEX_PATTERN As String = "{0}\(\""(\w+)\""(?:,\s*_?\s*\""([\w\[\]\.]+)\""(?:,\s*_?\s*\""([\w\[\]\.]+)\"")?)?\)"
    '------------------------------------------------------------------------------------------------------
    Dim regexPattern As String
    Dim matches As VBA.Collection
    Dim match As Variant
    '------------------------------------------------------------------------------------------------------
    Dim dataTypeName As String
    Dim dbField As String
    '------------------------------------------------------------------------------------------------------
    
    regexPattern = F.Strings.Format(REGEX_PATTERN, ADD_ENUM_METHOD_NAME)
    Set matches = F.regex.getMatchArraysCollection(methodBody, regexPattern)
    
    Set pExistingDataTypes = F.dictionaries.Create(False)
    For Each match In matches
        dataTypeName = match(LBound(match))
        dbField = match(LBound(match) + 1)
        Call F.dictionaries.addItem(pExistingDataTypes, dataTypeName, dbField)
    Next match
    
End Sub

Private Sub loadExistingMethods(moduleBody As String)
    Const REGEX_PATTERN As String = "Public Function (\w*)\(\) As enumDataType: Set \1 = ByName\(""\1""\): End Function"
    '------------------------------------------------------------------------------------------------------
    Dim matches As VBA.Collection
    Dim match As Variant
    Dim dataTypeName As String
    '------------------------------------------------------------------------------------------------------
    
    Set matches = F.regex.getMatchArraysCollection(moduleBody, REGEX_PATTERN)
    Set pOneLineMethods = F.dictionaries.Create(False)
    
    For Each match In matches
        dataTypeName = match(LBound(match))
        Call F.dictionaries.addItem(pOneLineMethods, dataTypeName, vbNullString)
    Next match
    
End Sub

Private Function clearPropsCollection(ByRef allProps As VBA.Collection) As ValidationObject
    Dim dataTypeDef As Variant
    Dim dataTypeName As String
    '------------------------------------------------------------------------------------------------------
    Dim duplicates As VBA.Collection
    Dim invalid As VBA.Collection
    Dim col As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    Dim warnings As VBA.Collection
    Dim warning As String
    Dim confirmation As Boolean
    '------------------------------------------------------------------------------------------------------
    
    
    Set warnings = New VBA.Collection
    Set duplicates = New VBA.Collection
    Set invalid = New VBA.Collection
    Set col = New VBA.Collection
    
    
    If F.Collections.isNonEmptyCollection(allProps) Then
        For Each dataTypeDef In allProps
            dataTypeName = F.arrays.getFirstItem(dataTypeDef)
            
            If Not isValidDataTypeDef(dataTypeDef) Then
                Call invalid.Add(dataTypeName)
            ElseIf pExistingDataTypes.Exists(dataTypeName) Then
                Call duplicates.Add(dataTypeName)
            Else
                Call col.Add(dataTypeDef)
            End If
            
        Next dataTypeDef
    End If
    
    
    If F.Collections.isNonEmptyCollection(invalid) Then
        warning = F.Strings.Format(msg.getText(INVALID_NAMES_WARNING), F.Collections.toString(invalid, , VBA.vbCrLf & VBA.vbTab & "- "))
        Call warnings.Add(warning)
    End If
    
    If F.Collections.isNonEmptyCollection(duplicates) Then
        warning = F.Strings.Format(msg.getText(DUPLICATE_NAMES_WARNING), F.Collections.toString(duplicates, , VBA.vbCrLf & VBA.vbTab & "- "))
        Call warnings.Add(warning)
    End If
    
    
    If F.Collections.isNonEmptyCollection(warnings) Then
        If Not pDisplayAlerts Then
            confirmation = True
        Else
            confirmation = F.Dialogs.getUserConfirmation(F.Collections.toString(warnings, , VBA.vbCrLf & VBA.vbCrLf & VBA.vbCrLf))
        End If
        Debug.Print F.Collections.toString(warnings, , VBA.vbCrLf)
    Else
        confirmation = True
    End If
    
    
    Set allProps = col
    Set clearPropsCollection = F.Create.ValidationObject(confirmation)
    
    
End Function


Private Function comparePropsAndMethodsList() As ValidationObject
    Dim noAccessMethods As Scripting.Dictionary
    Dim noEnumOptions As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim warnings As VBA.Collection
    Dim warning As String
    '------------------------------------------------------------------------------------------------------
    
    Set warnings = New VBA.Collection
    Set noAccessMethods = F.dictionaries.getElementsWithoutMatch(pExistingDataTypes, pOneLineMethods)
    Set noEnumOptions = F.dictionaries.getElementsWithoutMatch(pOneLineMethods, pExistingDataTypes)
    
    '[Props without enum option]
    If F.dictionaries.isNonEmptyDictionary(noEnumOptions) Then
        warning = F.Strings.Format(msg.getText(NO_ENUM_OPTION_WARNING), _
                    F.dictionaries.toString(noEnumOptions, DictPart_KeyOnly, , VBA.vbCrLf, openWrapper:=" - "))
        Call warnings.Add(warning)
    End If
    
    '[Props without one-line access method]
    If F.dictionaries.isNonEmptyDictionary(noAccessMethods) Then
        warning = F.Strings.Format(msg.getText(NO_ACCESS_METHOD_WARNING), _
                    F.dictionaries.toString(noAccessMethods, DictPart_KeyOnly, , VBA.vbCrLf, openWrapper:=" - "))
        Call warnings.Add(warning)
    End If
    
    
    If F.Collections.isNonEmptyCollection(warnings) Then
        If Not pDisplayAlerts Then
            Set comparePropsAndMethodsList = F.Create.ValidationObject(True)
        Else
            warning = F.Collections.toString(warnings, , VBA.vbCrLf & VBA.vbCrLf)
            Call VBA.MsgBox(warning, vbExclamation)
            Set comparePropsAndMethodsList = F.Create.ValidationObject(False, warning)
        End If
    Else
        Set comparePropsAndMethodsList = F.Create.ValidationObject(True)
    End If
    
    
End Function

Private Function isValidDataTypeDef(dataTypeDef As Variant) As Boolean
    Const NAME_REGEX_PATTERN As String = "^\w+$"
    Const DB_TABLE_NAME_REGEX_PATTERN As String = "^(?:(?:\[dbo\]\.)?\[\w+\]|\w+)$"
    '------------------------------------------------------------------------------------------------------
    Dim name As String
    Dim writeTableName As String
    Dim readTableName As String
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If F.arrays.isDefined(dataTypeDef) Then
        name = dataTypeDef(LBound(dataTypeDef))
        writeTableName = dataTypeDef(LBound(dataTypeDef) + 1)
        readTableName = dataTypeDef(LBound(dataTypeDef) + 2)
        
        If Not F.regex.checkIfMatch(name, NAME_REGEX_PATTERN) Then
            isValidDataTypeDef = False
        ElseIf Not F.regex.checkIfMatch(writeTableName, DB_TABLE_NAME_REGEX_PATTERN) Then
            isValidDataTypeDef = False
        ElseIf Not F.regex.checkIfMatch(readTableName, DB_TABLE_NAME_REGEX_PATTERN) Then
            isValidDataTypeDef = False
        Else
            isValidDataTypeDef = True
        End If
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    isValidDataTypeDef = False
    GoTo ExitPoint
    
End Function




Private Sub updateInsertEnumOptionsMethod(newProps As VBA.Collection, category As String)
    Dim newMethodBody As String
    '------------------------------------------------------------------------------------------------------
    Const REGEX_PATTERN As String = "(Private Sub $\(\)(?:.|\n)*?End Sub)"
    '------------------------------------------------------------------------------------------------------
    Dim regexPattern As String
    '------------------------------------------------------------------------------------------------------
    
    newMethodBody = getModifiedInsertEnumOptionsMethod(newProps, category)
    regexPattern = VBA.Replace(REGEX_PATTERN, "$", INSERT_METHOD_NAME)
    pContent = F.regex.Replace(pContent, regexPattern, newMethodBody)

End Sub

Private Function getModifiedInsertEnumOptionsMethod(newProps As VBA.Collection, category As String) As String
    Const REGEX_PATTERN_FIND_CATEGORY As String = "^[\s\t]*'\[{0}\]"
    Const REGEX_PATTERN_METHOD_CATEGORY_GROUP As String = "('\[{0}\](?:[\s<<NEW_LINE>>]*Call addEnumOption\(""\w+""(?:, ""\w+"")?\))*)"
    Const REGEX_PATTERN_UNKNOWN_CATEGORY As String = "(\nEnd Sub$)"
    Const CODE_LINE_PATTERN As String = "\tCall {0}(""{1}"", ""{2}"", ""{3}"")"
    Const DB_FIELD_PART_PATTERN As String = ", ""{0}"""
    Const CATEGORY_LABEL_PATTERN As String = "\t'[{0}]\n"
    '------------------------------------------------------------------------------------------------------
    Dim regexPattern As String
    Dim categoryFound As Boolean
    '------------------------------------------------------------------------------------------------------
    Dim lines As VBA.Collection
    Dim line As String
    Dim prop As Variant
    Dim dataTypeName As String
    Dim readTableName As String
    Dim writeTableName As String
    '------------------------------------------------------------------------------------------------------
    Dim replacementText As String
    '------------------------------------------------------------------------------------------------------
    
    
    '[Check if given category already exists in this method.] ---------------------------------------|
    regexPattern = F.Strings.Format(REGEX_PATTERN_FIND_CATEGORY, category)                          '|
    categoryFound = F.regex.checkIfMatch(pMethodBody, regexPattern)                                 '|
    '------------------------------------------------------------------------------------------------|
    
    
    '[Create collection of exact code lines to be inserted.] ----------------------------------------|
    Set lines = New VBA.Collection                                                                  '|
    For Each prop In newProps                                                                       '|
        dataTypeName = prop(LBound(prop))                                                           '|
        readTableName = prop(LBound(prop) + 1)                                                      '|
        writeTableName = prop(LBound(prop) + 2)                                                     '|
        If Not pExistingDataTypes.Exists(dataTypeName) Then                                         '|
            line = F.Strings.Format(CODE_LINE_PATTERN, ADD_ENUM_METHOD_NAME, _
                                    dataTypeName, readTableName, writeTableName)                    '|
            Call lines.Add(line)                                                                    '|
        End If                                                                                      '|
    Next prop                                                                                       '|
    '------------------------------------------------------------------------------------------------|
    
    
    If F.Collections.isNonEmptyCollection(lines) Then
        replacementText = VBA.IIf(categoryFound, vbNullString, F.Strings.Format(CATEGORY_LABEL_PATTERN, category)) & _
                            F.Collections.toString(lines, , VBA.vbCrLf)
    End If
    
    

    '[Create final string]
    If categoryFound Then
        regexPattern = VBA.Replace(F.Strings.Format(REGEX_PATTERN_METHOD_CATEGORY_GROUP, category), _
                                    NEW_LINE_REPLACEMENT_TAG, NEW_LINE_SYMBOL)
        getModifiedInsertEnumOptionsMethod = F.regex.Replace(pMethodBody, regexPattern, "$1" & VBA.vbCrLf & replacementText)
    Else
        regexPattern = REGEX_PATTERN_UNKNOWN_CATEGORY
        getModifiedInsertEnumOptionsMethod = F.regex.Replace(pMethodBody, regexPattern, replacementText & VBA.vbCrLf & "$1")
    End If
    
End Function



Private Sub updateAccessMethods(newProps As VBA.Collection, category As String)
    Dim newMethodBody As String
    '------------------------------------------------------------------------------------------------------
    Const REGEX_PATTERN_METHOD_CATEGORY_GROUP As String = _
                    "(^\s*'\[{0}\](?:[\s<<NEW_LINE>>]*" & _
                    "Public Function \w*\(\) As enumDataType: Set \w* = ByName\(""\w*""\): End Function)+)"
    Const CODE_LINE_PATTERN As String = _
                    "Public Function {0}() As enumDataType: Set {0} = ByName(""{0}""): End Function"
    Const CATEGORY_LABEL_PATTERN As String = "'[{0}]\n"
    '------------------------------------------------------------------------------------------------------
    Dim regexPattern As String
    Dim categoryFound As Boolean
    '------------------------------------------------------------------------------------------------------
    Dim lines As VBA.Collection
    Dim line As String
    Dim prop As Variant
    Dim dataTypeName As String
    '------------------------------------------------------------------------------------------------------
    Dim replacementText As String
    '------------------------------------------------------------------------------------------------------
    
    '[Check if given category already exists in this method.] ---------------------------------------|
    regexPattern = VBA.Replace(F.Strings.Format(REGEX_PATTERN_METHOD_CATEGORY_GROUP, category), _
                                                    NEW_LINE_REPLACEMENT_TAG, NEW_LINE_SYMBOL)      '|
    categoryFound = F.regex.checkIfMatch(pContent, regexPattern)                                    '|
    '------------------------------------------------------------------------------------------------|


    '[Create collection of exact code lines to be inserted.] ----------------------------------------|
    Set lines = New VBA.Collection                                                                  '|
    For Each prop In newProps                                                                       '|
        dataTypeName = prop(LBound(prop))                                                           '|
        If Not pExistingDataTypes.Exists(dataTypeName) Then                                         '|
            line = F.Strings.Format(CODE_LINE_PATTERN, dataTypeName)                                '|
            Call lines.Add(line)                                                                    '|
        End If                                                                                      '|
    Next prop                                                                                       '|
    '------------------------------------------------------------------------------------------------|
    
    
    If F.Collections.isNonEmptyCollection(lines) Then
        replacementText = VBA.IIf(categoryFound, vbNullString, F.Strings.Format(CATEGORY_LABEL_PATTERN, category)) & _
                            F.Collections.toString(lines, , VBA.vbCrLf)
    End If
        
    
    '[Create final string]
    If categoryFound Then
        regexPattern = VBA.Replace(F.Strings.Format(REGEX_PATTERN_METHOD_CATEGORY_GROUP, category), _
                                    NEW_LINE_REPLACEMENT_TAG, NEW_LINE_SYMBOL)
        pContent = F.regex.Replace(pContent, regexPattern, "$1" & VBA.vbCrLf & replacementText)
    Else
        pContent = pContent & VBA.vbCrLf & VBA.vbCrLf & replacementText
    End If
    
    
End Sub
