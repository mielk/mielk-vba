VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropsGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "PropsGenerator"
'----------------------------------------------------------------------------------------------------------
Private Const PROPERTIES_CLASS_NAME As String = "CProperties"
Private Const INSERT_METHOD_NAME As String = "insertEnumOptions"
Private Const ADD_ENUM_METHOD_NAME As String = "addEnumOption"
Private Const DEFAULT_CATEGORY_NAME As String = "Other"
'[Captions]
Private Const NO_ACCESS_METHOD_WARNING As String = "VBE.Warnings.AddProps.NoAccessMethods"
Private Const NO_ENUM_OPTION_WARNING As String = "VBE.Warnings.AddProps.NoEnumOptions"
Private Const DUPLICATE_NAMES_WARNING As String = "VBE.Warnings.AddProps.DuplicateNames"
Private Const INVALID_NAMES_WARNING As String = "VBE.Warnings.AddProps.InvalidNames"
'----------------------------------------------------------------------------------------------------------
Private pCoreProject As VBIDE.VBProject
Private pPropsComponent As VBIDE.VBComponent
'----------------------------------------------------------------------------------------------------------
Private pContent As String
Private pMethodBody As String
Private pExistingProps As Scripting.Dictionary
Private pOneLineMethods As Scripting.Dictionary
'[Settings] -----------------------------------------------------------------------------------------------
Private pDisplayAlerts As Boolean
'----------------------------------------------------------------------------------------------------------


Private Sub Class_Initialize()
    pDisplayAlerts = True
End Sub


'[SETTERS]
Public Function setDisplayAlerts(value As Boolean) As PropsGenerator
    Set setDisplayAlerts = Me
    pDisplayAlerts = value
End Function



Public Sub addProps(project As VBIDE.VBProject, ByVal newProps As VBA.Collection, _
                    Optional ByVal category As String = DEFAULT_CATEGORY_NAME)
    Dim body As String
    '----------------------------------------------------------------------------------------------------------
    
    If VBA.Len(category) = 0 Then category = DEFAULT_CATEGORY_NAME
    
    Call setup(project)
    Call parseCodeModule
    
    If Not clearPropsCollection(newProps).getStatus Then
        ' cancelled by user
    ElseIf Not comparePropsAndMethodsList.getStatus Then
        ' cancelled by user
    ElseIf F.Collections.isNonEmptyCollection(newProps) Then
        Call updateInsertEnumOptionsMethod(newProps, category)
        Call updateAccessMethods(newProps, category)
        Call F.Developer.overrideCodeModuleContent(pPropsComponent.CodeModule, pContent)
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
    Set pPropsComponent = F.Developer.getVbComponent(pCoreProject, PROPERTIES_CLASS_NAME)
End Sub


Private Sub parseCodeModule()
    Const REGEX_PATTERN As String = "(Private Sub $\(\)(?:.|\n)*?End Sub)"
    '------------------------------------------------------------------------------------------------------
    Dim regexPattern As String
    '------------------------------------------------------------------------------------------------------
    
    pContent = F.Developer.getCodeModuleContent(pPropsComponent.CodeModule)
    
    '[Loading existing enum options.] ---------------------------------------------------------------|
    regexPattern = VBA.Replace(REGEX_PATTERN, "$", INSERT_METHOD_NAME)                              '|
    pMethodBody = F.regex.getFirstGroupMatch(pContent, regexPattern)                                '|
    Call extractExistingProps(pMethodBody)                                                          '|
    Call loadExistingMethods(pContent)                                                              '|
    '------------------------------------------------------------------------------------------------|
    
End Sub

Private Sub extractExistingProps(methodBody As String)
    Const REGEX_PATTERN As String = "{0}\(\""(\w+)\""(?:, \""(\w+)\"")?\)"
    '------------------------------------------------------------------------------------------------------
    Dim regexPattern As String
    Dim matches As VBA.Collection
    Dim match As Variant
    '------------------------------------------------------------------------------------------------------
    Dim propName As String
    Dim dbField As String
    '------------------------------------------------------------------------------------------------------
    
    regexPattern = F.Strings.Format(REGEX_PATTERN, ADD_ENUM_METHOD_NAME)
    Set matches = F.regex.getMatchArraysCollection(methodBody, regexPattern)
    
    Set pExistingProps = F.dictionaries.Create(False)
    For Each match In matches
        propName = match(LBound(match))
        dbField = match(LBound(match) + 1)
        Call F.dictionaries.addItem(pExistingProps, propName, dbField)
    Next match
    
End Sub

Private Sub loadExistingMethods(moduleBody As String)
    Const REGEX_PATTERN As String = "Public Function (\w*)\(\) As enumProperty: Set \1 = ByName\(""\1""\): End Function"
    '------------------------------------------------------------------------------------------------------
    Dim matches As VBA.Collection
    Dim match As Variant
    Dim propName As String
    '------------------------------------------------------------------------------------------------------
    
    Set matches = F.regex.getMatchArraysCollection(moduleBody, REGEX_PATTERN)
    Set pOneLineMethods = F.dictionaries.Create(False)
    
    For Each match In matches
        propName = match(LBound(match))
        Call F.dictionaries.addItem(pOneLineMethods, propName, vbNullString)
    Next match
    
End Sub

Private Function comparePropsAndMethodsList() As ValidationObject
    Dim noAccessMethods As Scripting.Dictionary
    Dim noEnumOptions As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim warnings As VBA.Collection
    Dim warning As String
    '------------------------------------------------------------------------------------------------------
    
    Set warnings = New VBA.Collection
    Set noAccessMethods = F.dictionaries.getElementsWithoutMatch(pExistingProps, pOneLineMethods)
    Set noEnumOptions = F.dictionaries.getElementsWithoutMatch(pOneLineMethods, pExistingProps)
    
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

Private Function clearPropsCollection(ByRef allProps As VBA.Collection) As ValidationObject
    Dim propDef As Variant
    Dim propName As String
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
        For Each propDef In allProps
            propName = F.arrays.getFirstItem(propDef)
            
            If Not isValidPropDef(propDef) Then
                Call invalid.Add(propName)
            ElseIf pExistingProps.Exists(propName) Then
                Call duplicates.Add(propName)
            Else
                Call col.Add(propDef)
            End If
            
        Next propDef
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

Private Function isValidPropDef(propDef As Variant) As Boolean
    Const NAME_REGEX_PATTERN As String = "^\w+$"
    Const DB_FIELD_REGEX_PATTERN As String = "^\w*$"
    '------------------------------------------------------------------------------------------------------
    Dim name As String
    Dim dbField As String
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If F.arrays.isDefined(propDef) Then
        name = propDef(LBound(propDef))
        dbField = propDef(LBound(propDef) + 1)
        
        If Not F.regex.checkIfMatch(name, NAME_REGEX_PATTERN) Then
            isValidPropDef = False
        ElseIf Not F.regex.checkIfMatch(dbField, DB_FIELD_REGEX_PATTERN) Then
            isValidPropDef = False
        Else
            isValidPropDef = True
        End If
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    isValidPropDef = False
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
    Const CODE_LINE_PATTERN As String = "\tCall {0}(""{1}""{2})"
    Const DB_FIELD_PART_PATTERN As String = ", ""{0}"""
    Const CATEGORY_LABEL_PATTERN As String = "\t'[{0}]\n"
    '------------------------------------------------------------------------------------------------------
    Dim regexPattern As String
    Dim categoryFound As Boolean
    '------------------------------------------------------------------------------------------------------
    Dim lines As VBA.Collection
    Dim line As String
    Dim prop As Variant
    Dim propName As String
    Dim dbFieldName As String
    Dim dbFieldPart As String
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
        propName = prop(LBound(prop))                                                               '|
        dbFieldName = prop(LBound(prop) + 1)                                                        '|
        If Not pExistingProps.Exists(propName) Then                                                 '|
            dbFieldPart = VBA.IIf(VBA.Len(dbFieldName), _
                                F.Strings.Format(DB_FIELD_PART_PATTERN, dbFieldName), vbNullString) '|
            line = F.Strings.Format(CODE_LINE_PATTERN, ADD_ENUM_METHOD_NAME, propName, dbFieldPart) '|
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
    
    Debug.Print getModifiedInsertEnumOptionsMethod
    
End Function



Private Sub updateAccessMethods(newProps As VBA.Collection, category As String)
    Dim newMethodBody As String
    '------------------------------------------------------------------------------------------------------
    Const REGEX_PATTERN_METHOD_CATEGORY_GROUP As String = _
                    "(^\s*'\[{0}\](?:[\s<<NEW_LINE>>]*" & _
                    "Public Function \w*\(\) As enumProperty: Set \w* = ByName\(""\w*""\): End Function)+)"
    Const CODE_LINE_PATTERN As String = _
                    "Public Function {0}() As enumProperty: Set {0} = ByName(""{0}""): End Function"
    Const CATEGORY_LABEL_PATTERN As String = "'[{0}]\n"
    '------------------------------------------------------------------------------------------------------
    Dim regexPattern As String
    Dim categoryFound As Boolean
    '------------------------------------------------------------------------------------------------------
    Dim lines As VBA.Collection
    Dim line As String
    Dim prop As Variant
    Dim propName As String
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
        propName = prop(LBound(prop))                                                               '|
        If Not pExistingProps.Exists(propName) Then                                                 '|
            line = F.Strings.Format(CODE_LINE_PATTERN, propName)                                    '|
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

