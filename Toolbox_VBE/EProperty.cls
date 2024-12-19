VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "EProperty"
'----------------------------------------------------------------------------------------------------------
Private pRepo As ERepoDefinition
Private pName As String
Private pType As String
Private pPropName As String
Private pDbField As String
Private pLinkedRepo As String
Private pExtraProps As String
'----------------------------------------------------------------------------------------------------------





'[SETTERS]
Public Function setRepo(value As ERepoDefinition) As EProperty
    Set setRepo = Me
    Set pRepo = value
End Function

Public Function setName(value As String) As EProperty
    Set setName = Me
    pName = value
End Function

Public Function setType(value As String) As EProperty
    Set setType = Me
    pType = value
End Function

Public Function setPropName(value As String) As EProperty
    Set setPropName = Me
    pPropName = value
End Function

Public Function setDbField(value As String) As EProperty
    Set setDbField = Me
    pDbField = value
End Function

Public Function setLinkedRepo(value As String) As EProperty
    Set setLinkedRepo = Me
    pLinkedRepo = value
End Function

Public Function setExtraProps(value As String) As EProperty
    Set setExtraProps = Me
    pExtraProps = value
End Function



'[GETTERS]
Public Function getRepo() As ERepoDefinition
    Set getRepo = pRepo
End Function

Public Function getName() As String
    getName = pName
End Function

Public Function getType() As String
    getType = pType
End Function

Public Function getPropName() As String
    getPropName = pPropName
End Function

Public Function getDbField() As String
    If VBA.Len(pDbField) Then
        getDbField = pDbField
    Else
        getDbField = pPropName
    End If
End Function

Public Function getLinkedRepo() As String
    getLinkedRepo = pLinkedRepo
End Function

Public Function getExtraProps() As String
    getExtraProps = pExtraProps
End Function

Public Function isInsertedToDb() As Boolean
    isInsertedToDb = (Not isRepoId)
End Function

Public Function isRepoId() As Boolean
    isRepoId = F.Strings.compareStrings(pRepo.getIdField, pPropName)
End Function

Public Function isNumber() As Boolean
    If pType = TYPE_LONG Then
        isNumber = True
    ElseIf pType = TYPE_DOUBLE Then
        isNumber = True
    End If
End Function

Public Function isDate() As Boolean
    If pType = TYPE_DATE Then
        isDate = True
    End If
End Function

Public Function isObject() As Boolean
    isObject = Not F.variables.isPrimitiveDataType(pType)
End Function

Public Function isNameProperty() As Boolean
    Dim nameProp As String
    '------------------------------------------------------------------------------------------------------
    
    nameProp = props.name.getName
    
    If pDbField = nameProp Then
        isNameProperty = True
    ElseIf pName = nameProp Then
        isNameProperty = True
    ElseIf pPropName = nameProp Then
        isNameProperty = True
    ElseIf pPropName = (pRepo.getEntityName & nameProp) Then
        isNameProperty = True
    End If
    
End Function

Public Function isExtraStringPropRequired() As Boolean
    If isRepoId Then        'Id is always a number but there is no need for its string representation
        isExtraStringPropRequired = False
    ElseIf isNumber Then
        isExtraStringPropRequired = True
    ElseIf isDate Then
        isExtraStringPropRequired = True
    Else
        isExtraStringPropRequired = False
    End If
End Function

Public Function getProperties() As VBA.Collection
    Set getProperties = New VBA.Collection
    With getProperties
        Call .Add(getPropName)
        If isExtraStringPropRequired Then
            Call .Add(getPropStringVersion)
        End If
    End With
End Function

Private Function getPropStringVersion() As String
    getPropStringVersion = pPropName & STRING_PROPERTY_SUFFIX
End Function

Public Function getValueWithCasting() As String
    If pType = TYPE_STRING Then
        getValueWithCasting = "VBA.CStr(value)"
    ElseIf pType = TYPE_BOOLEAN Then
        getValueWithCasting = "VBA.CBool(value)"
    ElseIf isRepoId Then
        getValueWithCasting = "VBA.CLng(value)"
    Else
        getValueWithCasting = "value"
    End If
End Function

Public Function isLinkedToRepo() As Boolean
    isLinkedToRepo = (VBA.Len(pLinkedRepo) > 0)
End Function

Public Function isEditable() As Boolean
    isEditable = (Not isRepoId)
End Function



'[CODE LINES GENERATOR]
Public Function getDeclarationLine() As String
    Const LINE_PATTERN As String = "Private p{0} As {1}"
    '------------------------------------------------------------------------------------------------------
    getDeclarationLine = F.Strings.Format(LINE_PATTERN, _
                                    F.Strings.convertLetterCasing(pName, LetterCasing_Sentence), pType)
End Function

Public Function getDtoLoadingLine() As String
    Const LINE_PATTERN As String = "    Call setValueFromDto(dto, Props_Project.{0})"
    '------------------------------------------------------------------------------------------------------
    getDtoLoadingLine = F.Strings.Format(LINE_PATTERN, _
                                    F.Strings.convertLetterCasing(pPropName, LetterCasing_Sentence))
End Function

Public Function getSetterMethod() As String
    Dim pattern As String
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(pLinkedRepo) Then
        pattern = RepoGenerator.getPattern("Setter:RepoObject")
        getSetterMethod = F.Strings.Format(pattern, _
                            F.Strings.convertLetterCasing(pName, LetterCasing_StartWithLower), _
                            F.Strings.convertLetterCasing(pName, LetterCasing_Sentence), _
                            pType, _
                            pLinkedRepo, _
                            pRepo.getEntityClassName)
    ElseIf isRepoId Then
        pattern = RepoGenerator.getPattern("Setter:Id")
        getSetterMethod = F.Strings.Format(pattern, _
                            F.Strings.convertLetterCasing(pName, LetterCasing_Sentence), _
                            pRepo.getEntityClassName)
    ElseIf pType = TYPE_DATE Then
        pattern = RepoGenerator.getPattern("Setter:Date")
        getSetterMethod = F.Strings.Format(pattern, _
                            F.Strings.convertLetterCasing(pName, LetterCasing_Sentence), _
                            pRepo.getEntityClassName)
    ElseIf pType = TYPE_DOUBLE Or pType = TYPE_LONG Then
        pattern = RepoGenerator.getPattern("Setter:Number")
        getSetterMethod = F.Strings.Format(pattern, _
                            F.Strings.convertLetterCasing(pName, LetterCasing_Sentence), _
                            getSetterValueDataType, _
                            pRepo.getEntityClassName)
    ElseIf isObject Then
        pattern = RepoGenerator.getPattern("Setter:Object")
        getSetterMethod = F.Strings.Format(pattern, _
                            F.Strings.convertLetterCasing(pName, LetterCasing_Sentence), _
                            getSetterValueDataType, _
                            pRepo.getEntityClassName)
Stop
    Else
        pattern = RepoGenerator.getPattern("Setter:Default")
        getSetterMethod = F.Strings.Format(pattern, _
                            F.Strings.convertLetterCasing(pName, LetterCasing_Sentence), _
                            getSetterValueDataType, _
                            pRepo.getEntityClassName)
    End If
    
End Function

Private Function getSetterValueDataType() As String
    If isRepoId Then
        getSetterValueDataType = TYPE_LONG
    ElseIf pType = TYPE_BOOLEAN Then
        getSetterValueDataType = TYPE_BOOLEAN
    ElseIf pType = TYPE_STRING Then
        getSetterValueDataType = TYPE_STRING
    Else
        getSetterValueDataType = TYPE_VARIANT
    End If
End Function

Public Function getSetValueMethodCases() As String
    Const LINE_PATTERN As String = _
                "            Case Props_Project.{0}.getLowerCaseName:                   Call set{1}({2})"
    Const ASSIGNMENT_POSITION As Long = 69
    '------------------------------------------------------------------------------------------------------
    Dim properties As Scripting.Dictionary
    Dim key As Variant
    Dim propName As String
    Dim varName As String
    Dim line As String
    Dim lines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    Set lines = New VBA.Collection
    Set properties = F.dictionaries.Create(False)
    
    With properties
        Call .Add(pPropName, pName)
        If isExtraStringPropRequired Then Call .Add(getPropStringVersion, pName)       'for string versions
    End With
    If VBA.Len(pExtraProps) Then Call addExtraPropsToSetterPropsList(properties)
    
    
    For Each key In properties.Keys
        propName = key
        varName = properties.item(key)
        
        line = F.Strings.Format(LINE_PATTERN, _
                                    F.Strings.convertLetterCasing(propName, LetterCasing_Sentence), _
                                    F.Strings.convertLetterCasing(varName, LetterCasing_Sentence), _
                                    getValueWithCasting)
        Call lines.Add(adjustSelectCaseLength(line, ASSIGNMENT_POSITION, "Call set"))
        
    Next key
    
    getSetValueMethodCases = F.Collections.toString(lines, , VBA.vbCrLf)
    
End Function

Private Function addExtraPropsToSetterPropsList(dict As Scripting.Dictionary)
    Dim extraProps As Scripting.Dictionary
    Dim propName As String
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set extraProps = getExtraPropsDictionary
    For Each key In extraProps.Keys
        Call F.dictionaries.addItem(dict, key, pName, duplicateBehaviour_Skip)
    Next key
    
End Function

Private Function getExtraPropsDictionary() As Scripting.Dictionary
    Const REGEX_PATTERN As String = "(\w+)(?::(\w*))?"
    '------------------------------------------------------------------------------------------------------
    Dim matches As VBA.Collection
    Dim match As Variant
    Dim propName As String
    Dim dataType As String
    '------------------------------------------------------------------------------------------------------
    
    Set getExtraPropsDictionary = F.dictionaries.Create(False)
    
    Set matches = F.regex.getMatchArraysCollection(pExtraProps, REGEX_PATTERN)
    For Each match In matches
        propName = pName & F.Strings.convertLetterCasing(VBA.CStr(match(LBound(match))), LetterCasing_Sentence)
        dataType = F.Strings.coalesce(VBA.CStr(match(LBound(match) + 1)), TYPE_VARIANT)
        Call F.dictionaries.addItem(getExtraPropsDictionary, propName, dataType, duplicateBehaviour_Skip)
    Next match
        
End Function

Public Function getGetterMethod() As String
    Dim pattern As String
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(pLinkedRepo) Then
        pattern = RepoGenerator.getPattern("Getter:Object")
        getGetterMethod = F.Strings.Format(pattern, _
                            F.Strings.convertLetterCasing(pName, LetterCasing_Sentence), _
                            pType)
                            
        If VBA.Len(pExtraProps) Then
            getGetterMethod = getGetterMethod & getGetterMethodsForExtraProps
        End If
    ElseIf isObject Then
        pattern = RepoGenerator.getPattern("Getter:Object")
        getGetterMethod = F.Strings.Format(pattern, pName, pType)
    ElseIf isRepoId Then
        pattern = RepoGenerator.getPattern("Getter:Id")
        getGetterMethod = F.Strings.Format(pattern, pName)
    ElseIf isNumber Then
        pattern = RepoGenerator.getPattern("Getter:Number")
        getGetterMethod = F.Strings.Format(pattern, pName, pType)
    ElseIf isDate Then
        pattern = RepoGenerator.getPattern("Getter:Date")
        getGetterMethod = F.Strings.Format(pattern, pName, pType)
    Else
        pattern = RepoGenerator.getPattern("Getter:Default")
        getGetterMethod = F.Strings.Format(pattern, pName, pType)
    End If

End Function

Private Function getGetterMethodsForExtraProps() As String
    Static pattern As String
    '------------------------------------------------------------------------------------------------------
    Dim dict As Scripting.Dictionary
    Dim key As Variant
    Dim propName As String
    Dim dataType As String
    '------------------------------------------------------------------------------------------------------
    Dim lines As VBA.Collection
    Dim line As String
    '------------------------------------------------------------------------------------------------------
    
    Set lines = New VBA.Collection
    
    If VBA.Len(pattern) = 0 Then
        pattern = RepoGenerator.getPattern("Getter:RepoObjectExtraProp")
    End If
    
    Set dict = getExtraPropsDictionary
    For Each key In dict.Keys
        propName = key
        dataType = dict.item(key)
        line = F.Strings.Format(pattern, _
                    F.Strings.convertLetterCasing(propName, LetterCasing_Sentence), _
                    pName, _
                    VBA.Replace(propName, pName, vbNullString), _
                    dataType)
        Call lines.Add(line)
    Next key
    
    If F.Collections.isNonEmptyCollection(lines) Then
        getGetterMethodsForExtraProps = F.Collections.toString(lines, , vbNullString)
    End If
    
End Function

Public Function getGetValueMethodCases() As String
    Dim line As String
    '------------------------------------------------------------------------------------------------------
    Dim extraProps As Scripting.Dictionary
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(pLinkedRepo) Then
        getGetValueMethodCases = getSingleGetValueMethodCaseLine(pPropName, pName, True)
        
        Set extraProps = getExtraPropsDictionary
        If F.dictionaries.isNonEmptyDictionary(extraProps) Then
            For Each key In extraProps.Keys
                getGetValueMethodCases = getGetValueMethodCases & VBA.vbCrLf & _
                                         getSingleGetValueMethodCaseLine(CStr(key), CStr(key), False)
            Next key
        End If
    ElseIf isObject Then
        getGetValueMethodCases = getSingleGetValueMethodCaseLine(pPropName, pName, True)
    Else
        getGetValueMethodCases = getSingleGetValueMethodCaseLine(pPropName, pName, False)
        If isExtraStringPropRequired Then
            getGetValueMethodCases = getGetValueMethodCases & VBA.vbCrLf & _
                                     getSingleGetValueMethodCaseLine(pPropName & "Str", pName & "Str", False)
        End If
    End If
    
End Function

Public Function getSingleGetValueMethodCaseLine(propName As String, name As String, _
                                                Optional isObject As Boolean = False) As String
    Const LINE_PATTERN As String = _
                "        Case Props_Project.{0}.getLowerCaseName:              {2}getValue = get{1}"
    Const ASSIGNMENT_POSITION As Long = 69
    '------------------------------------------------------------------------------------------------------
    
    getSingleGetValueMethodCaseLine = F.Strings.Format(LINE_PATTERN, _
                                                    propName, _
                                                    F.Strings.convertLetterCasing(name, LetterCasing_Sentence), _
                                                    VBA.IIf(isObject, "Set ", VBA.vbNullString))
    
    getSingleGetValueMethodCaseLine = adjustSelectCaseLength( _
                                            getSingleGetValueMethodCaseLine, _
                                            ASSIGNMENT_POSITION, _
                                            "(?:Set )?getValue")
    
End Function

Public Function getCloneLine() As String
    Const LINE_PATTERN As String = "        Call .setValue(Props_Project.{0}.getName, get{1})"
    '------------------------------------------------------------------------------------------------------
    getCloneLine = F.Strings.Format(LINE_PATTERN, _
                                    F.Strings.convertLetterCasing(pPropName, LetterCasing_Sentence), _
                                    F.Strings.convertLetterCasing(pName, LetterCasing_Sentence))
End Function

Public Function getInsertDtoLine() As String
    Const LINE_PATTERN As String = "        Call .Add(Props_Project.{0}.getDbField, get{1})"
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(pLinkedRepo) Then
        getInsertDtoLine = F.Strings.Format(LINE_PATTERN, _
                                    F.Strings.convertLetterCasing(pPropName & "Id", LetterCasing_Sentence), _
                                    F.Strings.convertLetterCasing(pName & "Id", LetterCasing_Sentence))
    Else
        getInsertDtoLine = F.Strings.Format(LINE_PATTERN, _
                                    F.Strings.convertLetterCasing(pPropName, LetterCasing_Sentence), _
                                    F.Strings.convertLetterCasing(pName, LetterCasing_Sentence))
    End If
    
End Function

Public Function getEditLinesDefinitionsDeclaration() As String
    Const LINE_PATTERN As String = "    Dim {0} As Scripting.Dictionary"
    '------------------------------------------------------------------------------------------------------
    getEditLinesDefinitionsDeclaration = F.Strings.Format(LINE_PATTERN, _
                                            F.Strings.convertLetterCasing(pLinkedRepo, LetterCasing_StartWithLower))
End Function

Public Function getEditLinesDefinitionsAssignment() As String
    Const LINE_PATTERN As String = "    Set {0} = Entities.{1}.toDictionary(Props_Project.{2}, True)"
    '------------------------------------------------------------------------------------------------------
    getEditLinesDefinitionsAssignment = F.Strings.Format(LINE_PATTERN, _
                                            F.Strings.convertLetterCasing(pLinkedRepo, LetterCasing_StartWithLower), _
                                            pLinkedRepo, _
                                            F.Strings.convertLetterCasing(pName, LetterCasing_StartWithLower) & "Name")
End Function

Public Function getEditLineDefinition() As String
    Dim pattern As String
    '------------------------------------------------------------------------------------------------------
    
    If isNameProperty Then
        pattern = RepoGenerator.getPattern("EditLine:Name")
        getEditLineDefinition = F.Strings.Format(pattern, pRepo.getName, pPropName, pName)
    ElseIf VBA.Len(pLinkedRepo) Then
        pattern = RepoGenerator.getPattern("EditLine:RepoObject")
        getEditLineDefinition = F.Strings.Format(pattern, pRepo.getName, pPropName, pName, _
                                    F.Strings.convertLetterCasing(pLinkedRepo, LetterCasing_StartWithLower))
    ElseIf pType = TYPE_DATE Then
        pattern = RepoGenerator.getPattern("EditLine:Date")
        getEditLineDefinition = F.Strings.Format(pattern, pRepo.getName, pPropName, pName)
    ElseIf pType = TYPE_DOUBLE Or pType = TYPE_LONG Then
        pattern = RepoGenerator.getPattern("EditLine:Number")
        getEditLineDefinition = F.Strings.Format(pattern, pRepo.getName, pPropName, pName)
    Else
        pattern = RepoGenerator.getPattern("EditLine:Default")
        getEditLineDefinition = F.Strings.Format(pattern, pRepo.getName, pPropName, pName)
Stop
    End If
    
End Function

Private Function adjustSelectCaseLength(line As String, expectedPosition As Long, substr As String) As String
    Const POSITION_REGEX_PATTERN As String = "^(.*?){0}"
    Const TRIM_REGEX_PATTERN As String = "{0}({1})"
    '------------------------------------------------------------------------------------------------------
    Dim assignmentPosition As Long
    Dim regexPattern As String
    '------------------------------------------------------------------------------------------------------
                                    
    '[Adjusting line length]
    regexPattern = F.Strings.Format(POSITION_REGEX_PATTERN, substr)
    assignmentPosition = VBA.Len(F.regex.getFirstGroupMatch(line, regexPattern))
    
    If assignmentPosition < expectedPosition Then
        regexPattern = F.Strings.Format(TRIM_REGEX_PATTERN, VBA.vbNullString, substr)
        On Error Resume Next
        adjustSelectCaseLength = F.regex.Replace(line, regexPattern, VBA.String(expectedPosition - assignmentPosition, " ") & "$1")
        On Error GoTo 0
    Else
        regexPattern = F.Strings.Format(TRIM_REGEX_PATTERN, VBA.String(assignmentPosition - expectedPosition, " "), substr)
        On Error Resume Next
        adjustSelectCaseLength = F.regex.Replace(line, regexPattern, "$1")
        On Error GoTo 0
    End If
    
End Function

Public Sub appendValidationFunction(dict As Scripting.Dictionary)
    Const METHOD_NAME_PATTERN As String = "validate{0}"
    '------------------------------------------------------------------------------------------------------
    Dim methodName As String
    Dim methodBody As String
    Dim pattern As String
    '------------------------------------------------------------------------------------------------------
    
    If Not isEditable Then
        ' skip
    ElseIf isNameProperty Then
        ' skip
    ElseIf isDate Then
        methodName = F.Strings.Format(METHOD_NAME_PATTERN, TYPE_DATE)
        pattern = RepoGenerator.getPattern("ValidationFunction:Date")
        methodBody = F.Strings.Format(pattern, _
                                        F.Strings.convertLetterCasing(TYPE_DATE, LetterCasing_Sentence))
                                        
    ElseIf isLinkedToRepo Then
        methodName = F.Strings.Format(METHOD_NAME_PATTERN, _
                                        F.Strings.convertLetterCasing(pName, LetterCasing_Sentence))
        pattern = RepoGenerator.getPattern("ValidationFunction:RepoObject")
        methodBody = F.Strings.Format(pattern, _
                                        F.Strings.convertLetterCasing(pName, LetterCasing_Sentence), _
                                        F.Strings.convertLetterCasing(pType, LetterCasing_Sentence), _
                                        VBA.UCase$(pName))
    Else
        methodName = F.Strings.Format(METHOD_NAME_PATTERN, _
                                        F.Strings.convertLetterCasing(pName, LetterCasing_Sentence))
        pattern = RepoGenerator.getPattern("ValidationFunction:Default")
        methodBody = F.Strings.Format(pattern, _
                                        F.Strings.convertLetterCasing(pName, LetterCasing_Sentence))
    End If
    
    If VBA.Len(methodBody) Then
        Call F.dictionaries.addItem(dict, methodName, methodBody, duplicateBehaviour_Skip)
    End If
    
End Sub
