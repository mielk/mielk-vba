VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ERibbonControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "ERibbonControl"
'----------------------------------------------------------------------------------------------------------
Private Const RIBBON_ON_ACTION_VALUE As String = "action_{0}"
Private Const RIBBON_CALLBACK_PROP_VALUE As String = "get{0}_{1}"
'----------------------------------------------------------------------------------------------------------
Private pProject As EProject
Private pId As String
Private pType As enumRibbonControlType  'RibbonControlTypeEnum
Private pLabel As String                'Scripting.Dictionary
Private pScreentip As String            'Scripting.Dictionary
Private pVisible As Variant             'to allow null
Private pEnabled As Variant             'to allow null
Private pSize As String
Private pImage As String
'----------------------------------------------------------------------------------------------------------
Private pChildren As VBA.Collection
Private pParent As ERibbonControl
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    Set pChildren = New VBA.Collection
End Sub



'[SETTERS]
Public Function setProject(value As EProject) As ERibbonControl
    Set setProject = Me
    Set pProject = value
End Function

Public Function setId(value As String) As ERibbonControl
    Set setId = Me
    pId = value
End Function

Public Function setType(value As Variant) As ERibbonControl
    Set setType = Me
    Set pType = RibbonControlTypes.getControlType(value)
End Function

Public Function setLabels(value As String) As ERibbonControl
    Set setLabels = Me
    pLabel = value
End Function

Public Function setScreentips(value As String) As ERibbonControl
    Set setScreentips = Me
    pScreentip = value
End Function

Public Function setVisible(value As Variant) As ERibbonControl
    Set setVisible = Me
    If Not VBA.IsEmpty(value) Then pVisible = VBA.CBool(value)
End Function

Public Function setEnabled(value As Variant) As ERibbonControl
    Set setEnabled = Me
    If Not VBA.IsEmpty(value) Then pEnabled = VBA.CBool(value)
End Function

Public Function setSize(value As String) As ERibbonControl
    Set setSize = Me
    pSize = value
End Function

Public Function setImage(value As String) As ERibbonControl
    Set setImage = Me
    pImage = value
End Function

Public Function setValue(propName As String, value As Variant) As Variant
    Dim prop As enumProperty
    '------------------------------------------------------------------------------------------------------
    
    Set prop = Props_Project.GetProperty(propName)
    
    If Not prop Is Nothing Then
        Select Case prop.getName
            Case props.id.getName:                  Call setId(VBA.CStr(value))
            Case Props_Project.RibbonType.getName:  Call setType(value)
            Case Props_Project.Label.getName:       Call setLabels(VBA.CStr(value))
            Case Props_Project.ScreenTip.getName:   Call setScreentips(VBA.CStr(value))
            Case Props_Project.Visible.getName:     Call setVisible(value)
            Case Props_Project.Enabled.getName:     Call setEnabled(value)
            Case Props_Project.Image.getName:       Call setImage(VBA.CStr(value))
            Case Props_Project.size.getName:        Call setSize(VBA.CStr(value))
        End Select
    End If
End Function

Public Sub loadFromDto(dto As Scripting.Dictionary)
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    For Each key In dto.Keys
        Call setValue(VBA.CStr(key), dto.item(key))
    Next key
End Sub


Public Function setParent(value As ERibbonControl) As ERibbonControl
    Set setParent = Me
    Set pParent = value
End Function

Public Sub addControl(value As ERibbonControl)
    Call pChildren.Add(value)
End Sub




'[GETTERS]
Public Function getId() As String
    getId = pId
End Function

Public Function getIdForXml(Optional toLowerCase As Boolean = True) As String
    Dim letterCasing As LetterCasingEnum:       letterCasing = VBA.IIf(toLowerCase, LetterCasing_Lower, LetterCasing_Sentence)
    Dim parts As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    Set parts = New VBA.Collection
    With parts
        Call .Add(f.Strings.convertLetterCasing(pType.getName, letterCasing))
        If pType Is RibbonControlTypes.Tab_ Then
            Call .Add(f.Strings.convertLetterCasing(pProject.getCodeName, letterCasing))
        ElseIf pType.isContainer Then
            Call .Add(f.Strings.convertLetterCasing(pId, letterCasing))
        Else
            Call .Add(f.Strings.convertLetterCasing(pParent.getId, letterCasing))
            Call .Add(f.Strings.convertLetterCasing(pId, letterCasing))
        End If
    End With
        
    getIdForXml = f.Collections.toString(parts, , ".")
    
End Function

Public Function getTagForMsgFile(prop As enumProperty) As String
    Dim parts As Variant
    '------------------------------------------------------------------------------------------------------
    parts = VBA.Array( _
                    props.Ribbon.getName, _
                    getIdForXml(False), _
                    prop.getName)
    getTagForMsgFile = VBA.Join(parts, ".")
End Function

Public Function getType() As enumRibbonControlType
    Set getType = pType
End Function

Public Function getLabelString() As String
    getLabelString = pLabel
End Function

Public Function getLabelsDictionary() As Scripting.Dictionary
    Stop
End Function

Public Function getScreentipString() As String
    getScreentipString = pScreentip
End Function

Public Function getScreentipsDictionary() As Scripting.Dictionary
    Stop
End Function

Public Function getVisible() As Variant
    getVisible = pVisible
End Function

Public Function getEnabled() As Variant
    getEnabled = pEnabled
End Function

Public Function getImage() As String
    getImage = pImage
End Function

Public Function getSize() As String
    getSize = pSize
End Function

Public Function getLevel() As Long
    getLevel = pType.getLevel
End Function

Public Function getValue(propName As String) As Variant
    Dim prop As enumProperty
    '------------------------------------------------------------------------------------------------------
    
    Set prop = Props_Project.GetProperty(propName)
    
    Select Case prop.getName
        Case props.id.getName:                          getValue = pId
        Case Props_Project.RibbonType.getName:          getValue = pType
        Case Props_Project.Label.getName:               getValue = getLabelString
        Case Props_Project.LabelsDict.getName:          Set getValue = getLabelsDictionary
        Case Props_Project.ScreenTip.getName:           getValue = getScreentipString
        Case Props_Project.ScreentipsDict.getName:      Set getValue = getScreentipsDictionary
        Case Props_Project.Visible.getName:             getValue = pVisible
        Case Props_Project.Enabled.getName:             getValue = pEnabled
        Case Props_Project.Image.getName:               getValue = pImage
        Case Props_Project.size.getName:                getValue = pSize
    End Select
End Function


Public Function getParent() As ERibbonControl
    Set getParent = pParent
End Function

Public Function getChildren() As VBA.Collection
    Set getChildren = pChildren
End Function

Public Function getDescendants(Optional includeMe As Boolean = True) As VBA.Collection
    Dim ctrl As ERibbonControl
    '------------------------------------------------------------------------------------------------------
    
    Set getDescendants = New VBA.Collection
    
    If includeMe Then Call getDescendants.Add(Me)
    
    For Each ctrl In pChildren
        Call getDescendants.Add(ctrl)
        Call f.Collections.appendCollection(getDescendants, ctrl.getDescendants)
    Next ctrl
    
End Function



'[XML]
Public Function getXml() As String
    Dim pattern As String
    Dim values As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    pattern = pType.getXmlLinePattern
    Set values = f.dictionaries.createWithItems(False, _
                                    KeyValue(Props_Project.Attributes.getName, getAttributesXmlString), _
                                    KeyValue(Props_Project.Children.getName, getChildrenXmlString))
    getXml = f.Strings.replaceByDictionary(pattern, values)
    
End Function

Private Function getAttributesXmlString() As String
    Const SINGLE_ENTRY_PATTERN As String = "{0}=""{1}"""
    '------------------------------------------------------------------------------------------------------
    Dim properties As VBA.Collection
    Dim prop As enumProperty
    Dim entries As VBA.Collection
    Dim tag As String
    Dim value As String
    '------------------------------------------------------------------------------------------------------

    Set entries = New VBA.Collection
    
    Set properties = pType.getProperties
    For Each prop In properties
        If Not shouldBeExcluded(prop) Then
            tag = getRibbonPropertyXmlTag(prop)
            value = getValueForXml(prop)
            If Not f.Variables.isUndefined(value) Then
                Call entries.Add(f.Strings.Format(SINGLE_ENTRY_PATTERN, tag, value))
            End If
        End If
    Next prop

    getAttributesXmlString = f.Collections.toString(entries, , " ")

End Function


'There are properties that should be included in an XML for a control type,
'but only in a specific conditions.
Private Function shouldBeExcluded(prop As enumProperty) As Boolean
    If prop Is Props_Project.size Then
        If pParent.getType Is RibbonControlTypes.Menu Then shouldBeExcluded = True
    End If
End Function

Private Function getValueForXml(prop As enumProperty) As String
    If prop Is props.id Then
        getValueForXml = getIdForXml
    ElseIf prop Is Props_Project.OnAction Then
        getValueForXml = f.Strings.Format(RIBBON_ON_ACTION_VALUE, pProject.getCodeName)
    ElseIf Not isCallbackProperty(prop) Then
        getValueForXml = getValue(prop.getName)
        If f.Variables.isUndefined(getValueForXml, True) Then
            getValueForXml = getRibbonPropertyDefaultValue(prop, pType)
        End If
    Else
        getValueForXml = f.Strings.Format(RIBBON_CALLBACK_PROP_VALUE, f.Strings.toSentenceCase(prop.getName), _
                                                                                        pProject.getCodeName)
    End If
End Function

Private Function getChildrenXmlString() As String
    Dim child As ERibbonControl
    Dim lines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    If pType.isContainer Then
        Set lines = New VBA.Collection
        For Each child In pChildren
            Call lines.Add(child.getXml)
        Next child
        getChildrenXmlString = f.Collections.toString(lines, , VBA.vbCrLf)
    End If
    
End Function
'[/// XML]







'[JSON]
Public Function getPropertiesDictionaryForJsonFile() As Scripting.Dictionary
    Dim prop As enumProperty
    '------------------------------------------------------------------------------------------------------

    Set getPropertiesDictionaryForJsonFile = f.dictionaries.create(False)
    
    For Each prop In pType.getProperties
        If Not isCallbackProperty(prop) Then
            ' Skip this property - it is to be defined directly in XML file.
        Else
            Call f.dictionaries.AddItem( _
                        getPropertiesDictionaryForJsonFile, _
                        VBA.LCase(prop.getName), _
                        getSinglePropertyEntryForJsonFile(prop))
        End If
    Next prop
    
End Function

Private Function getSinglePropertyEntryForJsonFile(prop As enumProperty) As Scripting.Dictionary
    Const CALLBACK_PATTERN As String = "^{(.*)}$"
    Dim VALUE_TAG As String:                            VALUE_TAG = VBA.LCase$(props.value.getName)
    '------------------------------------------------------------------------------------------------------
    Dim value As Variant
    Dim functionName As String
    '------------------------------------------------------------------------------------------------------
    
    Set getSinglePropertyEntryForJsonFile = f.dictionaries.create(False)
    
    value = getValue(prop.getName)
    If VBA.IsEmpty(value) Then value = getRibbonPropertyDefaultValue(prop)
    functionName = f.regex.getFirstGroupMatch(VBA.CStr(value), CALLBACK_PATTERN)
    
    With getSinglePropertyEntryForJsonFile
        If VBA.Len(functionName) Then
            Call .Add(Props_Project.Functions.getName, createRibbonJsonEntryForCallback(functionName))
        ElseIf isStringProperty(prop) Then
            Call .Add(VALUE_TAG, getTagForMsgFile(prop))
        ElseIf Not VBA.IsEmpty(value) Then
            Call .Add(VALUE_TAG, getAdjustedRibbonProperty(prop, value))
        End If
    End With

End Function

Private Function createRibbonJsonEntryForCallback(functionName As String) As VBA.Collection
    Set createRibbonJsonEntryForCallback = New VBA.Collection
    Call createRibbonJsonEntryForCallback.Add(f.Collections.create(functionName))
End Function

'[///JSON]





'[MSG]
Public Function getCaptions() As Scripting.Dictionary
    Dim prop As enumProperty
    Dim rawText As String
    '------------------------------------------------------------------------------------------------------
    
    Set getCaptions = f.dictionaries.create(False)
    
    For Each prop In pType.getProperties
        If isStringProperty(prop) Then
            rawText = getValue(prop.getName)
            Call getCaptions.Add(VBA.LCase$(prop.getName), getCaptionsDictionaryFromString(rawText))
        End If
    Next prop

End Function

Private Function getCaptionsDictionaryFromString(text As String) As Scripting.Dictionary
    Const REGEX_PATTERN As String = "^(\w{2}):(.*)$"
    '------------------------------------------------------------------------------------------------------
    Dim matches As VBA.Collection
    Dim match As Variant
    '------------------------------------------------------------------------------------------------------
    Dim languages As Scripting.Dictionary
    Dim language As language
    Dim i As Long
    '------------------------------------------------------------------------------------------------------

    Set getCaptionsDictionaryFromString = f.dictionaries.create(False)

    Set languages = Msg.getLanguagesDictionary
    Set matches = f.regex.getMatchArraysCollection(text, REGEX_PATTERN)

    If f.Collections.isNonEmptyCollection(matches) Then
        For Each match In matches
            Call f.dictionaries.AddItem(getCaptionsDictionaryFromString, _
                                            match(LBound(match)), match(LBound(match) + 1))
        Next match
    Else
        Set language = Msg.getCurrentLanguage
        Call f.dictionaries.AddItem(getCaptionsDictionaryFromString, language.getKey, text)
    End If

End Function

'[///MSG]



