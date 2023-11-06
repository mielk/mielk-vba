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
Private pId As String
Private pType As RibbonControlTypeEnum
Private pLabels As Scripting.Dictionary
Private pScreentips As Scripting.Dictionary
Private pVisible As Variant         'to allow null
Private pEnabled As Variant         'to allow null
Private pSize As String
Private pImage As String
'----------------------------------------------------------------------------------------------------------



'[SETTERS]
Public Function setId(value As String) As ERibbonControl
    Set setId = Me
    pId = value
End Function

Public Function setType(value As Variant) As ERibbonControl
    Set setType = Me
    If VBA.IsNumeric(value) Then
        pType = value
    ElseIf f.Variables.isString(value) Then
        pType = getRibbonControlTypeFromString(VBA.CStr(value))
    End If
End Function

Public Function setLabels(value As Variant) As ERibbonControl
    Set setLabels = Me
    If f.dictionaries.isNonEmptyDictionary(value) Then
        Set pLabels = value
    ElseIf f.Variables.isString(value) Then
        Call loadCaptionsFromString(pLabels, VBA.CStr(value))
    End If
End Function

Public Function setScreentips(value As Variant) As ERibbonControl
    Set setScreentips = Me
    If f.dictionaries.isNonEmptyDictionary(value) Then
        Set pScreentips = value
    ElseIf f.Variables.isString(value) Then
        Call loadCaptionsFromString(pScreentips, VBA.CStr(value))
    End If
End Function

Public Function setVisible(value As Variant) As ERibbonControl
    Set setVisible = Me
    pEnabled = Null
    On Error Resume Next
    pEnabled = VBA.CBool(value)
End Function

Public Function setEnabled(value As Variant) As ERibbonControl
    Set setEnabled = Me
    pEnabled = Null
    On Error Resume Next
    pEnabled = VBA.CBool(value)
End Function

Public Function setSize(value As String) As ERibbonControl
    Set setSize = Me
    pSize = value
End Function

Public Function setImage(value As String) As ERibbonControl
    Set setImage = Me
    pImage = value
End Function

Private Sub loadCaptionsFromString(ByRef dict As Scripting.Dictionary, text As String)
    Const REGEX_PATTERN As String = "^(\w{2}):(.*)$"
    '------------------------------------------------------------------------------------------------------
    Dim matches As VBA.Collection
    Dim match As Variant
    '------------------------------------------------------------------------------------------------------
    Dim languages As Scripting.Dictionary
    Dim language As language
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If dict Is Nothing Then Set dict = f.dictionaries.create(False)
    
    Set languages = Msg.getLanguagesDictionary
    Set matches = f.regex.getMatchArraysCollection(text, REGEX_PATTERN)
    
    If f.Collections.isNonEmptyCollection(matches) Then
        For Each match In matches
            Call f.dictionaries.addItem(dict, match(LBound(match)), match(LBound(match) + 1))
        Next match
    Else
        Set language = Msg.getCurrentLanguage
        Call f.dictionaries.addItem(dict, language.getKey, text)
    End If
    
End Sub

Public Function setValue(propName As String, value As Variant) As Variant
    Dim prop As enumProperty
    '------------------------------------------------------------------------------------------------------
    
    Set prop = props_Project.GetProperty(propName)
    
    If Not prop Is Nothing Then
        Select Case prop.getName
            Case props.ID.getName:                  Call setId(VBA.CStr(value))
            Case props_Project.RibbonType.getName:  Call setType(value)
            Case props_Project.Labels.getName:      Call setLabels(value)
            Case props_Project.Screentips.getName:  Call setScreentips(value)
            Case props_Project.IsVisible.getName:   Call setVisible(value)
            Case props_Project.IsEnabled.getName:   Call setEnabled(value)
            Case props_Project.Image.getName:       Call setImage(VBA.CStr(value))
            Case props_Project.Size.getName:        Call setSize(VBA.CStr(value))
        End Select
    End If
End Function

Public Sub loadFromDto(dto As Scripting.Dictionary)
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    For Each key In dto.Keys
        Call setValue(VBA.CStr(key), dto.Item(key))
    Next key
End Sub




'[GETTERS]
Public Function getId() As String
    getId = pId
End Function

Public Function getType() As RibbonControlTypeEnum
    getType = pType
End Function

Public Function getLabels() As Scripting.Dictionary
    Set getLabels = pLabels
End Function

Public Function getScreentips() As Scripting.Dictionary
    Set getScreentips = pScreentips
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

Public Function getValue(propName As String) As Variant
    Dim prop As enumProperty
    '------------------------------------------------------------------------------------------------------
    
    Set prop = props_Project.GetProperty(propName)
    
    Select Case prop.getName
        Case props.ID.getName:                  getValue = pId
        Case props_Project.RibbonType.getName:  getValue = pType
        Case props_Project.Labels.getName:      Set getValue = pLabels
        Case props_Project.Screentips.getName:  Set getValue = pScreentips
        Case props_Project.IsVisible.getName:   getValue = pVisible
        Case props_Project.IsEnabled.getName:   getValue = pEnabled
        Case props_Project.Image.getName:       getValue = pImage
        Case props_Project.Size.getName:        getValue = pSize
    End Select
End Function
