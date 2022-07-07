VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MsgService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITranslator

Private Const CLASS_NAME As String = "MsgService"
    '----------------------------------------------------------------------------------------------------------
Private Const JSON_TAG_LANGUAGES As String = "languages"
Private Const JSON_TAG_MESSAGES As String = "messages"
Private Const JSON_TAG_NAME As String = "name"
Private Const JSON_TAG_IS_DEFAULT As String = "default"
'-----------------------------------------------------
Private pLanguages As Scripting.Dictionary
Private pCurrentLanguage As Language
Private pSourceFolderPath As String
'[Events] -------------------------------------------------------------------------------------------------
Event AfterLanguageChange()
    '----------------------------------------------------------------------------------------------------------






'[SETTERS]
Public Sub setCurrentLanguage(Language As Language)
    If Not Language Is Nothing Then
        Set pCurrentLanguage = Language
        RaiseEvent AfterLanguageChange
    End If
End Sub

Public Sub setCurrentLanguageByKey(Key As String)
    Call setCurrentLanguage(getLanguageByKey(Key))
End Sub



'[GETTERS]
Public Function getCurrentLanguage() As Language
    Set getCurrentLanguage = pCurrentLanguage
End Function

Public Function getText(tag As String, Optional ByVal Language As String) As String
    Dim objLanguage As Language
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(Language) Then
        Set objLanguage = findLanguage(Language)
    Else
        Set objLanguage = pCurrentLanguage
    End If
    
    If Not objLanguage Is Nothing Then
        getText = objLanguage.getText(tag)
    Else
        getText = tag
    End If
    
End Function

Private Function findLanguage(Language As Variant) As Language
    If VBA.IsObject(Language) Then
        If TypeOf Language Is Language Then
            Set findLanguage = Language
        End If
    ElseIf f.Variables.isString(Language) Then
        Set findLanguage = getLanguageByKey(VBA.CStr(Language))
    End If
End Function

Private Function getLanguageByKey(Key As String) As Language
    Set getLanguageByKey = f.dictionaries.getObject(pLanguages, Key)
End Function

Public Function getLanguagesCollection() As VBA.Collection
    Set getLanguagesCollection = f.dictionaries.toCollection(pLanguages, DictPart_ValueOnly)
End Function

Public Function getLanguagesDictionary() As Scripting.Dictionary
    Set getLanguagesDictionary = f.dictionaries.clone(pLanguages)
End Function






'[Loading items]
Public Sub loadJsonsFromFolder(folderPath As String)
    Const FILE_NAME_PATTERN As String = ".*\.json$"
    '------------------------------------------------------------------------------------------------------
    Dim folder As Scripting.folder
    Dim files As VBA.Collection
    Dim file As Scripting.file
    '------------------------------------------------------------------------------------------------------
    
    Set pLanguages = f.dictionaries.Create(False)
    Set pCurrentLanguage = Nothing
    pSourceFolderPath = folderPath
    Set folder = f.files.createFolder(folderPath)
    Set files = f.files.getFolderFiles(folderPath, True, FILE_NAME_PATTERN)
    
    For Each file In files
        Debug.Print "Loading messages from " & file.path
        Call loadFromJson(file.path)
    Next file
    
End Sub


Private Sub loadFromJson(filepath As String)
    Dim json As String
    Dim obj As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If pLanguages Is Nothing Then
        Set pLanguages = f.dictionaries.Create(False)
    End If
    
    On Error Resume Next
    json = f.TextFiles.readTextFile(filepath)
    Set obj = f.json.ParseJson(json)
    On Error GoTo 0
    
    If f.dictionaries.isNonEmptyDictionary(obj) Then
        Call loadLanguages(f.dictionaries.getObject(obj, JSON_TAG_LANGUAGES))
        Call loadTexts(f.dictionaries.getObject(obj, JSON_TAG_MESSAGES), vbNullString)
    End If
    
End Sub

Private Sub loadLanguages(dict As Scripting.Dictionary)
    Dim varKey As Variant
    Dim Language As Language
    '------------------------------------------------------------------------------------------------------
    
    If f.dictionaries.isNonEmptyDictionary(dict) Then
        For Each varKey In dict.Keys
            If Not pLanguages.exists(varKey) Then
                Set Language = createLanguage(VBA.CStr(varKey))
                Call f.dictionaries.addItem(pLanguages, varKey, Language)
            Else
                Set Language = pLanguages.Item(varKey)
            End If
            
            Call loadLanguageProperties(Language, dict.Item(varKey))

        Next varKey
    End If
    
End Sub

Private Sub loadTexts(dict As Scripting.Dictionary, prefix As String)
    Dim varKey As Variant
    Dim Language As Language
    Dim subDict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If f.dictionaries.isNonEmptyDictionary(dict) Then
        For Each varKey In dict.Keys
            Set subDict = f.cast.toDictionary(dict.Item(varKey))
            If Not subDict Is Nothing Then
                If subDict.count Then
                    Call loadTexts(dict.Item(varKey), prefix & VBA.IIf(VBA.Len(prefix), ".", vbNullString) & varKey)
                End If
            Else
                Set Language = f.dictionaries.getObject(pLanguages, varKey)
                If Language Is Nothing Then
                    Set Language = createLanguage(VBA.CStr(varKey))
                    Call pLanguages.add(varKey, Language)
                End If
                If Not Language Is Nothing Then
                    Call Language.addText(prefix, dict.Item(varKey))
                End If
            End If
        Next varKey
    End If
    
End Sub

Private Function createLanguage(Key As String) As Language
    Set createLanguage = New Language
    Call createLanguage.setKey(Key)
End Function

Private Sub loadLanguageProperties(Language As Language, Props As Scripting.Dictionary)
    Dim name As Variant
    Dim isDefault As Variant
    '------------------------------------------------------------------------------------------------------
    
    Call f.dictionaries.assignValueFromDictToVariable(Props, name, JSON_TAG_NAME)
    Call f.dictionaries.assignValueFromDictToVariable(Props, isDefault, JSON_TAG_IS_DEFAULT)
    
    With Language
        If Not VBA.IsEmpty(name) Then Call .setName(VBA.CStr(name))
        If Not VBA.IsEmpty(isDefault) Then Call .setIsDefault(VBA.CBool(isDefault))
    End With
        
    If Language.isDefault Or pCurrentLanguage Is Nothing Then
        Set pCurrentLanguage = Language
    End If

End Sub


Public Sub reload()
    If f.files.folderExists(pSourceFolderPath) Then
        Call loadJsonsFromFolder(pSourceFolderPath)
    End If
End Sub



'[Implementation of ITranslator]
Private Function ITranslator_getText(tag As String, Optional ByVal Language As String) As String
    ITranslator_getText = getText(tag, Language)
End Function
