VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "SConfig"
'----------------------------------------------------------------------------------------------------------
Private pConfigs As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------


Public Sub loadJsonsFromFolder()
    Const REGEX_PATTERN As String = "^{0}\\\\?(.*)\.json$"
    Dim regexPattern As String
    '------------------------------------------------------------------------------------------------------
    Dim files As VBA.Collection
    Dim file As Scripting.file
    Dim fileTag As String
    Dim json As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    Set pConfigs = F.dictionaries.Create(False)
    
    regexPattern = F.Strings.Format(REGEX_PATTERN, VBA.Replace(Paths.ConfigFolderPath, "\", "\\"))
    Set files = F.files.getFolderFiles(Paths.ConfigFolderPath, True, JSON_FILE_REGEX_PATTERN)
    For Each file In files
        fileTag = F.regex.getFirstGroupMatch(file.path, regexPattern)
        Set json = F.json.getJsonObjectFromFile(file.path)
        Call F.dictionaries.AddItem(pConfigs, fileTag, json, duplicateBehaviour_Override)
    Next file
    
End Sub




Public Function getConfig(tag As String) As Scripting.Dictionary
    Set getConfig = F.dictionaries.GetObject(pConfigs, tag)
End Function

Public Function getValueByTag(configName As String, tags As String) As Variant
    Dim dict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    Set dict = getConfig(configName)
    If Not dict Is Nothing Then
        Call F.dictionaries.assignValueFromDictToVariable(dict, getValueByTag, tags)
    End If
    
End Function

Public Function getValueByTags(configName As String, ParamArray tags() As Variant) As Variant
    Dim tag As String
    '------------------------------------------------------------------------------------------------------
    On Error Resume Next
    tag = VBA.Join(tags, ".")
    Call F.Variables.assign(getValueByTags, getValueByTag(configName, tag))
End Function
