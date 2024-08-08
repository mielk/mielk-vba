VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FConfig"
'----------------------------------------------------------------------------------------------------------

Public Sub assignProperty(dto As Scripting.Dictionary, ByRef property As Variant, ByVal propertyTag As String, _
                            Optional defaultValues As Scripting.Dictionary, Optional item As IConfigurable)
    Const HAS_DYNAMIC_PART_REGEX As String = "{\w+}"
    '------------------------------------------------------------------------------------------------------
    
    If f.dictionaries.hasCombinedKey(dto, propertyTag) Then
        Call f.dictionaries.assignValueFromDictToVariable(dto, property, propertyTag)
    ElseIf Not defaultValues Is Nothing Then
        Call f.dictionaries.assignValueFromDictToVariable(defaultValues, property, propertyTag)
        
        '[If value contains placeholders replace them with given item's properties.] ----------------|
        If VBA.IsObject(property) Then                                                              '|
            ' skip - not possible for objects.                                                      '|
        ElseIf item Is Nothing Then                                                                 '|
            ' skip - not possible if item is not specified.                                         '|
        ElseIf f.regex.checkIfMatch(VBA.CStr(property), HAS_DYNAMIC_PART_REGEX) Then                '|
            property = modifyDefaultProperty(item, VBA.CStr(property))                              '|
        End If                                                                                      '|
        '--------------------------------------------------------------------------------------------|
        
    End If
    
End Sub

Private Function modifyDefaultProperty(item As IConfigurable, value As String) As String
    Const DYNAMIC_PART_REGEX As String = "{(\w+)}"
    Const UPPER_CASE_REGEX As String = "^[A-Z]*$"
    '------------------------------------------------------------------------------------------------------
    Dim tags As VBA.Collection
    Dim arr As Variant
    Dim propName As String
    Dim replacement As String
    '------------------------------------------------------------------------------------------------------
    
    modifyDefaultProperty = value
    
    Set tags = f.regex.getMatchArraysCollection(value, DYNAMIC_PART_REGEX)
    For Each arr In tags
        If f.arrays.isDefined(arr) Then
            propName = f.arrays.getFirstItem(arr)
            replacement = item.getValue(propName)
            If f.strings.isInUpperCase(propName) Then replacement = VBA.UCase$(replacement)
            modifyDefaultProperty = VBA.replace(modifyDefaultProperty, "{" & propName & "}", replacement)
        End If
    Next arr
    
End Function


