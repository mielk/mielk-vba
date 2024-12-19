VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVbeProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "CVbeProperties"
'----------------------------------------------------------------------------------------------------------
Private pByNames As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------



'[INITIALIZING]
Private Sub Class_Initialize()
    Call initializeDictionaries
    Call insertEnumOptions
End Sub

Private Sub initializeDictionaries()
    Set pByNames = F.dictionaries.Create(False)
End Sub

Private Sub insertEnumOptions()
    
    '[Variables]
    Call addEnumOption("IsConst")
    Call addEnumOption("VariableType")
    Call addEnumOption("IsOptional")
    Call addEnumOption("PassMode")
    Call addEnumOption("IsArray")
    Call addEnumOption("IsParamArray")
    Call addEnumOption("DataType")
    Call addEnumOption("DefaultValue")
    Call addEnumOption("ArraySize")
    
End Sub

Private Sub addEnumOption(name As String, Optional ByVal dbField As String)
    Dim enumProperty As enumProperty
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(dbField) = 0 Then dbField = name
    Set enumProperty = F.Create.prop(name, dbField)
    
    Call F.dictionaries.addItem(pByNames, enumProperty.getName, enumProperty, duplicateBehaviour_WarningInImmediateWindow)
    
End Sub



'[ACCESS]
Public Function ByName(name As String) As enumProperty
    Set ByName = F.dictionaries.GetObject(pByNames, name)
End Function

Public Function getProperty(ByVal prop As Variant) As enumProperty
    If VBA.IsArray(prop) Then
        'not possible to get property
    ElseIf Not VBA.isObject(prop) Then
        Set getProperty = F.dictionaries.GetObject(pByNames, prop)
    ElseIf prop Is Nothing Then
        'Nothing should be returned
    ElseIf TypeOf prop Is enumProperty Then
        Set getProperty = prop
    End If

    If getProperty Is Nothing Then
        Set getProperty = Props_Project.getProperty(prop)
    End If
    
End Function


'[INTELLISENSE]
Public Function isConst() As enumProperty: Set isConst = ByName("IsConst"): End Function
Public Function variableType() As enumProperty: Set variableType = ByName("VariableType"): End Function
Public Function IsOptional() As enumProperty: Set IsOptional = ByName("IsOptional"): End Function
Public Function PassMode() As enumProperty: Set PassMode = ByName("PassMode"): End Function
Public Function IsArray() As enumProperty: Set IsArray = ByName("IsArray"): End Function
Public Function IsParamArray() As enumProperty: Set IsParamArray = ByName("IsParamArray"): End Function
Public Function dataType() As enumProperty: Set dataType = ByName("DataType"): End Function
Public Function DefaultValue() As enumProperty: Set DefaultValue = ByName("DefaultValue"): End Function
Public Function ArraySize() As enumProperty: Set ArraySize = ByName("ArraySize"): End Function
