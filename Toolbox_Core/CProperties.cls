VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "CProperties"
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
    
    '[PROJECT-SPECIFIC PROPERTIES]
    Call addEnumOption("ProjectCodeNameUCase")
    Call addEnumOption("ProjectLibFolderPath")
    Call addEnumOption("ProjectToolFolderPath")
    '<...>

    '[Ribbon]
    Call addEnumOption("RibbonType")
    Call addEnumOption("Labels")
    Call addEnumOption("Screentips")
    Call addEnumOption("IsVisible")
    Call addEnumOption("IsEnabled")
    Call addEnumOption("Image")
    Call addEnumOption("Size")
    
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
    ElseIf Not VBA.IsObject(prop) Then
        Set getProperty = F.dictionaries.GetObject(pByNames, prop)
    ElseIf prop Is Nothing Then
        'Nothing should be returned
    ElseIf TypeOf prop Is enumProperty Then
        Set getProperty = prop
    End If
    
    If getProperty Is Nothing Then
        Set getProperty = props.getProperty(prop)
    End If
    
End Function


'[INTELLISENSE]
'[Employees]
Public Function ProjectCodeNameUCase() As enumProperty: Set ProjectCodeNameUCase = ByName("ProjectCodeNameUCase"): End Function
Public Function ProjectLibFolderPath() As enumProperty: Set ProjectLibFolderPath = ByName("ProjectLibFolderPath"): End Function
Public Function ProjectToolFolderPath() As enumProperty: Set ProjectToolFolderPath = ByName("ProjectToolFolderPath"): End Function

'[Ribbon]
Public Function RibbonType() As enumProperty: Set RibbonType = ByName("RibbonType"): End Function
Public Function Labels() As enumProperty: Set Labels = ByName("Labels"): End Function
Public Function Screentips() As enumProperty: Set Screentips = ByName("Screentips"): End Function
Public Function IsVisible() As enumProperty: Set IsVisible = ByName("IsVisible"): End Function
Public Function IsEnabled() As enumProperty: Set IsEnabled = ByName("IsEnabled"): End Function
Public Function Image() As enumProperty: Set Image = ByName("Image"): End Function
Public Function Size() As enumProperty: Set Size = ByName("Size"): End Function
