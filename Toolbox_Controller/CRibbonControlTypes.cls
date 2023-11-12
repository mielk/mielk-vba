VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRibbonControlTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "CRibbonControlTypes"
'----------------------------------------------------------------------------------------------------------
Private pByName As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------




'[INITIALIZING]
Private Sub Class_Initialize()
    Call initializeDictionaries
    Call insertEnumOptions
End Sub

Private Sub initializeDictionaries()
    Set pByName = f.dictionaries.create(False)
End Sub

Private Sub insertEnumOptions()
    Call addEnumOption(createNewEnumOption("tab").setLevel(1).setProperties(props.id, Props_Project.Label).setIsContainer(True))
    Call addEnumOption(createNewEnumOption("group").setLevel(2). _
                                        setProperties(props.id, Props_Project.Visible, Props_Project.Label).setIsContainer(True))
    Call addEnumOption(createNewEnumOption("menu").setLevel(3). _
                                        setProperties(props.id, Props_Project.Visible, Props_Project.Label, _
                                                        Props_Project.size, Props_Project.Image). _
                                        setIsContainer(True))
    Call addEnumOption(createNewEnumOption("label").setLevel(4).setXmlTag("labelControl"). _
                                        setProperties(props.id, Props_Project.Visible, Props_Project.Label))
    Call addEnumOption(createNewEnumOption("button").setLevel(4). _
                                        setProperties(props.id, Props_Project.Visible, Props_Project.Enabled, _
                                                        Props_Project.Label, Props_Project.Image, Props_Project.size, _
                                                        Props_Project.ScreenTip, Props_Project.OnAction))
    Call addEnumOption(createNewEnumOption("separator").setLevel(4).setProperties(props.id, Props_Project.Visible))
End Sub

Private Function createNewEnumOption(ByVal name As String) As enumRibbonControlType
    '------------------------------------------------------------------------------------------------------
    Set createNewEnumOption = New enumRibbonControlType
    With createNewEnumOption
        Call .setName(name)
    End With
End Function

Private Sub addEnumOption(item As enumRibbonControlType)
    Call f.dictionaries.AddItem(pByName, item.getName, item)
End Sub
    




'[ACCESS]
Public Function ByName(tableName As String) As enumRibbonControlType
    Set ByName = f.dictionaries.GetObject(pByName, tableName)
End Function

Public Function getControlType(ByVal value As Variant) As enumRibbonControlType
    If VBA.IsArray(value) Then
        'not possible to get property
    ElseIf Not VBA.IsObject(value) Then
        Set getControlType = f.dictionaries.GetObject(pByName, value)
    ElseIf value Is Nothing Then
        'Nothing should be returned
    ElseIf TypeOf value Is enumRibbonControlType Then
        Set getControlType = value
    End If
End Function




'[INTELLISENSE]
Public Function Tab_() As enumRibbonControlType: Set Tab_ = ByName("Tab"): End Function
Public Function Group() As enumRibbonControlType: Set Group = ByName("Group"): End Function
Public Function Menu() As enumRibbonControlType: Set Menu = ByName("Menu"): End Function
Public Function Label() As enumRibbonControlType: Set Label = ByName("Label"): End Function
Public Function Button() As enumRibbonControlType: Set Button = ByName("Button"): End Function
Public Function Separator() As enumRibbonControlType: Set Separator = ByName("Separator"): End Function
