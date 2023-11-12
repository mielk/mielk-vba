VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "enumRibbonControlType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "enumRibbonControlType"
'----------------------------------------------------------------------------------------------------------
Private Const CONTAINER_NODE_PATTERN As String = "{0}<{1} {2}>\n{3}\n{0}</{1}>"
Private Const NON_CONTAINER_NODE_PATTERN As String = "{0}<{1} {2} />"
Private Const CALLBACK_FUNCTION_PATTERN As String = "get{0}"
'----------------------------------------------------------------------------------------------------------
Private pName As String
Private pXmlTag As String
Private pLevel As Long
Private pProperties As VBA.Collection
Private pIsClickable As Boolean
Private pIsContainer As Boolean
'----------------------------------------------------------------------------------------------------------



'[SETTERS]
Public Function setName(value As String) As enumRibbonControlType
    Set setName = Me
    pName = value
End Function

Public Function setXmlTag(value As String) As enumRibbonControlType
    Set setXmlTag = Me
    pXmlTag = value
End Function

Public Function setLevel(value As Long) As enumRibbonControlType
    Set setLevel = Me
    pLevel = value
End Function

Public Function setProperties(ParamArray props() As Variant) As enumRibbonControlType
    Dim prop As Variant
    '------------------------------------------------------------------------------------------------------
    Set setProperties = Me
    Set pProperties = New VBA.Collection
    For Each prop In props
        Call pProperties.Add(prop)
    Next prop
End Function

Public Function setIsClickable(value As Boolean) As enumRibbonControlType
    Set setIsClickable = Me
    pIsClickable = value
End Function

Public Function setIsContainer(value As Boolean) As enumRibbonControlType
    Set setIsContainer = Me
    pIsContainer = value
End Function






'[GETTERS]
Public Function getName() As String
    getName = pName
End Function

Public Function getXmlTag() As String
    getXmlTag = VBA.IIf(VBA.Len(pXmlTag), pXmlTag, pName)
End Function

Public Function getLevel() As Long
    getLevel = pLevel
End Function

Public Function getProperties() As VBA.Collection
    Set getProperties = pProperties
End Function

Public Function isClickable() As Boolean
    isClickable = pIsClickable
End Function

Public Function isContainer() As Boolean
    isContainer = pIsContainer
End Function

Public Function getIndentationForXmlFile() As String
    getIndentationForXmlFile = VBA.String(pLevel + 2, VBA.vbTab)
End Function

Public Function getXmlLinePattern() As String
    Dim indentation As String
    Dim pattern As String
    '------------------------------------------------------------------------------------------------------
    
    indentation = VBA.String(pLevel + 2, VBA.vbTab)
    pattern = VBA.IIf(pIsContainer, CONTAINER_NODE_PATTERN, NON_CONTAINER_NODE_PATTERN)
    getXmlLinePattern = f.Strings.Format(pattern, indentation, getXmlTag, _
                                        Props_Project.Attributes.getName, Props_Project.Children.getName)
    
End Function
