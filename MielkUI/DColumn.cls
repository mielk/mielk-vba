VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "DColumn"
'----------------------------------------------------------------------------------------------------------
Private pName As String
Private pIndex As Long
Private pLabel As String
Private pPicture As Variant
Private pInactivePicture As Variant
Private pWidth As Single
Private pIsSortable As Boolean
Private pIsScalable As Boolean
Private pIsExportable As Boolean
Private pControlType As ControlTypeEnum
Private pValuePropertyName As String
Private pSortValuePropertyName As String
Private pEnableFunction As String
Private pTextAlignment As fmTextAlign
Private pTooltip As String
Private pForeColor As Long
Private pIsUnderlined As Boolean
Private pMousePointer As Object
Private pHasRightBorder As Boolean
Private pNumberFormat As String
Private pLeft As Single
Private pIndent As Single
Private pFilterRegex As String
'----------------------------------------------------------------------------------------------------------
Private pFilter As IFilter
Private pActions As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------




Private Sub Class_Initialize()

    '[Set initial values]
    pIsSortable = True
    pIsScalable = True
    pIsExportable = True
    pTextAlignment = fmTextAlignLeft
    pHasRightBorder = True
    pNumberFormat = "@"
    
    '[Initialize containers]
    Set pActions = F.dictionaries.Create(False)
    
End Sub




'[SETTERS]

'<simple setters>
Public Function setName(value As String) As DColumn
    Set setName = Me
    pName = value
End Function

Public Function setIndex(value As Long) As DColumn
    Set setIndex = Me
    pIndex = value
End Function

Public Function setLabel(value As String) As DColumn
    Set setLabel = Me
    pLabel = value
End Function

Public Function setControlType(value As ControlTypeEnum) As DColumn
    Set setControlType = Me
    pControlType = value
End Function

Public Function setPicture(value As Variant) As DColumn
    Set setPicture = Me
    Set pPicture = value
End Function

Public Function setInactivePicture(value As Variant) As DColumn
    Set setInactivePicture = Me
    Set pInactivePicture = value
End Function

Public Function setWidth(value As Single) As DColumn
    Set setWidth = Me
    pWidth = value
End Function

Public Function setFilter(value As IFilter) As DColumn
    Set setFilter = Me
    Set pFilter = value
End Function

Public Function setIsScalable(value As Boolean) As DColumn
    Set setIsScalable = Me
    pIsScalable = value
End Function

Public Function setIsSortable(value As Boolean) As DColumn
    Set setIsSortable = Me
    pIsSortable = value
End Function

Public Function setIsExportable(value As Boolean) As DColumn
    Set setIsExportable = Me
    pIsExportable = value
End Function

Public Function setValueFunction(value As String) As DColumn
    Set setValueFunction = Me
    pValuePropertyName = value
End Function

Public Function setSortValueFunction(value As String) As DColumn
    Set setSortValueFunction = Me
    pSortValuePropertyName = value
End Function

Public Function setEnableFunction(value As String) As DColumn
    Set setEnableFunction = Me
    pEnableFunction = value
End Function

Public Function setTextAlignment(value As fmTextAlign) As DColumn
    Set setTextAlignment = Me
    pTextAlignment = value
End Function

Public Function setFontColor(value As Long) As DColumn
    Set setFontColor = Me
    pForeColor = value
End Function

Public Function setIsUnderlined(value As Boolean) As DColumn
    Set setIsUnderlined = Me
    pIsUnderlined = value
End Function

Public Function setMousePointer(value As Object) As DColumn
    Set setMousePointer = Me
    Set pMousePointer = value
End Function

Public Function setHasRightBorder(value As Boolean) As DColumn
    Set setHasRightBorder = Me
    pHasRightBorder = value
End Function

Public Function setTooltip(value As String) As DColumn
    Set setTooltip = Me
    pTooltip = value
End Function

Public Function setNumberFormat(value As String) As DColumn
    Set setNumberFormat = Me
    pNumberFormat = value
End Function

Public Function setLeft(value As Single) As DColumn
    Set setLeft = Me
    pLeft = value
End Function

Public Function setIndent(value As Single) As DColumn
    Set setIndent = Me
    pIndent = value
End Function

Public Function setFilterRegex(value As String) As DColumn
    Set setFilterRegex = Me
    pFilterRegex = value
End Function

'</simple setters>




'[GETTERS]

'<simple getters>
Public Function getName() As String
    getName = pName
End Function

Public Function getIndex() As Long
    getIndex = pIndex
End Function

Public Function getLabel() As String
    getLabel = pLabel
End Function

Public Function getControlType() As ControlTypeEnum
    getControlType = pControlType
End Function

Public Function getPicture() As Variant
    Set getPicture = pPicture
End Function

Public Function getInactivePicture() As picture
    If VBA.IsEmpty(pInactivePicture) Then
        Set getInactivePicture = pPicture
    Else
        Set getInactivePicture = pInactivePicture
    End If
End Function

Public Function getWidth() As Single
    getWidth = pWidth
End Function

Public Function isScalable() As Boolean
    isScalable = pIsScalable
End Function

Public Function isExportable() As Boolean
    isExportable = pIsExportable
End Function

Public Function isSortable() As Boolean
    isSortable = pIsSortable
End Function

Public Function getValuePropertyName() As String
    getValuePropertyName = pValuePropertyName
End Function

Public Function getSortValuePropertyName() As String
    If VBA.Len(pSortValuePropertyName) Then
        getSortValuePropertyName = pSortValuePropertyName
    Else
        getSortValuePropertyName = pValuePropertyName
    End If
End Function

Public Function getEnableFunction() As String
    getEnableFunction = pEnableFunction
End Function

Public Function getFilter() As IFilter
    Set getFilter = pFilter
End Function

Public Function hasFilter() As Boolean
    hasFilter = (Not pFilter Is Nothing)
End Function

Public Function getTextAlignment() As fmTextAlign
    getTextAlignment = pTextAlignment
End Function

Public Function isUnderlined() As Boolean
    isUnderlined = pIsUnderlined
End Function

Public Function hasRightBorder() As Boolean
    hasRightBorder = pHasRightBorder
End Function

Public Function getFontColor() As Long
    getFontColor = pForeColor
End Function

Public Function getMousePointer() As Object
    Set getMousePointer = pMousePointer
End Function

Public Function getTooltip() As String
    getTooltip = pTooltip
End Function

Public Function getActionName(eventName As String) As String
    getActionName = F.dictionaries.getItem(pActions, eventName)
End Function

Public Function getNumberFormat() As String
    getNumberFormat = pNumberFormat
End Function

Public Function getLeft() As Single
    getLeft = pLeft
End Function

Public Function getIndent() As Single
    getIndent = pIndent
End Function

Public Function getFilterRegex() As String
    getFilterRegex = pFilterRegex
End Function


'</simple getters>





Public Function addAction(eventName As String, actionName As String) As DColumn
    Set addAction = Me
    Call F.dictionaries.addItem(pActions, eventName, actionName)
End Function

