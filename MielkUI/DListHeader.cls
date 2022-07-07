VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DListHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "DListHeader"
'----------------------------------------------------------------------------------------------------------
Private pCaption As String
Private pProperty As Variant
Private pFilterable As Boolean
Private pSortable As Boolean
Private pWidth As Single
'[Events] -------------------------------------------------------------------------------------------------
Event AfterCaptionChanged()
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    '[Default values]
    pFilterable = False
    pSortable = True
    pWidth = 1
End Sub




'[SETTERS]
Public Function setCaption(value As String) As DListHeader
    Set setCaption = Me
    pCaption = value
    RaiseEvent AfterCaptionChanged
End Function

Public Function setProperty(value As Variant) As DListHeader
    Set setProperty = Me
    pProperty = value
End Function

Public Function setIsFilterable(value As Boolean) As DListHeader
    Set setIsFilterable = Me
    pFilterable = value
End Function

Public Function setIsSortable(value As Boolean) As DListHeader
    Set setIsSortable = Me
    pSortable = value
End Function

Public Function setWidth(value As Single) As DListHeader
    Set setWidth = Me
    pWidth = value
End Function



'[GETTERS]
Public Function getCaption() As String
    getCaption = pCaption
End Function

Public Function getProperty() As Variant
    getProperty = pProperty
End Function

Public Function isFilterable() As Boolean
    isFilterable = pFilterable
End Function

Public Function isSortable() As Boolean
    isSortable = pSortable
End Function

Public Function getWidth() As Single
    getWidth = pWidth
End Function
