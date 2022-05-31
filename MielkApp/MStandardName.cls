VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MStandardName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
'
'Private Const CLASS_NAME As String = "MStandardName"
''----------------------------------------------------------------------------------------------------------
'Private pId As Long
'Private pName As String
'Private pIsUnassigned As Boolean
'Private pRawNames As MielkCore.HashMap
''[Events] -------------------------------------------------------------------------------------------------
'Event AfterContentChange()
''----------------------------------------------------------------------------------------------------------
'
'
'
''[INITIALIZER]
'Private Sub Class_Initialize()
'    Call initializeContainers
'End Sub
'
'Private Sub initializeContainers()
'    Set pRawNames = MielkCore.createHashMap
'End Sub
'
'
'
'
''[SETTERS]
'Public Function setId(value As Long) As MStandardName
'    Set setId = Me
'    pId = value
'End Function
'
'Public Function setName(value As String) As MStandardName
'    Set setName = Me
'    pName = value
'End Function
'
'Public Function setIsUnassigned(value As Boolean) As MStandardName
'    Set setIsUnassigned = Me
'    pIsUnassigned = value
'End Function
'
'
''[GETTERS]
'Public Function getName() As String
'    getName = pName
'End Function
'
'Public Function getId() As Long
'    getId = pId
'End Function
'
'Public Function isUnassigned() As Boolean
'    isUnassigned = pIsUnassigned
'End Function
'
'Public Function getRawNames() As VBA.Collection
'    Set getRawNames = pRawNames.toCollection
'End Function
'
'Public Function hasRawName(name As String) As Boolean
'    hasRawName = pRawNames.Exists(name, DictEntry_Key)
'End Function
'
'
'Public Sub ClearContent()
'    Call pRawNames.Clear
'End Sub
'
'Public Sub addRawName(value As Variant)
'    Call pRawNames.Add(MielkCore.trimFull(MielkCore.stringify(value)))
'    RaiseEvent AfterContentChange
'End Sub
'
'Public Sub addRawNames(container As Variant)
'    Call pRawNames.addAll(container)
'    RaiseEvent AfterContentChange
'End Sub
'
'Public Sub deleteRawName(value As Variant)
'    Call pRawNames.Remove(value, DictEntry_Key)
'    RaiseEvent AfterContentChange
'End Sub
'
'Public Sub deleteRawNames(values As VBA.Collection)
'    Dim varValue As Variant
'    '------------------------------------------------------------------------------------------------------
'    For Each varValue In values
'        Call pRawNames.Remove(varValue, DictEntry_Key)
'    Next varValue
'    RaiseEvent AfterContentChange
'End Sub
'
