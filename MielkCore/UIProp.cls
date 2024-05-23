VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "UIProp"
'----------------------------------------------------------------------------------------------------------
Private pId As Long
Private pName As String
Private pIsSizeProperty As Boolean
Private pIsPositionProperty As Boolean
Private pIsLayoutProperty As Boolean
Private pIsLookProperty As Boolean
Private pInheritable As Boolean
Private pAliases As Scripting.Dictionary
'#If VBA7 Then
'    Private pConvertionFunction As LongPtr
'#Else
'    Private pConvertionFunction As Long
'#End If
Private pConvertionFunction As String
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    pInheritable = True
    Set pAliases = f.dictionaries.Create(False)
End Sub


'[SETTERS]
Friend Function setId(value As Long) As UIProp
    Set setId = Me
    pId = value
End Function

Friend Function setName(value As String) As UIProp
    Set setName = Me
    pName = value
End Function

Friend Function setIsInheritable(value As Boolean) As UIProp
    Set setIsInheritable = Me
    pInheritable = value
End Function

Friend Function setIsSizeProperty(value As Boolean) As UIProp
    Set setIsSizeProperty = Me
    pIsSizeProperty = value
End Function

Friend Function setIsPositionProperty(value As Boolean) As UIProp
    Set setIsPositionProperty = Me
    pIsPositionProperty = value
End Function

Friend Function setIsLayoutProperty(value As Boolean) As UIProp
    Set setIsLayoutProperty = Me
    pIsLayoutProperty = value
End Function

Friend Function setIsLookProperty(value As Boolean) As UIProp
    Set setIsLookProperty = Me
    pIsLookProperty = value
End Function

Friend Function setConvertionFunction(functionName As String) As UIProp
    Set setConvertionFunction = Me
    pConvertionFunction = functionName
End Function

Friend Function addAlias(value As String) As UIProp
    Set addAlias = Me
    With pAliases
        If Not .exists(value) Then
            Call .Add(value, value)
        End If
    End With
End Function




'[GETTERS]
Public Function getId() As Long
    getId = pId
End Function

Public Function getName() As String
    getName = pName
End Function

Public Function getLowerCaseName() As String
    getLowerCaseName = VBA.LCase$(pName)
End Function

Public Function isInheritable() As Boolean
    isInheritable = pInheritable
End Function

Public Function isSizeProperty() As Boolean
    isSizeProperty = pIsSizeProperty
End Function

Public Function isPositionProperty() As Boolean
    isPositionProperty = pIsPositionProperty
End Function

Public Function isLayoutProperty() As Boolean
    isLayoutProperty = pIsLayoutProperty
End Function

Public Function isLookProperty() As Boolean
    isLookProperty = pIsLookProperty
End Function

Public Function hasConvertionFunction() As Boolean
    hasConvertionFunction = VBA.Len(pConvertionFunction) > 0
End Function

Public Function getConvertionFunction() As String
    getConvertionFunction = pConvertionFunction
End Function

Public Function convertValue(value As Variant) As Variant
    If VBA.Len(pConvertionFunction) Then
        Call f.Variables.assign(convertValue, Excel.Application.run(pConvertionFunction, value))
    Else
        Call f.Variables.assign(convertValue, value)
    End If
End Function

Public Function getValueFromPropertiesDictionary(dict As Scripting.Dictionary) As Variant
    Dim keyAlias As Variant
    '----------------------------------------------------------------------------------------------------------
    
    If f.dictionaries.isNonEmptyDictionary(dict) Then
        For Each keyAlias In pAliases.Keys
            Call f.dictionaries.assignValueFromDictToVariable(dict, getValueFromPropertiesDictionary, _
                                    VBA.CStr(keyAlias), f.Reflection.getFunctionFullName(pConvertionFunction))
            If Not f.Variables.isUndefined(getValueFromPropertiesDictionary) Then Exit For
        Next keyAlias
    End If
    
End Function

Public Function getConvertedValue(value As Variant) As Variant
    If VBA.Len(pConvertionFunction) Then
        On Error GoTo ConversionException
        Call f.Variables.assign(getConvertedValue, Application.run(f.Reflection.getFunctionFullName(pConvertionFunction), value))
    Else
        Call f.Variables.assign(getConvertedValue, value)
    End If
    
'@@@
'-------------------------
ConversionException:
    
End Function

Public Function getAliasesString(Optional separator As String = ";") As String
    If Not pAliases Is Nothing Then
        getAliasesString = f.dictionaries.toString(pAliases, DictPart_KeyOnly, StringifyMode_Normal, _
                                                                separator, vbNullString, vbNullString)
    End If
End Function
