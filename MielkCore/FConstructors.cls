VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FConstructors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FConstructors"
'----------------------------------------------------------------------------------------------------------

Public Function Repository() As Repository
    Set Repository = New Repository
End Function

Public Function RibbonManager() As RibbonManager
    Set RibbonManager = New RibbonManager
End Function

Public Function DoubleDictionary() As DoubleDictionary
    Set DoubleDictionary = New DoubleDictionary
End Function

Public Function Coordinate(x As Single, y As Single) As Coordinate
    With Coordinate
        .x = x
        .y = y
    End With
End Function

Public Function Limit(Optional ByVal low As Variant, Optional ByVal high As Variant, _
                      Optional ByVal excludeLow As Boolean = False, _
                      Optional ByVal excludeHigh As Boolean = False) As Limit
    Set Limit = New Limit
    With Limit
        If Not VBA.IsMissing(low) Then Call .setLowValue(low, excludeLow)
        If Not VBA.IsMissing(high) Then Call .setHighValue(high, excludeHigh)
    End With
End Function

Public Function ValidationObject(Optional status As Boolean = True, Optional message As String) As ValidationObject
    Set ValidationObject = New ValidationObject
    With ValidationObject
        Call .setStatus(status)
        Call .setMessage(message)
    End With
End Function




Public Function ErrorObject(Optional className As String, Optional methodName As String, _
                      Optional number As Long, Optional description As String) As ErrorObject
    Set ErrorObject = New ErrorObject
    With ErrorObject
        Call .setClassName(className)
        Call .setMethodName(methodName)
        Call .setErrorNumber(number)
        Call .setErrorDescription(description)
    End With
End Function




Public Function ExcelFilter(sheet As Excel.Worksheet) As ExcelFilter
    Set ExcelFilter = New ExcelFilter
    With ExcelFilter
        Call .setSheet(sheet)
    End With
End Function


'
'Public Function ErrorManager() As ErrorManager
'    Set ErrorManager = New ErrorManager
'End Function
'
'Public Function Callback_____(methodName As String, Optional methodObject As Object) As Callback_____
'    Set Callback_____ = New Callback_____
'    With Callback_____
'        Call .setMethodName(methodName)
'        Call .setObject(methodObject)
'    End With
'End Function
'
'
