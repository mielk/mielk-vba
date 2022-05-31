VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FReflection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FReflection"
    '----------------------------------------------------------------------------------------------------------

Public Function getFunctionFullName(functionName As String, Optional sourceBook As Excel.Workbook) As String
    Const NAME_TEMPLATE As String = "'{0}'!{1}"
    '------------------------------------------------------------------------------------------------------
    Dim wkb As Excel.Workbook
    '------------------------------------------------------------------------------------------------------
    If VBA.Len(functionName) Then
        Set wkb = VBA.IIf(sourceBook Is Nothing, Excel.ThisWorkbook, sourceBook)
        getFunctionFullName = F.strings.format(NAME_TEMPLATE, wkb.name, functionName)
    End If
End Function


Public Function modifyValue(originalValue As Variant, ByVal functionName As String, _
                            Optional functionInvoker As Object, _
                            Optional params As Scripting.Dictionary) As Variant
    Dim invoker As Object
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(functionName) Then
        If F.strings.startsWith(functionName, LOCAL_METHOD) Then
            If functionInvoker Is Nothing Then Set invoker = originalValue Else Set invoker = functionInvoker
            functionName = VBA.mid$(functionName, VBA.Len(LOCAL_METHOD) + 1)
            If params Is Nothing Then
                Call F.Variables.assign(modifyValue, VBA.CallByName(invoker, functionName, VbMethod))
            Else
                Call F.Variables.assign(modifyValue, VBA.CallByName(invoker, functionName, VbMethod, params))
            End If
        Else
            If params Is Nothing Then
                Call F.Variables.assign(modifyValue, Excel.Application.Run(functionName, originalValue))
            Else
                Call F.Variables.assign(modifyValue, Excel.Application.Run(functionName, originalValue, params))
            End If
        End If
    Else
        Call F.Variables.assign(modifyValue, originalValue)
    End If
    
End Function

