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
        getFunctionFullName = f.strings.format(NAME_TEMPLATE, wkb.name, functionName)
    End If
End Function


Public Function modifyValue(originalValue As Variant, ByVal functionName As String, _
                            Optional functionInvoker As Object, _
                            Optional params As Scripting.Dictionary) As Variant
    Dim invoker As Object
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(functionName) Then
        If f.strings.startsWith(functionName, LOCAL_METHOD) Then
            If functionInvoker Is Nothing Then Set invoker = originalValue Else Set invoker = functionInvoker
            functionName = VBA.mid$(functionName, VBA.Len(LOCAL_METHOD) + 1)
            If Not invoker Is Nothing Then
                If params Is Nothing Then
                    Call f.Variables.assign(modifyValue, VBA.CallByName(invoker, functionName, VbMethod))
                Else
                    Call f.Variables.assign(modifyValue, VBA.CallByName(invoker, functionName, VbMethod, params))
                End If
            End If
        Else
            If params Is Nothing Then
                Call f.Variables.assign(modifyValue, Excel.Application.run(functionName, originalValue))
            Else
                Call f.Variables.assign(modifyValue, Excel.Application.run(functionName, originalValue, params))
            End If
        End If
    Else
        Call f.Variables.assign(modifyValue, originalValue)
    End If
    
End Function



Public Function runMethod(item As Variant, ByVal functionName As String, Optional functionInvoker As Object, _
                            Optional params As Scripting.Dictionary) As Variant
    Dim invoker As Object
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(functionName) Then
        If f.strings.startsWith(functionName, LOCAL_METHOD) Then
            If functionInvoker Is Nothing Then Set invoker = item Else Set invoker = functionInvoker
            functionName = VBA.mid$(functionName, VBA.Len(LOCAL_METHOD) + 1)
            If params Is Nothing Then
                Call VBA.CallByName(invoker, functionName, VbMethod)
            Else
                Call VBA.CallByName(invoker, functionName, VbMethod, params)
            End If
        Else
            If params Is Nothing Then
                Call Excel.Application.run(functionName, item)
            Else
                Call Excel.Application.run(functionName, item, params)
            End If
        End If
    End If
    
End Function

Public Function convertToLocalFormula(formula As String, Optional isR1C1Formula As Boolean = True) As String
    Dim cell As Excel.Range
    Dim alerts As Boolean
    '------------------------------------------------------------------------------------------------------
    
    Set cell = wksOperations.cells(1, 1)
    
    alerts = Application.displayAlerts
    Application.displayAlerts = False
    
    On Error Resume Next
    
    If isR1C1Formula Then
        cell.Formula2R1C1 = formula
        convertToLocalFormula = cell.Formula2R1C1Local
    Else
        cell.Formula2 = formula
        convertToLocalFormula = cell.Formula2Local
    End If
    On Error GoTo 0
    
    
    Application.displayAlerts = alerts
    
End Function
