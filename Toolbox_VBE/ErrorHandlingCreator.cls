VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandlingCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "ErrorHandlingCreator"
'----------------------------------------------------------------------------------------------------------

Public Sub addErrorHandlingToCurrentMethod()
    Dim method As EVbaMethod
    '------------------------------------------------------------------------------------------------------
    Set method = getTestMethod 'getActiveMethod
    Call method.analyzeBody
End Sub

Private Function getActiveMethod() As EVbaMethod
    Dim selection As ECodeSelection
    '------------------------------------------------------------------------------------------------------
    
    Set getActiveMethod = New EVbaMethod
    
    Set selection = Fn.Create.currentCodeSelection
    
    Stop
    
End Function

Private Function getTestMethod() As EVbaMethod
    Dim component As VBIDE.VBComponent
    Dim cm As VBIDE.CodeModule
    '------------------------------------------------------------------------------------------------------
    
    Set getTestMethod = New EVbaMethod
    
    Set component = F.Developer.getVbComponent(ThisWorkbook.VBProject, "ContextManager")
    Set cm = component.CodeModule
    Set getTestMethod = createVbaMethodByLine(cm, 35)
    
End Function
