Attribute VB_Name = "modConstructors"
Option Explicit

Private Const CLASS_NAME As String = "modConstructors"
'----------------------------------------------------------------------------------------------------------



'[SERVICES]

Public Function ContextManager() As ContextManager
    Static instance As ContextManager
    '------------------------------------------------------------------------------------------------------
    
    If instance Is Nothing Then
        Set instance = New ContextManager
    End If
    Set ContextManager = instance
    
End Function

Public Function Fn() As Functions
    Static instance As Functions
    '------------------------------------------------------------------------------------------------------
    
    If instance Is Nothing Then
        Set instance = New Functions
    End If
    Set Fn = instance
    
End Function

Public Function Props_Vbe() As CVbeProperties
    Static instance As CVbeProperties
    '------------------------------------------------------------------------------------------------------
    
    If instance Is Nothing Then
        Set instance = New CVbeProperties
    End If
    Set Props_Vbe = instance
    
End Function

Public Function ErrorHandling() As ErrorHandlingCreator
    Static instance As ErrorHandlingCreator
    '------------------------------------------------------------------------------------------------------
    
    If instance Is Nothing Then
        Set instance = New ErrorHandlingCreator
    End If
    Set ErrorHandling = instance
    
End Function

Public Function CodeFrameGenerator() As CodeFrameGenerator
    Static instance As CodeFrameGenerator
    '------------------------------------------------------------------------------------------------------
    
    If instance Is Nothing Then
        Set instance = New CodeFrameGenerator
    End If
    Set CodeFrameGenerator = instance
    
End Function

Public Function VbaCodeParser() As VbaCodeParser
    Static instance As VbaCodeParser
    '------------------------------------------------------------------------------------------------------
    
    If instance Is Nothing Then
        Set instance = New VbaCodeParser
    End If
    Set VbaCodeParser = instance
    
End Function

Public Function ClassGenerator() As ClassGenerator
    Static instance As ClassGenerator
    '------------------------------------------------------------------------------------------------------
    
    If instance Is Nothing Then
        Set instance = New ClassGenerator
    End If
    Set ClassGenerator = instance
    
End Function

Public Function ModulesExporter() As ModulesExporter
    Static instance As ModulesExporter
    '------------------------------------------------------------------------------------------------------
    
    If instance Is Nothing Then
        Set instance = New ModulesExporter
    End If
    Set ModulesExporter = instance
    
End Function





'[ENTITY INSTANCES]
Public Function createVbaProjectFromPath(path As String) As EVbaProject
    Set createVbaProjectFromPath = New EVbaProject
    With createVbaProjectFromPath
        Call .setPath(path)
    End With
End Function

Public Function createVbaProjectFromBook(book As Excel.Workbook) As EVbaProject
    Set createVbaProjectFromBook = New EVbaProject
    With createVbaProjectFromBook
        Call .setBook(book)
    End With
End Function

Public Function createVbaModule(ByVal component As Object) As EVbaModule
    Set createVbaModule = New EVbaModule
    
    If TypeOf component Is EVbaModule Then
        Set createVbaModule = component
    ElseIf TypeOf component Is VBIDE.CodeModule Then
        Call createVbaModule.setVbComponent(component.parent)
    ElseIf TypeOf component Is VBIDE.VBComponent Then
        Call createVbaModule.setVbComponent(component)
    End If
    
End Function

Public Function createVbaModuleByName(project As EVbaProject, name As String) As EVbaModule
    Set createVbaModuleByName = New EVbaModule
    With createVbaModuleByName
        Call .setProject(project)
        Call .setVbComponentByName(name)
    End With
End Function

Public Function createVbaMethod() As EVbaMethod
    Set createVbaMethod = New EVbaMethod
End Function

Public Function createVbaMethodByLine(ByVal component As Object, line As Long) As EVbaMethod
    Set createVbaMethodByLine = createVbaMethod. _
                                    setModule(createVbaModule(component)). _
                                    readByLineNumber(line)
End Function

Public Function createVbaVariable(method As EVbaMethod) As EVbaVariable
    Set createVbaVariable = New EVbaVariable
    With createVbaVariable
        Call .setMethod(method)
    End With
End Function
