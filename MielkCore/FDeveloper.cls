VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FDeveloper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FDeveloper"
    '----------------------------------------------------------------------------------------------------------


'[CONVERTERS]
Public Function procTypeToString(procType As Long) As String
    Select Case procType
        Case vbext_pk_Get: procTypeToString = "GET"
        Case vbext_pk_Set: procTypeToString = "SET"
        Case vbext_pk_Let: procTypeToString = "LET"
        Case vbext_pk_Proc: procTypeToString = "PROC"
    End Select
End Function

Private Function isComponentExportable(component As VBIDE.VBComponent) As Boolean
    If component.Type = vbext_ct_Document Then
        isComponentExportable = False
    ElseIf component.Type = vbext_ct_ActiveXDesigner Then
        isComponentExportable = False
    Else
        isComponentExportable = True
    End If
End Function

Private Function getComponentExtension(component As VBIDE.VBComponent) As String
    Select Case component.Type
        Case vbext_ct_ClassModule:      getComponentExtension = ".cls"
        Case vbext_ct_StdModule:        getComponentExtension = ".bas"
        Case vbext_ct_MSForm:           getComponentExtension = ".frm"
        Case vbext_ct_Document:         getComponentExtension = vbNullString
    End Select
End Function

Public Function getProcTypes() As Variant
    getProcTypes = VBA.Array(vbext_pk_Proc, vbext_pk_Let, vbext_pk_Get, vbext_pk_Set)
End Function


'[CHECKING DATA TYPE]
Public Function isVbProject(value As Variant) As Boolean
    Const VB_PROJECT_CLASS_NAME As String = "VbProject"
    '------------------------------------------------------------------------------------------------------
    isVbProject = f.strings.compareStrings(VBA.TypeName(value), VB_PROJECT_CLASS_NAME)
End Function

Public Function isVbComponent(value As Variant) As Boolean
    Const VB_COMPONENT_CLASS_NAME As String = "VbComponent"
    '------------------------------------------------------------------------------------------------------
    isVbComponent = f.strings.compareStrings(VBA.TypeName(value), VB_COMPONENT_CLASS_NAME)
End Function

Public Function isVbCodeModule(value As Variant) As Boolean
    Const VB_CODE_MODULE_CLASS_NAME As String = "CodeModule"
    '------------------------------------------------------------------------------------------------------
    isVbCodeModule = f.strings.compareStrings(VBA.TypeName(value), VB_CODE_MODULE_CLASS_NAME)
End Function




'[PROJECTS]

Public Function getVbProjectByName(name As String, Optional ByVal instance As Excel.Application) As Object 'VBIDE.VBProject
    Dim wkb As Excel.Workbook
    '------------------------------------------------------------------------------------------------------
    
    If instance Is Nothing Then Set instance = Excel.Application
    
    On Error Resume Next
    Set wkb = instance.Workbooks(name)
    If wkb Is Nothing Then
        Set wkb = f.Books.open_(name, False, instance)
    End If
    
    If Not wkb Is Nothing Then
        Set getVbProjectByName = getVbProject(wkb)
    End If
    
End Function

Public Function getVbProject(Optional project As Variant) As Object 'VBIDE.VBProject
    If VBA.IsMissing(project) Then
        Set getVbProject = Excel.ThisWorkbook.vbProject
    ElseIf TypeOf project Is Excel.Workbook Then
        Set getVbProject = project.vbProject
    ElseIf isVbProject(project) Then
        Set getVbProject = project
    End If
End Function

Public Function getVbComponentsDictionary(project As Variant) As Scripting.Dictionary
    Dim projectObject As Object         'VBIDE.VBProject
    Dim component As Object             'VBIDE.VBComponent
    '------------------------------------------------------------------------------------------------------
    
    Set getVbComponentsDictionary = VBA.CreateObject("Scripting.Dictionary")
    Set projectObject = getVbProject(project)
    If Not projectObject Is Nothing Then
        For Each component In projectObject.VBComponents
            Call getVbComponentsDictionary.Add(component.name, component)
        Next component
    End If
        
End Function

Public Function getVbComponent(project As Variant, componentName As String) As Object 'VBIDE.VBComponent
    Dim component As Object             'VBIDE.VBComponent
    '------------------------------------------------------------------------------------------------------
    
    If Not project Is Nothing Then
        For Each component In project.VBComponents
            If f.strings.compareStrings(componentName, component.name) Then
                Set getVbComponent = component
                Exit For
            End If
        Next component
    End If
    
End Function

Public Function getVbCodeModuleByProjectAndName(project As Variant, componentName As String) As Object 'VBIDE.codeModule
    Dim component As Object             'VBIDE.VBComponent
    '------------------------------------------------------------------------------------------------------
    
    Set component = getVbComponent(project, componentName)
    If Not component Is Nothing Then
        Set getVbCodeModuleByProjectAndName = component.codeModule
    End If
    
End Function

Public Function getVbCodeModuleByObject(object As Variant) As Object 'VBIDE.codeModule
    If Not object Is Nothing Then
        If isVbCodeModule(object) Then
            Set getVbCodeModuleByObject = object
        ElseIf isVbComponent(object) Then
            Set getVbCodeModuleByObject = object.codeModule
        End If
    End If
End Function

Public Function getCodeModuleContent(codeModule As Object) As String
    With codeModule
        getCodeModuleContent = .lines(1, .CountOfLines)
    End With
End Function

Public Function getCodeModuleMethods(cm As Object) As Scripting.Dictionary
    Dim PROC_TYPES As Variant: PROC_TYPES = VBA.Array(vbext_pk_Proc, vbext_pk_Let, vbext_pk_Get, vbext_pk_Set)
    Dim varProcType As Variant
    '------------------------------------------------------------------------------------------------------
    Dim previousLine As Long
    Dim line As Long
    Dim methodName As String
    Dim methodBodyLines As Integer
    Dim methodBody As String
    Dim procLines As Long
    '------------------------------------------------------------------------------------------------------
    
    Set getCodeModuleMethods = New Scripting.Dictionary
    Set cm = getVbCodeModuleByObject(cm)
    
    If Not cm Is Nothing Then
        line = cm.CountOfDeclarationLines + 1
        With cm
            Do While line < .CountOfLines
                previousLine = line
                methodName = .ProcOfLine(line, vbext_pk_Proc)
                methodBodyLines = 0
                methodBody = vbNullString
                
                For Each varProcType In PROC_TYPES
                    On Error Resume Next
                    methodBodyLines = .ProcCountLines(methodName, varProcType) - .ProcBodyLine(methodName, varProcType) + .ProcStartLine(methodName, varProcType)
                    If methodBodyLines Then
                        methodBody = .lines(.ProcBodyLine(methodName, varProcType), methodBodyLines)
                        line = line + .ProcCountLines(methodName, varProcType)
                        Exit For
                    End If
                Next varProcType
                
                Call f.dictionaries.addItem(getCodeModuleMethods, methodName, methodBody, duplicateBehaviour_Skip)
                
                If line = previousLine Then line = line + 1
                
            Loop
        End With
        
    End If
End Function


Public Function getProjectMethods(book As Excel.Workbook) As Scripting.Dictionary
    Dim project As Object
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set project = getVbProject(book)
    If Not project Is Nothing Then
        Set getProjectMethods = getVbComponentsDictionary(project)
        For Each key In getProjectMethods.Keys
            Set getProjectMethods.Item(key) = getCodeModuleMethods(getProjectMethods.Item(key))
        Next key
    End If
    
End Function

Public Sub exportModules(book As Excel.Workbook, folderPath As String)
    Dim project As VBIDE.vbProject
    Dim component As VBIDE.VBComponent
    '------------------------------------------------------------------------------------------------------
    Dim subfolder As Scripting.folder
    Dim FileName As String
    '------------------------------------------------------------------------------------------------------
    
    Set project = getVbProject(book)
    If Not project Is Nothing Then
        Set subfolder = f.files.createFolder(f.files.concatenatePath(folderPath, project.name))
        For Each component In project.VBComponents
            If isComponentExportable(component) Then
                FileName = f.files.concatenatePath(subfolder.path, component.name & getComponentExtension(component))
                Call component.Export(FileName)
            End If
        Next component
    End If

End Sub


Public Function getReferenceType(ref As VBIDE.Reference, projectTag As String) As ReferenceTypeEnum
    Const FILE_NAME_PATTERN As String = "\\\\{0}-[^\\\\]*\.xl[as]m$"
    '------------------------------------------------------------------------------------------------------
    Dim RefName As String
    Dim libraryRegexPattern As String
    Dim projectRegexPattern As String
    '------------------------------------------------------------------------------------------------------
    
    RefName = ref.FullPath
    libraryRegexPattern = f.strings.format(FILE_NAME_PATTERN, LIBRARY_NAME)
    projectRegexPattern = f.strings.format(FILE_NAME_PATTERN, projectTag)
    
    If f.regex.checkIfMatch(RefName, libraryRegexPattern) Then
        getReferenceType = ReferenceType_InternalLibrary
    ElseIf f.regex.checkIfMatch(RefName, projectRegexPattern) Then
        getReferenceType = ReferenceType_Project
    Else
        getReferenceType = ReferenceType_ExternalLibrary
    End If
    
End Function
