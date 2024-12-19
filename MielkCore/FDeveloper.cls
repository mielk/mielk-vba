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


'[VALIDATIONS]
Public Function isValidComponent(component As Variant) As Boolean
    Const ILLEGAL_COMPONENT_NAME As String = "1?!"
    '------------------------------------------------------------------------------------------------------
    Dim name As String:                 name = ILLEGAL_COMPONENT_NAME
    '------------------------------------------------------------------------------------------------------
    
    If Not VBA.IsObject(component) Then
        isValidComponent = False
    ElseIf component Is Nothing Then
        isValidComponent = False
    ElseIf Not TypeOf component Is VBIDE.VBComponent Then
        isValidComponent = False
    Else
        On Error Resume Next
        name = component.name
        isValidComponent = (name <> ILLEGAL_COMPONENT_NAME)
    End If
    
End Function


'[CONVERTERS]
Public Function procTypeToString(procType As Long) As String
    Select Case procType
        Case vbext_pk_Get: procTypeToString = "GET"
        Case vbext_pk_Set: procTypeToString = "SET"
        Case vbext_pk_Let: procTypeToString = "LET"
        Case vbext_pk_Proc: procTypeToString = "PROC"
    End Select
End Function

Public Function getProcTypeFromName(name As String) As VBIDE.vbext_ProcKind
    Select Case VBA.LCase(name)
        Case "sub", "function":             getProcTypeFromName = vbext_pk_Proc
        Case "property get":                getProcTypeFromName = vbext_pk_Get
        Case "property let":                getProcTypeFromName = vbext_pk_Let
        Case "property set":                getProcTypeFromName = vbext_pk_Set
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




'[PROPERTIES]
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

Public Function getVbProjectByProjectName(name As String) As Object
    Dim names As VBA.Collection
    Dim varName As Variant
    Dim addIn As Excel.addIn
    Dim wkb As Excel.Workbook
    Dim project As VBIDE.VBProject
    '------------------------------------------------------------------------------------------------------
    
    
    '[Collect names of all add-ins and regular Excel books.] ----------------------------------------|
    Set names = New VBA.Collection                                                                  '|
    For Each addIn In Excel.Application.AddIns2                                                     '|
        Call names.Add(addIn.name)                                                                  '|
    Next addIn                                                                                      '|
                                                                                                    '|
    For Each wkb In Excel.Application.Workbooks                                                     '|
        Call names.Add(wkb.name)                                                                    '|
    Next wkb                                                                                        '|
    '------------------------------------------------------------------------------------------------|
    
    
    For Each varName In names
        
        '[Try to set reference to a file with such a name.] -----------------------------------------|
        On Error Resume Next                                                                        '|
        Set wkb = Application.Workbooks(varName)                                                    '|
        On Error GoTo 0                                                                             '|
        '--------------------------------------------------------------------------------------------|
        
        If f.Books.IsValid(wkb) Then
            Set project = wkb.VBProject
            If project Is Nothing Then
                ' error when creating reference to [project]
            ElseIf f.strings.compareStrings(project.name, name) Then
                Set getVbProjectByProjectName = project
                Exit For
            End If
        End If
        
    Next varName
    
    
End Function



Public Function getVbProject(Optional project As Variant) As Object 'VBIDE.VBProject
    If VBA.IsMissing(project) Then
        Set getVbProject = Excel.ThisWorkbook.VBProject
    ElseIf TypeOf project Is Excel.Workbook Then
        Set getVbProject = project.VBProject
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

Public Function getCodeModuleContent(codeModule As Object, _
                                Optional ByVal startLine As Long = 1, _
                                Optional ByVal linesCounter As Long = 0) As String
    Dim totalLines As Long
    '------------------------------------------------------------------------------------------------------
    
    With codeModule
        totalLines = .CountOfLines
        
        '[Normalize start line parameter] -----------------------|
        If startLine <= 0 Then                                  '|
            startLine = 1                                       '|
        ElseIf startLine > totalLines Then                      '|
            startLine = 1                                       '|
        End If                                                  '|
        '--------------------------------------------------------|
        
        '[Normalize lines counter parameter] --------------------|
        If linesCounter <= 0 Then                               '|
            linesCounter = totalLines - startLine + 1           '|
        ElseIf linesCounter + startLine > totalLines Then       '|
            linesCounter = totalLines - startLine + 1           '|
        End If                                                  '|
        '--------------------------------------------------------|
        
        getCodeModuleContent = .lines(startLine, linesCounter)
        
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
            Set getProjectMethods.item(key) = getCodeModuleMethods(getProjectMethods.item(key))
        Next key
    End If
    
End Function

Public Sub exportModules(book As Excel.Workbook, folderPath As String)
    Dim project As VBIDE.VBProject
    Dim component As VBIDE.VBComponent
    '------------------------------------------------------------------------------------------------------
    Dim subfolder As Scripting.folder
    Dim fileName As String
    '------------------------------------------------------------------------------------------------------
    
    Set project = getVbProject(book)
    If Not project Is Nothing Then
        Set subfolder = f.files.createFolder(f.files.concatenatePath(folderPath, project.name))
        For Each component In project.VBComponents
            If isComponentExportable(component) Then
                fileName = f.files.concatenatePath(subfolder.path, component.name & getComponentExtension(component))
                Call component.Export(fileName)
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


Public Function getProcType(cm As codeModule, line As Long) As VBIDE.vbext_ProcKind
    Const OPENING_REGEX_PATTERN As String = "^(?:|Public |Private |Friend )(Sub|Function|Property (?:G|L|S)et)"
    '------------------------------------------------------------------------------------------------------
    Dim text As String
    Dim lines As Variant
    Dim i As Long
    Dim openingText As String
    '------------------------------------------------------------------------------------------------------
    
    text = getCodeModuleContent(cm)
    lines = VBA.Split(text, VBA.vbCrLf)
    
    If f.arrays.isInBounds(line, 1, lines) Then
        For i = line To LBound(lines) Step -1
            openingText = f.regex.getFirstGroupMatch(VBA.CStr(lines(i)), OPENING_REGEX_PATTERN)
            If VBA.Len(openingText) Then
                getProcType = getProcTypeFromName(openingText)
                Exit For
            End If
        Next i
    Else
        'Invalid [line] index
        Stop
    End If
    
End Function




Public Sub clearCodeModule(cm As VBIDE.codeModule)
    If Not cm Is Nothing Then
        With cm
            Call .DeleteLines(1, .CountOfLines)
        End With
    End If
End Sub

Public Sub overrideCodeModuleContent(cm As VBIDE.codeModule, content As String)
    If Not cm Is Nothing Then
        With cm
            Call .DeleteLines(1, .CountOfLines)
            Call .AddFromString(content)
        End With
    End If
End Sub

Public Sub appendToCodeModule(cm As VBIDE.codeModule, content As String)
    If Not cm Is Nothing Then
        With cm
            Call .AddFromString(content)
        End With
    End If
End Sub



Public Sub changeClassInstancing(cl As VBIDE.VBComponent, Optional toPublic As Boolean = True)
    Const INSTANCING_PROPERTY_NAME As String = "Instancing"
    '------------------------------------------------------------------------------------------------------
    Dim value As Long
    Dim prop As Variant
    '------------------------------------------------------------------------------------------------------
    
    value = VBA.IIf(toPublic, 2, 1)
    
    For Each prop In cl.properties
        If prop.name = INSTANCING_PROPERTY_NAME Then
            prop.value = value
        End If
    Next prop
    
End Sub
