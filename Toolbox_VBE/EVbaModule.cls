VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EVbaModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "EVbaModule"
'----------------------------------------------------------------------------------------------------------
Private pProject As EVbaProject
Private pComponent As VBIDE.VBComponent
Private pName As String
'----------------------------------------------------------------------------------------------------------
Private pModuleDeclarations As String
Private pMethods As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------


'[SETTERS]
Public Function setProject(value As EVbaProject) As EVbaModule
    Set setProject = Me
    Set pProject = value
End Function

Public Function setName(value As String) As EVbaModule
    Set setName = Me
    pName = value
End Function

Public Function setVbComponentByName(name As String) As EVbaModule
    Set setVbComponentByName = Me
    If Not pProject Is Nothing Then
        Set pComponent = pProject.getComponentByName(name)
    End If
End Function

Public Function setVbComponent(component As VBIDE.VBComponent) As EVbaModule
    Set setVbComponent = Me
    If F.Dev.isValidComponent(component) Then
        Set pComponent = component
        pName = component.name
    End If
End Function



'[GETTERS]
Public Function getName() As String
    getName = pName
End Function

Public Function getComponent() As VBIDE.VBComponent
    Set getComponent = pComponent
End Function

Public Function getCodeModule() As VBIDE.CodeModule
    Set getCodeModule = pComponent.CodeModule
End Function

Public Function getDeclarations() As String
    getDeclarations = pModuleDeclarations
End Function

Public Function getMethods() As Scripting.Dictionary
    Set getMethods = pMethods
End Function

Public Function getMethod(key As Variant) As EVbaMethod
    Set getMethod = F.dictionaries.GetObject(pMethods, key)
End Function

Public Function getJsonDict() As Scripting.Dictionary
    Set getJsonDict = F.dictionaries.Create(False)
    With getJsonDict
        Call .Add(META_TAG_NAME, getName)
        Call .Add(META_TAG_DECLARATIONS, pModuleDeclarations)
        Call .Add(META_TAG_METHODS, getMethodsJsonDict)
    End With
End Function

Private Function getMethodsJsonDict() As Scripting.Dictionary
    Dim key As Variant
    Dim method As EVbaMethod
    '------------------------------------------------------------------------------------------------------
    Set getMethodsJsonDict = F.dictionaries.Create(False)
    For Each key In pMethods.Keys
        Set method = pMethods.item(key)
        Call getMethodsJsonDict.Add(key, method.getJsonDict)
    Next key
End Function

Public Function hasMethod(methodName As String) As Boolean
    hasMethod = pMethods.Exists(methodName)
End Function




'[LOADING]
Public Function load() As EVbaModule
    Set load = Me
    Call loadModuleDeclarations
    Call loadMethods
End Function

Private Sub loadModuleDeclarations()
    With getCodeModule
        If .CountOfDeclarationLines Then
            pModuleDeclarations = .lines(1, .CountOfDeclarationLines)
        End If
    End With
End Sub

Private Sub loadMethods()
    Dim totalLines As Long
    Dim line As Long
    Dim method As EVbaMethod
    '------------------------------------------------------------------------------------------------------
    Dim methodName As String
    Dim cm As VBIDE.CodeModule:             Set cm = getCodeModule
    '------------------------------------------------------------------------------------------------------
    
    Set pMethods = F.dictionaries.Create(False)
    
    With cm
        totalLines = .CountOfLines
        line = .CountOfDeclarationLines + 1
    End With
    
    Do Until line >= totalLines
        Set method = createVbaMethodByLine(Me, line)
        With method
            If .getProcType < 0 Or pMethods.Exists(.getKey) Then
                Do
                    line = line + 1
                    methodName = cm.ProcOfLine(line, vbext_pk_Proc)
                    If Not pMethods.Exists(methodName & "|0") Then
                        Exit Do
                    ElseIf Not pMethods.Exists(methodName & "|1") Then
                        Exit Do
                    ElseIf Not pMethods.Exists(methodName & "|2") Then
                        Exit Do
                    ElseIf Not pMethods.Exists(methodName & "|3") Then
                        Exit Do
                    ElseIf line >= totalLines Then
                        Exit Do
                    End If
                Loop
            Else
                Call pMethods.Add(.getKey, method)
                line = .getLastLineIndex
            End If
        End With
    Loop
    
End Sub

Public Sub loadFromDict(dict As Scripting.Dictionary)
    pModuleDeclarations = VBA.Replace( _
                                VBA.Replace( _
                                        F.dictionaries.GetItem(dict, META_TAG_DECLARATIONS), VBA.vbCrLf, vbNullString), VBA.vbLf, VBA.vbCrLf)
    Set pMethods = createMethodsDictionary(F.dictionaries.GetObject(dict, META_TAG_METHODS))
End Sub

Private Function createMethodsDictionary(dict As Scripting.Dictionary) As Scripting.Dictionary
    Dim key As Variant
    Dim subdict As Scripting.Dictionary
    Dim method As EVbaMethod
    '------------------------------------------------------------------------------------------------------
    Set createMethodsDictionary = F.dictionaries.Create(False)
    If F.dictionaries.isNonEmptyDictionary(dict) Then
        For Each key In dict.Keys
            Set subdict = dict.item(key)
            Set method = New EVbaMethod
            Stop
            'Call method.loadFromDict(subdict)
            Call createMethodsDictionary.Add(key, method)
        Next key
    End If
End Function


Public Function toString() As String
    Dim key As Variant
    Dim method As EVbaMethod
    '------------------------------------------------------------------------------------------------------
    
    toString = pModuleDeclarations
    For Each key In pMethods.Keys
        Set method = pMethods.item(key)
        toString = toString & VBA.vbCrLf & VBA.vbCrLf & method.toString(True, True, False)
    Next key
    
End Function




Public Function compare(module As EVbaModule) As ValidationObject
    Stop
End Function




Public Function appendErrorHandling(entryPoints As Scripting.Dictionary) As Scripting.Dictionary
    Const ERR_HANDLING_ALREADY_EXIST As String = "Error handling skipped - already existed | "
    Const ERR_HANDLING_INTERFACE_METHOD As String = "Error handling skipped - interface method"
    Const ERR_HANDLING_SIMPLE_METHOD As String = "Error handling skipped - simple method"
    Const ERR_HANDLING_ANNOTATION_NO_HANDLING As String = "Error handling skipped - no error handling annotation"
    Const ERR_HANDLING_ANNOTATION_CUSTOM_HANDLING As String = "Error handling skipped - custom error handling annotation"
    Const ERR_HANDLING_APPENDED As String = "Error handling appended"
    Const ERR_HANDLING_ONE_LINER As String = "Error handling skipped - one line method"
    '------------------------------------------------------------------------------------------------------
    Dim methods As VBA.Collection
    Dim method As EVbaMethod
    Dim methodFullName As String
    Dim isEntryPoint As Boolean
    Dim existingErrHandlingStatus As String
    '------------------------------------------------------------------------------------------------------
    Dim methodBodyWithErrHandling As String
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Set appendErrorHandling = F.dictionaries.Create(False)
    
    Set methods = F.Collections.Sort( _
                            F.dictionaries.toCollection(pMethods), _
                            False, _
                            "$.getLastLineIndex")
    
    For Each method In methods
        methodFullName = pName & "." & method.getName
        If method.hasNoErrorHandlingAnnotation Then
            Call F.dictionaries.addItem(appendErrorHandling, method.getName, ERR_HANDLING_ANNOTATION_NO_HANDLING)
        ElseIf method.hasCustomErrorHandlingAnnotation Then
            Call F.dictionaries.addItem(appendErrorHandling, method.getName, ERR_HANDLING_ANNOTATION_CUSTOM_HANDLING)
        ElseIf method.isOneLiner Then
            Call F.dictionaries.addItem(appendErrorHandling, method.getName, ERR_HANDLING_ONE_LINER)
        ElseIf method.hasAnyErrorHandling Then
            existingErrHandlingStatus = method.getErrorHandlingStatus
            Call F.dictionaries.addItem(appendErrorHandling, method.getName, ERR_HANDLING_ALREADY_EXIST & existingErrHandlingStatus)
        ElseIf method.isInterfaceMethod Then
            Call F.dictionaries.addItem(appendErrorHandling, method.getName, ERR_HANDLING_INTERFACE_METHOD)
        ElseIf method.isSimpleMethod Then
            Call F.dictionaries.addItem(appendErrorHandling, method.getName, ERR_HANDLING_SIMPLE_METHOD)
        Else

            methodBodyWithErrHandling = method.getBodyWithErrorHandlingAppended(isEntryPoint)
            Call method.replaceWithNewVersion(methodBodyWithErrHandling)
            
            Call F.dictionaries.addItem(appendErrorHandling, method.getName, ERR_HANDLING_APPENDED)
        End If
        
    Next method
    
ExitPoint:
    Exit Function

ErrHandler:
    Stop
    
End Function


