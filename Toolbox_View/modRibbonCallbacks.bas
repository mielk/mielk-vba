Attribute VB_Name = "modRibbonCallbacks"
Option Explicit

Private Const CLASS_NAME As String = "modRibbonCallbacks"
'----------------------------------------------------------------------------------------------------------


'[Common generic callback functions]

Public Function getUserNameRibbonLabel() As String
    getUserNameRibbonLabel = F.System.getUserUid
End Function

Public Function getEnvironmentName() As String
    Const TEXT_PATTERN As String = "{0} ({1})"
    '------------------------------------------------------------------------------------------------------
    Dim path As String
    Dim envName As String
    '------------------------------------------------------------------------------------------------------
    
    path = Paths.EnvironmentNameFilePath
    If F.files.FileExists(path) Then
        envName = F.TextFiles.readTextFile(path)
    End If
    
    If VBA.Len(envName) Then
        Call setParentApp           'Make sure that it is already established
        getEnvironmentName = F.Strings.Format(TEXT_PATTERN, envName, App.getVersion)
    Else
        getEnvironmentName = "?????"
    End If
    
End Function









'[Permissions & actions]
Public Function checkUserPermission(controlId As String) As Boolean
    Select Case controlId
        Case "control.id":                                                  checkUserPermission = True
    End Select
End Function

Private Sub action_toolbox(ByVal control As IRibbonControl)
    Static inProgress As Boolean
    '----------------------------------------------------------------------------------------------------------
    
    If Not inProgress Then
        inProgress = True
        
        Call ErrorManager.Clear
        Call setParentApp               'To make sure that the information was not cleared.
        
        Select Case control.ID
            '[Code]
            Case "button.code.createNewProject":                        Call Toolbox.createNewProject(getSheetsDictionary)
            Case "button.code.compactFiles":                            Call CodeCompactor.run
            Case "button.code.compareCode":                             Call CodeComparisonManager.run
            '[Access]
            Case "button.access.printDbStructure":                      Stop
            Case "button.access.compareDatabases":                      Stop
            'Case "button.access.relinkDatabase":                        Call DbRelinker.run
        End Select
        inProgress = False
    Else
        Debug.Print "Action in progress"
    End If
    
End Sub
