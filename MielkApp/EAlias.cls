VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EAlias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IEditItem

Private Const CLASS_NAME As String = "EAlias"
'----------------------------------------------------------------------------------------------------------
Private Const ADD_ALIAS_WINDOW_TITLE As String = "AliasesPanel.AddAlias.WindowTitle"
Private Const EDIT_ALIAS_WINDOW_TITLE As String = "AliasesPanel.EditAlias.WindowTitle"
Private Const ERROR_EMPTY_FIELD As String = "AliasesPanel.EditAlias.ValidationErrors.EmptyAlias"
Private Const WARNING_ALREADY_ASSIGNED As String = "AliasesPanel.EditAlias.ValidationErrors.AlreadyAssigned"
'----------------------------------------------------------------------------------------------------------
Private pParent As IAliasable
Private pId As Long
Private pAlias As String
Private pConfirmWhenOverriding As Boolean
'----------------------------------------------------------------------------------------------------------



'[Setters]
Public Function setParent(value As Variant) As EAlias
    Set setParent = Me
    
    If F.Variables.isUndefined(value) Then
    ElseIf Not VBA.IsObject(value) Then
    ElseIf Not TypeOf value Is IAliasable Then
    Else
        Set pParent = value
    End If
End Function

Public Function setAlias(value As String) As EAlias
    Set setAlias = Me
    pAlias = value
End Function

Public Function setConfirmWhenOverriding(value As Boolean) As EAlias
    Set setConfirmWhenOverriding = Me
    pConfirmWhenOverriding = value
End Function

Public Function setValue(propName As String, value As Variant) As EAlias
    Select Case VBA.LCase$(propName)
        Case VBA.LCase$(props.alias.getName):               Call setAlias(VBA.CStr(value))
        Case VBA.LCase$(props.parent.getName):              Call setParent(value)
    End Select
End Function




'[Getters]
Public Function getParent() As IAliasable
    Set getParent = pParent
End Function

Public Function getAlias() As String
    getAlias = pAlias
End Function

Public Function getValue(propName As String) As Variant
    Select Case VBA.LCase$(propName)
        Case VBA.LCase$(TAG_ALIAS):                     getValue = getAlias
        Case VBA.LCase$(TAG_PARENT):                    Set getValue = getParent
        Case VBA.LCase$(TAG_CONFIRM_IF_WARNINGS):       getValue = confirmOverriding
    End Select
End Function







'[Editing]
Public Function clone(Optional includeId As Boolean = True) As EAlias
    Set clone = New EAlias
    With clone
        Call .setValue(TAG_PARENT, getParent)
        Call .setValue(TAG_ALIAS, getAlias)
        Call .setConfirmWhenOverriding(pConfirmWhenOverriding)
    End With
End Function

Public Function edit() As ValidationObject
    Dim editForm As WEditEntity
    Dim editedItem As EAlias
    Dim dto As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------

    On Error GoTo ErrHandler

    Set editForm = getEditForm
    If Not editForm Is Nothing Then
        With editForm
            Call .run(vbModal, UI.Create.LocationPropsForCurrentCursor)
            If Not .isCancelledByUser Then
                pAlias = .getEditItem.getValue(TAG_ALIAS)
            End If
            Set edit = F.Create.ValidationObject(Not .isCancelledByUser)
        End With
    End If

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then
        Set edit = F.Create.ValidationObject(False, errDescription)
    End If
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = Err.Number: errDescription = Err.Description
    'call ErrorManager.addError(CLASS_NAME, metho
    GoTo ExitPoint

End Function

Private Function getEditForm() As WEditEntity
    Set getEditForm = New WEditEntity
    With getEditForm
        Call .setWindowTitle(getWindowTitle)
        Call .setLinesDefinitions(getEditLinesDefinitions)
        Call .setItem(Me)
    End With
End Function

Public Function confirmOverriding() As Boolean
    Const QUESTION_TITLE As String = "AliasesPanel.EditAlias.OverrideQuestionTitle"
    Const MESSAGE_PATTERN As String = "AliasesPanel.EditAlias.OverrideQuestion"
    '------------------------------------------------------------------------------------------------------
    Dim display As String
    Dim displayPrevious As String
    Dim message As String
    '------------------------------------------------------------------------------------------------------
    
    If pConfirmWhenOverriding Then
        display = pParent.getLabel
        displayPrevious = pParent.getAliasCurrentParent(pAlias, True)
        message = F.Strings.format(MsgService.getText(MESSAGE_PATTERN), pAlias, displayPrevious, display)
        confirmOverriding = F.Dialogs.getUserConfirmation(message, MsgService.getText(QUESTION_TITLE))
    Else
        confirmOverriding = True
    End If
    
End Function

Private Function getWindowTitle() As String
    If pId Then
        getWindowTitle = F.Strings.format(MsgService.getText(EDIT_ALIAS_WINDOW_TITLE), getAlias)
    Else
        getWindowTitle = MsgService.getText(ADD_ALIAS_WINDOW_TITLE)
    End If
End Function

Private Function getEditLinesDefinitions() As VBA.Collection
    Dim validationParams As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------

    Set validationParams = F.dictionaries.createSingleEntryDict(TAG_ALIASES, pParent.getAliasesForAllObjects)
    
    Set getEditLinesDefinitions = New VBA.Collection
    With getEditLinesDefinitions
        '[Alias]
        Call .Add(F.dictionaries.createWithItems(False, _
                        KeyValue(TAG_CONTROL_TYPE, ControlType_TextBox), _
                        KeyValue(TAG_CAPTION, MsgService.getText("AliasesPanel.EditAlias.Labels.Alias")), _
                        KeyValue(TAG_PROPERTY, TAG_ALIAS), _
                        KeyValue(TAG_VALIDATION_FUNCTION, "validateName"), _
                        KeyValue(TAG_VALIDATION_PARAMS, validationParams), _
                        KeyValue(TAG_CONFIRM_WITH_ENTER, True)))
    End With

End Function

Private Function getAliasesPanel() As XAliasesPanel
    Set getAliasesPanel = New XAliasesPanel
    With getAliasesPanel
        Call .setItem(Me)
    End With
End Function



'[Validation]
Public Function getValidation(validationFunction As String, Optional validationParams As Scripting.Dictionary) As ValidationObject
    Set getValidation = VBA.CallByName(Me, validationFunction, VbMethod, validationParams)
End Function

Public Function validateName(Optional validationParams As Scripting.Dictionary) As ValidationObject
    Static aliases As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim name As String
    Dim id As Long
    Dim item As IAliasable
    Dim currentAliasAssignment As Variant
    '------------------------------------------------------------------------------------------------------

    If aliases Is Nothing Then
        Set aliases = F.dictionaries.GetObject(validationParams, TAG_ALIASES)
    End If
    name = F.dictionaries.getItem(validationParams, VALUE_____)

    If VBA.Len(name) = 0 Then
        Set validateName = F.Create.ValidationObject(False, MsgService.getText(ERROR_EMPTY_FIELD))
    Else
        If Not aliases.Exists(name) Then
            Set validateName = F.Create.ValidationObject(True)
        Else
            Set item = F.dictionaries.GetObject(aliases, name)
            If item Is Nothing Then
                Set validateName = F.Create.ValidationObject(True)
            Else
                id = item.getId
                If id = pParent.getId Then
                    Set validateName = F.Create.ValidationObject(True)
                Else
                    currentAliasAssignment = pParent.getAliasCurrentParent(name, True)
                    Set validateName = F.Create.ValidationObject(WARNING_CODE, _
                                            F.Strings.format(MsgService.getText(WARNING_ALREADY_ASSIGNED), currentAliasAssignment))
                End If
            End If
        End If
    End If

End Function





'[Implementation of IEditItem]
Private Function IEditItem_clone() As MielkCore.IEditItem
    Set IEditItem_clone = clone
End Function

Private Sub IEditItem_delete()
    'N/A
End Sub

Private Function IEditItem_edit(Optional params As Scripting.Dictionary) As MielkCore.ValidationObject
    Set IEditItem_edit = edit
End Function

Private Function IEditItem_getId() As Long
    IEditItem_getId = pId
End Function

Private Function IEditItem_getName() As String
    IEditItem_getName = getAlias
End Function

Private Function IEditItem_getObjectsAffected() As Collection
    Set IEditItem_getObjectsAffected = pParent.getObjectsAffectedByReassignment
End Function

Private Function IEditItem_getValidation(validationFunction As String, Optional validationParams As Scripting.IDictionary) As MielkCore.ValidationObject
    Set IEditItem_getValidation = getValidation(validationFunction, validationParams)
End Function

Private Function IEditItem_getValue(valueKey As String) As Variant
    Call F.Variables.assign(IEditItem_getValue, getValue(valueKey))
End Function

Private Sub IEditItem_runAction(actionName As String, Optional params As Scripting.IDictionary)
    Call VBA.CallByName(Me, actionName, VbMethod, params)
End Sub

Private Sub IEditItem_updateValue(valueKey As String, value As Variant)
    Call setValue(valueKey, value)
End Sub

