VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "Main"
'[Services] -----------------------------------------------------------------------------------------------
Private pConstructors As FConstructors
Private pFunctions As FFunctions
Private pStyling As FStyling
'[References] ---------------------------------------------------------------------------------------------
Private pRepoEditBook As Excel.Workbook
'[References] ---------------------------------------------------------------------------------------------
Private pRepoManagers As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    '[Services]
    Set pConstructors = New FConstructors
    Set pFunctions = New FFunctions
    Set pStyling = New FStyling
    '[Collections]
    Set pRepoManagers = F.dictionaries.Create(False)
End Sub


'[Services]
Public Function Create() As FConstructors
    If pConstructors Is Nothing Then Set pConstructors = New FConstructors
    Set Create = pConstructors
End Function

Public Function Functions() As FFunctions
    If pFunctions Is Nothing Then Set pFunctions = New FFunctions
    Set Functions = pFunctions
End Function

Public Function Styling() As FStyling
    If pStyling Is Nothing Then Set pStyling = New FStyling
    Set Styling = pStyling
End Function



'[References]
Public Sub setRepoEditBook(value As Excel.Workbook)
    Set pRepoEditBook = value
End Sub

Public Function getRepoEditBook() As Excel.Workbook
    Set getRepoEditBook = pRepoEditBook
End Function



'[Collections]
Friend Sub registerRepoManager(manager As IRepoManager)
    Dim name As String
    Dim prevManager As IRepoManager
    '----------------------------------------------------------------------------------------------------------
    name = manager.getName
    Set prevManager = F.dictionaries.GetObject(pRepoManagers, name)
    If Not prevManager Is Nothing Then Call prevManager.destroy
    Call F.dictionaries.addItem(pRepoManagers, name, manager)
End Sub

Friend Sub unregisterRepoManager(manager As IRepoManager)
    Call F.dictionaries.removeItem(pRepoManagers, manager.getName)
    Call manager.destroy
End Sub

Friend Sub clearRepoManagersCollection()
    Dim key As Variant
    Dim manager As IRepoManager
    '----------------------------------------------------------------------------------------------------------
    For Each key In pRepoManagers.Keys
        Set manager = pRepoManagers.item(key)
        Call manager.destroy
    Next key
    Set pRepoManagers = F.dictionaries.Create(False)
End Sub




'[Language change]
Public Function selectLanguage() As Language
    Const WINDOW_TITLE As String = "Languages.ChangeLanguage.WindowTitle"
    '----------------------------------------------------------------------------------------------------------
    Dim languages As Scripting.Dictionary
    Dim selection As VBA.Collection
    Dim selectedLanguage As Language
    Dim dictInitialSelection As Scripting.Dictionary
    '----------------------------------------------------------------------------------------------------------
    
    Set languages = MsgService.getLanguagesDictionary
    If F.dictionaries.isNonEmptyDictionary(languages) Then
        Set languages = F.dictionaries.changeKeys(languages, "$.getNameForCurrentLanguage")
        Set selectedLanguage = MsgService.getCurrentLanguage
        Set dictInitialSelection = F.dictionaries.createSingleEntryDict(selectedLanguage.getNameForCurrentLanguage, selectedLanguage)
        Set selection = UI.Dialogs.selectWindow(languages). _
                                    setInitialSelection(dictInitialSelection). _
                                    setTitle(MsgService.getText(WINDOW_TITLE)). _
                                    setLocationProps(UI.Create.LocationPropsForScreenCenter). _
                                    Execute
        If F.Collections.isNonEmptyCollection(selection) Then
            Set selectLanguage = selection.item(1)
        End If
    End If
    
End Function
