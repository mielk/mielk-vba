VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EVbaProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "EVbaProject"
'----------------------------------------------------------------------------------------------------------
Private pBook As Excel.Workbook
Private pProject As VBIDE.VBProject
Private pName As String
Private pPath As String
Private pComponents As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------



'[SETTERS]
Public Function setPath(value As String) As EVbaProject
    Set setPath = Me
    Call readFromPath(value)
End Function

Public Function setBook(value As Excel.Workbook) As EVbaProject
    Set setBook = Me
    Set pBook = value
    Call read(pBook)
End Function



'[GETTERS]
Public Function getPath() As String
    getPath = pPath
End Function

Public Function getName() As String
    getName = pName
End Function

Public Function countComponents() As Long
    If F.dictionaries.isNonEmptyDictionary(pComponents) Then
        countComponents = pComponents.Count
    End If
End Function

Public Function getComponents() As VBA.Collection
    Set getComponents = F.dictionaries.toCollection(pComponents)
End Function

Public Function getComponentsDict() As Scripting.Dictionary
    Set getComponentsDict = F.dictionaries.Clone(pComponents)
End Function

Public Function getComponentByName(name As String) As VBIDE.VBComponent
    Set getComponentByName = F.dictionaries.GetObject(pComponents, name)
End Function

Private Function read(wkb As Excel.Workbook) As EVbaProject
    Dim project As VBIDE.VBProject
    Dim component As VBIDE.VBComponent
    Dim module As EVbaModule
    Dim counter As Long
    Dim totalItems As Long
    '------------------------------------------------------------------------------------------------------
    
    Set read = Me
    Set pProject = wkb.VBProject
    pName = pProject.name
    pPath = wkb.FullName
    Set pComponents = F.dictionaries.Create(False)
    
    Set project = F.Developer.getVbProject(wkb)
    
    If project.Protection = vbext_pp_locked Then
        Debug.Print "File " & getPath & " is protected"
    Else
        Debug.Print project.VBComponents.Count
        totalItems = project.VBComponents.Count
        For Each component In project.VBComponents
        
            If filterComponentByName(component.name) Then
        
                Debug.Print counter & ": " & component.name
                Set module = New EVbaModule
                With module
                    Call .setVbComponent(component)
                    Call .load
                    Call F.dictionaries.addItem(pComponents, .getName, module)
                End With
                counter = counter + 1
    Call ProgressBar.setProgress(counter / totalItems)
        
            End If

        Next component
    End If
    
End Function



Private Function readFromPath(filepath As String) As EVbaProject
    Dim wkb As Excel.Workbook
    Dim wasOpen As Boolean
    Dim xls As Excel.Application
    '------------------------------------------------------------------------------------------------------
    
    Set readFromPath = Me
    
    Set xls = F.Excel.getBackgroundInstance(False, False)
    Set wkb = F.Books.open_(filepath, True, xls, wasOpen:=wasOpen)
    
    If F.Books.IsValid(wkb) Then
        Call read(wkb)
        Call wkb.Close(False)
    End If
    
End Function
