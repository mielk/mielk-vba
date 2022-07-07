VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRepoParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "IRepoParent"
'----------------------------------------------------------------------------------------------------------


Public Function getName() As String
End Function

Public Function getSheetName() As String
End Function

Public Function isAddingAllowed() As Boolean
End Function

Public Function isDeletingAllowed() As Boolean
End Function

Public Function isFilteringAllowed() As Boolean
End Function

Public Function isSortingAllowed() As Boolean
End Function

Public Function getIdField() As String
End Function

Public Function getColumns() As VBA.Collection
End Function

Public Function getSortingRules() As Scripting.Dictionary
End Function

Public Function getDeleteSuccessInfo() As String
End Function

Public Function getDeleteItemQuestion() As String
End Function

Public Function getDeleteItemWindowTitle() As String
End Function


'[Actions]
Public Function createNewItem() As IEditItem
End Function

Public Function runAfterEditAction(ByRef updateView As Boolean)
End Function
