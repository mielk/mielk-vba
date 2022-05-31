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


Public Function getWindowTitle() As String
End Function

Public Function getFrameHeight() As Single
End Function

Public Function getFrameWidth() As Single
End Function

Public Function getColumnsDictionary() As Scripting.Dictionary
End Function

Public Function getFilterColumns() As VBA.Collection
End Function

Public Function getSortColumnKey() As String
End Function

Public Function getDeleteSuccessInfo() As String
End Function

Public Function getDeleteItemQuestion() As String
End Function

Public Function getDeleteItemWindowTitle() As String
End Function

Public Function hasRecordsCounter() As Boolean
End Function

Public Function isAddingAllowed() As Boolean
End Function

Public Function getModalState() As FormShowConstants
End Function


'[Actions]
Public Function createNewItem() As IEdititem
End Function

Public Function runAfterEditAction(ByRef updateView As Boolean)
End Function
