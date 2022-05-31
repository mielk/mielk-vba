VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITreeViewItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Interface for objects that can be displayed on tree view.

Private Const CLASS_NAME As String = "ITreeViewItem"
'----------------------------------------------------------------------------------------------------------

Public Function getUuid() As String
End Function

Public Function getTreeNodeCaption() As String
End Function

Public Function getObject() As Object
End Function

Public Function equals(comparedItem As Variant) As Boolean
End Function

Public Sub delete()
End Sub
