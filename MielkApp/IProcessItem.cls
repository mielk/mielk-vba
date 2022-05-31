VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProcessItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "IProcessItem"
'----------------------------------------------------------------------------------------------------------


Public Function setItemsProcessor(value As ItemsProcessor) As IProcessItem
End Function

Public Function setParams(params As Scripting.Dictionary) As IProcessItem
End Function



Public Function getName() As String
End Function

Public Function getCaption() As String
End Function

Public Function getStatus() As Long
End Function

Public Function getMessage() As Variant
End Function



Public Sub process()
End Sub

Public Sub preview()
End Sub
