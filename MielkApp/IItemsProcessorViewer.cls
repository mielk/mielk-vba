VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IItemsProcessorViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "IItemsProcessorViewer"
'----------------------------------------------------------------------------------------------------------

Public Function setProcessor(value As ItemsProcessor) As IItemsProcessorViewer
End Function

Public Function getProcessor() As ItemsProcessor
End Function

Public Function getItems() As VBA.Collection
End Function

Public Function getCanvas() As Variant
End Function

Public Function getStatusSettings(value As Long) As Scripting.Dictionary
End Function

Public Function getStatusesDictionary() As Scripting.Dictionary
End Function

Public Sub arrangeLayout()
End Sub

Public Sub display()
End Sub

Public Sub hide()
End Sub

Public Sub AfterConfirm()
End Sub

Public Sub afterCancel()
End Sub
