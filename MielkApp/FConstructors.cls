VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FConstructors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FConstructors"
'----------------------------------------------------------------------------------------------------------

Public Function ItemsProcessor(Optional items As Scripting.Dictionary, Optional viewer As IItemsProcessorViewer) As ItemsProcessor
    Set ItemsProcessor = New ItemsProcessor
    With ItemsProcessor
        Call .setItems(items)
        Call .setViewer(viewer)
    End With
End Function

Public Function DefaultItemsProcessorViewer() As WDefaultItemsProcessorViewer
    Set DefaultItemsProcessorViewer = New WDefaultItemsProcessorViewer
End Function

Public Function processItemsList() As XProcessItemsList
    Set processItemsList = New XProcessItemsList
End Function

Public Function processItemLine(processItemsList As XProcessItemsList, item As IProcessItem, index As Long) As xxProcessItemLine
    Set processItemLine = New xxProcessItemLine
    With processItemLine
        Call .setParent(processItemsList)
        Call .setItem(item)
        Call .setName(item.getName)
        Call .setIndex(index)
    End With
End Function

Public Function processSummaryPanel() As XProcessSummaryPanel
    Set processSummaryPanel = New XProcessSummaryPanel
End Function

Public Function processSummaryControl(parent As XProcessSummaryPanel) As xxProcessSummaryControl
    Set processSummaryControl = New xxProcessSummaryControl
    With processSummaryControl
        Call .setParent(parent)
    End With
End Function

Public Function Repository2() As Repository2
    Set Repository2 = New Repository2
End Function
