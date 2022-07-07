VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FConstructors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Friend Function processItemsList() As XProcessItemsList
    Set processItemsList = New XProcessItemsList
End Function

Friend Function processItemLine(processItemsList As XProcessItemsList, item As IProcessItem, index As Long) As xxProcessItemLine
    Set processItemLine = New xxProcessItemLine
    With processItemLine
        Call .setParent(processItemsList)
        Call .setItem(item)
        Call .setName(item.getName)
        Call .setIndex(index)
    End With
End Function

Friend Function processSummaryPanel() As XProcessSummaryPanel
    Set processSummaryPanel = New XProcessSummaryPanel
End Function

Friend Function processSummaryControl(parent As XProcessSummaryPanel) As xxProcessSummaryControl
    Set processSummaryControl = New xxProcessSummaryControl
    With processSummaryControl
        Call .setParent(parent)
    End With
End Function

Public Function Repository2() As Repository2
    Set Repository2 = New Repository2
End Function


Public Function AliasesPanel() As XAliasesPanel
    Set AliasesPanel = New XAliasesPanel
End Function

Public Function EditEntityWindow() As WEditEntity
    Set EditEntityWindow = New WEditEntity
End Function

Public Function DatesRange(Optional byMonths As Boolean = False, Optional maxRange As Long = 0) As EDatesRange
    Set DatesRange = New EDatesRange
    With DatesRange
        Call .setByMonths(byMonths)
        Call .setMaxRange(maxRange)
    End With
End Function


Public Function Standarizer() As Standarizer
    Set Standarizer = New Standarizer
End Function

Friend Function StandarizingAction(alias As String, from As Variant, into As Variant) As StandarizingAction
    Set StandarizingAction = New StandarizingAction
    With StandarizingAction
        Call .setAlias(alias)
        Call .setBefore(from)
        Call .setAfter(into)
    End With
End Function
