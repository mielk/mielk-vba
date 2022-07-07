VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemsProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "ItemsProcessor"
'[Status] -------------------------------------------------------------------------------------------------
Private pMetaObject As Scripting.Dictionary
Private pItems As Scripting.Dictionary
Private pStatusCounters As Scripting.Dictionary
Private pProcessId As Variant
'[UI components] ------------------------------------------------------------------------------------------
Private pStatusViewer As IItemsProcessorViewer
'[Events] -------------------------------------------------------------------------------------------------
Event AfterConfirm()
Event AfterFinish()
Event AfterProcessingItem(item As IProcessItem)
Event BeforeProcessingItem(item As IProcessItem)
Event CancelledByUser()
Event DetailProgressChange(currentOperation As String, currentProgress As Single)
Event AfterStatusCountersChange(previousStatus As Variant, newStatus As Variant)
'[User form actions]
Event AfterExpandCollapse(item As IProcessItem)
Event AfterPreviewFileClicked(item As IProcessItem)
Event AfterReprocessClicked(item As IProcessItem)
Event AfterItemRemoved(item As IProcessItem)
'----------------------------------------------------------------------------------------------------------



'[INITIALIZER & TERMINATOR]
Private Sub Class_Initialize()
    Call initializeContainers
End Sub

Private Sub initializeContainers()
    Set pItems = F.dictionaries.Create(False)
End Sub

Private Sub Class_Terminate()
    Set pItems = Nothing
End Sub


'[SETTERS]
Public Function setMetaObject(value As Scripting.Dictionary) As ItemsProcessor
    Set setMetaObject = Me
    Set pMetaObject = value
End Function

Public Function setViewer(viewer As IItemsProcessorViewer) As ItemsProcessor
    Set setViewer = Me
    Set pStatusViewer = viewer
    Call viewer.setProcessor(Me)
End Function

Public Function setProcessId(value As Variant) As ItemsProcessor
    Set setProcessId = Me
    Call F.Variables.assign(pProcessId, value)
End Function

Public Function setItems(value As Scripting.Dictionary) As ItemsProcessor
    Set setItems = Me
    Set pItems = value
End Function




'[GETTERS]
Public Function countItems() As Long
    countItems = pItems.Count
End Function

Public Function getProcessId() As Variant
    getProcessId = pProcessId
End Function

Public Function getItemsDictionary() As Scripting.Dictionary
    Set getItemsDictionary = pItems
End Function

Public Function getItemsCollection() As VBA.Collection
    Set getItemsCollection = F.dictionaries.toCollection(pItems, DictPart_ValueOnly)
End Function

Public Function getMetaProperty(propertyTag As String) As Variant
    Call F.Variables.assign(getMetaProperty, F.dictionaries.getItem(pMetaObject, propertyTag))
End Function

Public Function getItem(key As String) As IProcessItem
    Dim obj As Object
    '------------------------------------------------------------------------------------------------------
    Set obj = F.dictionaries.GetObject(pItems, key)
    If Not obj Is Nothing Then
        If TypeOf obj Is IProcessItem Then
            Set getItem = F.dictionaries.GetObject(pItems, key)
        End If
    End If
End Function

Public Function getStatusCounter(status As Variant) As Long
    If status < 0 Then          '[ALL] status
        getStatusCounter = pItems.Count
    Else
        getStatusCounter = VBA.Val(F.dictionaries.getItem(pStatusCounters, status))
    End If
End Function



'[Items]
Public Function addItems(value As Scripting.Dictionary) As ItemsProcessor
    Set addItems = Me
    Call F.dictionaries.appendItemsFromContainer(pItems, value)
End Function

Public Function addItem(item As IProcessItem, key As String) As ItemsProcessor
    Set addItem = Me
    Call F.dictionaries.addItem(pItems, key, item)
End Function




'[API]
Public Sub run(Optional processId As Variant = vbNullString)
    Dim key As Variant
    Dim item As IProcessItem
    '------------------------------------------------------------------------------------------------------
    
    Set pStatusCounters = createCountersDictionary
    pProcessId = processId
    
    If Not pStatusViewer Is Nothing Then Call pStatusViewer.display
    For Each key In pItems.Keys
        Set item = pItems.item(key)
        Call processItem(item)
    Next key
    
    RaiseEvent AfterFinish
    
End Sub

Private Sub processItem(item As IProcessItem)
    Dim statusBefore As Variant:                        statusBefore = item.getStatus
    '------------------------------------------------------------------------------------------------------
    
    Call item.setItemsProcessor(Me)
    RaiseEvent BeforeProcessingItem(item)
        DoEvents
    Call item.process
        DoEvents
    Call updateCounters(statusBefore, item.getStatus)
    
    RaiseEvent AfterProcessingItem(item)
    
End Sub

Private Function createCountersDictionary() As Scripting.Dictionary
    Dim key As Variant
    Dim item As IProcessItem
    Dim status As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set createCountersDictionary = F.dictionaries.Create(False)
    
    For Each key In pItems.Keys
        Set item = pItems.item(key)
        status = item.getStatus
        
        With createCountersDictionary
            If .Exists(status) Then
                .item(status) = .item(status) + 1
            Else
                Call .Add(status, 1)
            End If
        End With
    Next key
    
End Function


Private Sub updateCounters(statusBefore As Variant, statusAfter As Variant)
    If statusBefore <> statusAfter Then
        With pStatusCounters
            'Decrement previous status' counter.
            If .Exists(statusBefore) Then
                .item(statusBefore) = .item(statusBefore) - 1
            End If
            
            'Increment new status' counter
            If VBA.IsEmpty(statusAfter) Then
                'Skip for <Empty>
            ElseIf .Exists(statusAfter) Then
                .item(statusAfter) = .item(statusAfter) + 1
            Else
                Call .Add(statusAfter, 1)
            End If
            
            RaiseEvent AfterStatusCountersChange(statusBefore, statusAfter)
            
        End With
    End If
End Sub


Public Sub removeItem(name As String)
    Dim item As IProcessItem
    Dim status As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set item = F.dictionaries.GetObject(pItems, name)
    If Not item Is Nothing Then
        status = item.getStatus
        Call updateCounters(status, Empty)
    End If
    
    Call F.dictionaries.removeItem(pItems, name)
    
    RaiseEvent AfterItemRemoved(item)

End Sub



'[Triggers]
Friend Sub triggerAfterExpandCollapseEvent(item As IProcessItem)
    RaiseEvent AfterExpandCollapse(item)
End Sub

Friend Sub triggerAfterPreviewItemEvent(item As IProcessItem)
    RaiseEvent AfterPreviewFileClicked(item)
End Sub

Friend Sub triggerAfterReprocessItemEvent(item As IProcessItem)
    Dim statusBefore As Variant:                statusBefore = item.getStatus
    '------------------------------------------------------------------------------------------------------
    RaiseEvent AfterReprocessClicked(item)
    With item
        Call .setItemsProcessor(Me)
        Call .process
        Call updateCounters(statusBefore, .getStatus)
    End With
End Sub

Public Sub triggerDetailStatusUpdate(action As String, progress As Single, Optional randomVariation As Single)
    Dim displayProgress As Single
    '------------------------------------------------------------------------------------------------------
    
    If randomVariation = 0 Then
        displayProgress = progress
    Else
        displayProgress = F.Numerics.getRandomBetween(progress * (1 - randomVariation), progress * 1)
    End If
    
    RaiseEvent DetailProgressChange(action, displayProgress)
    DoEvents
    
End Sub







'[CONFIRM & CANCEL]
Public Function confirm()
    RaiseEvent AfterConfirm
End Function

Public Function Cancel()
    RaiseEvent CancelledByUser
End Function

Public Sub hide()
    Call pStatusViewer.hide
End Sub
