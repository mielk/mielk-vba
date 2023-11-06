VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WSelectFromTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "WSelectFromTree"
'[Control names] ------------------------------------------------------------------------------------------
Private Const TREE_VIEW_NAME As String = "tvItems"
Private Const TREE_VIEW_WIDTH As Single = 267
Private Const TREE_VIEW_HEIGHT As Single = 327
Private Const TREE_VIEW_LEAF_COLOR As Long = 7950891
'[UI constants] -------------------------------------------------------------------------------------------
Private Const DEFAULT_TITLE As String = "Select items"
Private Const DEFAULT_LIST_BOX_WIDTH As Single = 200
Private Const DEFAULT_LIST_BOX_HEIGHT As Single = 200
Private Const DEFAULT_LIST_BOX_BORDER_COLOR As Long = 11842740
Private Const DEFAULT_HAS_TITLE_PANEL As Boolean = True
Private Const DEFAULT_MIN_ITEMS_SELECTED As Long = 0
Private Const DEFAULT_VERTICAL_MARGIN As Single = 6
'[Logic components] ---------------------------------------------------------------------------------------
Private pListener As IListener
Private pDataStructure As Scripting.Dictionary
Private pInitialSelection As Scripting.Dictionary
'[UI components] ------------------------------------------------------------------------------------------
Private WithEvents pForm As WTemplate
Attribute pForm.VB_VarHelpID = -1
Private pFrame As MSForms.frame
Private WithEvents pTreeView As clsTreeView
Attribute pTreeView.VB_VarHelpID = -1
'[UI properties] ------------------------------------------------------------------------------------------
Private pTitle As String
Private pIsRendered As Boolean
'[Tree view properties] -----------------------------------------------------------------------------------
Private pRootItem As ITreeViewItem
Private pNodesMap As Scripting.Dictionary
Private pItemsMap As DoubleDictionary
'----------------------------------------------------------------------------------------------------------


Private Sub Class_Initialize()

    Set pItemsMap = F.Create.DoubleDictionary
    Set pNodesMap = F.dictionaries.Create(False)
    
    '[Default values]
    pTitle = DEFAULT_TITLE
    
End Sub



'[SETTERS]
Public Function setListener(value As IListener) As WSelectFromTree
    Set setListener = Me
    Set pListener = value
End Function

Public Function setDataStructure(value As Scripting.Dictionary) As WSelectFromTree
    Set setDataStructure = Me
    Set pDataStructure = value
End Function

Public Function setInitialSelection(value As Scripting.Dictionary) As WSelectFromTree
    Set setInitialSelection = Me
    Set pInitialSelection = value
End Function

Public Function setTitle(value As String) As WSelectFromTree
    Set setTitle = Me
    pTitle = value
End Function




'[GETTERS]
Public Function getSelection() As VBA.Collection
    Dim key As Variant
    Dim node As clsNode
    Dim items As Scripting.Dictionary
    Dim item As Object
    '------------------------------------------------------------------------------------------------------
    
    Set getSelection = New VBA.Collection
    
    For Each key In pNodesMap.keys
        Set items = pItemsMap.getDictionaryById
        Set node = pNodesMap.item(key)
        If node.Checked Then
            Set item = F.dictionaries.getObject(items, key)
            If Not item Is Nothing Then
                Call getSelection.Add(item)
            End If
        End If
    Next key
    
End Function





'[RENDERING]
Public Sub display(Optional ByVal locationProps As DLocation)
    If locationProps Is Nothing Then Set locationProps = UI.Create.LocationPropsForScreenCenter
    
    Call createForm
    Call render
    Call pForm.arrangeComponents
    Call pForm.locateRelativeTo(locationProps)
    Call pForm.display(vbModal, False)
    
End Sub


Private Sub createForm()
    Set pForm = UI.Create.WindowTemplate
    With pForm
        Call .setCaption(pTitle)
        Call .render
    End With
End Sub

Private Sub render()
    If Not pIsRendered Then
        Call insertFrame
        Call insertTreeView
        Call loadItemsToTreeView
        pIsRendered = True
    End If
End Sub

Private Sub insertFrame()
    Dim frame As MSForms.frame
    '----------------------------------------------------------------------------------------------------------
    
    Set frame = pForm.getContentFrame
    Set pFrame = frame.controls.Add(CREATE_FRAME_ID, "frmTreeView", True)
    With pFrame
        .SpecialEffect = fmSpecialEffectFlat
        .caption = vbNullString
        .borderStyle = fmBorderStyleNone
        .width = TREE_VIEW_WIDTH
        .height = TREE_VIEW_HEIGHT
        .top = 6
        .left = 0
    End With
    
End Sub

Private Sub insertTreeView()
    Set pTreeView = UI.Create.TreeView
    With pTreeView
        Set .TreeControl = pFrame
        Call .NodesClear
        .AppName = App.getName
        .CheckBoxes(True) = True
        .RootButton = True
        .LabelEdit(True) = False
        .EnableLabelEdit(True) = False
        .FullWidth = False
        .Indentation = 21
        .NodeHeight = 15
        .ShowLines = True
        .ShowExpanders = True
        Call .ExpanderImage(Icons.controls("Win7Minus").picture, _
                            Icons.controls("Win7Plus2").picture)
    End With
    
End Sub

Private Sub loadItemsToTreeView()
    Dim key As Variant
    Dim caption As String
    Dim node As clsNode
    '----------------------------------------------------------------------------------------------------------
    
    With pTreeView
        Call .NodesClear
        If F.dictionaries.isNonEmptyDictionary(pDataStructure) Then
            For Each key In pDataStructure.keys
                caption = VBA.CStr(key)
                Set node = addNode(True, Nothing, caption, Nothing, pDataStructure.item(key))
            Next key
            Call .Refresh
            Call updateItemsSelection
        End If
    End With
    
End Sub

Private Function addNode(isRoot As Boolean, parentNode As clsNode, caption As String, item As Object, subitems As Object) As clsNode
    Dim key As Variant
    Dim dict As Scripting.Dictionary
    Dim col As VBA.Collection
    Dim colItem As Variant
    '------------------------------------------------------------------------------------------------------
    
    key = F.Crypto.createUUID
    
    If isRoot Then
        Set addNode = pTreeView.AddRoot(sKey:=VBA.CStr(key), vCaption:=caption)
    Else
        Set addNode = parentNode.AddChild(VBA.CStr(key), caption)
    End If
    
    Call formatNode(addNode, isRoot, Not F.Variables.isContainer(subitems))
    Call F.dictionaries.addItem(pNodesMap, key, addNode)
    
    If F.Variables.isUndefined(subitems) Then
        '
    ElseIf F.dictionaries.isDictionary(subitems) Then
        Set dict = F.Cast.toDictionary(subitems)
        For Each key In dict.keys
            Call addNode(False, addNode, VBA.CStr(key), Nothing, dict.item(key))
        Next key
    ElseIf F.Collections.isCollection(subitems) Then
        Set col = F.Cast.toCollection(subitems)
        For Each colItem In col
            Call addNode(False, addNode, F.Strings.stringify(colItem), F.Cast.toObject(colItem), Nothing)
        Next colItem
    ElseIf VBA.IsObject(subitems) Then
        Call pItemsMap.addItem(key, subitems)
    End If
    
End Function



Private Sub formatNode(node As clsNode, isRoot As Boolean, isLastLevel As Boolean)
    Dim item As ITreeViewItem:          Set item = node.getObject
    '------------------------------------------------------------------------------------------------------
    With node
        .Bold = isRoot
        .ForeColor = VBA.IIf(isLastLevel, TREE_VIEW_LEAF_COLOR, vbBlack)
        Call .setTreeView(pTreeView)
    End With
End Sub

Private Sub updateItemsSelection()
    Dim itemsMap As Scripting.Dictionary
    Dim key As Variant
    Dim node As clsNode
    Dim item As Object
    '------------------------------------------------------------------------------------------------------
    
    If F.dictionaries.isNonEmptyDictionary(pInitialSelection) Then
        Set itemsMap = pItemsMap.getDictionaryById
        For Each key In itemsMap.keys
            Set item = itemsMap.item(key)
            If pInitialSelection.Exists(item) Then
                Set node = pNodesMap.item(key)
                If Not node Is Nothing Then
                    node.Checked = True
                End If
            End If
        Next key
    End If
    
End Sub
