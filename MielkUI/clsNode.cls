VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Build 026
'***************************************************************************
'
' Authors:  JKP Application Development Services, info@jkp-ads.com, http://www.jkp-ads.com
'           Peter Thornton, pmbthornton@gmail.com
'
' (c)2013-2017, all rights reserved to the authors
'
' You are free to use and adapt the code in these modules for
' your own purposes and to distribute as part of your overall project.
' However all headers and copyright notices should remain intact
'
' You may not publish the code in these modules, for example on a web site,
' without the explicit consent of the authors
'***************************************************************************

'-------------------------------------------------------------------------
' Module    : clsNode
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 15-01-2013
' Purpose   : Holds all information of a node of the tree
'-------------------------------------------------------------------------
Option Explicit

Implements IListener

Private mbExpanded As Boolean

Private mcolChildNodes As Collection

Private moParentNode As clsNode
Private moLastActiveNode As clsNode
Private moTree As clsTreeView

Private msKey As String
Private mvCaption
Private msControlTipText As Variant

Private mlChecked As Long        ' PT checkbox tristate boolean 0/-1 or 1 for null
'Private mbVisible As Boolean        ' PT determines if the node can be displayed
Private mnIndex As Long             ' PT order added to Treeview's mcolNodes, won't change
Private mlVisIndex As Long          ' PT the visible order in the current view, changes with expand/collapse
Private mvIconMainKey               ' PT string name or numeric index as icon Key for the Image collection
Private mvIconExpandedKey           ' PT ditto for expanded icon
Private mlIconCnt As Long           ' PT number of icons availabel for this node 0, 1 or 2
Private msngTextWidth As Single     ' PT autosized text width before the node is widened beyond the frame
Private mlBackColor As Long         ' PT
Private mbBold As Boolean           ' PT
Private mlForeColor As Long         ' PT
Private mvTag
Private mbHighlightPreviousActiveNode As Boolean

Private WithEvents mctlControl As MSForms.label
Attribute mctlControl.VB_VarHelpID = -1
Private WithEvents mctlExpander As MSForms.label
Attribute mctlExpander.VB_VarHelpID = -1
Private WithEvents moEditBox As MSForms.textBox     ' PT editbox
Attribute moEditBox.VB_VarHelpID = -1
Private WithEvents mctlCheckBox As MSForms.label    ' PT checkbox
Attribute mctlCheckBox.VB_VarHelpID = -1

Private mctlExpanderBox As MSForms.label
Private mctlVLine As MSForms.label  ' PT vertical line, only the first child node with children will have a vertical line
Private mctlHLine As MSForms.label  ' PT horizontal line
Private mctlIcon As MSForms.image   ' PT separate icon image control

Public Enum ndSortOrder
    ndAscending = 1
    ndDescending = 2
End Enum
Public Enum ndCompareMethod
    ndBinaryCompare = 0
    ndTextCompare = 1
End Enum
Public Enum ndMouse
    ndDown = 1
    ndUp = 2
    ndMove = 3
    ndBeforeDragOver = 4
    ndBeforeDropOrPaste = 5
End Enum

#If Mac Then
    Const mcFullWidth As Long = 800
#Else
    Const mcFullWidth As Long = 600
#End If

Private moLabSizer As MSForms.label
Private Const mcBreak As String = "¶ " '"• "


'----------------------------------------------------------------------------------------------------------
Private Const BUTTONS_LEFT_MARGIN As Single = 0
Private Const BUTTONS_SPACE As Single = 3
Private Const SPECIAL_BUTTON_SIZE As Single = 10.5
Private Const MOVE_LAYER_MARGIN As Single = 3
Private Const MOVE_LAYER_PADDING As Single = 6
Private Const MOVE_LAYER_BACK_COLOR As Long = 11462647
Private Const MOVE_LAYER_BORDER_COLOR As Long = 1489127
Private Const DROP_VIEW_BACK_COLOR As Long = 9296036
Private Const DROP_VIEW_BORDER_COLOR As Long = 2651456
Private Const CTRL_RIGHT_PADDING As Single = 6
'----------------------------------------------------------------------------------------------------------
Private pTreeView As clsTreeView
Private pObject As ITreeViewItem
Private pCurrentHover As Boolean
Private pSpecialActionDtos As VBA.Collection
Private pSpecialButtons As VBA.Collection
'---------------------------------------------------------------------------------------------------------
Private WithEvents pHoverBackground As MSForms.label
Attribute pHoverBackground.VB_VarHelpID = -1
Private WithEvents pMoveCaptionLabel As MSForms.label
Attribute pMoveCaptionLabel.VB_VarHelpID = -1
Private WithEvents pMoveLabel As MSForms.label
Attribute pMoveLabel.VB_VarHelpID = -1
Private WithEvents pDropCaptionLabel As MSForms.label
Attribute pDropCaptionLabel.VB_VarHelpID = -1
Private WithEvents pDropLabel As MSForms.label
Attribute pDropLabel.VB_VarHelpID = -1
Private pMoveMode As Boolean
Attribute pMoveMode.VB_VarHelpID = -1
Private pMoveAnchorPoint As Point
'----------------------------------------------------------------------------------------------------------


'*********************
'* Public Properties *
'*********************

Public Property Get backColor() As Long

    backColor = mlBackColor ' if zero the treecaller will apply the frame container's backcolor

End Property

Public Property Let backColor(lColor As Long)
'PT if lColor is written as 0/black, change it to 1 as 0 means default
    mlBackColor = lColor
    If mlBackColor = 0 Then mlBackColor = 1
    If Not mctlControl Is Nothing Then
        mctlControl.backColor = lColor
    End If
End Property

Public Property Get Bold() As Boolean
    Bold = mbBold
End Property

Public Property Let Bold(bBold As Boolean)
    mbBold = bBold
    If Not mctlControl Is Nothing Then
        mctlControl.font.Bold = mbBold
    End If
End Property

Public Property Get caption()
    caption = mvCaption
End Property

Public Property Let caption(ByVal vCaption)
Dim sngWd As Single
    mvCaption = vCaption
    If Not mctlControl Is Nothing Then
        
        mctlControl.caption = CStr(vCaption)
        If mctlControl.AutoSize Then
            msngTextWidth = mctlControl.width
        Else
            sngWd = mctlControl.width
            mctlControl.AutoSize = True
            msngTextWidth = mctlControl.width
            mctlControl.AutoSize = False
            Call locateSpecialButtons
        End If
    End If
End Property

Public Property Get Checked()    ' PT
     ' Checked values are -1 true, 0 false, +1 mixed
     ' If TriState is enabled be careful not to return a potential +1 to a boolean or it'll coerce to True
    Checked = mlChecked
End Property

Public Property Let Checked(vChecked)  ' PT
    Dim bFlag As Boolean, bTriState As Boolean
    Dim lChecked As Long
    Dim cChild As clsNode

    ' Checked values are -1 true, 0 false, +1 mixed
    ' if vChecked is a boolean Checked will coerce to -1 or 0
    ' if vChecked is Null Checked is set as +1

    If VarType(vChecked) = vbBoolean Then
        lChecked = vChecked
    ElseIf IsNull(vChecked) Then
        lChecked = 1
    ElseIf vChecked >= -1 And vChecked <= 1 Then
        lChecked = vChecked
    End If

    bFlag = lChecked <> mlChecked
    mlChecked = lChecked

    If Not mctlCheckBox Is Nothing And bFlag Then
        moTree.changed = True
        UpdateCheckbox
    End If
    
    If Not moTree Is Nothing Then    ' eg during clone
        bFlag = moTree.CheckBoxes(bTriState)
        If bTriState Then
            If parentNode.caption <> "RootHolder" Then
                parentNode.CheckTriStateParent
            End If
            
            If Not childNodes Is Nothing Then
                For Each cChild In childNodes
                    cChild.CheckTriStateChildren mlChecked
                Next
            End If
        End If
    End If
    
End Property

Public Property Get child() As clsNode
' PT Returns a reference to the first Child node, if any
    On Error Resume Next
    Set child = mcolChildNodes(1)
End Property

Public Property Get childNodes() As Collection
    Set childNodes = mcolChildNodes
End Property

Public Property Set childNodes(colChildNodes As Collection)
    Set mcolChildNodes = colChildNodes
End Property

Public Property Get ControlTipText() As String
    ControlTipText = msControlTipText
End Property

Public Property Let ControlTipText(ByVal sControlTipText As String)
    msControlTipText = sControlTipText
    If Not mctlControl Is Nothing Then
        mctlControl.ControlTipText = msControlTipText
    End If
End Property

Public Property Get Expanded() As Boolean
    Expanded = mbExpanded
End Property

Public Property Let Expanded(ByVal bExpanded As Boolean)
    mbExpanded = bExpanded
    If Not Me.Expander Is Nothing Then
        UpdateExpanded bControlOnly:=False
    ElseIf Not Me.control Is Nothing Then
        UpdateExpanded bControlOnly:=True
    End If
End Property

Public Property Get ForeColor() As Long
    ForeColor = mlForeColor
End Property

Public Property Let ForeColor(lColor As Long)
'PT if lColor is written as 0/black, change it to 1 as 0 means default
    mlForeColor = lColor
    If mlForeColor = 0 Then mlForeColor = 1
    If Not mctlControl Is Nothing Then
        mctlControl.ForeColor = lColor
    End If
End Property

Public Property Get FirstSibling() As clsNode
    If Not moParentNode Is Nothing Then    ' PT Root has no parent
        Set FirstSibling = moParentNode.GetChild(1)
    End If
End Property

Public Property Get LastSibling() As clsNode
    If Not moParentNode Is Nothing Then    ' PT Root has no parent
        Set LastSibling = moParentNode.GetChild(-1)    ' -1 flags GetChild to return the last Child
    End If
End Property

Public Property Get ImageExpanded()
' PT string name or numeric index for the main icon key
    ImageExpanded = mvIconExpandedKey
End Property

Public Property Let ImageExpanded(vImageExpanded)
' PT string name or numeric index for an expanded icon key
    On Error GoTo errExit
    If Not IsMissing(vImageExpanded) Then
        If Not IsEmpty(vImageExpanded) Then
            If Len(mvIconMainKey) = 0 Then
                mvIconMainKey = vImageExpanded
            End If
            mvIconExpandedKey = vImageExpanded
            mlIconCnt = 2
        End If
    End If
errExit:
End Property

Public Property Get ImageMain()
' PT string name or numeric index for the main icon key
    ImageMain = mvIconMainKey
End Property

Public Property Let ImageMain(vImageMain)
' PT string name or numeric index for the main icon key
    On Error GoTo errExit
    If Not IsMissing(vImageMain) Then
        If Not IsEmpty(vImageMain) Then
            mvIconMainKey = vImageMain
            If mlIconCnt = 0 Then mlIconCnt = 1
        End If
    End If
errExit:
End Property

Public Sub ImageUpdate()
Dim bFullWidth As Boolean
Dim vKey
Dim pic As StdPicture

    On Error GoTo errH
    If Me.hasIcon(vKey) Then
    '  error if Icon or Control is not yet created, silent abort, it'll update in the next BuildRoot
        If moTree.GetNodeIcon(vKey, pic, bFullWidth) Then
            If bFullWidth Then
                Me.icon.picture = pic
            Else
                Me.control.picture = pic
            End If
        ElseIf bFullWidth Then
            Set Me.icon.picture = Nothing
        Else
            Set Me.control.picture = Nothing
        End If
    End If
    
    Exit Sub
errH:
    #If DebugMode = 1 Then
        Stop
        Resume
    #End If
End Sub


Public Property Get key() As String
    key = msKey
End Property

Public Property Let key(ByVal sKey As String)
    Dim bIsInMainCol As Boolean
    Dim i As Long
    Dim cTmp As clsNode

    On Error GoTo errH

    If tree Is Nothing Then
        msKey = sKey
        Exit Property
    ElseIf msKey = sKey Or Len(sKey) = 0 Then
        Exit Property
    End If

    On Error Resume Next
    Set cTmp = tree.Nodes.item(sKey)
    On Error GoTo errH

    If Not cTmp Is Nothing Then
        Err.Raise 457    ' standard duplicate key error
    End If

    ' to change the Key, remove Me and add Me back where it was with the new key
    For Each cTmp In tree.Nodes
        i = i + 1
        If cTmp Is Me Then
            bIsInMainCol = True
            Exit For
        End If
    Next

    If bIsInMainCol Then
        With tree.Nodes
            .remove i
            If i <= .Count Then
                .add Me, sKey, i
            Else
                .add Me
            End If
        End With
    Else
        ' Let Key  called by via move/copy
    End If

    msKey = sKey

    Exit Property
errH:
    Err.Raise Err.Number, "Let Key", Err.description
End Property

Public Property Get Level() As Long
    Dim lLevel As Long
    Dim cNode As clsNode

    On Error GoTo errH
    lLevel = -1
    Set cNode = Me.parentNode
    While Not cNode Is Nothing
        lLevel = lLevel + 1
        Set cNode = cNode.parentNode
    Wend
    Level = lLevel
    Exit Property
errH:
    #If DebugMode = 1 Then
        Stop
        Resume
    #End If
End Property

Public Property Get NextNode() As clsNode    ' can't name this proc 'Next' in VBA
' PT return the next sibling if there is one
    Dim i As Long
    Dim cNode As clsNode

    With Me.parentNode
        For Each cNode In .childNodes
            i = i + 1
            If cNode Is Me Then
                Exit For
            End If
        Next
        If .childNodes.Count > i Then
            Set NextNode = .childNodes(i + 1)
        End If
    End With
End Property

Public Property Get parentNode() As clsNode
    Set parentNode = moParentNode
End Property

Public Property Set parentNode(oParentNode As clsNode)
    Set moParentNode = oParentNode
End Property

Public Property Get Previous() As clsNode
' PT return the previous sibling if there is one
    Dim i As Long
    Dim cNode As clsNode

    With Me.parentNode
        For Each cNode In Me.parentNode.childNodes
            i = i + 1
            If cNode Is Me Then
                Exit For
            End If
        Next
        If i > 1 Then
            Set Previous = .childNodes(i - 1)
        End If
    End With
End Property

Public Property Get root() As clsNode
    Dim cTmp As clsNode
    Set cTmp = Me
    Do While Not cTmp.parentNode.parentNode Is Nothing
        Set cTmp = cTmp.parentNode
    Loop
    Set root = cTmp
End Property

Public Property Get tag()
    tag = mvTag
End Property

Public Property Let tag(vTag)
    mvTag = vTag
End Property


'*****************************
'* Public subs and functions *
'*****************************

Public Function sort(Optional ByVal ndOrder As ndSortOrder = ndAscending, _
                     Optional ByVal ndCompare As ndCompareMethod = ndTextCompare) As Boolean
' PT Sorts the child nodes,
'    returns True if the order has changed to flag Refresh should be called
    Dim sCaptions() As String
    Dim lStart As Long, lLast As Long, i As Long
    Dim colNodes As New Collection
    Dim bIsUnSorted As Boolean

    On Error GoTo errExit
    lStart = 1
    lLast = childNodes.Count    ' error if no childnodes to sort

    If lLast = 1 Then
        ' nothing to sort
        Exit Function
    End If

    ReDim idx(lStart To lLast) As Long
    ReDim sCaptions(lStart To lLast) As String
    For i = lStart To lLast
        idx(i) = i
        sCaptions(i) = childNodes.item(i).caption
    Next

    If ndOrder <> ndAscending Then ndOrder = -1    ' descending
    If ndCompare <> ndTextCompare Then ndCompare = ndBinaryCompare

    Call BinarySortIndexText(sCaptions(), lStart, lLast, idx, ndOrder, ndCompare)

    For i = lStart To lLast - 1
        If idx(i) <> idx(i + 1) - 1 Then
            bIsUnSorted = True
            Exit For
        End If
    Next

    If bIsUnSorted Then
        For i = lStart To lLast
            colNodes.add childNodes(idx(i))
        Next
        Set childNodes = colNodes
        sort = True
    End If

errExit:
'   Probably(?) any error was because there were no childnodes, no need to raise an error
End Function

Public Function AddChild(Optional sKey As String, _
                         Optional vCaption, _
                         Optional vImageMain, _
                         Optional vImageExpanded) As clsNode
    Dim bTriState As Boolean
    Dim cChild As clsNode

    On Error GoTo errH
    Set cChild = New clsNode

    With moTree.Nodes

        If Len(sKey) Then
100         .add cChild, sKey
101
            cChild.key = sKey
        Else
            .add cChild
        End If

        cChild.index = .Count
    End With

    If mcolChildNodes Is Nothing Then
        Set mcolChildNodes = New Collection
    End If

    mcolChildNodes.add cChild

    With cChild
        If Not IsMissing(vImageMain) Then
            If Len(vImageMain) Then
                .ImageMain = vImageMain
            End If
        End If

        If Not IsMissing(vImageExpanded) Then
            If Len(vImageExpanded) Then
                .ImageExpanded = vImageExpanded
            End If
        End If

        .caption = vCaption
        
        If mlChecked = -1 Then ' -1 = true, +1 = mixed
            If moTree.CheckBoxes(bTriState) Then
                If bTriState Then
                    .Checked = True
                End If
            End If
        End If
        
        Set .tree = moTree
        Set .parentNode = Me
    End With

    Set AddChild = cChild

    Exit Function
errH:
    #If DebugMode = 1 Then
        Stop
        Resume
    #End If

    If Erl = 100 And Err.Number = 457 Then
        Err.Raise vbObjectError + 1, "clsNode.AddChild", "Duplicate key: '" & sKey & "'"
    Else
        Err.Raise Err.Number, "clsNode.AddChild", Err.description
    End If
End Function

Public Function ChildIndex(sKey As String) As Long
'-------------------------------------------------------------------------
' Procedure : ChildIndex
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 15-01-2013
' Purpose   : Returns the index of a childnode using its key
'-------------------------------------------------------------------------
    Dim cNode As clsNode
    Dim lCt As Long
    For Each cNode In mcolChildNodes
        lCt = lCt + 1
        If sKey = cNode.key Then
            ChildIndex = lCt
            Set cNode = Nothing
            Exit Function
        End If
    Next
    Set cNode = Nothing
End Function

Public Function FullPath() As String
' PT, get all the grand/parent keys
' assumes use of key

    Dim s As String
    Dim cNode As clsNode

    On Error GoTo errDone
    s = Me.key
    Set cNode = Me

    While Err.Number = 0
        Set cNode = cNode.parentNode
        s = cNode.key & "\" & s
    Wend

errDone:
    FullPath = s
End Function

Public Function GetChild(vKey As Variant) As clsNode
'-------------------------------------------------------------------------
' Procedure : GetChild
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 15-01-2013
' Purpose   : Returns a childnode using its key
'-------------------------------------------------------------------------
    Dim cNode As clsNode
    Dim lIdx As Long

    If VarType(vKey) = vbString Then

        For Each cNode In mcolChildNodes
            If vKey = cNode.key Then
                Set GetChild = cNode
                Set cNode = Nothing
                Exit Function
            End If
        Next

    ElseIf Not mcolChildNodes Is Nothing Then
        lIdx = vKey
        If lIdx = -1 Then
            lIdx = mcolChildNodes.Count
        End If
        If lIdx > 0 Then
            Set GetChild = mcolChildNodes(lIdx)
        Else: Set mcolChildNodes = Nothing
        End If
    End If

    Set cNode = Nothing
End Function


'*************************************************************************
'*    Friend Properties, Subs & Funtions                                 *
'*    ** these procedures are visible throughout the project but should  *
'*    ** only be used to communicate with the TreeView, ie clsTreeView   *
'*************************************************************************

Friend Property Get control() As MSForms.label
    Set control = mctlControl
End Property

Friend Property Set control(ctlControl As MSForms.label)
    Set mctlControl = ctlControl
    If Not mctlControl Is Nothing Then
        If Not moTree Is Nothing Then
            Set mctlControl.font = moTree.TreeControl.font
            mctlControl.BackStyle = fmBackStyleTransparent
        Else
            Stop
        End If
    End If
End Property

Friend Property Get index() As Long    ' PT
    index = mnIndex
End Property

Friend Property Let index(idx As Long)
' PT Index: the order this node was added to Treeview's collection mcolNodes
'    Index will never increase but may decrement if previously added nodes are removed
    mnIndex = idx
End Property

Friend Property Let VisIndex(lVisIndex As Long)
    mlVisIndex = lVisIndex
End Property

Friend Property Get VisIndex() As Long    ' PT
    VisIndex = mlVisIndex
End Property

Friend Property Get tree() As clsTreeView
    Set tree = moTree
End Property

Friend Property Set tree(oTree As clsTreeView)
    Set moTree = oTree
End Property

Friend Property Get Checkbox() As MSForms.control
    Set Checkbox = mctlCheckBox
End Property

Friend Property Set Checkbox(oCtl As MSForms.control)
    Set mctlCheckBox = oCtl
End Property

Friend Property Get Expander() As MSForms.label
    Set Expander = mctlExpander
End Property

Friend Property Set Expander(ctlExpander As MSForms.label)
    Set mctlExpander = ctlExpander
End Property

Friend Property Get ExpanderBox() As MSForms.label
    Set ExpanderBox = mctlExpanderBox
End Property

Friend Property Set ExpanderBox(ctlExpanderBox As MSForms.label)
    Set mctlExpanderBox = ctlExpanderBox
End Property

Friend Property Set HLine(ctlHLine As MSForms.label)
    Set mctlHLine = ctlHLine
End Property

Friend Property Get HLine() As MSForms.label
    Set HLine = mctlHLine
End Property

Friend Property Set icon(ctlIcon As MSForms.image)
    Set mctlIcon = ctlIcon
End Property

Friend Property Get icon() As MSForms.image
    Set icon = mctlIcon
End Property

Friend Property Get TextWidth() As Single
    TextWidth = msngTextWidth
End Property

Friend Property Let TextWidth(sngTextWidth As Single)
    msngTextWidth = sngTextWidth
End Property

Friend Property Get VLine() As MSForms.label
    Set VLine = mctlVLine
End Property

Friend Property Set VLine(ctlVLine As MSForms.label)
    Set mctlVLine = ctlVLine
End Property

Friend Sub CheckTriStateParent()
' PT set triState value of parent according to its childnodes' values
    Dim alChecked(-1 To 1) As Long
    Dim cChild As clsNode

    If Not childNodes Is Nothing Then
        For Each cChild In childNodes
            alChecked(cChild.Checked) = alChecked(cChild.Checked) + 1
        Next
        If alChecked(1) Then
            alChecked(1) = 1
        ElseIf alChecked(-1) = childNodes.Count Then
            alChecked(1) = -1
        ElseIf alChecked(0) = childNodes.Count Then
            alChecked(1) = 0
        Else
            alChecked(1) = 1
        End If
        
        If mlChecked <> alChecked(1) Then
            mlChecked = alChecked(1)
            UpdateCheckbox
        End If

    End If
    
    If Not Me.caption = "RootHolder" Then
        If Not parentNode.parentNode Is Nothing Then
            parentNode.CheckTriStateParent
        End If
    End If

End Sub

Friend Sub CheckTriStateChildren(lChecked As Long)
' PT, make checked values of children same as parent's
'     only called if triState is enabled
Dim cChild As clsNode

    If mlChecked <> lChecked Then
        mlChecked = lChecked
        UpdateCheckbox
    End If
    
    If Not childNodes Is Nothing Then
        For Each cChild In childNodes
            cChild.CheckTriStateChildren lChecked
        Next
    End If
End Sub

Friend Function hasIcon(vKey) As Boolean
' PT get the appropriate icon key/index, if any
    If mlIconCnt = 2 And mbExpanded Then
        vKey = mvIconExpandedKey
        hasIcon = True    'Not IsEmpty(vKey) '(True
    ElseIf mlIconCnt Then
        vKey = mvIconMainKey
        hasIcon = True    'Not IsEmpty(vKey)
    End If
End Function

'Friend Sub EditBox(bEnterEdit As Boolean)    '  PT new in 006PT2 ,,move to clsTreView?
''-------------------------------------------------------------------------
'' Procedure : moCtl_Click
'' Author    : Peter Thornton
'' Created   : 20-01-2013
'' Purpose   : Enter/exit Editmode, show/hide the edit textbox
''-------------------------------------------------------------------------
'    On Error Resume Next
'    Set moEditBox = moTree.TreeControl.Controls("EditBox")
'    On Error GoTo 0
'
'    If bEnterEdit Then
'
'        If moEditBox Is Nothing Then
'            Set moEditBox = moTree.TreeControl.Controls.Add("forms.textbox.1", False)
'            moEditBox.Name = "EditBox"
'        End If
'
'        With moEditBox
'            .Left = Control.Left - 3
'            .Top = Control.Top - 1.5
'            .AutoSize = True
'            .BorderStyle = fmBorderStyleSingle
'            .Text = Caption
'            Control.Visible = False    ' hide the node label while editing
'            .ZOrder 0
'            .Visible = True
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .SetFocus
'        End With
'
'    ElseIf Not moEditBox Is Nothing Then
'        ' exit editmode
'        If Not moEditBox Is Nothing Then
'            ' error if moEditBox has already been removed
'            On Error Resume Next
'            moEditBox.Visible = False
'            moEditBox.Text = ""
'            Set moEditBox = Nothing
'        End If
'        Control.Visible = True
'
'    End If
'End Sub
Friend Sub EditBox(bEnterEdit As Boolean)    '  PT new in 006PT2 ,,move to clsTreView?
'-------------------------------------------------------------------------
' Procedure : moCtl_Click
' Author    : Peter Thornton
' Created   : 20-01-2013. Ammended 026 to resize/reposition the editbox within the confines of the frame container
' Purpose   : Enter/exit Editmode, show/hide the edit textbox
'-------------------------------------------------------------------------

Dim wd As Single, ht As Single
    On Error Resume Next
    Set moEditBox = moTree.TreeControl.controls("EditBox")
    Set moLabSizer = moTree.TreeControl.controls("LabSizer")
    On Error GoTo 0

    If bEnterEdit Then

        If moEditBox Is Nothing Then
            Set moEditBox = moTree.TreeControl.controls.add("forms.textbox.1", "EditBox")
            Set moLabSizer = moTree.TreeControl.add("forms.Label.1", "LabSizer")
            moLabSizer.left = moTree.TreeControl.left
            moLabSizer.top = moTree.TreeControl.top - 12
            moLabSizer.WordWrap = False
            moLabSizer.AutoSize = True
            moLabSizer.visible = False
        End If

        With moEditBox
            .left = mctlControl.left - 3
            .top = mctlControl.top - 1.5
            .AutoSize = False
            .width = moTree.TreeControl.width
            .borderStyle = fmBorderStyleSingle
            If Len(Me.caption) Then
                .text = Me.caption
            Else
                moEditBox_Change
            End If
            wd = .width
            ht = .height
            .ZOrder 0
            .visible = True
            mctlControl.visible = False
            .SelStart = 0
            .SelLength = Len(.text)

            .setFocus
            .width = wd
            .height = ht
        End With

    ElseIf Not moEditBox Is Nothing Then
        ' exit editmode
        If Not moEditBox Is Nothing Then
            ' error if moEditBox has already been removed
            On Error Resume Next
            moEditBox.visible = False
            moEditBox.text = ""
'            If Len(moEditBox.Tag) Then
'                moTree.TreeControl.ScrollTop = CSng(moEditBox.Tag)
'                moEditBox.Tag = ""
'            End If
            Set moEditBox = Nothing
            Set moLabSizer = Nothing
        End If
        mctlControl.visible = True
    End If

End Sub

Friend Function RemoveChild(cNode As clsNode, Optional bReverse As Boolean) As Boolean
'PT remove a node from the collection,
'   note, this is only one part of the process of removing a node

Dim lCt As Long
Dim i As Long
Dim cTmp As clsNode
    On Error GoTo errH

    If bReverse Then
        For i = mcolChildNodes.Count To 1 Step -1
            Set cTmp = mcolChildNodes(i)
            If mcolChildNodes(i) Is cNode Then
                mcolChildNodes.remove i
                RemoveChild = True
                Exit For
            End If
        Next
    Else
        For Each cTmp In mcolChildNodes
            lCt = lCt + 1
            If cTmp Is cNode Then
                mcolChildNodes.remove lCt
                RemoveChild = True
                Exit For
            End If
        Next
    End If
    
    If mcolChildNodes.Count = 0 Then
        Set mcolChildNodes = Nothing
        Me.Expanded = False
    End If

    Exit Function
errH:
    Err.Raise vbObjectError, "RemoveChild", Err.description
End Function

Friend Sub RemoveNodeControls()
    Dim cChild As clsNode
    If Not childNodes Is Nothing Then
        For Each cChild In childNodes
            cChild.RemoveNodeControls
        Next
    End If
    DeleteNodeControls False
End Sub

Friend Sub TerminateNode(Optional bDeleteNodeControls As Boolean)
'-------------------------------------------------------------------------
' Procedure : TerminateNode
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 15-01-2013
' Purpose   : Terminates the class instance
'-------------------------------------------------------------------------
    Dim cChild As clsNode
    'Instead of the Terminate event of the class we use this public
    'method so it can be explicitly called by parent classes.
    'This is done because to break the two way or circular references
    'between the parent child classes.
    
    'The most important call in this routine is to destroy the reference
    'between this node class and the parent treeview class -
    '    < Set moTree = Nothing >
    'Once all the moTree references to have been destroyed everything else will
    ' 'tear down' normally

    If Not childNodes Is Nothing Then
        For Each cChild In childNodes
            ' recursively drill down to all child nodes in this branch
            cChild.TerminateNode bDeleteNodeControls
        Next
    End If

    ' If deleting individual nodes while the treeview is running we also want to
    ' remove all associated controls as well as removing references
    
    If bDeleteNodeControls Then
        DeleteNodeControls True
        If bDeleteNodeControls Then
            index = -1
        End If
    End If

    Set mcolChildNodes = Nothing
    Set moTree = Nothing
End Sub


'******************************
'* Private subs and functions *
'******************************

Private Sub BinarySortIndexText(sCaptions() As String, ByVal lStart As Long, ByVal lEnd As Long, ByRef idx() As Long, ndOrder As Long, ndCompare As ndCompareMethod)
' PT sorts the index array based on the string array
    Dim lSmall As Long, lLarge As Long, sMid As String, lTmp As Long

    lSmall = lStart
    lLarge = lEnd
    sMid = sCaptions(idx((lSmall + lLarge) / 2))

    Do While lSmall <= lLarge
        Do While (StrComp(sCaptions(idx(lSmall)), sMid, ndCompare) = -ndOrder And lSmall < lEnd)
            lSmall = lSmall + 1
        Loop
        Do While (StrComp(sCaptions(idx(lLarge)), sMid, ndCompare) = ndOrder And lLarge > lStart)
            lLarge = lLarge - 1
        Loop
        If lSmall <= lLarge Then
            lTmp = idx(lSmall)
            idx(lSmall) = idx(lLarge)
            idx(lLarge) = lTmp
            lSmall = lSmall + 1
            lLarge = lLarge - 1
        End If
    Loop

    If lStart <= lLarge Then
        Call BinarySortIndexText(sCaptions(), lStart, lLarge, idx, ndOrder, ndCompare)
    End If
    If lSmall <= lEnd Then
        Call BinarySortIndexText(sCaptions(), lSmall, lEnd, idx, ndOrder, ndCompare)
    End If
End Sub

Private Sub DeleteNodeControls(bClearIndex As Boolean)
'PT Delete all controls linked to this node

    On Error GoTo errH

    With moTree.TreeControl.controls
        If Not mctlControl Is Nothing Then
            .remove mctlControl.name
            Set mctlControl = Nothing
            If Not mctlHLine Is Nothing Then
                .remove mctlHLine.name
                Set mctlHLine = Nothing
            End If
            If Not mctlIcon Is Nothing Then
                .remove mctlIcon.name
                Set mctlIcon = Nothing
            End If
            If Not mctlIcon Is Nothing Then
                .remove mctlIcon.name
                Set mctlIcon = Nothing
            End If
        End If

        Call removeCustomControls

        If Not mctlExpander Is Nothing Then
            .remove mctlExpander.name
            Set mctlExpander = Nothing
        End If
        If Not mctlExpanderBox Is Nothing Then
            .remove mctlExpanderBox.name
            Set mctlExpanderBox = Nothing
        End If
        If Not mctlVLine Is Nothing Then
            .remove mctlVLine.name
            Set mctlVLine = Nothing
        End If


        If Not moEditBox Is Nothing Then
            .remove moEditBox.name
            Set moEditBox = Nothing
        End If
        If Not mctlCheckBox Is Nothing Then
            .remove mctlCheckBox.name
            Set mctlCheckBox = Nothing
        End If

        If Not Me.parentNode Is Nothing Then
            ' if Me is the last child delete parent's expander and VLine (if it has one)
            If FirstSibling Is LastSibling Then

                If Not Me.parentNode.VLine Is Nothing Then
                    .remove Me.parentNode.VLine.name
                    Set Me.parentNode.VLine = Nothing
                End If
                
                If Not Me.parentNode.ExpanderBox Is Nothing Then
                    .remove Me.parentNode.ExpanderBox.name
                    Set Me.parentNode.ExpanderBox = Nothing
                End If

                If Not Me.parentNode.Expander Is Nothing Then
                    .remove Me.parentNode.Expander.name
                    Set Me.parentNode.Expander = Nothing
                End If

                Me.parentNode.Expanded = False

            End If

        End If

    End With

    If bClearIndex Then
        Me.index = -1  ' flag this node to be removed from mcolNodes in NodeRemove
    End If

    Exit Sub
errH:
    ' Stop
    Resume Next
End Sub

Private Function UpdateCheckbox()
Dim pic As StdPicture

    If Not mctlCheckBox Is Nothing Then
        With mctlCheckBox
            If moTree.GetCheckboxIcon(mlChecked, pic) Then
                .picture = pic
            Else
                .caption = IIf(mlChecked, "a", "")
                If (mlChecked = 1) <> (.ForeColor = RGB(180, 180, 180)) Then
                    .ForeColor = IIf(mlChecked = 1, RGB(180, 180, 180), vbWindowText)
                End If
            End If

        End With
    End If
    
    moTree.NodeEventRouter Me, "Checkbox", 1
    
End Function

Private Sub UpdateExpanded(bControlOnly As Boolean)
'-------------------------------------------------------------------------
' Procedure : UpdateExpanded
' Author    : Peter Thornton
' Created   : 27-01-2013
' Purpose   : Called via an Expander click or arrow keys
'             Updates the Expanded property and changes +/- caption
'-------------------------------------------------------------------------
    Dim bFullWidth As Boolean
    Dim vKey
    Dim pic As StdPicture

    If Not bControlOnly Then
        With Me.Expander
            If moTree.GetExpanderIcon(mbExpanded, pic) Then
                .picture = pic
            Else
                If mbExpanded Then
                    .caption = "-"
                Else
                    .caption = "+"
                End If
            End If
        End With
    End If

    On Error GoTo errExit
    If Me.hasIcon(vKey) Then
        If moTree.GetNodeIcon(vKey, pic, bFullWidth) Then
            If bFullWidth Then
                Me.icon.picture = pic   ' potential error if Icon is nothing, let error abort
            Else
                Me.control.picture = pic
            End If
        End If
    End If
errExit:
End Sub



'***********************
'*   Node Events       *
'***********************

Private Sub mctlCheckBox_Click()    ' PT new in 006PT2
'-------------------------------------------------------------------------
' Procedure : moCtl_Click
' Author    : Peter Thornton
' Created   : 20-01-2013
' Purpose   : Event fires when a Checkbox label is clicked
'-------------------------------------------------------------------------
    If moTree.EditMode(Me) Then
        ' exit editmode if in editmode
        moTree.EditMode(Me) = False
    End If
    If mlChecked = 0 Then

        Checked = -1
    Else
        Checked = 0
    End If

    If Not moTree.ActiveNode Is Me Then
        Set moTree.ActiveNode = Me

        moTree.NodeEventRouter Me, "Caption", 1    ' share the checkbox click event
    End If
End Sub

Private Sub mctlControl_Click()
'-------------------------------------------------------------------------
' Procedure : mctlControl_Click
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 15-01-2013
' Purpose   : Event fires when a treebranch is clicked
'-------------------------------------------------------------------------

' PT the call to NodeClick will raise the click event to the form
Dim bFlag As Boolean

    If Not moLastActiveNode Is Nothing Then
        Call moLastActiveNode.deactivate
        moLastActiveNode.control.borderStyle = fmBorderStyleNone
        Set moLastActiveNode = Nothing
        bFlag = True
    End If

    If moTree.ActiveNode Is Nothing Then
        Set moTree.ActiveNode = Me
        bFlag = True
    ElseIf Not bFlag Then
        bFlag = mctlControl.borderStyle <> fmBorderStyleNone
    End If
    
    If Not moTree.ActiveNode Is Me Or bFlag Then
        ' only raise the event the first time the node is activated
         moTree.NodeEventRouter Me, "Caption", 1 ' tvClick
         ' if preferred the click event is always raised to the form (even if the
         ' node was previously active) simply comment or remove this If/EndIf check
    End If

End Sub

Private Sub mctlControl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' PT  a node label has been double-clicked, enter edit-mode if manual editing is enabled
    Dim bDummy As Boolean

        If moTree.EnableLabelEdit(bDummy) Then
            moTree.EditMode(Me) = True
            EditBox bEnterEdit:=True
        Else
            ' moTree.NodeEventRouter Me, "Caption", tvDblClick
        End If

End Sub

Private Sub mctlControl_MouseDown(ByVal Button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
'PT temporarily activate and highlight the MouseDown node and a grey border to the previous activenode
'   MouseUp and Click events will confirm the action or reset the previous active node
    Dim bFlag As Boolean
    Dim previousActiveNode As clsNode:         Set previousActiveNode = moTree.ActiveNode
    

    If moTree.ActiveNode Is Me Then
        bFlag = Me.control.backColor = vbHighlight
        'bFlag = bFlag Or Me.control.borderStyle = fmBorderStyleSingle ' in Access this should be uncommented
    End If
    
    If Not bFlag Then
        Set moLastActiveNode = moTree.ActiveNode
        Set moTree.ActiveNode = Me
        If mbHighlightPreviousActiveNode Then
            If Not moLastActiveNode Is Nothing Then
                moLastActiveNode.control.borderStyle = fmBorderStyleSingle
                moLastActiveNode.control.borderColor = RGB(200, 200, 200)
            End If
        End If
    End If

    If moTree.EditMode(Me) Then
        ' if any node is in edit mode exit edit mode
        moTree.EditMode(Me) = False
    End If
    
    ' moTree.NodeEventRouter Me, "Caption", tvDown, Button, Shift, X, Y
End Sub



Private Sub mctlControl_MouseMove(ByVal Button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
    Call setAsHoveredNode
End Sub

Private Sub mctlControl_MouseUp(ByVal Button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
' PT MouseUp fires before the Click event, at this point we don't know 100% if user
'    definately wants to activate the MouseDown node. If user drags the mouse off the MouseDown node the
'    Click event will not fire which means user wants to cancel and revert to the previous activenode.
'
'    If MouseUp occurs with the cursor not over the node reset the original activenode

Dim bFlag As Boolean
Dim bMouseIsOver As Boolean
Dim bMoveCopy As Boolean

    If Not moLastActiveNode Is Nothing Then
        With Me.control
            ' is the mouse over the node or within a pixel of it
            bMouseIsOver = (x >= -1 And x <= .width + 1) And (y >= -1 And y <= .height + 1)
        End With
        
        If Not bMouseIsOver Then
            ' if the last-activenode was marked for MoveCopy we will need to reset it
            bFlag = moLastActiveNode Is moTree.MoveCopyNode(bMoveCopy)

            ' reset the original activenode
            moLastActiveNode.control.borderStyle = fmBorderStyleNone
            Set moTree.ActiveNode = moLastActiveNode

            If bFlag Then
                Set moTree.MoveCopyNode(bMoveCopy) = moLastActiveNode
            End If

            Set moLastActiveNode = Nothing
            
        ElseIf Button = 2 Then
            ' the click event doesn't fire with right click so explicitly call it
            moTree.NodeEventRouter Me, "Caption", tvRightClick
        End If
    End If
    
    ' moTree.NodeEventRouter Me, "Caption", tvUp, Button, Shift, X, Y

End Sub

Private Sub mctlExpander_Click()
'
    Expanded = Not Expanded
    If moTree.EditMode(Me) Then
        ' if any node is in edit mode exit edit mode
        moTree.EditMode(Me) = False
    End If
    moTree.NodeEventRouter Me, "Expander", 1
End Sub

Private Sub moEditBox_Change()
' PT resize horizontally & vertically to accommodate the changed text but restrict size to a bit less than the container frame
'    Autosize is good for width only if multiline=false, but it's problematic to size both width & height to text without this fix
Dim i As Long, ub As Long
Dim lSelSt As Long
Dim lRows As Long
Dim ht As Single, wd As Single
Dim sngMaxW As Single, sngMaxHt As Single, sngMaxR As Single, sngVisR
Dim sngIdealL, sngIdealTop As Single
Dim sEdit As String
Dim splt, v, splt2
Static bExit As Boolean
Const cPad As Long = 6

    If bExit Then Exit Sub

    On Error GoTo errH

    With moTree.TreeControl
        sngMaxW = .InsideWidth - cPad
        sngMaxHt = IIf(.ScrollHeight, .ScrollHeight, .InsideHeight) - cPad
        sngMaxR = IIf(.ScrollWidth > .InsideWidth, .ScrollWidth, .InsideWidth) - cPad
        sngVisR = .ScrollLeft + .InsideWidth
    End With

    sngIdealL = mctlControl.left - 1.5
    sngIdealTop = mctlControl.top - 1.5

    sEdit = moEditBox

    If InStr(1, sEdit, vbNewLine) Then
        splt = Split(sEdit, vbNewLine)
    ElseIf InStr(1, sEdit, vbCr) Then
        splt = Split(sEdit, vbCr)
    ElseIf InStr(1, sEdit, vbLf) Then
        splt = Split(sEdit, vbLf)
    Else
        ReDim splt(0)
        splt(0) = sEdit
    End If
    ub = UBound(splt)
    
    With moLabSizer

        .caption = "A"
        ht = .height
        .caption = ""
        lRows = 1

        For i = 0 To ub
            splt2 = Split(splt(i), " ")
            For Each v In splt2
                .caption = .caption & v & " "
                .AutoSize = True
                If .width > sngMaxW - 9 Then
                    lRows = lRows + 1
                    .caption = v & " "
                    If .width > wd Then
                        wd = .width
                        If .width > sngMaxW Then
                            lRows = lRows + 1
                        End If
                    End If
                    
                ElseIf .width > wd Then
                    wd = .width
                End If
            Next
            If i < ub Then
                lRows = lRows + 1
                .caption = ""
            End If
        Next
        If .width > wd Then wd = .width
        If wd + 15 > sngMaxW Then wd = sngMaxW Else wd = wd + 15
    End With

    bExit = True

    With moEditBox

        If lRows > 1 Then
            .AutoSize = True
            .MultiLine = True
            lSelSt = .SelStart
            .SelStart = 0
            .SelStart = lSelSt

            If ht * lRows + 9 > sngMaxHt Then
                .height = sngMaxHt - 3
            Else
                .height = ht * lRows + 9
            End If

            If sngIdealTop + .height > sngMaxHt - 3 Then
                .top = sngMaxHt - .height + 3
            Else
                .top = sngIdealTop
            End If
            .width = wd

        Else
            .MultiLine = False
            .AutoSize = True
            .MultiLine = True
            .top = mctlControl.top - 1.5
        End If

        
        If sngIdealL + .width > sngMaxR Then
            .left = sngMaxR - .width + 3
        Else
            .left = sngIdealL
        End If
        sngVisR = (.left + .width) - sngVisR
        If sngVisR > 0 Then
            moTree.TreeControl.ScrollLeft = moTree.TreeControl.ScrollLeft + sngVisR
        End If
          
    End With
done:
    bExit = False
    Exit Sub
    
errH:
    #If DebugMode = 1 Then
        Stop
        Resume
    #End If
    
    EditBox False
    Resume done
End Sub


Private Sub moEditBox_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)    'PT
' PT Textbox key events to Enter or Esc the Editbox,   006PT2
' PT ammended 026 to resize/reposition the editbox within the confines of the frame container
'    remove any CR & LF to ensure the node label is a single line but retain line breaks in the property if bMultiLine

    Dim bCancel As Boolean
    Dim bSort As Boolean, bMultiLine As Boolean
    Dim sNewText As String
    

    If keyCode = vbKeyReturn And shift = 0 Then
        sNewText = moEditBox
        If sNewText = Me.caption Then
            keyCode = vbKeyEscape
        Else
            ' note: sNewText could be parsed where the RaiseAfterLabelEdit is received
            bCancel = moTree.RaiseAfterLabelEdit(Me, sNewText)
            If Not bCancel Then
                bCancel = moTree.EnableLabelEdit(bSort, bMultiLine)
                
                If InStr(1, sNewText, vbCr) Or InStr(1, sNewText, vbLf) Then

                    #If Mac Then
                        '' Mac doesn't seem to use vbCrLf/vbNewLine, only vbCr or vbLf (?)
                        If bMultiLine Then
                            Me.caption = Replace(Replace(sNewText, vbCr, mcBreak), vbLf, mcBreak)
                        Else
                            Me.caption = Replace(Replace(sNewText, vbCr, "  "), vbLf, "  ")
                        End If
                    #Else
                        If bMultiLine Then
                            Me.caption = Replace(Replace(sNewText, vbCrLf, mcBreak), vbLf, mcBreak)
                        Else
                            Me.caption = Replace(Replace(sNewText, vbCrLf, "  "), vbLf, "  ")
                        End If
                    #End If

                    If bMultiLine Then
                    ' Drectly write the original multiline to the Caption, but don't use Let-Caption Property
                        mvCaption = sNewText
                    End If
                    
                Else
                    Me.caption = sNewText
                    
                End If

                mctlControl.WordWrap = False
                mctlControl.AutoSize = True
                TextWidth = mctlControl.width
                
                If TextWidth < mcFullWidth And moTree.FullWidth Then
                    control.width = mcFullWidth
                End If
                
                moTree.changed = True

                If bSort Then
                    If Me.parentNode.sort Then
                        moTree.Refresh
                    End If
                End If
                
                moTree.SetScrollbars True ' replaced moTree.SetScrollbars in 026
            End If
            EditBox False
        End If
    End If
    
    If keyCode = vbKeyEscape Then
        moTree.EditMode(Me) = False
        EditBox False
    End If

End Sub


Private Sub Class_Initialize()

    Set pSpecialButtons = New VBA.Collection
    
' default properties
    mbExpanded = True  ' default
    
    #If DebugMode = 1 Then
        gClsNodeInit = gClsNodeInit + 1    ' PT, for testing only, remove, see ClassCounts() in modDemo
    #End If
End Sub

Private Sub Class_Terminate()
    #If DebugMode = 1 Then
        gClsNodeTerm = gClsNodeTerm + 1    ' PT, for testing,
    #End If
    Set moTree = Nothing
    Set pSpecialButtons = Nothing
End Sub




'[mielk]
'Public Function isLeaf() As Boolean
'    isLeaf = Not F.Collections.isNonEmptyCollection(childNodes)
'End Function
'
'Public Function isVisible() As Boolean
'    If UI.Controls.isValid(mctlControl) Then
'        isVisible = mctlControl.visible
'    End If
'End Function
'
'Public Function getBottomY() As Single
'    If UI.Controls.isValid(mctlControl) Then
'        getBottomY = UI.Controls.getBottom(mctlControl)
'    End If
'End Function

Public Function setObject(value As ITreeViewItem) As clsNode
    Set setObject = Me
    Set pObject = value
End Function

Public Function setTreeView(tv As clsTreeView) As clsNode
    Set setTreeView = Me
    Set pTreeView = tv
End Function

Public Function setSpecialActions(dtos As VBA.Collection) As clsNode
    Set setSpecialActions = Me
    Set pSpecialActionDtos = dtos
    If UI.controls.isValid(mctlControl) Then
        Call addHoverBackground
        Call addSpecialButtons(pSpecialActionDtos)
    End If
End Function

Private Sub addSpecialButtons(buttonDtos As VBA.Collection)
    Dim dto As Scripting.Dictionary
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    Dim Button As UIIcon
    Dim left As Single:             left = getRightX + BUTTONS_LEFT_MARGIN
    '------------------------------------------------------------------------------------------------------
    
    Set pSpecialButtons = New VBA.Collection
    
    If F.Collections.isNonEmptyCollection(buttonDtos) Then
        For Each dto In buttonDtos
            Set Button = UI.Create.UIIcon(getFrame). _
                            setListener(Me). _
                            setActionName(F.dictionaries.getItem(dto, UI_PROP_ACTION_NAME)). _
                            render. _
                            Move(left). _
                            setHeight(SPECIAL_BUTTON_SIZE). _
                            setWidth(SPECIAL_BUTTON_SIZE). _
                            setTooltip(F.dictionaries.getItem(dto, UI_PROP_TOOLTIP)). _
                            setPicture(F.dictionaries.getItem(dto, UI_PROP_PICTURE)). _
                            alignVertically(mctlControl)
                            
                left = Button.getRight + 0 ' BUTTONS_SPACE
                Call pSpecialButtons.add(Button)
        Next dto
    End If
    
    If UI.controls.isValid(pHoverBackground) Then
        pHoverBackground.width = left - pHoverBackground.left
    End If
    
End Sub

Private Sub addHoverBackground()
    If pHoverBackground Is Nothing Then
        Set pHoverBackground = pTreeView.TreeControl.controls.add(CREATE_LABEL_ID, visible:=False)
        With pHoverBackground
            .width = mctlControl.width + 2 * MOVE_LAYER_MARGIN
            .height = mctlControl.height + 2 * MOVE_LAYER_MARGIN
            .top = mctlControl.top - MOVE_LAYER_MARGIN
            .left = mctlControl.left - MOVE_LAYER_MARGIN
            .BackStyle = fmBackStyleTransparent
            .borderStyle = fmBorderStyleNone
            .visible = True
        End With
    End If
End Sub

Public Sub updateSpecialButtons()
    If Not mctlControl Is Nothing Then
        Call addHoverBackground
        Call addSpecialButtons(pSpecialActionDtos)
    End If
End Sub

Public Sub setHoveredStatus(value As Boolean)
    If Not pTreeView Is Nothing Then
        If Not pTreeView.isInDragDropMode Then
            pCurrentHover = value
            Call showHideSpecialButtons(value)
            Call createMoveLabel(True)
        End If
    End If
End Sub

Private Sub showHideHoveBackground(visibility As Boolean)
    On Error Resume Next
    pHoverBackground.visible = visibility
End Sub

Private Sub showHideSpecialButtons(visibility As Boolean)
    Dim label As UIIcon
    Dim right As Single
    '------------------------------------------------------------------------------------------------------
    
    For Each label In pSpecialButtons
        Call label.alignVertically(mctlControl).setVisibility(visibility)
        right = label.getRight
    Next label
    
    If Not pHoverBackground Is Nothing Then
        pHoverBackground.width = F.Numerics.getMax(0, right - pHoverBackground.left + MOVE_LAYER_MARGIN)
    End If

End Sub

Public Function getObject() As ITreeViewItem
    Set getObject = pObject
End Function

Public Function getRightX() As Single
    If UI.controls.isValid(mctlControl) Then
        getRightX = mctlControl.left + mctlControl.width
    End If
End Function

Private Function getFrame() As MSForms.frame
    If moTree Is Nothing Then
        If Not pTreeView Is Nothing Then
            Set getFrame = pTreeView.TreeControl
        End If
    Else
        Set getFrame = moTree.TreeControl
    End If
End Function

Private Sub removeCustomControls()
    Call destroyMoveLabels
    Call destroyDropLabels
    Call destroySpecialButtons
End Sub

Private Sub destroyMoveLabels()
    On Error Resume Next
    Call UI.Forms.removeControls(pTreeView.TreeControl, pMoveLabel, pMoveCaptionLabel)
    Set pMoveLabel = Nothing
    Set pMoveCaptionLabel = Nothing
End Sub

Private Sub destroyDropLabels()
    On Error Resume Next
    Call UI.Forms.removeControls(pTreeView.TreeControl, pDropLabel, pDropCaptionLabel)
    Set pDropLabel = Nothing
    Set pDropCaptionLabel = Nothing
End Sub

Private Sub destroySpecialButtons()
    Dim btn As UIIcon
    '------------------------------------------------------------------------------------------------------
    For Each btn In pSpecialButtons
        Call btn.delete
    Next btn
End Sub



'[Drag & Drop]
Public Sub deactivate()
    Call destroyMoveLabels
    pMoveMode = False
End Sub

Private Sub createMoveLabel(Optional reset As Boolean = False)
    If reset Then Call destroyMoveLabels
    If pMoveLabel Is Nothing Then
        Set pMoveLabel = moTree.TreeControl.controls.add(CREATE_LABEL_ID, visible:=False)
    End If
    If pMoveCaptionLabel Is Nothing Then
        Set pMoveCaptionLabel = moTree.TreeControl.controls.add(CREATE_LABEL_ID, visible:=False)
    End If
    Call showHideMoveLabel(False)
End Sub

Private Sub createDropLabel(Optional reset As Boolean = False)
    If reset Then Call destroyDropLabels
    If pDropLabel Is Nothing Then
        Set pDropLabel = moTree.TreeControl.controls.add(CREATE_LABEL_ID, visible:=False)
    End If
    If pDropCaptionLabel Is Nothing Then
        Set pDropCaptionLabel = moTree.TreeControl.controls.add(CREATE_LABEL_ID, visible:=False)
    End If
    Call showHideDropLabel(False)
End Sub

Public Sub setDragView(value As Boolean)
    If value Then
        Call createMoveLabel
        Call showHideMoveLabel(True)
    Else
        Call destroyMoveLabels
    End If
End Sub

Public Sub setDropView(value As Boolean)
    If value Then
        Call createDropLabel
        Call showHideDropLabel(True)
    Else
        Call destroyDropLabels
    End If
End Sub

Public Sub quitDragMode()
    Call destroyMoveLabels
    pMoveMode = False
End Sub

Public Sub showEventsLayer(value As Boolean)
    If UI.controls.isValid(pHoverBackground) Then
        pHoverBackground.visible = value
    End If
End Sub

Private Sub showHideDropLabel(value As Boolean)
    Call showHideSpecialLabels(value, pDropLabel, pDropCaptionLabel, DROP_VIEW_BACK_COLOR, DROP_VIEW_BORDER_COLOR, True)
End Sub

Private Sub showHideMoveLabel(value As Boolean)
    Call showHideSpecialLabels(value, pMoveLabel, pMoveCaptionLabel, MOVE_LAYER_BACK_COLOR, MOVE_LAYER_BORDER_COLOR)
End Sub

Private Sub showHideSpecialLabels(value As Boolean, mainLabel As MSForms.label, _
                captionLabel As MSForms.label, backColor As Long, borderColor As Long, _
                Optional handCursor As Boolean = False)
    With mainLabel
        .width = mctlControl.width + 2 * MOVE_LAYER_MARGIN
        .height = mctlControl.height + 2 * MOVE_LAYER_MARGIN
        .top = mctlControl.top - MOVE_LAYER_MARGIN
        .left = mctlControl.left - MOVE_LAYER_MARGIN
        .caption = vbNullString
        .visible = value
        
        If value Then
            .BackStyle = fmBackStyleOpaque
            .backColor = backColor
            .borderStyle = fmBorderStyleSingle
            .borderColor = borderColor
            If handCursor Then
                .MousePointer = fmMousePointerCustom
                Set .MouseIcon = Icons.getMouseIcon
            Else
                .MousePointer = fmMousePointerDefault
            End If
        Else
            .BackStyle = fmBackStyleTransparent
            .borderStyle = fmBorderStyleNone
            .MousePointer = fmMousePointerDefault
        End If
    End With
    
    With captionLabel
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        If handCursor Then
            .MousePointer = fmMousePointerCross
            Set .MouseIcon = Icons.getMouseIcon
        Else
            .MousePointer = fmMousePointerDefault
        End If
        .caption = mctlControl.caption
        .font.size = mctlControl.font.size
        .font.Bold = mctlControl.font.Bold
        .ForeColor = vbBlack
        .width = mainLabel.width
        .AutoSize = True
        .AutoSize = False
        .left = (mctlControl.width - .width) / 2 + mctlControl.left
        .top = (mctlControl.height - .height) / 2 + mctlControl.top
        .visible = value
    End With
    
    Call locateSpecialButtons
    
End Sub

Public Sub setFontSize(size As Single)
    With mctlControl
        .font.size = size
        .AutoSize = True
        .AutoSize = False
        .width = .width + CTRL_RIGHT_PADDING
        .height = F.Numerics.roundTo(.height, PIXEL_SIZE) + PIXEL_SIZE / 3
        .font.size = F.Numerics.roundTo(.font.size, PIXEL_SIZE) + PIXEL_SIZE / 3
        .top = F.Numerics.roundTo(.top, PIXEL_SIZE) + PIXEL_SIZE / 3
    End With
End Sub



'[Implementation of IListener]
Private Sub IListener_catchEvent(src As Variant, ev As Variant, Optional properties As Scripting.IDictionary)
    Call moTree.NodeEventRouter(Me, src.getActionName, tvClick)
End Sub

Public Sub updatePosition(x As Single, y As Single)
    Dim offsetY As Single
    Dim offsetX As Single
    '------------------------------------------------------------------------------------------------------
    
    If pMoveMode Then
        offsetX = x - pMoveAnchorPoint.x
        offsetY = y - pMoveAnchorPoint.y
        
        With pMoveLabel
            .top = y
            .left = x
        End With
        
        With pMoveCaptionLabel
            .top = y
            .left = x
        End With
    End If
    
End Sub

Private Sub pDropCaptionLabel_Click()
    Call pTreeView.setAsSelectedDropNode(Me)
End Sub

Private Sub pDropLabel_Click()
    Call pTreeView.setAsSelectedDropNode(Me)
End Sub

Private Sub pHoverBackground_MouseMove(ByVal Button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
    Call setAsHoveredNode
End Sub

Private Sub setAsHoveredNode()
    If Not pMoveMode Then
        If Not pCurrentHover Then
            Call moTree.setHoveredNode(Me)
            Call setHoveredStatus(True)
        End If
    End If
End Sub

Public Sub locateSpecialButtons()
    Dim img As UIIcon
    Dim backgroundTop As Single
    Dim iconsTop As Single
    Dim left As Single:             left = getRightX + BUTTONS_LEFT_MARGIN
    '------------------------------------------------------------------------------------------------------
    
    If Not F.Collections.isNonEmptyCollection(pSpecialButtons) Then
        '
    ElseIf pHoverBackground Is Nothing Then
        '
    ElseIf mctlControl.visible Then
        With mctlControl
            backgroundTop = F.Numerics.roundTo(.top + (.height - pHoverBackground.height) / 2, PIXEL_SIZE)
            iconsTop = F.Numerics.roundTo(.top + (.height - SPECIAL_BUTTON_SIZE) / 2, PIXEL_SIZE)
        End With
        For Each img In pSpecialButtons
            With img
                Call .setEnability(True)
                Call .Move(x:=left, y:=iconsTop)
                left = .getRight + BUTTONS_SPACE
            End With
        Next img
        
        With pHoverBackground
            .visible = True
            .top = backgroundTop
            .width = left - .left
        End With
    Else
        pHoverBackground.visible = False
        For Each img In pSpecialButtons
            Call img.setEnability(False)
        Next img
    End If
End Sub



Public Function getLeafs() As Scripting.Dictionary
    Dim childNodes As VBA.Collection
    Dim child As clsNode
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set getLeafs = F.dictionaries.Create(False)
    Set childNodes = Me.childNodes
    
    If F.Collections.isNonEmptyCollection(childNodes) Then
        For Each child In childNodes
            Call F.dictionaries.appendDictionary(getLeafs, child.getLeafs)
        Next child
    Else
        If F.Numerics.IsNumber(Me.key) Then key = F.Numerics.toNumber(Me.key) Else key = Me.key
        Call getLeafs.add(key, Me.Checked)
    End If
    
End Function
