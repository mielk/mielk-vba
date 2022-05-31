VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IListParentWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Interface for UI objects that can contain any implementation of IList.

Private Const CLASS_NAME As String = "IListParentWindow"
'----------------------------------------------------------------------------------------------------------

Public Sub restoreFromBackground()
End Sub

Public Sub sendToBackground()
End Sub

Public Function getDefaultLocationProps() As DLocation
End Function
