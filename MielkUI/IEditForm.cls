VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Interface groupping all classes that represents window to edit a record.

Private Const CLASS_NAME As String = "IEditForm"
'----------------------------------------------------------------------------------------------------------


Public Function isCancelledByUser() As Boolean
End Function

Public Sub run(Optional mode As FormShowConstants = vbModal, Optional locationProps As DLocation)
End Sub
