VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FNormalize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FNormalize"
'----------------------------------------------------------------------------------------------------------
Private pScaling As Coordinate
'----------------------------------------------------------------------------------------------------------





Private Sub Class_Initialize()
    pScaling = F.Screen.getDpiScaling
End Sub



Public Function position(ByVal value As Single) As Single
    position = F.Numerics.roundTo(VBA.CDbl(value), PIXEL_SIZE / pScaling.y)
End Function

Public Function size(ByVal value As Single) As Single
    size = F.Numerics.roundTo(VBA.CDbl(value), PIXEL_SIZE / pScaling.y)
End Function

Public Function fontSize(ByVal value As Single) As Single
    fontSize = F.Numerics.roundTo(VBA.CDbl(value), PIXEL_SIZE)
End Function
