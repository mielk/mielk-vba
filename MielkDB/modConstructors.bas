Attribute VB_Name = "modConstructors"
Option Explicit

Private Const CLASS_NAME As String = "fnCreators"
'----------------------------------------------------------------------------------------------------------


Public Function D() As DbEngine
    Static instance As DbEngine
    '----------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then Set instance = New DbEngine
    Set D = instance
End Function


