Attribute VB_Name = "modActions"
Option Explicit

Private Const CLASS_NAME As String = "modActions"
'----------------------------------------------------------------------------------------------------------


Public Sub addErrorHandlingToCurrentMethod()
    Call ErrorHandling.addErrorHandlingToCurrentMethod
End Sub

Public Sub createFramedSection()
    Call CodeFrameGenerator.addFrame
End Sub

Public Sub addSeparatorLine()
    Call CodeFrameGenerator.addSeparatorLine
End Sub

Public Sub addClass()
    Call ClassGenerator.addClass
End Sub

Public Sub addSettersAndGetters()
    Call ClassGenerator.addSettersAndGetters
End Sub
