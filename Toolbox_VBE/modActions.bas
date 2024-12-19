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

Public Sub addProps()
    Call ClassGenerator.addProps
End Sub

Public Sub addDataTypes()
    Dim project As VBIDE.VBProject
    Dim col As VBA.Collection
    Dim arr As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set project = Fn.getActiveProject
    
    Set col = New VBA.Collection
    Call col.Add(VBA.Array("Channels", "[dbo].[channels]", "[dbo].[channels_read]"))
    Call col.Add(VBA.Array("AliasesChannels", "[dbo].[AliasesChannels]", "[dbo].[aliases_channels]"))
    
    Call DataTypesGenerator.addDataTypes(project, col)
    
End Sub

Public Sub addRepository()
    Call RepoGenerator.run
End Sub
