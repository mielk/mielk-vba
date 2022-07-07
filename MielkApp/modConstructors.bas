Attribute VB_Name = "modConstructors"
Option Explicit

Private Const CLASS_NAME As String = "modConstructors"
'----------------------------------------------------------------------------------------------------------



Public Function MApp() As MApp
    Static instance As MApp
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then
        Set instance = New MApp
    End If
    Set MApp = instance
End Function


Public Function ActionListener() As ActionListener
    Static instance As ActionListener
    '------------------------------------------------------------------------------------------------------
    If instance Is Nothing Then
        Set instance = New ActionListener
    End If
    Set ActionListener = instance
End Function


'
''[Constructors]
'Public Function createSettingsManager(parent As IApplication, _
'                                        registryBaseKey As Long, _
'                                        registryProjectKey As String) As AppSettingsManager
'    Set createSettingsManager = New AppSettingsManager
'    With createSettingsManager
'        Call .setParent(parent)
'        Call .setRegistryKeys(registryBaseKey, registryProjectKey)
'    End With
'End Function
'
'Public Function createSetting(nameTag As String, registryKey As String) As CSetting
'    Set createSetting = New CSetting
'    With createSetting
'        Call .setTag(nameTag)
'        Call .setRegistryKey(registryKey)
'    End With
'End Function
'
'
'
'
''[Processing items]
'Public Function createItemsProcessor() As MItemsProcessor
'    Set createItemsProcessor = New MItemsProcessor
'End Function
'
'
'
''[Standarizer]
'Public Function createStandarizer(translator As ITranslator) As MStandarizer
'    Set createStandarizer = New MStandarizer
'    With createStandarizer
'        Call .setTranslator(translator)
'        Call .initialize
'    End With
'End Function
'
'Public Function createStandardName(id As Long, name As String) As MStandardName
'    Set createStandardName = New MStandardName
'    With createStandardName
'        Call .setId(id)
'        Call .setName(name)
'    End With
'End Function
'
'Public Function createStandarizerView(parent As MStandarizer) As MStandarizerView
'    Set createStandarizerView = New MStandarizerView
'    With createStandarizerView
'        Call .setParent(parent)
'    End With
'End Function
