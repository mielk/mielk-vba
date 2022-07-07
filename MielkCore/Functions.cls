VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Functions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "Functions"
'----------------------------------------------------------------------------------------------------------
Private pArrays As FArrays
Private pBooks As FBooks
Private pCast As FCast
Private pCollections As FCollections
Private pColors As FColors
Private pConstructors As FConstructors
Private pContextMenu As FContextMenu
Private pCrypto As FCryptography
Private pDatabase As FDatabase
Private pDates As FDates
Private pDeveloper As FDeveloper
Private pDialogs As FDialogs
Private pDictionaries As FDictionaries
Private pExcel As FExcel
Private pExcelNames As FExcelNames
Private pFiles As FFiles
Private pJson As FJSON
Private pMeta As FMeta
Private pNumerics As FNumerics
Private pObjects As FObjects
Private pRanges As FRanges
Private pReflection As FReflection
Private pRegex As FRegex
Private pRegistry As FRegistry
Private pScreen As FScreen
Private pSheets As FSheets
Private pStrings As FStrings
Private pSystem As FSystem
Private pTextFiles As FTextFile
Private pUtils As FUtils
Private pVariables As FVariables
'----------------------------------------------------------------------------------------------------------


'[Initializer]
Private Sub Class_Initialize()
    Set pArrays = New FArrays
    Set pBooks = New FBooks
    Set pCast = New FCast
    Set pCollections = New FCollections
    Set pContextMenu = New FContextMenu
    Set pColors = New FColors
    Set pConstructors = New FConstructors
    Set pCrypto = New FCryptography
    Set pDates = New FDates
    Set pDatabase = New FDatabase
    Set pDeveloper = New FDeveloper
    Set pDialogs = New FDialogs
    Set pDictionaries = New FDictionaries
    Set pExcel = New FExcel
    Set pExcelNames = New FExcelNames
    Set pFiles = New FFiles
    Set pJson = New FJSON
    Set pMeta = New FMeta
    Set pNumerics = New FNumerics
    Set pObjects = New FObjects
    Set pRanges = New FRanges
    Set pReflection = New FReflection
    Set pRegex = New FRegex
    Set pRegistry = New FRegistry
    Set pScreen = New FScreen
    Set pSheets = New FSheets
    Set pStrings = New FStrings
    Set pSystem = New FSystem
    Set pTextFiles = New FTextFile
    Set pUtils = New FUtils
    Set pVariables = New FVariables
End Sub

'[Submodules]
Public Function arrays() As FArrays
    Set arrays = pArrays
End Function

Public Function Books() As FBooks
    Set Books = pBooks
End Function

Public Function cast() As FCast
    Set cast = pCast
End Function

Public Function Collections() As FCollections
    Set Collections = pCollections
End Function

Public Function ContextMenu() As FContextMenu
    Set ContextMenu = pContextMenu
End Function

Public Function colors() As FColors
    Set colors = pColors
End Function

Public Function Create() As FConstructors
    Set Create = pConstructors
End Function

Public Function Crypto() As FCryptography
    Set Crypto = pCrypto
End Function

Public Function Dates() As FDates
    Set Dates = pDates
End Function

Public Function Db() As FDatabase
    Set Db = pDatabase
End Function

Public Function Developer() As FDeveloper
    Set Developer = pDeveloper
End Function

Public Function Dialogs() As FDialogs
    Set Dialogs = pDialogs
End Function

Public Function dictionaries() As FDictionaries
    Set dictionaries = pDictionaries
End Function

Public Function Excel() As FExcel
    Set Excel = pExcel
End Function

Public Function ExcelNames() As FExcelNames
    Set ExcelNames = pExcelNames
End Function

Public Function files() As FFiles
    Set files = pFiles
End Function

Public Function json() As FJSON
    Set json = pJson
End Function

Public Function Math() As FNumerics
    Set Math = pNumerics
End Function

Public Function Meta() As FMeta
    Set Meta = pMeta
End Function

Public Function Numerics() As FNumerics
    Set Numerics = pNumerics
End Function

Public Function Objects() As FObjects
    Set Objects = pObjects
End Function

Public Function ranges() As FRanges
    Set ranges = pRanges
End Function

Public Function Reflection() As FReflection
    Set Reflection = pReflection
End Function

Public Function regex() As FRegex
    Set regex = pRegex
End Function

Public Function Registry() As FRegistry
    Set Registry = pRegistry
End Function

Public Function Screen() As FScreen
    Set Screen = pScreen
End Function

Public Function Sheets() As FSheets
    Set Sheets = pSheets
End Function

Public Function strings() As FStrings
    Set strings = pStrings
End Function

Public Function System() As FSystem
    Set System = pSystem
End Function

Public Function TextFiles() As FTextFile
    Set TextFiles = pTextFiles
End Function

Public Function Utils() As FUtils
    Set Utils = pUtils
End Function

Public Function Variables() As FVariables
    Set Variables = pVariables
End Function
