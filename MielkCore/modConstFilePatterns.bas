Attribute VB_Name = "modConstFilePatterns"
Option Explicit

Private Const CLASS_NAME As String = "modConstFilePatterns"
'[Extensions] ---------------------------------------------------------------------------------------------
Public Const EXTENSION_CSV As String = ".csv"
Public Const EXTENSION_TXT As String = ".txt"
Public Const EXTENSION_EXCEL_ADDIN As String = ".xlam"
Public Const EXTENSION_EXCEL_MACRO_FILE As String = ".xlsm"
Public Const EXTENSION_JSON As String = ".json"
Public Const EXTENSION_ZIP As String = ".zip"
'[File patterns] ------------------------------------------------------------------------------------------
Public Const FILES_PATTERN_ACCESS As String = "Access files, *.mdb; *.mde; *.accdb; *.accde"
Public Const FILES_PATTERN_CSV As String = "CSV files, *.csv"
Public Const FILES_PATTERN_EXCEL As String = "Excel files, *.xls; *.xlsm; *.xlsx; *.xlsb"
Public Const FILES_PATTERN_EXCEL_NO_MACRO As String = "Excel files, *.xls; *.xlsx; *.xlsb"
Public Const FILES_PATTERN_EXCEL_MACRO As String = "Excel macro files, *.xlsm; *.xlsb; *.xla; *.xlam"
Public Const FILES_PATTERN_EXCEL_XLSM As String = "Excel Macro Enabled Workbook (*.xlsm), *.xlsm"
Public Const FILES_PATTERN_JSON As String = "JSON files, *.json"
'[File types codes] ---------------------------------------------------------------------------------------
Public Const FILE_TYPE_CODE_CSV As String = "csv"
Public Const FILE_TYPE_CODE_EXCEL As String = "xls"
'----------------------------------------------------------------------------------------------------------
