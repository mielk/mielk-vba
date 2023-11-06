Attribute VB_Name = "modRegexPatterns"
Option Explicit

Private Const CLASS_NAME As String = "modRegexPatterns"
'[Dates] --------------------------------------------------------------------------------------------------
Public Const MONTH_YEAR_REGEX_PATTERN As String = "^\s*(\d{1,2})\/(\d{4})\s*$"
'[Excel] --------------------------------------------------------------------------------------------------
Public Const MACRO_FILE_REGEX_PATTERN As String = "^[^~].*\.xl(s|a)m$"
'----------------------------------------------------------------------------------------------------------
