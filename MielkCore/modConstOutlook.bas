Attribute VB_Name = "modConstOutlook"
Option Explicit

Private Const CLASS_NAME As String = "modConstOutlook"

'[Outlook] ------------------------------------------------------------------------------------------------
Public Const OUTLOOK_APP As String = "Outlook.Application"
Public Const OL_CLASS_NAME_MAIL_ITEM As String = "MailItem"
'----------------------------------------------------------------------------------------------------------
Public Const olMailItem As Long = 0
'__ Mail importance __
Public Const olImportanceLow = 0
Public Const olImportanceNormal = 1
Public Const olImportanceHigh = 2
'__ Folders __
Public Const olFolderDeletedItems = 3
Public Const olFolderDrafts = 16
'__ Other Outlook constants __
Public Const olFormatHtml = 2
'----------------------------------------------------------------------------------------------------------
