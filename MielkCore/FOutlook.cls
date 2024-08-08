VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const VALID_MAIL_PATTERN As String = "([a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*)"
Private Const RECIPIENT_TYPE_RECIPIENT As Long = 1
Private Const RECIPIENT_TYPE_CC As Long = 2


Public Function isValidMailAddress(address As String) As Boolean
    isValidMailAddress = f.regex.checkIfMatch(address, "^" & VALID_MAIL_PATTERN & "$")
End Function

Public Function isMailItem(value As Variant) As Boolean
    If Not VBA.IsObject(value) Then
        isMailItem = False
    ElseIf value Is Nothing Then
        isMailItem = False
    Else
        isMailItem = f.strings.compareStrings(VBA.TypeName(value), OL_CLASS_NAME_MAIL_ITEM)
    End If
End Function

Public Function getMailAddressesFromString(addresses As String) As VBA.Collection
    Dim matches As VBA.Collection
    Dim match As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set getMailAddressesFromString = New VBA.Collection
    Set matches = f.regex.getMatchArraysCollection(addresses, VALID_MAIL_PATTERN)
    For Each match In matches
        If f.arrays.is1D(match) Then
            Call getMailAddressesFromString.Add(match(LBound(match)))
        ElseIf f.Variables.isString(match) Then
            Call getMailAddressesFromString.Add(match)
        End If
    Next match
End Function

Public Function getMapiFolder() As Object
    Dim ol As Object                    'Outlook.Application
    Dim namespace As Object             'Outlook.Namespace
    Dim folder As Object                'Outlook.Folder
    '------------------------------------------------------------------------------------------------------
    
    Set ol = VBA.CreateObject(OUTLOOK_APP)
    Set namespace = ol.GetNamespace("MAPI")
    
    For Each folder In namespace.Folders
        getMapiFolder = folder
    Next folder
    
End Function


Public Function sendMail(content As String, subject As String, recipients As String, _
                        Optional importance As Variant = olImportanceNormal)
    Dim ol As Object                    'Outlook.Application
    Dim mail As Object                  'Outlook.MailItem
    '------------------------------------------------------------------------------------------------------
    
    Set ol = VBA.CreateObject(OUTLOOK_APP)
    Set mail = ol.CreateItem(olMailItem)
    With mail
        .body = content
        .subject = subject
        .To = recipients
        .importance = importance
        Call .send
    End With

End Function


Public Function createMailDraft(subject As String, body As String, recipients As String, _
                                    Optional attachments As VBA.Collection) As Object
    Dim ol As Object                    'Outlook.Application
    Dim mail As Object                  'Outlook.MailItem
    Dim attachmentFile As Scripting.file
    '------------------------------------------------------------------------------------------------------
    
    Set ol = VBA.CreateObject(OUTLOOK_APP)
    Set createMailDraft = ol.CreateItem(olMailItem)
    With createMailDraft
        .body = body
        .subject = subject
        .To = recipients
        
        If f.Collections.isNonEmptyCollection(attachments) Then
            For Each attachmentFile In attachments
                Call .attachments.Add(attachmentFile.path)
            Next attachmentFile
        End If
        
        Call .save
        
    End With

End Function



Public Function getMailDraft(Optional subject As Variant, Optional recipients As Variant, _
                             Optional cc As Variant, Optional body As Variant) As Object
    Dim ol As Object                    'Outlook.Application
    Dim namespace As Object             'Outlook.Namespace
    Dim drafts As Object                'Outlook.folder
    Dim draft As Object                 'Outlook.MailItem
    '------------------------------------------------------------------------------------------------------
    Dim colRecipients As VBA.Collection
    Dim colCc As VBA.Collection
    Dim colDraftRecipients As VBA.Collection
    Dim colDraftCc As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    Set ol = VBA.CreateObject(OUTLOOK_APP)
    Set namespace = ol.GetNamespace("MAPI")
    Set drafts = namespace.GetDefaultFolder(olFolderDrafts)
    
    'Convert addresses with names to collection of addresses only
    Set colRecipients = getMailAddressesFromString(VBA.CStr(recipients))
    Set colCc = getMailAddressesFromString(VBA.CStr(cc))
    
    
    
    If Not drafts Is Nothing Then
        For Each draft In drafts.items
            If f.strings.compareStrings(VBA.TypeName(draft), OL_CLASS_NAME_MAIL_ITEM) Then
            
                Call getMailRecipients(draft, colDraftRecipients, colDraftCc)
                
                If (Not f.Variables.isUndefined(subject) And draft.subject <> subject) Then
                    ' subject is defined and different than in draft mail
                ElseIf (Not f.Variables.isUndefined(recipients) And Not compareRecipients(colDraftRecipients, colRecipients)) Then
                    ' recipients are defined and different than in draft mail
                ElseIf (Not f.Variables.isUndefined(cc) And Not compareRecipients(colDraftCc, colCc)) Then
                    ' cc are defined and different than in draft mail
                ElseIf (Not f.Variables.isUndefined(body) And clearBodyText(draft.body) <> clearBodyText(VBA.CStr(body))) Then
                    ' body is defined and different than in draft mail
                Else
                    Set getMailDraft = draft
                    Exit For
                End If
            
            End If
        Next draft
    End If
    
    
End Function


Private Function getMailRecipients(mail As Object, ByRef outRecipients As VBA.Collection, ByRef outCc As VBA.Collection) As VBA.Collection
    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    '------------------------------------------------------------------------------------------------------
    Dim recipients As Object            'Outlook.Recipients
    Dim recipient As Object             'Outlook.Recipient
    Dim propAccessor As Object          'Outlook.PropertyAccessor
    Dim address As String
    '------------------------------------------------------------------------------------------------------
    
    Set getMailRecipients = New VBA.Collection
    Set outRecipients = New VBA.Collection
    Set outCc = New VBA.Collection
    
    Set recipients = mail.recipients
    For Each recipient In recipients
        address = f.regex.getFirstGroupMatch(recipient.name, VALID_MAIL_PATTERN)
        If VBA.Len(address) = 0 Then
            On Error Resume Next
            Set propAccessor = recipient.PropertyAccessor
            address = propAccessor.getProperty(PR_SMTP_ADDRESS)
            On Error GoTo 0
        End If
        
        address = f.strings.coalesce(address, recipient.name)
        
        If recipient.Type = RECIPIENT_TYPE_RECIPIENT Then
            Call outRecipients.Add(address)
        Else
            Call outCc.Add(address)
        End If
        Call getMailRecipients.Add(address)
        
    Next recipient
        
End Function


Private Function clearBodyText(text As String) As String
    Dim SUFFIX_PATTERN As String: SUFFIX_PATTERN = VBA.Chr(32) & VBA.Chr(13) & VBA.Chr(10)
    clearBodyText = f.strings.cutFromEnd(text, SUFFIX_PATTERN)
End Function


Private Function compareRecipients(recipients1 As VBA.Collection, recipients2 As VBA.Collection)
    Dim dict1 As Scripting.Dictionary
    Dim dict2 As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    Set dict1 = f.Collections.toDictionary(recipients1)
    Set dict2 = f.Collections.toDictionary(recipients2)
    
    compareRecipients = f.dictionaries.compare(dict1, dict2)
    
End Function

Private Function createRecipientsAddressesDictionary(recipients As String) As Scripting.Dictionary
    Dim arr As Variant
    Dim i As Long
    Dim address As String
    '------------------------------------------------------------------------------------------------------
    
    Set createRecipientsAddressesDictionary = f.dictionaries.Create(False)
    arr = VBA.Split(recipients, ";")
    If f.arrays.isDefined(arr) Then
        For i = LBound(arr) To UBound(arr)
            address = extractRecipientAddress(VBA.CStr(arr(i)))
            Call f.dictionaries.addItem(createRecipientsAddressesDictionary, address, address)
        Next i
    End If
    
End Function

Private Function extractRecipientAddress(fullRecipient As String) As String
    extractRecipientAddress = f.strings.substring(fullRecipient, "<", ">")
    If VBA.Len(extractRecipientAddress) = 0 Then extractRecipientAddress = f.strings.substring(fullRecipient, "'", "'")
    If VBA.Len(extractRecipientAddress) = 0 Then extractRecipientAddress = VBA.trim(fullRecipient)
End Function


Public Function openMailFromFile(filepath As String) As Object                  ' MailItem
    Dim ol As Object                    'Outlook.Application
    '------------------------------------------------------------------------------------------------------
    Set ol = VBA.CreateObject(OUTLOOK_APP)
    Set openMailFromFile = ol.session.openSharedItem(filepath)
End Function

Public Function addMailFromFileToDraftFolder(filepath As String) As Object      ' MailItem
    Dim copyPath As String
    '------------------------------------------------------------------------------------------------------
    Dim ol As Object                    'Outlook.Application
    Dim namespace As Object             'Outlook.namespace
    Dim drafts As Object                'Outlook.folder
    '------------------------------------------------------------------------------------------------------
    
    copyPath = f.files.uniqueFilePath(filepath)
    Call fso.CopyFile(filepath, copyPath)
    
    Set ol = VBA.CreateObject(OUTLOOK_APP)
    Set addMailFromFileToDraftFolder = ol.session.openSharedItem(copyPath)
    Set namespace = ol.GetNamespace("MAPI")
    Set drafts = namespace.GetDefaultFolder(olFolderDrafts)
    
    Call addMailFromFileToDraftFolder.Move(drafts)
    
    Call f.files.deleteFile(copyPath)
    
End Function


Public Sub moveMailToToDraftFolder(mail As Object)
    Const MAX_ATTEMPTS As Long = 5
    '------------------------------------------------------------------------------------------------------
    Dim ol As Object                    'Outlook.Application
    Dim namespace As Object             'Outlook.namespace
    Dim drafts As Object                'Outlook.folder
    Dim attempt As Long
    Dim entryId As String
    '------------------------------------------------------------------------------------------------------
    
    Set ol = VBA.CreateObject(OUTLOOK_APP)
    Set namespace = ol.GetNamespace("MAPI")
    
    Do
        Call Sleep(150)
        Set drafts = namespace.GetDefaultFolder(olFolderDrafts)
        Call mail.Move(drafts)
        
        attempt = attempt + 1
        If attempt > MAX_ATTEMPTS Then Exit Do
        
    Loop Until isMailInDraftFolder(mail)

End Sub



Private Function isMailInDraftFolder(mail As Object) As Boolean
    Dim ol As Object                    'Outlook.Application
    Dim namespace As Object             'Outlook.namespace
    Dim drafts As Object                'Outlook.folder
    Dim conversationIndex As String
    Dim obj As Object
    '------------------------------------------------------------------------------------------------------
    
    Set ol = VBA.CreateObject(OUTLOOK_APP)
    Set namespace = ol.GetNamespace("MAPI")
    Set drafts = namespace.GetDefaultFolder(olFolderDrafts)
    conversationIndex = mail.conversationIndex
    
    For Each obj In drafts.items
        If isMailItem(obj) Then
            If obj.conversationIndex = conversationIndex Then
                isMailInDraftFolder = True
                Exit For
            End If
        End If
    Next obj
    
End Function


Public Sub removeAttachmentsFromMail(mail As Object)
    Dim Counter As Long
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    Counter = mail.attachments.count
    For i = Counter To 1 Step -1
        Call mail.attachments.Remove(i)
    Next i
End Sub

Private Sub removeMailPermanently(subject As String)
    Dim ol As Object                    'Outlook.Application
    Dim namespace As Object             'Outlook.namespace
    Dim deletedFolder As Object         'Outlook.folder
    Dim items As Object                 'Outlook.Items
    Dim mail As Object                  'Outlook.MailItem
    Dim item As Object
    '------------------------------------------------------------------------------------------------------

    On Error Resume Next
    Set ol = VBA.CreateObject(OUTLOOK_APP)
    Set namespace = ol.GetNamespace("MAPI")
    Set deletedFolder = namespace.GetDefaultFolder(olFolderDeletedItems)
    Set items = deletedFolder.items
    On Error GoTo 0

    If Not items Is Nothing Then
        Call items.sort("[ReceivedTime]", True)
        For Each item In items
            If isMailItem(item) Then
                If item.subject = subject Then
                    Call item.delete
                End If
            End If
        Next item
    End If

End Sub



Public Function rangeToHtml(rng As Excel.Range) As String
    Dim filepath As String
    Dim wkb As Excel.Workbook
    Dim wks As Excel.Worksheet
    Dim pubObject As Excel.PublishObject
    Dim stream As Scripting.TextStream
    '------------------------------------------------------------------------------------------------------
    
    filepath = VBA.Environ$("temp") & "\" & VBA.format(Now, "yyyy-mm-dd hh-mm-ss") & ".htm"
    
    Call rng.copy
    
    Set wkb = Workbooks.Add(1)
    Set wks = wkb.Worksheets(1)
    With wks
        Call .cells(1).PasteSpecial(Paste:=8)
        Call .cells(1).PasteSpecial(xlPasteValues, , False, False)
        Call .cells(1).PasteSpecial(xlPasteFormats, , False, False)
        Call .cells(1).Select
        
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.visible = True
        Call .DrawingObjects.delete
        On Error GoTo 0
        
        Set pubObject = wkb.PublishObjects.Add(xlSourceRange, filepath, .name, .UsedRange.address, xlHtmlStatic)
        Call pubObject.Publish(True)
        
    End With
    
    
    Set stream = fso.getFile(filepath).OpenAsTextStream(ForReading, TristateMixed)
    With stream
        rangeToHtml = .ReadAll
        Call .Close
    End With
    rangeToHtml = VBA.replace(rangeToHtml, "align=center x:publishsource=", "align=left x:publishsource=")
    
    '[Clean up]
    Call wkb.Close(False)
    Call VBA.Kill(filepath)
    Set stream = Nothing
    Set wkb = Nothing
    
End Function
