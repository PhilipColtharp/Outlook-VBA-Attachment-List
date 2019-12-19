'Core original code: Diane Poremsky
'https://www.slipstick.com/developer/code-samples/add-attachments-names-to-message-before-sending/

'Other Code adaptations and inspirations from
' * https://www.mrexcel.com/board/threads/check-to-see-if-outlook-has-an-attachment-and-is-going-to-an-outside-domain.1100867/#post5294426
' * https://www.datanumen.com/blogs/batch-find-replace-text-multiple-outlook-emails/

Public Sub AddAttachmentNames(oItem As MailItem)

  Dim oAtt As Attachment
  Dim strAtt As String
  Dim HTMLImgSrc As String
    Dim olInspector As Outlook.Inspector
    Dim olDocument As Word.Document
    Dim olSelection As Word.Selection

  If oItem.BodyFormat = olFormatHTML Then
  
     If oItem.Attachments.Count > 0 Then
       HTMLImgSrc = getHTMLImgSrc(oItem)
     
       strAtt = "Message Attachments:" & vbCrLf
        
       For Each oAtt In oItem.Attachments
         Select Case (oAtt.Type)
           Case olOLE
             'Avoiding the error: "Outlook cannot perform this action on this type of attachment"
             'OlAttachmentType enumeration, olOLE=6, attachment is an OLE document.
             'https://docs.microsoft.com/en-us/office/vba/api/outlook.olattachmenttype
           Case Else
             If IsFileAttachment(oAtt, HTMLImgSrc) _
               Then strAtt = strAtt & "  •  " & oAtt.FileName & vbCrLf
         End Select
       Next oAtt
           
     Else
       strAtt = ""
     End If
    
     Set olInspector = Application.ActiveInspector()
     Set olDocument = olInspector.WordEditor
     Set olSelection = olDocument.Application.Selection
       olSelection.HomeKey Unit:=wdStory
    
    'source code: https://www.datanumen.com/blogs/batch-find-replace-text-multiple-outlook-emails/
    
     With olSelection.Find
       .ClearFormatting
       .Text = "<attachment list>"
       .Replacement.ClearFormatting
       .Replacement.Text = strAtt
       .Forward = True
       .Wrap = wdFindStop
       .Format = False
       .MatchCase = True
       .MatchWholeWord = False
       .Execute Replace:=wdReplaceOne
     End With
  End If
  
Set oItem = Nothing
End Sub

'inspired from https://www.mrexcel.com/board/threads/check-to-see-if-outlook-has-an-attachment-and-is-going-to-an-outside-domain.1100867/#post5294426

Private Function getHTMLImgSrc(ByRef outMailItem As MailItem) As String
  Dim HTMLdoc As Object
  Dim HTMLImgSrc_ As String
  Dim fName As String
    
  'Get src attribute in every HTML img tag
  Set HTMLdoc = CreateObject("HTMLfile")
  HTMLdoc.Open
  HTMLdoc.Write outMailItem.HTMLBody
  HTMLdoc.Close
  For Each img In HTMLdoc.getElementsByTagName("IMG")
      HTMLImgSrc_ = HTMLImgSrc_ & img.src & " "
  Next
  'MsgBox HTMLImgSrc_
  getHTMLImgSrc = HTMLImgSrc_
  
End Function


Private Function IsFileAttachment(outAttachment As Outlook.Attachment, allImgSrc As String) As Boolean
    'Returns True is the attachment is a file attachment, or False if it's a embedded or inline attachment
    IsFileAttachment_ = True
    If IsFileAttachment_ Then
      If InStr(1, allImgSrc, "cid:" & outAttachment.FileName, vbTextCompare) _
        Then IsFileAttachment_ = False
    End If
    IsFileAttachment = IsFileAttachment_
End Function
