'Core original code: Diane Poremsky
'https://www.slipstick.com/developer/code-samples/add-attachments-names-to-message-before-sending/

'Other Code adaptations and inspirations from
' * https://www.mrexcel.com/board/threads/check-to-see-if-outlook-has-an-attachment-and-is-going-to-an-outside-domain.1100867/#post5294426
' * https://www.datanumen.com/blogs/batch-find-replace-text-multiple-outlook-emails/

Public Sub AddAttachmentNames(oItem As MailItem)
  Dim Att_List As String
  Dim AttCount As Integer
     
  If oItem.BodyFormat = olFormatHTML Then
  
    Get_Attachments_List oItem, Att_List, AttCount
                       
    If AttCount > 0 Then
        Att_List = "Message Attachments:" & vbCrLf & Att_List
      Else
        Att_List = ""
    End If
    
    Find_and_Replace "<attachment list>", Att_List
  End If
  
  Set oItem = Nothing
End Sub

Private Sub Find_and_Replace(Find_ As String, _
                             Replace_ As String)
  Dim olInspector As Outlook.Inspector
  Dim olDocument As Word.Document
  Dim olSelection As Word.Selection
    
  Set olInspector = Application.ActiveInspector()
  Set olDocument = olInspector.WordEditor
  Set olSelection = olDocument.Application.Selection
  
  olSelection.HomeKey Unit:=wdStory

  With olSelection.Find
    .ClearFormatting
    .Text = Find_
    .Replacement.ClearFormatting
    .Replacement.Text = Replace_
    .Forward = True
    .Wrap = wdFindStop
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .Execute replace:=wdReplaceOne
  End With

  olSelection.HomeKey Unit:=wdStory
  
  Set olInspector = Nothing
  Set olDocument = Nothing
  Set olSelection = Nothing

End Sub

Private Sub Get_Attachments_List(ByRef oItem As MailItem, _
                                 ByRef strAtt As String, _
                                 ByRef AttCount As Integer)
  Dim HTMLImgSrc As String
  Dim oAtt As Attachment

  strAtt = ""
  AttCount = 0

  If oItem.Attachments.Count > 0 Then
    HTMLImgSrc = getHTMLImgSrc(oItem)
  
    AttCount = 0
    For Each oAtt In oItem.Attachments
      Select Case (oAtt.Type)
        Case olOLE
          'Avoiding the error: "Outlook cannot perform this action on this type of attachment"
          'OlAttachmentType enumeration, olOLE=6, attachment is an OLE document.
          'https://docs.microsoft.com/en-us/office/vba/api/outlook.olattachmenttype
          'This will effectively skip a file that the user may expect to be in attachmet list.
        Case Else
          If IsFileAttachment(oAtt, HTMLImgSrc) Then
            AttCount = AttCount + 1
            strAtt = strAtt & "  •  " & oAtt.FileName & vbCrLf
          End If
      End Select
    Next oAtt
    
  End If

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


Public Sub oItem_Reply(ByRef oItem As MailItem, Reply_All As Boolean)
  
  Dim oAtt As Attachment
  Dim Att_List As String
  Dim AttCount As Integer
      Dim olInspector As Outlook.Inspector
      Dim olDocument As Word.Document
      Dim olSelection As Word.Selection
  
  If bDiscardEvents Or oItem.Attachments.Count = 0 Then
         Exit Sub
     End If
     
     Cancel = True
     bDiscardEvents = True
  
  If oItem.BodyFormat = olFormatHTML Then
    Get_Attachments_List oItem, Att_List, AttCount
    If AttCount > 0 Then
      Att_List = "<attachment list>" & _
                    vbCrLf & vbCrLf & _
                    "Attachments sent " & _
                    Format(oItem.SentOn, "MM/dd/yy") & _
                    Format(oItem.SentOn, "hh:mm AM/PM") & _
                    vbCrLf & Att_List
     End If
   End If
   
    Dim oResponse As MailItem
    If Reply_All Then
        Set oResponse = oItem.ReplyAll
      Else
        Set oResponse = oItem.Reply
      End If
      
    oResponse.Display
      
  If oItem.BodyFormat = olFormatHTML And AttCount > 0 Then
      Find_and_Replace "<attachment list>", Att_List
  End If
     
  Set oItem = Nothing
      
End Sub


