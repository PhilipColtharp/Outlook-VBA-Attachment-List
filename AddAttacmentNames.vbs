'original code: Diane Poremsky
'https://www.slipstick.com/developer/code-samples/add-attachments-names-to-message-before-sending/


Public Sub AddAttachmentNames(oItem As MailItem)

  Dim oAtt As Attachment
  Dim strAtt As String
    Dim olInspector As Outlook.Inspector
    Dim olDocument As Word.Document
    Dim olSelection As Word.Selection


  If oItem.Attachments.Count > 0 Then
  
    strAtt = "Message Attachments:" & vbCrLf
     
    For Each oAtt In oItem.Attachments
      Select Case (oAtt.Type)
        Case olOLE
          'Avoiding the error: "Outlook cannot perform this action on this type of attachment"
          'OlAttachmentType enumeration, olOLE=6, attachment is an OLE document.
          'https://docs.microsoft.com/en-us/office/vba/api/outlook.olattachmenttype
        Case Else
          strAtt = strAtt & "  •  " & oAtt.FileName & vbCrLf
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
 
Set oItem = Nothing
End Sub


