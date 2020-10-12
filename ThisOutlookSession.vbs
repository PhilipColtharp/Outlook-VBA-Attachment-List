'based on code from Dian Pormskyhttps://www.slipstick.com/developer/code-samples/insert-attachment-names-replying/
Option Explicit

Private WithEvents oExpl As Explorer
Private WithEvents oItem As MailItem
Private bDiscardEvents As Boolean
Private Const Reply = False
Private Const ReplyAll = True

Private Sub Application_Startup()
   Set oExpl = Application.ActiveExplorer
   bDiscardEvents = False
End Sub

Private Sub oExpl_SelectionChange()
   On Error Resume Next
   Set oItem = oExpl.Selection.Item(1)
End Sub

' Reply
Private Sub oItem_Reply(ByVal Response As Object, Cancel As Boolean)

  If bDiscardEvents Or oItem.Attachments.Count = 0 Then
         Exit Sub
     End If

     Cancel = True
     bDiscardEvents = True

  AddAttachmentNames_.oItem_Reply oItem, Reply

  bDiscardEvents = False

End Sub

' Reply All
Private Sub oItem_ReplyAll(ByVal Response As Object, Cancel As Boolean)

  If bDiscardEvents Or oItem.Attachments.Count = 0 Then
         Exit Sub
     End If

     Cancel = True
     bDiscardEvents = True

  AddAttachmentNames_.oItem_Reply oItem, ReplyAll

  bDiscardEvents = False
End Sub

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
  AssertSignatureFormat Item
  AddAttachmentNames Item
End Sub


