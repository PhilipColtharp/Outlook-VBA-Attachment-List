'this subroutine should be added to your existing Application_ItemSend() function

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
  AddAttachmentNames Item
End Sub
