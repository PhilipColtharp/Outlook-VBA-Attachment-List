'******change module "ThisOutlookSession" and add reference to function *******
'
'Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
'
'  'other code...
'
'  AssertSignatureFormat Item
'
'  'other code...
'
'End Sub


Public Sub AssertSignatureFormat(oItem As MailItem)
    ' Thanks to ilcaa72 and msofficeforums.com - https://www.msofficeforums.com/word-vba/19714-loop-through-each-line-word.html
    ' Thanks to GuruKay and stackoverflow.com - https://stackoverflow.com/a/43166192/6524470
  
  Dim olInspector As Outlook.Inspector
  Dim olDocument As Word.Document
  Dim wdParagraph As Paragraph
  Dim regEx As Variant
  Dim Text As String
  
  Set olInspector = Application.ActiveInspector()
  Set olDocument = olInspector.WordEditor
     
  Set regEx = CreateObject("vbscript.regexp") 'Initialize the regex object, see AssertSignatureFormat
    
    ' uses reference library "Microsoft VBScript Regular Expressions 5.5"
        Dim strPattern As String: strPattern = "[^a-zA-Z0-9]" 'The regex pattern to find special characters
        Dim strReplace As String: strReplace = "" 'The replacement for the special characters
        'uses regEx already created to be equal to CreateObject("vbscript.regexp") 'Initialize the regex object
    ' Configure the regex object
          With regEx
              .Global = True
              .MultiLine = True
              .IgnoreCase = False
              .Pattern = "[^a-zA-Z0-9]"
          End With
    ' Perform the regex replacement. Example: strang = regEx.Replace(strang , strReplace)
    MsgBoxReturns = vbOK
    LineCount = 0
    For Each wdParagraph In olDocument.Paragraphs
      Text = regEx.replace(wdParagraph.Range.Text, strReplace)
      'If MsgBoxReturns <> vbCancel _
         Then MsgBoxReturns = MsgBox("""" + Text + """", vbOKCancel)
        
      With wdParagraph.Range.Font
        Select Case (True)
          Case bsc(Text, "PhilipColtharp")
            'Signature_Name Style
            .Name = "Calibri"
            .Bold = True
            .Italic = False
            .Size = 12
            LineCount = 1
          End Select
          If LineCount > 0 Then
            Select Case (True)
              Case (bsc(Text, "GovernmentOperationsConsultantII") Or _
                    bsc(Text, "Office8507173646") Or _
                    bsc(Text, "Fax8504881967") _
                   ) And LineCount
                'Sig_Contanct Style
                .Name = "Calibri"
                .Bold = False
                .Italic = False
                .Size = 11
              Case (bsc(Text, "InspiringSuccessbyTransformingOneLifeataTime") _
                   ) And LineCount
                'Sig_Contanct Style
                .Name = "Arial"
                .Bold = True
                .Italic = True
                .Size = 11
              Case (bsc(Text, "CONFIDENTIALITYPUBLICRECORDSNOTICE:") _
                   ) And LineCount
                'Sig_Note Style
                .Name = "Calibri"
                .Bold = True
                .Italic = False
                .Size = 9
              Case Else
            End Select
          End If
      End With
      Select Case (LineCount)
        Case Is > 15
          LineCount = 0
        Case Is > 0
          LineCount = LineCount + 1
        Case Else
      End Select
    Next
End Sub

Private Function bsc(leeft As String, reight As String)  'Bountded String Compare
  l = min(Len(leeft), Len(reight))
  bsc = (Left(leeft, l) = Left(reight, l))
End Function

Private Function min(x, y As Variant) As Variant
   min = IIf(x < y, x, y)
End Function

