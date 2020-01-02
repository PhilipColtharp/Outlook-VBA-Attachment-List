# Outlook-VBA-Attachment-List
Use a tag in your signatures to place a list of attachments when sending messages in Outlook.

Thanks to:
  * Diane Poremsky, Slipstick Systems
    - https://www.slipstick.com/developer/code-samples/add-attachments-names-to-message-before-sending/
  * Datanumen
    - https://www.datanumen.com/blogs/batch-find-replace-text-multiple-outlook-emails/

Instructions:
1. Add code the ```ThisOutlookSession```.
2. In Outlook options, check the "Open replies and forwards in a new window" checkbox.  See  [support.office.com - Article: Reply to or forward an email message](https://support.office.com/en-us/article/reply-to-or-forward-an-email-message-a843f8d3-01b0-48da-96f5-a71f70d0d7c8), FAQ: Can I have all replies and forwards open in a new window?
2. Intendend use: add the string "```<attachment list>```" to user's Outlook e-mail signatures
3. Alternate use: mannualy put the string "```<attachment list>```" in an e-mail
4. When e-mail is sent, the first occurance of the string  "```<attachment list>```" is replaced.  It is replaced by a list of e-mail attacments -or- it is replaced by nothing, if no attachments are present.

*Does not work with OLE objects
*Ignores file attacmentents that are in image elements in the message body.
