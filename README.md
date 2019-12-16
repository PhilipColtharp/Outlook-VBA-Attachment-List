# Outlook-VBA-Attachment-List
Use a tag in your signatures to place a list of attachments when sending messages in Outlook.

Thanks to:
  * Diane Poremsky, Slipstick Systems
    - https://www.slipstick.com/developer/code-samples/add-attachments-names-to-message-before-sending/
  * Datanumen
    - https://www.datanumen.com/blogs/batch-find-replace-text-multiple-outlook-emails/

Instructions:
1. Add code the ThisOutlookSession 
2. Intendend use: add the string "```<attachment list>```" to users Outlook e-mail signatures
3. Alternate use: mannualy put the string "```<attachment list>```" in an e-mail
4. When e-mail is sent, the first occurance of the string  "```<attachment list>```" is replaced.  It is replaced by a list of e-mail attacments -or- it is replaced by nothing, if no attachments are present.

