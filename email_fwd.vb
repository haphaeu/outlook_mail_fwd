'
' Batch forward emails from defined mail boxes to an external email account.
'
' This is done as an Excel macro to avoid blocking Outlook. If the same
' macro is run from within Outlook, it seems to block execution and 
' Outboxs is flooded. It took several days to forward around 20 000 emails.
'
' Progress is saved to a file and also shown in excel sheet.
'
' Note that emails with attachments often create a pop up message saying
' "Are you sure you want to send potentially unsafe attachments?", which
' blocks the execution of the macro. To prevent that, the following key 
' has been added/modified in the registry:
'
'     HKCU\Software\Policies\Microsoft\Office\16.0\outlook\security
'     DontPromptLevel1AttachSend (REG_DWORD) = 0
'

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub doit()

    Dim objNS As Outlook.Namespace: Set objNS = GetNamespace("MAPI")
    Dim olFolder As Outlook.MAPIFolder
    Dim emailTo As String
    Dim fromDate As Date
    Dim toDate As Date
    
	' Emails received within this date range will be forwarded.
	' This is done to reduce load if sending years of emails.
	'
    ' Running through the weekend using
    '   fromDate = "01.01.2010 00:00:00"
    '   toDate = "01.07.2021 00:00:00"
    fromDate = "01.01.2010 00:00:00"
    toDate = "01.07.2021 00:00:00"
    
	' address to forward emails to
    emailTo = "user.name@emailserver.com"
    
	' email boxes to be forwarded
	' To find exactly which matches, some debugging is needed. To do that,
	' just initialise `objNS` and see the contents of the array `objNS.Folders`
    Dim mailboxNames As Variant
    mailboxNames = Array( _
                   Array("user.name@company.com", "Sent Items"), _
                   Array("user.name@company.com", "Inbox"), _
                   Array("Online Archive - user.name@oldco.com", "Sent Items"), _
                   Array("Online Archive - user.name@oldco.com", "Inbox"))
                   
    WriteProgress "Starting at " & Now()
    
    For Each pair In mailboxNames
        folder = pair(0)
        subfolder = pair(1)
        WriteProgress "Mailbox " & folder & " folder " & subfolder
        Cells(1, 1) = folder ' progress report in excel sheet
        Cells(2, 1) = subfolder ' idem
        Set olFolder = objNS.Folders(folder).Folders(subfolder)
        Call forward_all_mails(olFolder, emailTo, fromDate, toDate)
    Next
    
    WriteProgress "Done at " & Now()

End Sub


' Forward all emails in the folder `olFolder` received between 
' `fromDate` and `toDate` to `emailTo`.
Sub forward_all_mails(olFolder As Outlook.MAPIFolder, _
                      emailTo As String, _
                      fromDate As Date, _
                      toDate As Date)
 
    ' first, count how many emails to be sent
    numberEmails = 0
    For Each Item In olFolder.Items
        If TypeOf Item Is Outlook.MailItem Then
            If Item.To <> emailTo _
               And Item.ReceivedTime >= fromDate _
               And Item.ReceivedTime <= toDate Then
                    numberEmails = numberEmails + 1
                    Cells(3, 1) = numberEmails
            End If
        End If
    Next
    WriteProgress "found " & numberEmails & " emails."
    WriteProgress ""
    
	' set this to True for a dry run only counting emails.
    If False Then
        GoTo alldone
    End If
    
    i = 0
    For Each Item In olFolder.Items
        If TypeOf Item Is Outlook.MailItem Then
            If Item.To <> emailTo _
               And Item.ReceivedTime >= fromDate _
               And Item.ReceivedTime <= toDate Then
                    i = i + 1
                    Dim oMail As Outlook.MailItem: Set oMail = Item
					
					' some email are sent without permission to forward.
                    If oMail.Permission = olDoNotForward Then
                        GoTo for_continue
                    End If
                    
                    Set objMail = oMail.Forward
                    objMail.To = emailTo
					
					' progress report to excel and to file
                    WriteProgress "[" & i & "/" & numberEmails & "] Subject:" & oMail.Subject
                    Cells(4, 1) = i
					
                    'objMail.Display  ' this would display the email window
                    objMail.Send
                    Sleep (2000)  ' wait a bit not to flood the Outbox
for_continue:
            End If
        End If
    Next
    
alldone:
End Sub


' Write `text` string to file and flush contents
' Used for progress report line by line.
Function WriteProgress(text As String)
    
    Open "C:\Temp\progress.txt" For Append As #1
    Print #1, text
    Close #1
    
End Function
