# Outlook email forward

## Batch forward emails in Outlook 

### Scenario is

You're changing jobs and due to company policy you'd lose 
access to years of emails in Outlook. Exporting PST is not 
allowed/disabled/encrypted.

### Solution

Create an email account for this purpose. For example yahoo 
offers 1 TB free storage.

Then forwad all your emails to yourself using this macro.

## How to use it

This macro is meant as an Excel macro. So just paste in Excel, edit
and run.

You'll need to change:

 - `emailTo`: email address to forward the emails to, 
 
 - `mailboxNames`: mailbox(es) folder(s) and subfolder(s) you want 
   forwarded. To do that you'll have to debug a little and find out
   the actual names of the mailboxes you have. This is found in the 
   arrays and sub-arrays `Folders` of the "MAPI" namespace:
   
``` VB
    Dim objNS As Outlook.Namespace: Set objNS = GetNamespace("MAPI")
    ' Debug/break here and watch contents of objNS
```

 - `fromDate`, `toDate`: emails received within this range is date will
   be forwarded. The format is "31.12.2019 19:30:00"

 - File path to save progress report to. See function `WriteProgress`

## Potentially unsafe attachments

**Forget about the below, as it doesn't work**

Note that emails with attachments often create a pop up message saying
*"Are you sure you want to send potentially unsafe attachments?"*, which
blocks the execution of the macro. 

To prevent that, the following key has been added/modified in the 
registry:

```
     HKCU\Software\Policies\Microsoft\Office\16.0\outlook\security
     DontPromptLevel1AttachSend (REG_DWORD) = 0
```

