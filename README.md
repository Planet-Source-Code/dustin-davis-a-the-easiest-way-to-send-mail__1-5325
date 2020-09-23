<div align="center">

## a The easiest way to send mail\!


</div>

### Description

I have seen some e-mail stuff on this site, but all are so freakin complicated. This is simple and VERY easy to use!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dustin Davis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dustin-davis.md)
**Level**          |Beginner
**User Rating**    |3.9 (55 globes from 14 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dustin-davis-a-the-easiest-way-to-send-mail__1-5325/archive/master.zip)





### Source Code

```
Dim sRes As String
Private Sub Command1_Click()
Winsock1.RemotePort = 25
Winsock1.RemoteHost = your_mail_server_here 'use your mail server
Winsock1.Connect
Do Until Winsock1.State = 7 '7=connected
  DoEvents
Loop
sRes = "0"
Winsock1.SendData "MAIL FROM: " & your_email_here & vbCrLf
Do Until sRes = "250"
  DoEvents
Loop
sRes = "0"
Winsock1.SendData "RCPT TO: " & someone_email_here & vbCrLf
Do Until sRes = "250"
  DoEvents
Loop
sRes = "0"
Winsock1.SendData "DATA" & vbCrLf
Do Until sRes = "354"
  DoEvents
Loop
Winsock1.SendData "FROM: " & your_name_here & vbCrLf
Winsock1.SendData "SUBJECT: " & subject_here & vbCrLf
Winsock1.SendData Text1.Text & vbCrLf & "." & vbCrLf
Do Until sRes = "250"
  DoEvents
Loop
Winsock1.Close
MsgBox "Mail sent!"
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
Dim Length As Long
Winsock1.GetData Data
Length = Len(Data)
sRes = Left$(Data, 3)
End Sub
```

