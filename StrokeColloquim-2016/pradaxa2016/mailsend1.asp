<%
sendUrl="http://schemas.microsoft.com/cdo/configuration/sendusing"
smtpUrl="http://schemas.microsoft.com/cdo/configuration/smtpserver"


' Set the mail server configuration
Set objConfig=CreateObject("CDO.Configuration")
objConfig.Fields.Item(sendUrl)=2 ' cdoSendUsingPort
objConfig.Fields.Item(smtpUrl)="relay-hosting.secureserver.net"
objConfig.Fields.Update


' Create and send the mail
Set objMail=CreateObject("CDO.Message")
' Use the config object created above
Set objMail.Configuration=objConfig
objMail.From="singh.guru@gmail.com"
objMail.ReplyTo="singh.guru@gmail.com"
objMail.To="singh.guru@gmail.com"
objMail.Subject="subject"
objMail.TextBody="body"
objMail.Send

response.write("email sent <br><br><br>")
set objMail=Nothing
set objConfig=Nothing
%>


<%
Set myMail = CreateObject("CDO.Message")
myMail.Subject = "Sending email with CDO"
myMail.From = "singh.guru@gmail.com"
myMail.To = "singh.guru@gmail.com"
myMail.TextBody = "This is a message."
myMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'Name or IP of remote SMTP server
myMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "relay-hosting.secureserver.net"
'Server port
myMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
myMail.Configuration.Fields.Update
myMail.Send
set myMail = nothing
%>