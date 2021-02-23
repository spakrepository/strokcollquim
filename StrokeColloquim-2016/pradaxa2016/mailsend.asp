<%
	Dim mbody 
Dim strSubject


If Request("pg")="ask" then

	mbody = "DATE OF Ask Question:"& Date() & " : TIME : " & time() & "<br><br>"
	mbody = mbody & "Name : " & request("name") & "<br><br>"
	mbody = mbody & "EMAIL : " & request("email") & "<br><br>"
	mbody = mbody & "Specialty : " & request("area") & "<br><br>"
	mbody = mbody & "Question: "  & request("que") & "<br><br>"

strSubject ="Ask a question"

ElseIf Request("pg")="con" then

	mbody = "DATE OF Contact Us:"& Date() & " : TIME : " & time() & "<br><br>"
	mbody = mbody & "Name : " & request("name") & "<br><br>"
	mbody = mbody & "Location : " & request("loc") & "<br><br>"
	mbody = mbody & "Feedback / Query: "  & request("comment") & "<br><br>"

strSubject ="Contact Us"



ElseIf Request("pg")="reg" then

	mbody = "DATE OF Registration:"& Date() & " : TIME : " & time() & "<br><br>"
	mbody = mbody & "Name : " & request("name") & "<br><br>"
	mbody = mbody & "City : " & request("city") & "<br><br>"

	mbody = mbody & "Email Id : " & request("email") & "<br><br>"
	mbody = mbody & "Mobile No : " & request("mobile") & "<br><br>"

	mbody = mbody & "Confirm my attendance : " & request("attendance") & "<br><br>"
	mbody = mbody & "Live Or Webcast : " & request("via") & "<br><br>"

	If request("via") = "Webcast" Then
		mbody = mbody & "Webcast : " & request("webcast") & "<br><br>"
	Else 
		mbody = mbody & "Live : " & request("live") & "<br><br>"
	End If
	
	

strSubject ="New Registration"

End If


%>
<%
Set myMail = CreateObject("CDO.Message")
myMail.Subject = strSubject
myMail.From = "shiv@spakcomm.com"
myMail.To = "bi.ask@spakcomm.com"
myMail.Bcc = "shivstuff@gmail.com"
myMail.HTMLBody = mbody
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
set myMail = Nothing


response.redirect("thankyou.html?pg=" & Request("pg") &"&visit="&request("attendance"))
%>