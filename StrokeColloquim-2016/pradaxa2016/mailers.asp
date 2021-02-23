
<!--#INCLUDE FILE="include/connection.asp"-->
<!--#INCLUDE FILE="Include/CommonFunctions.asp"-->
<%
opencon

IF Request.ServerVariables("REQUEST_METHOD")="POST" Then

sql = "select top "& getval("cmbsendto") &" u.userid, Name, email, convert(nvarchar(26), u.regDate, 106) as RegDate from tblConfUsers u "
'sql = "select top 1 u.userid, Name, email, convert(nvarchar(26), u.regDate, 106) as RegDate from tblConfUsers u "
sql = sql & " left outer join tblMailSent m on m.userId = U.UserID "
sql = sql & " where m.mailsentOn is null "
'sql = sql & " where u.userid = 1"

mailbody = getval("mailbody")
'mailbody = Replace(mailbody, vbCrLf, "<br>")

response.write (sql& "<br><br>")

openrs
rs.open sql,con,1,3
if not rs.eof Then
	While Not Rs.Eof
		mailbody = getval("mailbody")
		'response.write(Replace(mailbody, "{name}", rs("Name"))& "<br><br>")
		mailbody = Replace(mailbody, "{name}", rs("Name"))
		mailbody = Replace(mailbody, "{uid}", rs("userid"))
		mailbody = Replace(mailbody, "{date}", rs("RegDate"))

		If mailbody <> "" Then
			Set myMail = CreateObject("CDO.Message")
			myMail.Subject = getval("txtsubject")
			myMail.From = getval("txtFrom")
			myMail.To = rs("email")
			myMail.HTMLBody = mailbody
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
			sql = "insert into tblMailSent (userId) values ("& rs("userid") &")"
			response.write(sql)
			con.execute sql

		End If
		response.write(mailbody & "<br><br>")
		rs.moveNext()
	Wend
%>




<%
End If
closers
End If

closecon%>