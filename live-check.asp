<!--#INCLUDE FILE="include/connection.asp"-->
<!--#INCLUDE FILE="Include/CommonFunctions.asp"-->
<%
opencon
'openrs
'rs.open "select * from tbladmin",con,1,3
'if not rs.eof Then
'response.write(rs.recordCount & "---- total count<br><br>") 
'While Not Rs.Eof
'response.write(rs("username")& "--- admin id<br><br>")
'rs.moveNext()
'Wend
'End If
'closers
%>


<%
IF Request.ServerVariables("REQUEST_METHOD")="POST" Then



sql = "insert into tblConfUsersnew (name,city, email, phone, confirmAttendance, webcastOrLive, Webcast, Live) values ('"& getval("name") &"','"&getval("city")&"', '"&getval("email")&"','"&getval("phone")&"','"&getval("attendance")&"','"&getval("via")&"','"&getval("webcast")&"', '"& getval("live") &"')"

con.execute sql

If  getval("liveorweb") = "L" then
	Response.redirect("thankyou.html?msg=Thanks for the registration&city="&GETQVAL("city"))
Else
	Response.redirect("thankyou.html?msg=Thanks for the registration&city="&GETQVAL("city"))

End If

End If

%>

<%closecon%>