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



sql = "insert into tblConfUsers (Name,Email, Phone, Speciality, City, LiveOrWebCast) values ('"& getval("name") &"','"&getval("email")&"','"&getval("phone")&"','"&getval("speciality")&"','"&getval("live")&"', '"& getval("liveorweb") &"')"

con.execute sql

If  getval("liveorweb") = "L" then
	Response.redirect("live.html?msg=Thanks for the registration&city="&GETQVAL("city"))
Else
	Response.redirect("webcast.html?msg=Thanks for the registration&city="&GETQVAL("city"))

End If

End If

%>

<%closecon%>