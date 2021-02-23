<!--#INCLUDE FILE="include/connection.asp"-->
<!--#INCLUDE FILE="Include/CommonFunctions.asp"-->
<%
opencon

%>


<%
IF Request.ServerVariables("REQUEST_METHOD")="POST" Then



sql = "insert into tblsurvyanswer (userid, Q1, Q2, Q3) values ("& GETQVAL("uid") &",'"&getval("q1")&"','"&getval("q2")&"','"&getval("q3")&"')"
con.execute sql
Response.redirect("thank-you-survey.html")

End If

%>

<%closecon%>