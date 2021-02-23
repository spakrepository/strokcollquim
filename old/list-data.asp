<%
Server.ScriptTimeout = 500
%>

<!--#INCLUDE FILE="include/connection.asp"-->
<!--#INCLUDE FILE="Include/CommonFunctions.asp"-->
<%
opencon

'	if not rs.eof Then
'		response.write(rs.recordCount & "---- total count<br><br>") 
'	
'	End If

%>


<table border="1" cellpadding="0" cellspacing="0">
<tr>
	<td>UserID</td>
	<td>Name</td>
	<td>City </td>
	<td>Email</td>
	<td>Phone no.</td>
	<td>Confirm My Attendance</td>
	<td>Live Or Web Cast</td>
	<td>Web Cast</td>
	<td>Live</td>
	<td>Date of Registration</td>
</tr>
<%openrs
rs.open "select  * from tblconfusersnew order by userid desc",con,1,3
if not rs.eof Then

	While Not Rs.Eof
			
%>
<tr>
	<td><%=(rs("userId"))%></td>
	<td><%=(rs("Name"))%></td>
	<td><%=(rs("city"))%> </td>
		<td><%=(rs("email"))%> </td>

	<td><%=(rs("Phone"))%></td>
	<td>
		<%
			If rs("confirmAttendance")  = "A" Then
				response.write("Attend")
			Else
				response.write(rs("confirmAttendance"))
			End If
		%>


	</td>
	<td><%=(rs("webcastOrLive"))%></td>
	<td><%=(rs("Webcast"))%></td>
	<td><%=(rs("live"))%></td>
	<td><%=(rs("Createddate"))%></td>
</tr>

<%
response.flush()
			rs.moveNext()
		Wend

Else %>

<tr>
	<td colspan="8">No record found</td>
</tr>

<%
End If

closers%>


</table>
<%closecon%>