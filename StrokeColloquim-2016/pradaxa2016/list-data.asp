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
	<td>Email </td>
	<td>Phone no.</td>
	<td>Speciality</td>
	<td>City</td>
	<td>Live Or Web Cast</td>
	<td>Date of Registration</td>
</tr>
<%openrs
rs.open "select * from tblConfUsers order by userid desc",con,1,3
if not rs.eof Then

	While Not Rs.Eof
			
%>
<tr>
	<td><%=(rs("userId"))%></td>
	<td><%=(rs("Name"))%></td>
	<td><%=(rs("Email"))%> </td>
	<td><%=(rs("Phone"))%></td>
	<td><%=(rs("Speciality"))%></td>
	<td><%=(rs("City"))%></td>
	<td><% If rs("LiveOrWebCast") = "L" Then 
	response.write("Live")
	Else 
	
	response.write("Webcast")
	End If
	%></td>
	<td><%=(rs("regDate"))%></td>
</tr>

<%
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