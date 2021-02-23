<%
if Session("User")="" then
	URL_ =Request.ServerVariables("SCRIPT_NAME")
	QS= Request.QueryString 
	Response.Redirect RootPath&"/Admin/default.asp?URL=" & Server.URLEncode(URL_ &"?"& QS)
End if
%>