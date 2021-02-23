<!--#INCLUDE FILE="ADOVBS.INC"-->
<%
response.write request.querystring("jjj")
Dim con,RS,com,JDServer,objFS,saUp,Mail
RootPath=Application("RootPath")
Sub OpenCon
	Set con=Server.CreateObject("ADODB.Connection")
	con.Open Session("ConnectStr")
'	con.execute "set dateformat dmy"
End Sub

Sub CloseCon
	if con.State =1 then
		con.Close 
	end if
	Set con=nothing
End Sub

SUB OpenCom(conn)
	Set com=Server.CreateObject("ADODB.Command")
	if IsObject(conn) then
		if conn.State = 1 then
			com.ActiveConnection = conn
		else 
			OpenCon
			com.ActiveConnection = con
		end if
	else 
		OpenCon
		com.ActiveConnection = con
	end if
End SUB

Sub CloseCom
	Set com=nothing
End Sub

SUB OpenRS
	Set RS=Server.CreateObject("ADODB.Recordset")
End SUB

Sub CloseRS
	Set RS=nothing
End Sub

SUB OpenRS1
	Set RS1=Server.CreateObject("ADODB.Recordset")
End SUB

Sub CloseRS1
	Set RS1=nothing
End Sub



Sub Check(str)
	Response.Write str
	Response.End 
End SUB

Sub println (str)
	Response.Write str &"<BR>"
End Sub

Sub printl (str)
	Response.Write str 
End Sub

Sub OpenObjFS
	Set objFS=Server.CreateObject("Scripting.FileSystemObject")
END SUB

SUB OpenSAFileUp
	Set saUp=Server.CreateObject("SoftArtisans.FileUp")
END SUB

SUB OpenCDONTS
	Set Mail=Server.CreateObject("CDONTS.NewMail")
END SUB
%><!--#INCLUDE FILE="CommonFunctions.asp"-->