<%
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open "DSN Name Removed to Help Protect my Databases on my Server","User Name Changed for Security Reasons","I am not telling you any passwords"
	lSQL="A query would appear here, but sorry, I do not want to give out table structures to the world."
	set rs=server.createobject("ADODB.RECORDSET")
	rs.open lSQL,conn,3,3
	if rs.eof=true then
		hitcnt=1
	else
		hitcnt=clng(rs.fields(0).Value)+1
	end if	
	if rs.eof=true then
		rs.addnew
	end if
	rs.fields(0)=hitcnt
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "[HITCOUNT]" & vbCrLf & "Count=" & hitCNT
	response.end
%>
<html>
<head>
</head>
<body>
</body>

</html>
