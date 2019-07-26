<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<% 
	sql="select * from gbook_set where id in(1)"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	gbookname=uni(rs("gbookname"))
	homehave=uni(rs("homehave"))
	homeurlu=uni(rs("homeurlu"))
	homeurlk=uni(rs("homeurlk"))
	homename=uni(rs("homename"))
	out=out&"<url gbookname='"&gbookname&"' gbookurl='"&gbookurl&"' homehave='"&homehave&"' homeurlu='"&homeurlu&"' homeurlk='"&homeurlk&"' homename='"&homename&"'></url>"	
	rs.close
	set rs=nothing
	Response.Write "<?xml version='1.0' encoding='utf-8'?>"
	Response.Write ""&out&""
%>