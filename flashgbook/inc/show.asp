<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<% 
if request("action")="show" then
	sql="select * from gbook where id in("&request("show_id")&")"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	show_name=uni(rs("name"))
	show_id=uni(rs("id"))
	show_title=uni(rs("title"))
	
	if rs("blog")<>"" then
	show_blog=uni(rs("blog"))
	end if
	if rs("homepage")<>"" then
	show_homepage=uni(rs("homepage"))
	end if
	if rs("gmcontent")<>"" then
	show_gmcontent=uni(rs("gmcontent"))
	end if
	if rs("gmdate")<>"" then
	show_gmdate=uni(rs("gmdate"))
	end if
	
	if rs("email")<>"" then
	show_email=uni(rs("email"))
	end if
	if rs("qq")<>"" then
	show_qq=uni(rs("qq"))
	end if
	show_content=uni(rs("content"))
	show_date=uni(rs("date"))
	out=out&"<info show_name='"&show_name&"' show_blog='"&show_blog&"'  show_homepage='"&show_homepage&"'  show_gmcontent='"&show_gmcontent&"'  show_gmdate='"&show_gmdate&"' show_id='"&show_id&"' show_title='"&show_title&"' show_email='"&show_email&"' show_qq='"&show_qq&"' show_content='"&show_content&"' show_date='"&show_date&"' />"	
	rs.close
	set rs=nothing
	Response.Write "<?xml version='1.0' encoding='utf-8'?>"
	Response.Write "<gbook>"&out&"</gbook>"
end if
%>