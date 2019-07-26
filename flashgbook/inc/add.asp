<!--#include file="conn.asp"-->
<% 
Session.CodePage="65001"
if request("action")="add" then
			sql="select * from gbook "
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,3,3
			rs.addnew		
			rs("name")=encodestr(request("w_name"))
			rs("title")=encodestr(request("w_title"))
			
				if encodestr(request("w_homepage"))="" then
				rs("homepage")=null
				else
				rs("homepage")=encodestr(request("w_homepage"))
				end if
				
				if encodestr(request("w_blog"))="" then
				rs("blog")=null
				else
				rs("blog")=encodestr(request("w_blog"))
				end if
				
				if encodestr(request("w_email"))="" then
				rs("email")=null
				else
				rs("email")=encodestr(request("w_email"))
				end if
				
				if encodestr(request("w_qq"))="" then
				rs("qq")=null
				else
				rs("qq")=encodestr(request("w_qq"))
				end if
				
			rs("content")=encodestr(request("w_content"))
			rs("date")=date()
			rs.update
			rs.close
			set rs=nothing
			response.write"&addok=ok"
end if
Session.CodePage="936"
%>
