<!--#include file="admin_conn.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<% 
	Session.TimeOut=30
if request("action")="login" then
		admin=trim(request.form("name"))
		for i=1 to len(admin)     '��MID������������admin��i λ�õ�һ���ַ�
			manage=mid(admin,i,1)
			if manage="'" or manage="%" or manage="<" or manage=">" or manage="&" then    '���admin�к���' % < > &�ַ���ת������ҳ��
				response.redirect "Error.asp"
				response.end
			end if
		next
		pwd=trim(request.form("pwd"))
		for i=1 to len(pwd)     '��MID������������pwd��i λ�õ�һ���ַ�
			pass=mid(pwd,i,1)
			if pass="'" or pass="%" or pass="<" or pass=">" or pass="&" then    '���pass�к���' % < > &�ַ���ת������ҳ��
				response.redirect "Error.asp"
				response.end
			end if
		next 
		if admin="" or pwd="" then
			Response.Redirect ("admin.asp")
		end if		
		set rs=server.createobject("adodb.recordset")
		sql="select * from admin where adminname='"&admin&"'and adminpwd='"&pwd&"'"
		rs.open sql,conn,1,1
		if not rs.eof then
			session("admin_name")=admin
			response.redirect"main.asp?action=info"
		else
			response.redirect"error.asp"		
		end if
end if 
if request("action")="logout" then
		session("admin_name")=""
		response.write "<script language=javascript>alert('�˳��ɹ���');location.href('admin.asp');</script>"
end if
%>