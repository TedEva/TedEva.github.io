<!--#include file="admin_conn.asp"-->
<% 
	if session("admin_name")="" then
		response.redirect"admin.asp"
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>STONE�ռ� flash���Ա�  ��̨����</title>
<meta name="keywords" content="STONE,STONE�ռ� flash���Ա�">
<meta http-equiv="STONE�ռ� flash���Ա�" content="http://stone-stone.vip.sina.com">
<meta name="description" content="�������� STONE ��Ʊ�д�������������뵽 STONE�ռ� flash���Ա� �����лл��">
<link href="css.css" rel="stylesheet" type="text/css">
</head>

<body>
<div align="center">
<!--#include file="top.asp" -->
<br>
<% 
dim bakfolder,bakname
'�������ݿ���ļ���
bakfolder="../data/bak"
'�������ݿ���ļ�
bakname="bak.asp"
bakdb=bakfolder&"/"&bakname
set fileobj=server.createobject("scripting.filesystemobject")
	if fileobj.FileExists(server.mappath(bakdb)) then
	bakdatar="����"
	end if
set fileobj=nothing
Function GetFileSize(FileName)
	set fso=server.createobject("scripting.filesystemobject")
	drvpath=server.mappath(FileName)
	set d=fso.getfile(drvpath)	
	size=d.size
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   showsize=round(size,2) & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=round(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=round(size,2) & "&nbsp;GB"	   
	end if   
	set fso=nothing
	GetFileSize = showsize
End Function
' ���������Ƿ�֧��ĳһ����
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
'�鿴�ļ��޸�ʱ��
Function GetFileModified(FileName)
	set fso=server.createobject("scripting.filesystemobject")
	set d=fso.getfile(server.mappath(FileName))	
	set fso=nothing
	GetFileModified = d.datelastmodified
End Function
%>
<% 
select case request("action")
	case "gopage"
		call manage()

	case "info"
		call info()
		
	case "canshu"
		call canshu()
		
	case "gmhuifu"
		call gmhuifu()
		
	case "canshu_edit"
		call canshu_edit()
		
	case "manage"
		call manage()
		
	case "edit_pwd"
		call edit_pwd()
		
	case "edit_pwd_save"
		call edit_pwd_save()
		
	case "view"
		call view()
		
	case "del"
		call del()
		
	case "shujuku"
		call shujuku()
		
	case "beifen"
		call beifen()
		
	case "yasuo"
		call yasuo()
		
	case "huifu"
		call huifu()
end select
sub canshu_edit()
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from gbook_set where id in(1)"
	rs.open sql,conn,3,2
	rs("gbookname")=trim(request.form("gbookname"))
	rs("homehave")=trim(request.form("homehave"))
	rs("homeurlu")=trim(request.form("homeurlu"))
	rs("homeurlk")=trim(request.form("homeurlk"))
	rs("homename")=trim(request.form("homename"))
	rs.update		
	rs.close
	set rs=nothing
	response.write"<script language='javascript'>alert('���óɹ���');location.href('?action=canshu');</script>"
end sub
sub gmhuifu()
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from gbook where id in("&request("id")&")"
	rs.open sql,conn,3,2
	rs("name")=trim(request.form("name"))
	
	if trim(request.form("blog"))="" then
	rs("blog")=null
	else
	rs("blog")=trim(request.form("blog"))
	end if
	
	if trim(request.form("homepage"))="" then
	rs("homepage")=null
	else
	rs("homepage")=trim(request.form("homepage"))
	end if
	
	if trim(request.form("qq"))="" then
	rs("qq")=null
	else
	rs("qq")=trim(request.form("qq"))
	end if
	
	if trim(request.form("email"))="" then
	rs("email")=null
	else
	rs("email")=trim(request.form("email"))
	end if
	
	rs("title")=trim(request.form("title"))
	rs("content")=trim(request.form("content"))
	
	if trim(request.form("gmcontent"))<>"" then
	rs("gmcontent")=trim(request.form("gmcontent"))
	rs("gmdate")=date()
	else
	rs("gmcontent")=null
	rs("gmdate")=null
	end if
	
	rs.update		
	rs.close
	set rs=nothing
	response.write"<script language='javascript'>alert('�༭�ɹ���');location.href('?action=manage');</script>"
end sub
sub edit_pwd_save()
	  admin=trim(request.form("name"))
      pwd=trim(request.form("pwd"))
	  pwd2=trim(request.form("pwd2"))	  
	  if admin="" or pwd="" then
	  	 response.write"<script language='javascript'>alert('�û��������벻��Ϊ�գ�');location.href('?action=edit_pwd');</script>"
	  end if
	  if pwd2<>pwd then
	  	 response.write"<script language='javascript'>alert('�������벻һ�£�');location.href('?action=edit_pwd');</script>"
	  end if
      set rs=server.createobject("adodb.recordset")
	  sql="select * from admin where id=1"
	  rs.open sql,conn,3,2
	  rs("adminname")=admin
	  rs("adminpwd")=pwd
	  rs.update
	  rs.close
	  set rs=nothing
	  response.write"<script language='javascript'>alert('�޸ĳɹ���');location.href('?action=info');</script>"
end sub
sub del()
		set rs=server.CreateObject("adodb.recordset")
		sql="select * from gbook where id="&request("id")&""
		rs.open sql,conn,3,2
		rs.delete		
		rs.close
		set rs=nothing
        response.write"<script language='javascript'>alert('ɾ���ɹ���');location.href('?action=manage');</script>"
end sub
'�������ݿ�
sub beifen()
	  Set Fso=server.createobject("scripting.filesystemobject")
	  If Fso.fileexists(server.mappath(db)) then
			Fso.copyfile server.mappath(db),server.mappath(bakdb)
	  end if
	  set Fso=nothing
	  	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.CopyFile server.mappath(bakdb),server.mappath(bakfolder) & "temp.mdb"
	Set Engine = CreateObject("JRO.JetEngine")
	Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.mappath(bakfolder) & "temp.mdb", _
  	"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.mappath(bakfolder) & "temp1.mdb"
	fso.CopyFile server.mappath(bakfolder) & "temp1.mdb",server.mappath(bakdb)
	fso.DeleteFile(server.mappath(bakfolder) & "temp.mdb")
	fso.DeleteFile(server.mappath(bakfolder) & "temp1.mdb")
	Set fso = nothing
	Set Engine = nothing

      response.write "<script language=javascript>alert('�������ݿ�ɹ���');location.href('?action=shujuku');</script>"
end sub
'ѹ�����ݿ�
sub yasuo()
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.CopyFile server.mappath(bakdb),server.mappath(bakfolder) & "temp.mdb"
	Set Engine = CreateObject("JRO.JetEngine")
	Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.mappath(bakfolder) & "temp.mdb", _
  	"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.mappath(bakfolder) & "temp1.mdb"
	fso.CopyFile server.mappath(bakfolder) & "temp1.mdb",server.mappath(bakdb)
	fso.DeleteFile(server.mappath(bakfolder) & "temp.mdb")
	fso.DeleteFile(server.mappath(bakfolder) & "temp1.mdb")
	Set fso = nothing
	Set Engine = nothing
    response.write "<script language=javascript>alert('ѹ�����ݿ�ɹ���');location.href('?action=shujuku');</script>"
end sub
'�ָ����ݿ�
sub huifu()
	  Set Fso=server.createobject("scripting.filesystemobject")
	  If Fso.fileexists(server.mappath(bakdb)) then
			Fso.copyfile server.mappath(bakdb),server.mappath(db)
	  end if
	  set Fso=nothing
      response.write "<script language=javascript>alert('�ָ����ݿ�ɹ���');location.href('?action=shujuku');</script>"
end sub
 %>
<table width="760" border="1" cellpadding="1" cellspacing="1" bordercolor="66ccff">
  <tr>
    <td><table width="100%" border="0" cellpadding=2 cellspacing=2 class="k1" style="border-collapse: collapse">
        <tr align="center" bgcolor="#3399CC" class="fontmenu2">
          <td colspan="2" bgcolor="#006699">�����ʹ���˴����Ա����뵽�ҵĲ����������ĵ�ַ���������º���֪ͨ����<br>
            <a href="http://www.stonemx.com/blog/archives/2006/flashgbook.html" target="_blank">http://www.stonemx.com/blog/archives/2006/flashgbook.html<br>
            </a>�����Ա��汾:2.0(2006-10-9)
            <table border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td>���°汾�� </td>
                  <td><iframe src="http://www.stonemx.com/design/flashgbook/bb.html" width="100" marginwidth="0" height="14" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
                </tr>
                        </table></td>
        </tr>
        
        <tr align="center" bgcolor="#3399CC" class="fontmenu2">
          <td height=12 colspan="2">��������������</td>
        </tr>
        <tr bgcolor="#eeeeee" class="fontmenu2">
          <td width="18%" height=25 bgcolor="#66CCFF" class="fonthei">&nbsp;����������</td>
          <td width='82%' bgcolor="#006699">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
        </tr>
        <tr bgcolor="#eeeeee" class="fontmenu2">
          <td height=25 bgcolor="#66CCFF" class="fonthei">&nbsp;������IP��</td>
          <td bgcolor="#006699">&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
        </tr>
        <tr bgcolor="#eeeeee" class="fontmenu2">
          <td height=25 bgcolor="#66CCFF" class="fonthei">&nbsp;�������˿ڣ�</td>
          <td bgcolor="#006699">&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
        </tr>
        <tr bgcolor="#eeeeee" class="fontmenu2">
          <td height=25 bgcolor="#66CCFF" class="fonthei">&nbsp;������ʱ�䣺</td>
          <td bgcolor="#006699">&nbsp;<%=now%></td>
        </tr>
        <tr bgcolor="#eeeeee" class="fontmenu2">
          <td height=25 bgcolor="#66CCFF" class="fonthei">&nbsp;IIS�汾��</td>
          <td bgcolor="#006699">&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
        </tr>
        <tr bgcolor="#eeeeee" class="fontmenu2">
          <td height=25 bgcolor="#66CCFF" class="fonthei">&nbsp;����������ϵͳ��</td>
          <td bgcolor="#006699">&nbsp;<%=Request.ServerVariables("OS")%></td>
        </tr>
        <tr bgcolor="#eeeeee" class="fontmenu2">
          <td height=25 bgcolor="#66CCFF" class="fonthei">&nbsp;�ű���ʱʱ�䣺</td>
          <td bgcolor="#006699">&nbsp;<%=Server.ScriptTimeout%> ��</td>
        </tr>
        <tr bgcolor="#eeeeee" class="fontmenu2">
          <td height=25 bgcolor="#66CCFF" class="fonthei">&nbsp;վ������·����</td>
          <td bgcolor="#006699">&nbsp;<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
        </tr>
        <tr bgcolor="#eeeeee" class="fontmenu2">
          <td height=25 bgcolor="#66CCFF" class="fonthei">&nbsp;������CPU������</td>
          <td bgcolor="#006699">&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</td>
        </tr>
        <tr bgcolor="#eeeeee" class="fontmenu2">
          <td height=25 bgcolor="#66CCFF" class="fonthei">&nbsp;�������������棺</td>
          <td bgcolor="#006699">&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
        </tr>
        <tr bgcolor="#eeeeee" class="fontmenu2">
          <td height=25 bgcolor="#66CCFF" class="fonthei">&nbsp;���ļ�·����</td>
          <td bgcolor="#006699"><%bwjlj=Request.ServerVariables("PATH_TRANSLATED")%>            <%=  Replace(bwjlj, "\", "/") %></td>
        </tr>
        <tr bgcolor="#eeeeee" class="fontmenu2">
          <td height=25 bgcolor="#66CCFF" class="fonthei">&nbsp;�ļ��д�С��</td>
          <td bgcolor="#006699">
<%If IsObjInstalled("Scripting.FileSystemObject") = False Then%>
		�˹���Ҫ�������֧���ļ�ϵͳ����FSO�������㵱ǰ�ķ�������֧�֣�
		
<%else
	
Set fso1 = CreateObject("Scripting.FileSystemObject")  
Set ff = fso1.GetFolder(server.MapPath("../")) 
%>
<%=  Replace(ff, "\", "/") %> &nbsp;&nbsp;��Ŀ¼�¹�&nbsp;&nbsp;
<% 
	size=ff.size
	showsize=round(size,2) & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   showsize=round(size,2) & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=Round(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=Round(size,2) & "&nbsp;GB"	   
	end if   
	set fso1=nothing
 %>
            <%=showsize%>
			<% End If %></td>
        </tr>
        
      </table></td>
  </tr>
</table>
<br>
<!--#include file="bottom.asp" -->
</div>
</body>
</html>
