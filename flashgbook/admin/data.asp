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
<% sub shujuku()%>
<table width="760" border="1" cellpadding="1" cellspacing="1" bordercolor="66ccff">
  <tr>
    <td align="center">
        <table width="100%" border="0" cellspacing="2" cellpadding="2">
			<form action="?action=beifen" method="post" name="beifen" id="beifen">
          <tr align="center" bgcolor="#3399CC">
            <td height="25" colspan="2">���ݾݿ����</td>
          </tr>
          <tr>
            <td width="26%" height="25" align="right" bgcolor="#66CCFF" class="fonthei">���ݿ⣺</td>
            <td align="left" valign="top" bgcolor="#006699">&nbsp;·����<%= db %></td>
            </tr>
          <tr>
            <td width="26%" height="25" align="right" bgcolor="#66CCFF" class="fonthei">&nbsp;</td>
            <td align="left" valign="top" bgcolor="#006699">&nbsp;�ļ���С��<% if IsObjInstalled("Scripting.FileSystemObject") = False Then %>��֧��FSO<% Else %>
<%= GetFileSize(db) %><% End If %>
</td>
          </tr>
          <tr bgcolor="#3399CC">
            <td height="25" colspan="2" align="right" valign="top" class="fonthei">&nbsp;</td>
            </tr>
		  <% If bakdatar="����" Then %>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">��&nbsp;&nbsp;�ݣ�</td>
            <td align="left" valign="top" bgcolor="#006699"> 
              &nbsp;·����<%= bakdb %> </td>
            </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">&nbsp;</td>
            <td align="left" valign="top" bgcolor="#006699">&nbsp;�ļ���С��<%if IsObjInstalled("Scripting.FileSystemObject") = False Then %>��֧��FSO<% Else %><%= GetFileSize(bakdb) %><% End If %>
</td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">&nbsp;</td>
            <td align="left" valign="top" bgcolor="#006699">&nbsp;��󱸷�ʱ�䣺<%if IsObjInstalled("Scripting.FileSystemObject") = False Then %>��֧��FSO<% Else %><%= GetFileModified(bakdb) %><% End If %>
 </td>
          </tr>
		  <% Else %>
          <tr align="center">
            <td height="25" colspan="2" bgcolor="#66CCFF" class="fonthei">��û�б������ݿ⣡</td>
            </tr>
			<% End If %>
          <tr bgcolor="#3399CC">
            <td height="25" align="right" class="fonthei">&nbsp;</td>
            <td align="left"><%if IsObjInstalled("Scripting.FileSystemObject") = False Then %>��֧��FSO<% Else %><input name="Submit" type="submit" class="inputbt" value="�� ��"><% End If %>
</td>
          </tr>
		   </form>
        </table>
        
   </td>
  </tr>
</table>
<br>
<% If bakdatar="����" Then %>
<table width="760" border="1" cellpadding="1" cellspacing="1" bordercolor="66ccff">
  <tr>
    <td align="center">
        <table width="100%" border="0" cellspacing="2" cellpadding="2">
		<form action="?action=huifu" method="post" name="huifu" id="huifu">
          <tr align="center" bgcolor="#3399CC">
            <td height="25" colspan="2">��ԭ�ݿ����</td>
          </tr>
          <tr>
            <td width="26%" height="25" align="right" bgcolor="#66CCFF" class="fonthei">˵����</td>
            <td width="74%" align="left" bgcolor="#006699"><%if IsObjInstalled("Scripting.FileSystemObject") = False Then %>��֧��FSO<% Else %>��ѡ�����ã������ݿ���ʱ��ָ����ݣ����ݿ⽫�ָ�����<%= GetFileModified(bakdb) %> <% End If %>
</td>
          </tr>
          <tr bgcolor="#3399CC">
            <td height="25" align="right" class="fonthei">&nbsp;</td>
            <td align="left"><%if IsObjInstalled("Scripting.FileSystemObject") = False Then %>��֧��FSO<% Else %><input name="Submit" type="submit" class="inputbt" value="�� ԭ"  onClick="{if (confirm('ȷ����ԭ���⽫��ԭ����󱸷�ʱ��!')){return true;}return false;}"><% End If %>
</td>
          </tr>
		  </form>
        </table>
    </td>
  </tr>
</table>
<br>
<% end if %>  
<% end sub %>  
<!--#include file="bottom.asp" -->
</div>
</body>
</html>
