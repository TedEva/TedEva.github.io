<!--#include file="admin_conn.asp"-->
<% 
	if session("admin_name")="" then
		response.redirect"admin.asp"
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>STONE空间 flash留言本  后台管理</title>
<meta name="keywords" content="STONE,STONE空间 flash留言本">
<meta http-equiv="STONE空间 flash留言本" content="http://stone-stone.vip.sina.com">
<meta name="description" content="本程序由 STONE 设计编写！程序有问题请到 STONE空间 flash留言本 提出，谢谢！">
<link href="css.css" rel="stylesheet" type="text/css">
</head>

<body>
<div align="center">
<!--#include file="top.asp" -->
<br>
<% 
dim bakfolder,bakname
'备份数据库的文件夹
bakfolder="../data/bak"
'备份数据库的文件
bakname="bak.asp"
bakdb=bakfolder&"/"&bakname
set fileobj=server.createobject("scripting.filesystemobject")
	if fileobj.FileExists(server.mappath(bakdb)) then
	bakdatar="存在"
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
' 检测服务器是否支持某一对象
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
'查看文件修改时间
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
'备份数据库
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

      response.write "<script language=javascript>alert('备份数据库成功！');location.href('?action=shujuku');</script>"
end sub
'恢复数据库
sub huifu()
	  Set Fso=server.createobject("scripting.filesystemobject")
	  If Fso.fileexists(server.mappath(bakdb)) then
			Fso.copyfile server.mappath(bakdb),server.mappath(db)
	  end if
	  set Fso=nothing
      response.write "<script language=javascript>alert('恢复数据库成功！');location.href('?action=shujuku');</script>"
end sub
 %>
<% sub shujuku()%>
<table width="760" border="1" cellpadding="1" cellspacing="1" bordercolor="66ccff">
  <tr>
    <td align="center">
        <table width="100%" border="0" cellspacing="2" cellpadding="2">
			<form action="?action=beifen" method="post" name="beifen" id="beifen">
          <tr align="center" bgcolor="#3399CC">
            <td height="25" colspan="2">备份据库管理</td>
          </tr>
          <tr>
            <td width="26%" height="25" align="right" bgcolor="#66CCFF" class="fonthei">数据库：</td>
            <td align="left" valign="top" bgcolor="#006699">&nbsp;路径：<%= db %></td>
            </tr>
          <tr>
            <td width="26%" height="25" align="right" bgcolor="#66CCFF" class="fonthei">&nbsp;</td>
            <td align="left" valign="top" bgcolor="#006699">&nbsp;文件大小：<% if IsObjInstalled("Scripting.FileSystemObject") = False Then %>不支持FSO<% Else %>
<%= GetFileSize(db) %><% End If %>
</td>
          </tr>
          <tr bgcolor="#3399CC">
            <td height="25" colspan="2" align="right" valign="top" class="fonthei">&nbsp;</td>
            </tr>
		  <% If bakdatar="存在" Then %>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">备&nbsp;&nbsp;份：</td>
            <td align="left" valign="top" bgcolor="#006699"> 
              &nbsp;路径：<%= bakdb %> </td>
            </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">&nbsp;</td>
            <td align="left" valign="top" bgcolor="#006699">&nbsp;文件大小：<%if IsObjInstalled("Scripting.FileSystemObject") = False Then %>不支持FSO<% Else %><%= GetFileSize(bakdb) %><% End If %>
</td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">&nbsp;</td>
            <td align="left" valign="top" bgcolor="#006699">&nbsp;最后备份时间：<%if IsObjInstalled("Scripting.FileSystemObject") = False Then %>不支持FSO<% Else %><%= GetFileModified(bakdb) %><% End If %>
 </td>
          </tr>
		  <% Else %>
          <tr align="center">
            <td height="25" colspan="2" bgcolor="#66CCFF" class="fonthei">还没有备份数据库！</td>
            </tr>
			<% End If %>
          <tr bgcolor="#3399CC">
            <td height="25" align="right" class="fonthei">&nbsp;</td>
            <td align="left"><%if IsObjInstalled("Scripting.FileSystemObject") = False Then %>不支持FSO<% Else %><input name="Submit" type="submit" class="inputbt" value="备 份"><% End If %>
</td>
          </tr>
		   </form>
        </table>
        
   </td>
  </tr>
</table>
<br>
<% If bakdatar="存在" Then %>
<table width="760" border="1" cellpadding="1" cellspacing="1" bordercolor="66ccff">
  <tr>
    <td align="center">
        <table width="100%" border="0" cellspacing="2" cellpadding="2">
		<form action="?action=huifu" method="post" name="huifu" id="huifu">
          <tr align="center" bgcolor="#3399CC">
            <td height="25" colspan="2">还原据库管理</td>
          </tr>
          <tr>
            <td width="26%" height="25" align="right" bgcolor="#66CCFF" class="fonthei">说明：</td>
            <td width="74%" align="left" bgcolor="#006699"><%if IsObjInstalled("Scripting.FileSystemObject") = False Then %>不支持FSO<% Else %>此选项慎用！当数据库损坏时请恢复数据！数据库将恢复到：<%= GetFileModified(bakdb) %> <% End If %>
</td>
          </tr>
          <tr bgcolor="#3399CC">
            <td height="25" align="right" class="fonthei">&nbsp;</td>
            <td align="left"><%if IsObjInstalled("Scripting.FileSystemObject") = False Then %>不支持FSO<% Else %><input name="Submit" type="submit" class="inputbt" value="还 原"  onClick="{if (confirm('确定还原？这将还原到最后备份时间!')){return true;}return false;}"><% End If %>
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
