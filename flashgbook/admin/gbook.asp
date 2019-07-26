<!--#include file="admin_conn.asp"-->
<% 
	if session("admin_name")="" then
		response.redirect"admin.asp"
	end if
'-----------------------------------------------------------------------
function encodestr(str)
	dim i
	str=trim(str)
	str=replace(str,"'","""")
	str=replace(str,vbCrLf&vbCrlf,"</p><p>")
	encodestr=replace(str,vbCrLf,"<br>")
end function
Function uni(Chinese)
	For j = 1 to Len (Chinese)
	a=Mid(Chinese, j, 1)
	uni= uni & "&#x" & Hex(Ascw(a)) & ";"
	next
End Function
'------------------------------------------------------------
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
select case request("action")
	case "gopage"
		call manage()

	case "gmhuifu"
		call gmhuifu()

		
	case "manage"
		call manage()
		
		
	case "view"
		call view()
		
	case "del"
		call del()
				
	case "huifu"
		call huifu()
		
	case "designup"
		call designup()
		
	case "designdown"
		call designdown()
end select
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
	rs("content")=encodestr(trim(request.form("content")))
	
	if trim(request.form("gmcontent"))<>"" then
	rs("gmcontent")=encodestr(trim(request.form("gmcontent")))
	rs("gmdate")=date()
	else
	rs("gmcontent")=null
	rs("gmdate")=null
	end if
	
	rs.update		
	rs.close
	set rs=nothing
	response.write"<script language='javascript'>alert('编辑成功！');location.href('?action=manage');</script>"
end sub
sub del()
		set rs=server.CreateObject("adodb.recordset")
		sql="select * from gbook where id="&request("id")&""
		rs.open sql,conn,3,2
		rs.delete		
		rs.close
		set rs=nothing
        response.write"<script language='javascript'>alert('删除成功！');location.href('?action=manage');</script>"
end sub
sub designup()
'排序向上
	set rs=server.createobject("adodb.recordset")
	sql="select pxid from gbook where id in("&request("id")&")"
	rs.open sql,conn,3,2
	rs(0)=rs(0)+1
	rs.update
	rs.close
	set rs=nothing
	response.write "<script language=javascript>alert('编辑成功！');location.href('?action=manage');</script>"
end sub
'排序向下
sub designdown()
	set rs=server.createobject("adodb.recordset")
	sql="select pxid from gbook where id in("&request("id")&")"
	rs.open sql,conn,3,2
	rs(0)=rs(0)-1
	rs.update
	rs.close
	set rs=nothing
	response.write "<script language=javascript>alert('编辑成功！');location.href('?action=manage');</script>"
end sub
 %>
<% sub manage() %>
  <%
dim idcount'记录总数
dim pages'每页条数
dim pagec'总页数
dim page'页码
dim pagenc
dim pagenb
dim datafrom'数据表名
dim taxis'排序的语句
'-------------------设置参数开始---------------------------------
'taxis="order by id asc" '正排序
taxis="order by pxid desc,id desc" '倒排序
pages=20'每页条数
datafrom="gbook"'数据表名
pagenb=7 '每页显示的分页页码数量 必须为单整数，例如 3 5 7 9
'-------------------设置参数结束---------------------------------
pagenc=(pagenb-1)/2
dim pagenmax '每页显示的分页的最大页码
dim pagenmin '每页显示的分页的最小页码
page=clng(request("page"))
dim sqlid'本页需要用到的id
dim myself'本页地址
myself = request.servervariables("path_info")
dim i'用于循环的整数
start=timer()
'获取记录总数
	sql="select count(id) as idcount from ["& datafrom &"]"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,0,1
idcount=rs("idcount")'获取记录总数

if(idcount>0) then'如果记录总数=0,则不处理
	if(idcount mod pages=0)then'如果记录总数除以每页条数有余数,则=记录总数/每页条数+1
		pagec=int(idcount/pages)'获取总页数
	else
		pagec=int(idcount/pages)+1'获取总页数
	end if
	'获取本页需要用到的id============================================
	'读取所有记录的id数值,因为只有id所以速度很快
	sql="select id from ["& datafrom &"] " & taxis
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1

	   rs.pagesize = pages '每页显示记录数
	   if page < 1 then page = 1
	   if page > pagec then page = pagec
	   if pagec > 0 then rs.absolutepage = page  

	for i=1 to rs.pagesize
	if rs.eof then exit for  
		if(i=1)then
			sqlid=rs("id")
		else
			sqlid=sqlid &","&rs("id")
		end if
	rs.movenext
	next
	'获取本页需要用到的id结束============================================
end if
%>
<%
if(idcount>0 and sqlid<>"") then'如果记录总数=0,则不处理
	'用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
	sql="select [id],[title],[name],[date],[pxid] from ["& datafrom &"] where id in("& sqlid &")"&taxis
	'sql="select [id],[aaaa],[bbbb],[cccc] from ["& datafrom &"] where id in("& sqlid &") "&taxis
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
%>

  <table width="760" border="1" cellpadding="1" cellspacing="1" bordercolor="66ccff">
  <tr>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
      <tr align="center" bgcolor="#3399CC">
        <td height="25" colspan="5">留言管理</td>
      </tr>
      <tr align="center" bgcolor="#006699">
        <td width="10%" height="25">ID</td>
        <td width="45%">标题</td>
        <td width="22%">姓名</td>
        <td width="13%">留言时间</td>
        <td width="10%">操作</td>
      </tr>
  <%
dim ii
ii=0
while(not rs.eof)'填充数据到表格
'if ii mod 5=0 then
'response.write"<tr>"
'end if
%>
      <tr align="center" class="fonthei">
        <td height="25" bgcolor="#006699"><%=rs("id")%><br>
<a href="?id=<%=rs("id")%>&action=designup">上</a> [<%=rs("pxid")%>] <a href="?id=<%=rs("id")%>&action=designdown">下</a> </td>
        <td align="left" bgcolor="#66CCFF"><a href="#" title="<%=rs("title")%>"><font class="fonthei"><%=rs("title")%></font></a></td>
        <td bgcolor="#66CCFF"><%=rs("name")%></td>
        <td bgcolor="#66CCFF"><%=rs("date")%></td>
        <td bgcolor="#006699"><a href="?id=<%=rs("id")%>&action=view">回复</a> <a href="?id=<%=rs("id")%>&action=del">删除</a></td>
      </tr>
      
  <%
		rs.movenext
		ii=ii+1
	wend
	%>
	  <tr  bgcolor="#3399CC" >
	    <td height="25" colspan="5">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td align="center"><span class="fontmenu1"><span class="fontlan">共有<strong><font color="#ff0000"><%=idcount%></font></strong>条数据,第<strong><font color="#ff0000"><%=pagec%></font></strong>/<strong><font color="#ff0000"><%=page%></font></strong>页,每页<strong><font color="#ff0000"><%=pages%></font></strong>条</span> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                <%
	'设置分页页码开始===============================
	pagenmin=page-pagenc'计算页码开始值
	pagenmax=page+pagenc'计算页码结束值
	if(pagenmin<1) then'如果页码开始值小于1则=1
	    pagenmin=1
	end if

	if(page>1) then'如果页码大于1则显示(第一页)
		response.write ("<a href='"& myself &"?action=manage&page=1'><FONT face=Webdings color='#FFFFFF'>9</FONT></a> ")	
	end if
	if(pagenmin>1) then'如果页码开始值大于1则显示(更前)
		response.write ("<a href='"& myself &"?action=manage&page="& page-(pagenc*2+1) &"'><FONT face=Webdings color='#FFFFFF'>7</FONT></a> ")
	end if

	if(pagenmax>pagec) then'如果页码结束值大于总页数,则=总页数
	    pagenmax=pagec
	end if

	for i = pagenmin to pagenmax'循环输出页码
	    if(i=page) then
		response.write ("<font color='#ff0000'><strong>["& i &"]</strong></font> ")
	    else
		response.write (" <a href="& myself &"?action=manage&page="& i &"><font color='#FFFFFF'>["& i &"]</font></a> ")
	    end if
	next
	if(pagenmax<pagec) then'如果页码结束值小于总页数则显示(更后)
		response.write ("<a href='"& myself &"?action=manage&page="& page+(pagenc*2+1) &"'><FONT face=webdings color='#FFFFFF'>8</FONT></a> ")
	end if
	if(page<pagec) then'如果页码小于总页数则显示(最后页)	
		response.write ("<a href='"& myself &"?action=manage&page="& pagec &"'><FONT face=webdings color='#ff0000'>:</FONT></A> ")
	end if
	'设置分页页码结束===============================
	%>
转到
<script language="javascript">
<!--
function gopage() {
window.location.href="<%=myself%>?action=manage&page="+ page.value;
}
//-->
      </script>
<input name="page" type="text" class="pagego" value="<%=page%>" size="2" onmouseover='this.focus();this.select()'>
页
<input name="submit" type="button" class="pagego" onClick="gopage()" value=" G O ">
            </span></td>
          </tr>
        </table>
		</td>

	    </tr>
		
    </table>
	</td>
  </tr>
</table>
  <% Else %>
  <table width="760" border="1" cellpadding="1" cellspacing="1" bordercolor="66ccff">
    <tr>
      <td align="center"><form action="?action=login" method="post" name="login" id="login">
          <table width="100%" border="0" cellspacing="2" cellpadding="2">
            <tr align="center" bgcolor="#3399CC">
              <td height="25" colspan="2">目前还没有留言</td>
            </tr>
            <tr>
              <td width="47%" height="25" align="right" bgcolor="#66CCFF" class="fonthei">&nbsp;</td>
              <td width="53%" align="left" bgcolor="#006699">&nbsp;</td>
            </tr>
            <tr>
              <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">&nbsp;</td>
              <td align="left" bgcolor="#006699">&nbsp;</td>
            </tr>
            <tr bgcolor="#3399CC">
              <td height="25" align="right" class="fonthei">&nbsp;</td>
              <td align="left">&nbsp;</td>
            </tr>
          </table>
      </form></td>
    </tr>
  </table>
  <% End If
			'endt=timer()
rs.close
set rs=nothing
 %>
  <br>
<% end sub %> 
<% sub view() %> 
<% sql="select * from gbook where id="&request("id")&""
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1 %>
<table width="760" border="1" cellpadding="1" cellspacing="1" bordercolor="66ccff">
  <tr>
    <td align="center"><form action="?id=<%=rs("id")%>&action=gmhuifu" method="post" name="delgbook" id="delgbook">
        <table width="100%" border="0" cellspacing="2" cellpadding="2">
          <tr align="center" bgcolor="#3399CC">
            <td height="25" colspan="3">查看留言			</td>
          </tr>
          <tr>
            <td width="10%" height="25" align="right" bgcolor="#66CCFF" class="fonthei">姓名：</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <input name="name" type="text" class="input" id="name" value="<%= rs("name") %>" size="15"></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">QQ：</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <input name="qq" type="text" class="input" id="qq" value="<%= rs("qq") %>" size="15"></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">E-mail：</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <input name="email" type="text" class="input" id="email" value="<%= rs("email") %>" size="45"></td>
          </tr>
		  <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">blog：</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <input name="blog" type="text" class="input" id="blog" value="<%= rs("blog") %>" size="45"></td>
          </tr>
		  <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">homepage：</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <input name="homepage" type="text" class="input" id="homepage" value="<%= rs("homepage") %>" size="45"></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">留言时间：</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;<%=rs("date")%></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">标题：</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <input name="title" type="text" class="input" id="title" value="<%= rs("title") %>" size="45"></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">内容：</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <textarea name="content" cols="45" rows="5" class="input" id="content"><%= rs("content") %></textarea></td>
          </tr>
		    <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">回复：</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <textarea name="gmcontent" cols="45" rows="5" class="input" id="gmcontent"><%= rs("gmcontent") %></textarea></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">回复时间：</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;<%=rs("gmdate")%></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">ID：</td>
            <td width="7%" align="left" bgcolor="#006699">&nbsp;<%=rs("id")%></td>
            <td width="83%" align="left" bgcolor="#006699"><span class="fonthei">
              &nbsp;
              <input name="Submit" type="submit" class="inputbt" value="提交">
              <input name="Submit" type="button" class="inputbt" value="返 回" onClick="javascript:history.back()">
            </span></td>
          </tr>
        </table>
    </form></td>
  </tr>
</table>
<br>
<%
rs.close
set rs=nothing
%>
<% end sub %>
<br>
<!--#include file="bottom.asp" -->
</div>
</body>
</html>
