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
	response.write"<script language='javascript'>alert('�༭�ɹ���');location.href('?action=manage');</script>"
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
sub designup()
'��������
	set rs=server.createobject("adodb.recordset")
	sql="select pxid from gbook where id in("&request("id")&")"
	rs.open sql,conn,3,2
	rs(0)=rs(0)+1
	rs.update
	rs.close
	set rs=nothing
	response.write "<script language=javascript>alert('�༭�ɹ���');location.href('?action=manage');</script>"
end sub
'��������
sub designdown()
	set rs=server.createobject("adodb.recordset")
	sql="select pxid from gbook where id in("&request("id")&")"
	rs.open sql,conn,3,2
	rs(0)=rs(0)-1
	rs.update
	rs.close
	set rs=nothing
	response.write "<script language=javascript>alert('�༭�ɹ���');location.href('?action=manage');</script>"
end sub
 %>
<% sub manage() %>
  <%
dim idcount'��¼����
dim pages'ÿҳ����
dim pagec'��ҳ��
dim page'ҳ��
dim pagenc
dim pagenb
dim datafrom'���ݱ���
dim taxis'��������
'-------------------���ò�����ʼ---------------------------------
'taxis="order by id asc" '������
taxis="order by pxid desc,id desc" '������
pages=20'ÿҳ����
datafrom="gbook"'���ݱ���
pagenb=7 'ÿҳ��ʾ�ķ�ҳҳ������ ����Ϊ������������ 3 5 7 9
'-------------------���ò�������---------------------------------
pagenc=(pagenb-1)/2
dim pagenmax 'ÿҳ��ʾ�ķ�ҳ�����ҳ��
dim pagenmin 'ÿҳ��ʾ�ķ�ҳ����Сҳ��
page=clng(request("page"))
dim sqlid'��ҳ��Ҫ�õ���id
dim myself'��ҳ��ַ
myself = request.servervariables("path_info")
dim i'����ѭ��������
start=timer()
'��ȡ��¼����
	sql="select count(id) as idcount from ["& datafrom &"]"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,0,1
idcount=rs("idcount")'��ȡ��¼����

if(idcount>0) then'�����¼����=0,�򲻴���
	if(idcount mod pages=0)then'�����¼��������ÿҳ����������,��=��¼����/ÿҳ����+1
		pagec=int(idcount/pages)'��ȡ��ҳ��
	else
		pagec=int(idcount/pages)+1'��ȡ��ҳ��
	end if
	'��ȡ��ҳ��Ҫ�õ���id============================================
	'��ȡ���м�¼��id��ֵ,��Ϊֻ��id�����ٶȺܿ�
	sql="select id from ["& datafrom &"] " & taxis
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1

	   rs.pagesize = pages 'ÿҳ��ʾ��¼��
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
	'��ȡ��ҳ��Ҫ�õ���id����============================================
end if
%>
<%
if(idcount>0 and sqlid<>"") then'�����¼����=0,�򲻴���
	'��inˢѡ��ҳ�����Ե�����,����ȡ��ҳ���������,�����ٶȿ�
	sql="select [id],[title],[name],[date],[pxid] from ["& datafrom &"] where id in("& sqlid &")"&taxis
	'sql="select [id],[aaaa],[bbbb],[cccc] from ["& datafrom &"] where id in("& sqlid &") "&taxis
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
%>

  <table width="760" border="1" cellpadding="1" cellspacing="1" bordercolor="66ccff">
  <tr>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
      <tr align="center" bgcolor="#3399CC">
        <td height="25" colspan="5">���Թ���</td>
      </tr>
      <tr align="center" bgcolor="#006699">
        <td width="10%" height="25">ID</td>
        <td width="45%">����</td>
        <td width="22%">����</td>
        <td width="13%">����ʱ��</td>
        <td width="10%">����</td>
      </tr>
  <%
dim ii
ii=0
while(not rs.eof)'������ݵ����
'if ii mod 5=0 then
'response.write"<tr>"
'end if
%>
      <tr align="center" class="fonthei">
        <td height="25" bgcolor="#006699"><%=rs("id")%><br>
<a href="?id=<%=rs("id")%>&action=designup">��</a> [<%=rs("pxid")%>] <a href="?id=<%=rs("id")%>&action=designdown">��</a> </td>
        <td align="left" bgcolor="#66CCFF"><a href="#" title="<%=rs("title")%>"><font class="fonthei"><%=rs("title")%></font></a></td>
        <td bgcolor="#66CCFF"><%=rs("name")%></td>
        <td bgcolor="#66CCFF"><%=rs("date")%></td>
        <td bgcolor="#006699"><a href="?id=<%=rs("id")%>&action=view">�ظ�</a> <a href="?id=<%=rs("id")%>&action=del">ɾ��</a></td>
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
            <td align="center"><span class="fontmenu1"><span class="fontlan">����<strong><font color="#ff0000"><%=idcount%></font></strong>������,��<strong><font color="#ff0000"><%=pagec%></font></strong>/<strong><font color="#ff0000"><%=page%></font></strong>ҳ,ÿҳ<strong><font color="#ff0000"><%=pages%></font></strong>��</span> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                <%
	'���÷�ҳҳ�뿪ʼ===============================
	pagenmin=page-pagenc'����ҳ�뿪ʼֵ
	pagenmax=page+pagenc'����ҳ�����ֵ
	if(pagenmin<1) then'���ҳ�뿪ʼֵС��1��=1
	    pagenmin=1
	end if

	if(page>1) then'���ҳ�����1����ʾ(��һҳ)
		response.write ("<a href='"& myself &"?action=manage&page=1'><FONT face=Webdings color='#FFFFFF'>9</FONT></a> ")	
	end if
	if(pagenmin>1) then'���ҳ�뿪ʼֵ����1����ʾ(��ǰ)
		response.write ("<a href='"& myself &"?action=manage&page="& page-(pagenc*2+1) &"'><FONT face=Webdings color='#FFFFFF'>7</FONT></a> ")
	end if

	if(pagenmax>pagec) then'���ҳ�����ֵ������ҳ��,��=��ҳ��
	    pagenmax=pagec
	end if

	for i = pagenmin to pagenmax'ѭ�����ҳ��
	    if(i=page) then
		response.write ("<font color='#ff0000'><strong>["& i &"]</strong></font> ")
	    else
		response.write (" <a href="& myself &"?action=manage&page="& i &"><font color='#FFFFFF'>["& i &"]</font></a> ")
	    end if
	next
	if(pagenmax<pagec) then'���ҳ�����ֵС����ҳ������ʾ(����)
		response.write ("<a href='"& myself &"?action=manage&page="& page+(pagenc*2+1) &"'><FONT face=webdings color='#FFFFFF'>8</FONT></a> ")
	end if
	if(page<pagec) then'���ҳ��С����ҳ������ʾ(���ҳ)	
		response.write ("<a href='"& myself &"?action=manage&page="& pagec &"'><FONT face=webdings color='#ff0000'>:</FONT></A> ")
	end if
	'���÷�ҳҳ�����===============================
	%>
ת��
<script language="javascript">
<!--
function gopage() {
window.location.href="<%=myself%>?action=manage&page="+ page.value;
}
//-->
      </script>
<input name="page" type="text" class="pagego" value="<%=page%>" size="2" onmouseover='this.focus();this.select()'>
ҳ
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
              <td height="25" colspan="2">Ŀǰ��û������</td>
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
            <td height="25" colspan="3">�鿴����			</td>
          </tr>
          <tr>
            <td width="10%" height="25" align="right" bgcolor="#66CCFF" class="fonthei">������</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <input name="name" type="text" class="input" id="name" value="<%= rs("name") %>" size="15"></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">QQ��</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <input name="qq" type="text" class="input" id="qq" value="<%= rs("qq") %>" size="15"></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">E-mail��</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <input name="email" type="text" class="input" id="email" value="<%= rs("email") %>" size="45"></td>
          </tr>
		  <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">blog��</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <input name="blog" type="text" class="input" id="blog" value="<%= rs("blog") %>" size="45"></td>
          </tr>
		  <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">homepage��</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <input name="homepage" type="text" class="input" id="homepage" value="<%= rs("homepage") %>" size="45"></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">����ʱ�䣺</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;<%=rs("date")%></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">���⣺</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <input name="title" type="text" class="input" id="title" value="<%= rs("title") %>" size="45"></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">���ݣ�</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <textarea name="content" cols="45" rows="5" class="input" id="content"><%= rs("content") %></textarea></td>
          </tr>
		    <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">�ظ���</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;
              <textarea name="gmcontent" cols="45" rows="5" class="input" id="gmcontent"><%= rs("gmcontent") %></textarea></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">�ظ�ʱ�䣺</td>
            <td colspan="2" align="left" bgcolor="#006699">&nbsp;<%=rs("gmdate")%></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">ID��</td>
            <td width="7%" align="left" bgcolor="#006699">&nbsp;<%=rs("id")%></td>
            <td width="83%" align="left" bgcolor="#006699"><span class="fonthei">
              &nbsp;
              <input name="Submit" type="submit" class="inputbt" value="�ύ">
              <input name="Submit" type="button" class="inputbt" value="�� ��" onClick="javascript:history.back()">
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
