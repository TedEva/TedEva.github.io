<!--#include file="admin_conn.asp"-->
<% 
	if session("admin_name")="" then
		response.redirect"admin.asp"
	end if
	sql="select * from web_set where id in(1)"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
web_name=rs("web_name")
web_email=rs("web_email")
web_qq=rs("web_qq")
rs.close
set rs=nothing
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
	case "edit"
		call edit()
end select
sub edit()
	set rstt=server.CreateObject("adodb.recordset")
	sqltt="select * from gbook_set where id in(1)"
	rstt.open sqltt,conn,3,2
	rstt("gbookname")=trim(request.form("gbookname"))
	rstt("homehave")=trim(request.form("homehave"))
	rstt("homeurlu")=trim(request.form("homeurlu"))
	rstt("homeurlk")=trim(request.form("homeurlk"))
	rstt("homename")=trim(request.form("homename"))
	rstt.update		
	rstt.close
	set rstt=nothing
	set rst=server.CreateObject("adodb.recordset")
	sqlt="select * from web_set where id in(1)"
	rst.open sqlt,conn,3,2
	rst("web_qq")=trim(request.form("web_qq"))	
	rst("web_email")=trim(request.form("web_email"))	
	rst("web_name")=trim(request.form("homename"))
	rst.update		
	rst.close
	set rst=nothing
	response.write"<script language='javascript'>alert('���óɹ���');location.href('web_set.asp');</script>"
end sub
 %>

<%
 	set rs=server.CreateObject("adodb.recordset")
	sql="select * from gbook_set where id in(1)"
	rs.open sql,conn,0,1
%>
  <table width="760" border="1" cellpadding="1" cellspacing="1" bordercolor="66ccff">
  <tr>
    <td>
      <table width="100%" border="0" cellpadding="2" cellspacing="2">
	  <form action="?action=edit" method="post" name="canshu" id="canshu">
        <tr align="center" bgcolor="#3399CC">
          <td height="25" colspan="4">FLASH���Ա���������</td>
        </tr>
        <tr>
          <td width="20%" height="25" align="right" bgcolor="#66CCFF" class="fonthei">���Ա����ƣ�</td>
          <td colspan="2" align="left" bgcolor="#006699"><input name="gbookname" type="text" class="input" id="gbookname" value="<%= rs("gbookname") %>" size="15"></td>
          <td width="40%" rowspan="7" align="left" bgcolor="#006699">��ҳ�����Ա��·���ʾ��������</td>
        </tr>
        
        <tr>
          <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">��վ���ƣ�</td>
          <td colspan="2" align="left" bgcolor="#006699"><input name="homename" type="text" class="input" id="homename" value="<%= rs("homename") %>" size="15">            </td>
          </tr>
        <tr>
          <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">վ��QQ��</td>
          <td colspan="2" align="left" bgcolor="#006699"><input name="web_qq" type="text" class="input" id="web_qq" value="<%= web_qq %>" size="15"></td>
          </tr>
        <tr>
          <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">վ��Email��</td>
          <td colspan="2" align="left" bgcolor="#006699"><input name="web_email" type="text" class="input" id="web_email" value="<%=web_email %>" size="25"></td>
          </tr>
        <tr>
          <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">�Ƿ�������ҳ��</td>
          <td width="20%" align="center" bgcolor="#006699">��
            <input name="homehave" type="radio" value="you" <%if rs("homehave")="you" then%>checked<%end if%>></td>
          <td width="20%" align="center" bgcolor="#006699">��
            
            <input name="homehave" type="radio" value="wu" <%if rs("homehave")="wu" then%>checked<%end if%>></td>
          </tr>
        <tr>
          <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">��ҳ��ַ��</td>
          <td colspan="2" align="left" bgcolor="#006699"><input name="homeurlu" type="text" class="input" id="homeurlu" value="<%= rs("homeurlu") %>" size="40"></td>
          </tr>
        <tr>
          <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">��ҳ�򿪷�ʽ��</td>
          <td align="center" bgcolor="#006699">ͬһ��ҳ
            <input name="homeurlk" type="radio" value="_self" <%if rs("homeurlk")="_self" then%>checked<%end if%>>            </td>
          <td align="center" bgcolor="#006699">�¿�����
            <input type="radio" name="homeurlk" value="_blank" <%if rs("homeurlk")="_blank" then%>checked<%end if%>>            </td>
          </tr>
        <tr>
          <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">���Ա�����ͼƬ��</td>
          <td align="left" bgcolor="#006699">
 <input name="bg_jpg" type="text" class="input" id="bg_jpg" size="15" value="" disabled>
            <a href="../images/bg.jpg" target="_blank">�鿴</a></td>
          <td align="center" bgcolor="#006699"><iframe name="I1" width="155" height="25" src="stone_uponefile.asp" scrolling="no" border="0" frameborder="0">�������֧��Ƕ��ʽ��ܣ�������Ϊ����ʾǶ��ʽ��ܡ�</iframe></td>
          <td align="left" bgcolor="#006699">788x430 ���� *.jpg �ļ�</td>
        </tr>
        <tr align="center" bgcolor="#3399CC">
          <td height="25" colspan="4" class="fonthei"><input name="Submit" type="submit" class="inputbt" value="�� ��">
            &nbsp;&nbsp;&nbsp;&nbsp;
            <input name="Submit" type="reset" class="inputbt" value="�� ��"></td>
          </tr>
	      </form>
      </table>
   
      </td>
  </tr>
</table>
  <br>
<% rs.close
set rs=nothing %>
<!--#include file="bottom.asp" -->
</div>
</body>
</html>
