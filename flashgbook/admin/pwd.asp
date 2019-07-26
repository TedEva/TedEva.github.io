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
select case request("action")
		
	case "edit_pwd_save"
		call edit_pwd_save()

end select
sub edit_pwd_save()
	  admin=trim(request.form("name"))
      pwd=trim(request.form("pwd"))
	  pwd2=trim(request.form("pwd2"))	  
	  if admin="" or pwd="" then
	  	 response.write"<script language='javascript'>alert('用户名和密码不能为空！');location.href('?action=edit_pwd');</script>"
	  end if
	  if pwd2<>pwd then
	  	 response.write"<script language='javascript'>alert('两次密码不一致！');location.href('?action=edit_pwd');</script>"
	  end if
      set rs=server.createobject("adodb.recordset")
	  sql="select * from admin where id=1"
	  rs.open sql,conn,3,2
	  rs("adminname")=admin
	  rs("adminpwd")=pwd
	  rs.update
	  rs.close
	  set rs=nothing
	  response.write"<script language='javascript'>alert('修改成功！');location.href('?action=info');</script>"
end sub
 %>
<table width="760" border="1" cellpadding="1" cellspacing="1" bordercolor="66ccff">
  <tr>
    <td align="center">
        <table width="100%" border="0" cellspacing="2" cellpadding="2">
			<form action="?action=edit_pwd_save" method="post" name="pwd" id="pwd">
          <tr align="center" bgcolor="#3399CC">
            <td height="25" colspan="2">修改密码</td>
          </tr>
          <tr>
            <td width="47%" height="25" align="right" bgcolor="#66CCFF" class="fonthei">管&nbsp;理&nbsp;员：</td>
            <td width="53%" align="left" bgcolor="#006699"><input name="name" type="text" class="input" id="name" size="15"></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">密&nbsp;&nbsp;&nbsp;&nbsp;码：</td>
            <td align="left" bgcolor="#006699"><input name="pwd" type="password" class="input" id="pwd" size="15"></td>
          </tr>
          <tr>
            <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">重复密码：</td>
            <td align="left" bgcolor="#006699"><input name="pwd2" type="password" class="input" id="pwd2" size="15"></td>
          </tr>
          <tr bgcolor="#3399CC">
            <td height="25" align="right" class="fonthei">&nbsp;</td>
            <td align="left"><input name="Submit" type="submit" class="inputbt" value="修 改">
              <input name="Submit" type="reset" class="inputbt" value="重 置"></td>
          </tr>
		   </form>
        </table>
   </td>
  </tr>
</table>
<br>
<!--#include file="bottom.asp" -->
</div>
</body>
</html>
