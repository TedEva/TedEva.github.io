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
	response.write"<script language='javascript'>alert('设置成功！');location.href('web_set.asp');</script>"
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
          <td height="25" colspan="4">FLASH留言本参数设置</td>
        </tr>
        <tr>
          <td width="20%" height="25" align="right" bgcolor="#66CCFF" class="fonthei">留言本名称：</td>
          <td colspan="2" align="left" bgcolor="#006699"><input name="gbookname" type="text" class="input" id="gbookname" value="<%= rs("gbookname") %>" size="15"></td>
          <td width="40%" rowspan="7" align="left" bgcolor="#006699">首页及留言本下方显示参数设置</td>
        </tr>
        
        <tr>
          <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">网站名称：</td>
          <td colspan="2" align="left" bgcolor="#006699"><input name="homename" type="text" class="input" id="homename" value="<%= rs("homename") %>" size="15">            </td>
          </tr>
        <tr>
          <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">站长QQ：</td>
          <td colspan="2" align="left" bgcolor="#006699"><input name="web_qq" type="text" class="input" id="web_qq" value="<%= web_qq %>" size="15"></td>
          </tr>
        <tr>
          <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">站长Email：</td>
          <td colspan="2" align="left" bgcolor="#006699"><input name="web_email" type="text" class="input" id="web_email" value="<%=web_email %>" size="25"></td>
          </tr>
        <tr>
          <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">是否链接主页：</td>
          <td width="20%" align="center" bgcolor="#006699">是
            <input name="homehave" type="radio" value="you" <%if rs("homehave")="you" then%>checked<%end if%>></td>
          <td width="20%" align="center" bgcolor="#006699">否
            
            <input name="homehave" type="radio" value="wu" <%if rs("homehave")="wu" then%>checked<%end if%>></td>
          </tr>
        <tr>
          <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">主页地址：</td>
          <td colspan="2" align="left" bgcolor="#006699"><input name="homeurlu" type="text" class="input" id="homeurlu" value="<%= rs("homeurlu") %>" size="40"></td>
          </tr>
        <tr>
          <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">主页打开放式：</td>
          <td align="center" bgcolor="#006699">同一网页
            <input name="homeurlk" type="radio" value="_self" <%if rs("homeurlk")="_self" then%>checked<%end if%>>            </td>
          <td align="center" bgcolor="#006699">新开窗口
            <input type="radio" name="homeurlk" value="_blank" <%if rs("homeurlk")="_blank" then%>checked<%end if%>>            </td>
          </tr>
        <tr>
          <td height="25" align="right" bgcolor="#66CCFF" class="fonthei">留言本背景图片：</td>
          <td align="left" bgcolor="#006699">
 <input name="bg_jpg" type="text" class="input" id="bg_jpg" size="15" value="" disabled>
            <a href="../images/bg.jpg" target="_blank">查看</a></td>
          <td align="center" bgcolor="#006699"><iframe name="I1" width="155" height="25" src="stone_uponefile.asp" scrolling="no" border="0" frameborder="0">浏览器不支持嵌入式框架，或被配置为不显示嵌入式框架。</iframe></td>
          <td align="left" bgcolor="#006699">788x430 像素 *.jpg 文件</td>
        </tr>
        <tr align="center" bgcolor="#3399CC">
          <td height="25" colspan="4" class="fonthei"><input name="Submit" type="submit" class="inputbt" value="修 改">
            &nbsp;&nbsp;&nbsp;&nbsp;
            <input name="Submit" type="reset" class="inputbt" value="重 置"></td>
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
