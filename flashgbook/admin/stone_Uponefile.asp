<% 
	if session("admin_name")="" then
		response.redirect"admin.asp"
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Admin_Upwork</title>
<style type="text/css">
<!--
body {
	background-color: #006699;
}
.inputbt {
	font-family: "����";
	font-size: 12px;
	border: 0px solid #000000;
	background-color: #66CCFF;
	padding-top: 2px;
	padding-left: 1px;
	height: 18px;
}
-->
</style></head>
<body topmargin="0" leftmargin="0" >
<form action="stone_Upfile.asp?action=onefile" method="POST" enctype="multipart/form-data" class="fontmenu2" onsubmit="up.disabled=true;up.value='�ϴ���,���Ժ򡭡�'">
	<div align="center">
	  <input name="onefile" type="file" class="inputbt" size="1">
	  <input name="up" type="submit" class="inputbt" value="�ϴ�" >
    </div>
</form>
</body>
</html>