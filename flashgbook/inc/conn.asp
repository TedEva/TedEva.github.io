<!--#include file="../data.asp" -->
<% 
db="../data/"&dbname
show_page = 13 '每页显示的纪录 
set conn=server.createobject("adodb.connection")
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath(db)
'-----------------------------------------------------------------------
'用途:将UTF-8编码汉字转为GB2312码，兼容英文和数字！ 
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