<!--#include file="../data.asp" -->
<% 
db="../data/"&dbname
show_page = 13 'ÿҳ��ʾ�ļ�¼ 
set conn=server.createobject("adodb.connection")
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath(db)
'-----------------------------------------------------------------------
'��;:��UTF-8���뺺��תΪGB2312�룬����Ӣ�ĺ����֣� 
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