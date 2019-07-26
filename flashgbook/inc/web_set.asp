<!--#include file="../data.asp" -->
<%
db="flashgbook/data/"&dbname
set conn=server.createobject("adodb.connection")
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath(db)
sql="select * from web_set where id in(1)"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
web_name=rs("web_name")
web_email=rs("web_email")
web_qq=rs("web_qq")
rs.close
set rs=nothing
%>
