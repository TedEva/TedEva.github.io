<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<%
if request("action")="showpage" then
	dim idcount'记录总数
	dim pages'每页条数
	dim pagec'总页数
	dim page'页码
	dim datafrom'数据表名
	dim taxis'排序的语句
	'-------------------设置参数开始---------------------------------
	
	'taxis="order by id asc" '正排序
	taxis="order by pxid desc,id desc" '倒排序
	pages=show_page'每页条数
	datafrom="gbook"'数据表名
	
	page=clng(request("page"))

	
	'-------------------设置参数结束---------------------------------
	dim pagenmax '每页显示的分页的最大页码
	dim pagenmin '每页显示的分页的最小页码
	dim sqlid'本页需要用到的id
	dim i'用于循环的整数
	
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
		sql="select [id],[title],[name],[date] from ["& datafrom &"] where id in("& sqlid &")"&taxis
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,0,1
		while(not rs.eof)'填充数据到表格
		page_id=rs("id")
		
		page_name=uni(rs("name"))
		
		if len(rs("title")) > 19 then '截取字符
		page_title=left(rs("title"),19)&".."
		else
		page_title=rs("title")
		end if
		
		page_title=uni(page_title)
		
		'page_email=rs("email")
		'page_qq=rs("qq")
		
		page_date=rs("date")
		
			out=out&"<info page_id='"&page_id&"' page_name='"&page_name&"' page_title='"&page_title&"' page_date='"&page_date&"' />"	
		
		rs.movenext
		wend
		rs.close
		set rs=nothing
		Response.Write "<?xml version='1.0' encoding='utf-8'?>"
		Response.Write "<gbook total='"&idcount&"' maxpage='"&pagec&"' page='"&page&"'>"&out&"</gbook>"
		Session.CodePage="936"
	
	else
		total=0
		maxpage=0
		page=0
		out=""
		wujilu="1"
		Response.Write "<?xml version='1.0' encoding='utf-8'?>"
		Response.Write "<gbook total='"&total&"' maxpage='"&maxpage&"' page='"&page&"' wujilu='"&wujilu&"'></gbook>"
		Session.CodePage="936"
	end if
end if'end action showpage
%>
