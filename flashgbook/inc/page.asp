<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<%
if request("action")="showpage" then
	dim idcount'��¼����
	dim pages'ÿҳ����
	dim pagec'��ҳ��
	dim page'ҳ��
	dim datafrom'���ݱ���
	dim taxis'��������
	'-------------------���ò�����ʼ---------------------------------
	
	'taxis="order by id asc" '������
	taxis="order by pxid desc,id desc" '������
	pages=show_page'ÿҳ����
	datafrom="gbook"'���ݱ���
	
	page=clng(request("page"))

	
	'-------------------���ò�������---------------------------------
	dim pagenmax 'ÿҳ��ʾ�ķ�ҳ�����ҳ��
	dim pagenmin 'ÿҳ��ʾ�ķ�ҳ����Сҳ��
	dim sqlid'��ҳ��Ҫ�õ���id
	dim i'����ѭ��������
	
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
		sql="select [id],[title],[name],[date] from ["& datafrom &"] where id in("& sqlid &")"&taxis
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,0,1
		while(not rs.eof)'������ݵ����
		page_id=rs("id")
		
		page_name=uni(rs("name"))
		
		if len(rs("title")) > 19 then '��ȡ�ַ�
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
