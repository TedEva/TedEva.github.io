<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Server.ScriptTimeout=999%>
<% 
	if session("admin_name")="" then
		response.redirect"admin.asp"
	end if
%>
<!--#include file="stone_Upload.asp" -->
<!-- �ϴ�pic_s��ʼ -->
<% 
	if request("action")="pic_s" then
		set upload=new upload_5xsoft
		set file=upload.file("pic")
		fileExt=lcase(right(file.filename,4))
		if  fileEXT<>".jpg" and fileEXT<>".jpeg" and fileEXT<>".swf"  then
			response.write"<script>alert('ͼƬ��ʽ���ԣ��������ϴ���');location='"&request.ServerVariables("HTTP_REFERER")&"'</script>"
			response.end
		end if 
		if file.fileSize>0 then
			'if file.fileSize>1000*1024 then'
			'else
			formPath="uploadfile/pic"			
			if right(formPath,1)<>"/" then 
				formPath=formPath&"/"
			end if			
		vfname = filename(now())
		fname = vfname & "." & GetExtendName(file.FileName)
		file.SaveAs Server.mappath(formPath&fname)   ''�����ļ�
		
%>
<script>
		parent.myform.pic_s.value='uploadfile/pic/<%=fname%>'
		//parent.frmadd.dreamcontent.value+='[img]upload/<%=ufp%>[/img]'
		location.replace('stone_uppic_s.asp')
		</script>
<%		'	end if
		end if
		set file=nothing
		set upload=nothing
		function filename(fname)
			fname = now()
			fname = replace(fname,"-","")
			fname = replace(fname," ","") 
			fname = replace(fname,":","")
			fname = replace(fname,"PM","")
			fname = replace(fname,"AM","")
			fname = replace(fname,"����","")
			fname = replace(fname,"����","")
			filename=fname
		end function 
		function GetExtendName(FileName)
			dim ExtName
			ExtName = LCase(FileName)
			ExtName = right(ExtName,3)
			ExtName = right(ExtName,3-Instr(ExtName,"."))
			GetExtendName = ExtName
		end function 
	end if'end action
%>
<!-- �ϴ�pic_s���� -->

<!-- �ϴ�pic_b��ʼ -->
<% 
	if request("action")="pic_b" then
		set upload=new upload_5xsoft
		set file=upload.file("pic")
		fileExt=lcase(right(file.filename,4))
		if  fileEXT<>".jpg" and fileEXT<>".jpeg"  then
			response.write"<script>alert('ͼƬ��ʽ���ԣ��������ϴ���');location='"&request.ServerVariables("HTTP_REFERER")&"'</script>"
			response.end
		end if 
		if file.fileSize>0 then
			formPath="uploadfile/pic"			
			if right(formPath,1)<>"/" then 
				formPath=formPath&"/"
			end if			
		vfname = filename(now())
		fname = vfname & "." & GetExtendName(file.FileName)
		file.SaveAs Server.mappath(formPath&fname)   ''�����ļ�
		
%>
<script>
		parent.myform.pic_b.value='uploadfile/pic/<%=fname%>'
		//parent.frmadd.dreamcontent.value+='[img]upload/<%=ufp%>[/img]'
		location.replace('stone_uppic_b.asp')
		</script>
<%			'end if
		end if
		set file=nothing
		set upload=nothing
		function filename(fname)
			fname = now()
			fname = replace(fname,"-","")
			fname = replace(fname," ","") 
			fname = replace(fname,":","")
			fname = replace(fname,"PM","")
			fname = replace(fname,"AM","")
			fname = replace(fname,"����","")
			fname = replace(fname,"����","")
			filename=fname
		end function 
		function GetExtendName(FileName)
			dim ExtName
			ExtName = LCase(FileName)
			ExtName = right(ExtName,3)
			ExtName = right(ExtName,3-Instr(ExtName,"."))
			GetExtendName = ExtName
		end function 
	end if'end action
%>
<!-- �ϴ�pic_b���� -->

<!-- �ϴ������ļ���ʼ -->
<% 
	if request("action")="onefile" then
		set upload=new upload_5xsoft
		set file=upload.file("onefile")
		fileExt=lcase(right(file.filename,4))
		if fileEXT<>".jpg" then
			response.write"<script>alert('������ .jpg ��ʽ��');location='"&request.ServerVariables("HTTP_REFERER")&"'</script>"
			response.end
		end if 
		if file.fileSize>0 then
			formPath="../images"
			if right(formPath,1)<>"/" then 
			formPath=formPath&"/"
			end if			
			vfname = "bg"'�̶��ļ���
			fname = vfname & "." & GetExtendName(file.FileName)
			file.SaveAs Server.mappath(formPath&fname)   ''�����ļ�
%>
		<script>
		parent.canshu.bg_jpg.value='�ϴ��ɹ�'
		//parent.frmadd.dreamcontent.value+='[img]upload/<%=ufp%>[/img]'
		location.replace('stone_Uponefile.asp')
		</script>
<%			
		end if
		set file=nothing
		set upload=nothing
		'function filename(fname)
			'fname = now()
			'fname = replace(fname,"-","")
			'fname = replace(fname," ","") 
			'fname = replace(fname,":","")
			'fname = replace(fname,"PM","")
			'fname = replace(fname,"AM","")
			'fname = replace(fname,"����","")
			'fname = replace(fname,"����","")
			'filename=fname
		'end function 
		function GetExtendName(FileName)
			dim ExtName
			ExtName = LCase(FileName)
			ExtName = right(ExtName,3)
			ExtName = right(ExtName,3-Instr(ExtName,"."))
			GetExtendName = ExtName
		end function 
	end if
%>
<!-- �ϴ������ļ����� -->