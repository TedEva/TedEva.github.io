<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Server.ScriptTimeout=999%>
<% 
	if session("admin_name")="" then
		response.redirect"admin.asp"
	end if
%>
<!--#include file="stone_Upload.asp" -->
<!-- 上传pic_s开始 -->
<% 
	if request("action")="pic_s" then
		set upload=new upload_5xsoft
		set file=upload.file("pic")
		fileExt=lcase(right(file.filename,4))
		if  fileEXT<>".jpg" and fileEXT<>".jpeg" and fileEXT<>".swf"  then
			response.write"<script>alert('图片格式不对，请重新上传！');location='"&request.ServerVariables("HTTP_REFERER")&"'</script>"
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
		file.SaveAs Server.mappath(formPath&fname)   ''保存文件
		
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
			fname = replace(fname,"上午","")
			fname = replace(fname,"下午","")
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
<!-- 上传pic_s结束 -->

<!-- 上传pic_b开始 -->
<% 
	if request("action")="pic_b" then
		set upload=new upload_5xsoft
		set file=upload.file("pic")
		fileExt=lcase(right(file.filename,4))
		if  fileEXT<>".jpg" and fileEXT<>".jpeg"  then
			response.write"<script>alert('图片格式不对，请重新上传！');location='"&request.ServerVariables("HTTP_REFERER")&"'</script>"
			response.end
		end if 
		if file.fileSize>0 then
			formPath="uploadfile/pic"			
			if right(formPath,1)<>"/" then 
				formPath=formPath&"/"
			end if			
		vfname = filename(now())
		fname = vfname & "." & GetExtendName(file.FileName)
		file.SaveAs Server.mappath(formPath&fname)   ''保存文件
		
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
			fname = replace(fname,"上午","")
			fname = replace(fname,"下午","")
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
<!-- 上传pic_b结束 -->

<!-- 上传单个文件开始 -->
<% 
	if request("action")="onefile" then
		set upload=new upload_5xsoft
		set file=upload.file("onefile")
		fileExt=lcase(right(file.filename,4))
		if fileEXT<>".jpg" then
			response.write"<script>alert('必须是 .jpg 格式！');location='"&request.ServerVariables("HTTP_REFERER")&"'</script>"
			response.end
		end if 
		if file.fileSize>0 then
			formPath="../images"
			if right(formPath,1)<>"/" then 
			formPath=formPath&"/"
			end if			
			vfname = "bg"'固定文件名
			fname = vfname & "." & GetExtendName(file.FileName)
			file.SaveAs Server.mappath(formPath&fname)   ''保存文件
%>
		<script>
		parent.canshu.bg_jpg.value='上传成功'
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
			'fname = replace(fname,"上午","")
			'fname = replace(fname,"下午","")
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
<!-- 上传单个文件结束 -->