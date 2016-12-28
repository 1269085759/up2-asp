<% @Language=vbscript Codepage=936 %>
<%
Option Explicit
Response.Buffer=True
%>
<!--#include file="UpLoadClass.asp"-->
<!--#include file="UrlDecode.asp"-->
<%
dim lngUpSize,uploader,intError
	Set uploader = new UpLoadClass
	uploader.TotalSize= 52428800'50MB,字节计算器：http://www.beesky.com/newsite/bit_byte.htm
	uploader.MaxSize  = 10000*1024
	uploader.FileType = "gif/jpg/png/bmp/doc/rar/zip/exe/txt/xls/docx/xlsx/ppt/pptx"
	uploader.Savepath = "upload/"
	
	'自动创建上传文件夹
	dim folder,fs
	folder = server.MapPath(uploader.Savepath)
	set fs = Server.CreateObject("Scripting.FileSystemObject")
	if(fs.FolderExists(folder) = false) then
		fs.CreateFolder(folder)
	end if

	lngUpSize = uploader.Open()
	intError = uploader.Form("photo2_Err")
	'Dim txt: txt = uploader.Form("txt")
	'txt = URLDecode(txt)
	'输出文件名称和路径：2011-09-10-5-52-255252.jpg'
	response.Write("upload/" & uploader.Form("ServerFileName"))
	if lngUpSize>uploader.MaxSize then
%>
		<script language="javascript">
		<!--
			alert("您上传的文件最大不能超过10M!!");
			history.back();
		//-->
		</script>
<%
		response.end
	end if
	if intError=-1 then
%>
		<script language="javascript">
		<!--
			alert("您没有上传任何文件，请重新上传!!");
			history.back();
		//-->
		</script>
<%
		response.end
	end if
	Set uploader = nothing
%>
