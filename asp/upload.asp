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
	uploader.TotalSize= 52428800'50MB,�ֽڼ�������http://www.beesky.com/newsite/bit_byte.htm
	uploader.MaxSize  = 10000*1024
	uploader.FileType = "gif/jpg/png/bmp/doc/rar/zip/exe/txt/xls/docx/xlsx/ppt/pptx"
	uploader.Savepath = "upload/"
	
	'�Զ������ϴ��ļ���
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
	'����ļ����ƺ�·����2011-09-10-5-52-255252.jpg'
	response.Write("upload/" & uploader.Form("ServerFileName"))
	if lngUpSize>uploader.MaxSize then
%>
		<script language="javascript">
		<!--
			alert("���ϴ����ļ�����ܳ���10M!!");
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
			alert("��û���ϴ��κ��ļ����������ϴ�!!");
			history.back();
		//-->
		</script>
<%
		response.end
	end if
	Set uploader = nothing
%>
