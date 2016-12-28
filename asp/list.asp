<%
'
'列出upload文件夹中所有图片
'更新记录：
'	2012-4-14 创建

	Dim fs,fname
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	Dim path : path = Server.MapPath("upload")
	Dim fo
	Set fo = fs.GetFolder(path)
	
	For Each x In fo.files
	  'Print the name of all files in the test folder
	  Response.Write("<img src=""upload/")
	  Response.Write(x.Name)
	  Response.Write(""" />")
	Next

	Set fs = Nothing
	Set fo = Nothing
%>