<%
'
'清空upload文件夹中所有图片
'更新记录：
'	2012-4-14 创建
	
	Dim fs
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	Dim path : path = Server.MapPath("upload")
	
	'文件夹存在
	If fs.FolderExists(path) Then
		Dim fo
		Set fo = fs.GetFolder(path)
		
		For Each x In fo.files
		  '删除文件
		  fs.DeleteFile(x.Path)
		Next
		
	End if
	
	Response.Write("<script language='javascript'>history.go(-1)</script>")
	Response.End()
%>