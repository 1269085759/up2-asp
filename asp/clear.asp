<%
'
'���upload�ļ���������ͼƬ
'���¼�¼��
'	2012-4-14 ����
	
	Dim fs
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	Dim path : path = Server.MapPath("upload")
	
	'�ļ��д���
	If fs.FolderExists(path) Then
		Dim fo
		Set fo = fs.GetFolder(path)
		
		For Each x In fo.files
		  'ɾ���ļ�
		  fs.DeleteFile(x.Path)
		Next
		
	End if
	
	Response.Write("<script language='javascript'>history.go(-1)</script>")
	Response.End()
%>