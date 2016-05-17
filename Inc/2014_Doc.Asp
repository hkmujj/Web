<%
'doc文件操作

'读Doc文件内容 20141218
'例：Call Rw( GetDocText("1.docx"))
Function GetDocText(ByVal FileName)
    On Error Resume Next
    Dim Fso, FText, OpenFile 
    'GetFileText = ""   '它默认返回的就是空， 这个是多此一举 (2013,9,30)
    Call HandlePath(FileName)    '获得完整路径 	
	Dim doc
	Set doc = CreateObject("AxWsc.Word.1.00")
	GetDocText = doc.GetDocContent(FileName)
    If Err Then doError Err.Description, "GetDocText 读取文件内容 函数出错，FileName=" & FileName 
End Function
%>