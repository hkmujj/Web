<!--#Include File = "Inc/Config.Asp"--> 
<%
dim s,s2,tempS,c,url
 
 
        Dim myOBJ1: Set myOBJ1 = CreateObject("xy2015.calldll")
		dim f1,f2,f3,f4,f5
		f1=handlePath("/1/1.txt")
		f2=handlePath("/1/2.gif")
		f3=handlePath("/1/3.txt")
        call myOBJ1.callHeBingFile(f1,f2,f3)
		
		f4=handlePath("/1/end.txt")
		f5=handlePath("/1/5.txt")
        call myOBJ1.callHeBingFile(f3,f4,f5)
 
  c=getFileBinaryContent(f5)
'  c=getFileBinaryContent("/postdata.txt")
 	 
	url="http://127.0.0.1/admin/upfile_photo.asp?returnInputName="
	call rw(XMLPost(url, c))
%>

 