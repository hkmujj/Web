<!--#Include File = "Inc/_Config.Asp"-->            
<!--#Include File = "web/function.Asp"-->      
<%

		dim code
		dim s,i
if 1=1 then
		code=getftext("1.html")
				code = ziphtml(code)                         '自定义
		call rw(code)


else

		s="123中国abc"
		call echo(s,len(s))
		for i = 1 to len(s)
			call echo(i,mid(s,i,1))
		next

end if
%>