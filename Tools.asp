<!--#Include virtual = "/Inc/_Config.Asp"--> 
<%

call mydd3()
sub mydd3()
	dim content
	content=request("content")
	call createaddfile("tools.txt",request("act") & "=" & request("content"))
	
	if request("act")="gbtoutf" then
		call rw("<pre>")
		call rwend(GBtoUTF8(request("content")))
	elseif request("act")="pinyin" then
		call pinYin(request("content"),"")
	elseif request("act")="chatrobot" then				'aa/tools.asp?act=chatrobot&content=����
		call chatRobot(request("content"))
	elseif request("act")="formatphp" or request("act")="formatjs" then 
		call rw("<pre>")
		call rw(replace(formattingJSPHP(request("content")),"<","&lt;"))
	elseif request("act")="formathtml" then 
		call rw("<pre>")
		'�Զ�������20160420
		if request("stype")="autohandlerow" then
			content=formatting(content, "")		
		end if
		content = htmlFormatting(content)                                                     '��			
		content=replace(content,"<","&lt;")
		call rw(content)
		
	end if
end sub
'���������
function chatRobot(content)
	dim s,url,c
	url="http://www.tuling123.com/openapi/api?key=3e242fb1429235f841789a591a81f108&info=" + escape(content)
	s=getHttpUrl(url,"utf-8")
	s=getstrcut(s,"""text"":""","""}",0)
	c="�ʣ�" & content & "<br>��" & s & "<hr>" 
	c=escape(c)
	call rw(c)
end function
%> 