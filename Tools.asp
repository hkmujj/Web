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
	'asp�������
	elseif request("act")="aspCodeConfusion" then 
		call rw("<pre>")	
		content=aspCodeConfusion(content)
		content=replace(content,"<","&lt;")
		call rwend(content)
	'js�������
	elseif request("act")="jsCodeConfusion" then 
		call rw("<pre>")	
		content=jsCodeConfusion(content)
		content=replace(content,"<","&lt;")
		call rwend(content)
	'php�������
	elseif request("act")="phpCodeConfusion" then 
		call rw("<pre>")	
		content=phpCodeConfusion(content)
		content=replace(content,"<","&lt;")
		call rwend(content)		
	'�����뷱�廥��
	elseif request("act")="simplifiedTab" then 
		'Ϊ����ת����
		if request("isFan")="1" then		
			call rwend(simplifiedTransfer(content))
		else
			call rwend(simplifiedChinese(content))
		end if
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