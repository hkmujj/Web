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
	elseif request("act")="chatrobot" then				'aa/tools.asp?act=chatrobot&content=美国
		call chatRobot(request("content"))
	elseif request("act")="formatphp" or request("act")="formatjs" then 
		call rw("<pre>")
		call rw(replace(formattingJSPHP(request("content")),"<","&lt;"))
	elseif request("act")="formathtml" then 
		call rw("<pre>")
		'自动处理行20160420
		if request("stype")="autohandlerow" then
			content=formatting(content, "")		
		end if
		content = htmlFormatting(content)                                                     '简单			
		content=replace(content,"<","&lt;")
		call rw(content)
	'asp代码混淆
	elseif request("act")="aspCodeConfusion" then 
		call rw("<pre>")	
		content=aspCodeConfusion(content)
		content=replace(content,"<","&lt;")
		call rwend(content)
	'js代码混淆
	elseif request("act")="jsCodeConfusion" then 
		call rw("<pre>")	
		content=jsCodeConfusion(content)
		content=replace(content,"<","&lt;")
		call rwend(content)
	'php代码混淆
	elseif request("act")="phpCodeConfusion" then 
		call rw("<pre>")	
		content=phpCodeConfusion(content)
		content=replace(content,"<","&lt;")
		call rwend(content)		
	'简体与繁体互换
	elseif request("act")="simplifiedTab" then 
		'为简体转繁体
		if request("isFan")="1" then		
			call rwend(simplifiedTransfer(content))
		else
			call rwend(simplifiedChinese(content))
		end if
	end if
end sub
'聊天机器人
function chatRobot(content)
	dim s,url,c
	url="http://www.tuling123.com/openapi/api?key=3e242fb1429235f841789a591a81f108&info=" + escape(content)
	s=getHttpUrl(url,"utf-8")
	s=getstrcut(s,"""text"":""","""}",0)
	c="问：" & content & "<br>答：" & s & "<hr>" 
	c=escape(c)
	call rw(c)
end function
%> 