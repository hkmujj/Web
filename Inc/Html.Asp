<% 
'Html 处理HTML代码 (2014,1,3)
'显示HTML结构        call rw(displayHTmL("<br>aasdfds<br>"))
'关闭显示HTML结构   call rwend(unDisplayHtml("&lt;br&gt;aasdfds&lt;br&gt;"))

'显示HTML结构
Function displayHtml(str)
    str = Replace(str, "<", "&lt;") 
    str = Replace(str, ">", "&gt;") 
    displayHtml = str 
End Function 
'关闭显示HTML结构
Function unDisplayHtml(str)
    str = Replace(str, "&lt;", "<") 
    str = Replace(str, "&gt;", ">") 
    unDisplayHtml = str 
End Function 

'处理闭合HTML标签(20150902)  比上面的更好用 第二种
Function handleCloseHtml(content, imgAddAlt, action)
    Dim i, endStr, s, s2, c, labelName, startLabel, endLabel 
    action = "|" & action & "|" 
    startLabel = "<" 
    endLabel = ">" 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        endStr = Mid(content, i) 
        If s = "<" Then
            If InStr(endStr, ">") > 0 Then
                s = Mid(endStr, 1, InStr(endStr, ">")) 
                i = i + Len(s) - 1 
                s = Mid(s, 2, Len(s) - 2) 
                s = phptrim(s) 
                If Right(s, 1) = "/" Then
                    s = phptrim(Left(s, Len(s) - 1)) 
                End If 
                endStr = Right(endStr, Len(endStr) - Len(s) - 2)                                '最后字符减去当前标签  -2是因为它有<>二个字符
                '注意之前放在labelName下面
                labelName = Mid(s, 1, InStr(s & " ", " ") - 1) 
                labelName = LCase(labelName) 
                'call eerr("s",s)

                If InStr(action, "|处理A链接|") > 0 Then
                    s = handleHtmlAHref(s, labelName, "http://127.0.0.1/TestWeb/Web/", "处理A链接") '处理干净html标签
                ElseIf InStr(action, "|处理A链接第二种|") > 0 Then
                    s = handleHtmlAHref(s, labelName, "http://127.0.0.1/debugRemoteWeb.asp?url=", "处理A链接") '处理干净html标签
                End If 
                'call echo(s,labelName)   param与embed是Flash用到，不过embed有结束标签的
                If InStr("|meta|link|embed|param|input|img|br|hr|rect|line|area|script|div|span|a|", "|" & labelName & "|") > 0 Then
                    s = Replace(Replace(Replace(Replace(s, " class=""""", ""), " alt=""""", ""), " title=""""", ""), " name=""""", "") '临时这么做一下，以后要完整系统的做
                    s = Replace(Replace(Replace(Replace(s, " class=''", ""), " alt=''", ""), " title=''", ""), " name=''", "") 

                    '给vb.net软件用的 要不然它会报错，晕
                    If labelName = "img" And imgAddAlt = True Then
                        If InStr(s, " alt") = False Then
                            s = s & " alt=""""" 
                        End If 
                        s = Trim(s) 
                        s = s & " /" 
                    '补齐<script>20160106  暂时不能用这个，等改进
                    ElseIf labelName = "script" Then
                        If InStr(s, " type") = False Then
                            s = s & " type=""text/javascript""" 
                        End If 
                    ElseIf Right(Trim(s), 1) <> "/" And InStr("|meta|link|embed|param|input|img|br|hr|rect|line|area|", "|" & labelName & "|") > 0 Then
                        s = Trim(s) 
                        s = s & " /" 
                    End If 
                End If 
                s = startLabel & s & endLabel 
                '处理javascript script部分
                If labelName = "script" Then
                    s2 = Mid(endStr, 1, InStr(endStr, "</script>") + 8) 

                    'call eerr("",s2)
                    i = i + Len(s2) 
                    s = s & s2 
                End If 
            'call echo("s",replace(s,"<","&lt;"))
            End If 
        End If 
        c = c & s 
    Next 
    handleCloseHtml = c 
End Function 
'处理htmlA标签的Href链接  配合上面函数
Function handleHtmlAHref(ByVal content, labelName, addToHttpUrl, action)
    Dim i, s, c, temp 
    Dim isValue                                                               '是否为内容值
    Dim valueStr                                                              '存储内容值
    Dim yinghaoLabel                                                          '引号类型如'"
    Dim parentName                                                            '参数名称
    Dim behindStr                                                             '后面全部字符
    Dim noDanYinShuangYinStr                                                  '不是单引号和双引号字符
    action = "|" & action & "|" 
    content = Replace(content & " ", vbTab, " ")                              '退格替换成空格，最后加一个空格，方便计算
    content = Replace(Replace(content, " =", "="), " =", "=") 
    isValue = False                                                           '默认内容为假，因为先是获得标签名称
    For i = 1 To Len(content)
        s = Mid(content, i, 1)                                                       '获得当前一个字符
        behindStr = Mid(content, i)                                                  '后面字符
        If s = "=" And isValue = False Then                                          '不是内容值，并为=号
            isValue = True 
            valueStr = "" 
            yinghaoLabel = "" 
            If c <> "" And Right(c, 1) <> " " Then c = c & " " 
            parentName = LCase(temp)                           '参数名称转小写
            c = c & parentName & s 
            temp = "" 
        '获得值第一个字符，因为它是引号类型
        ElseIf isValue = True And yinghaoLabel = "" Then
            If s <> " " Then
                If s <> "'" And s <> """" Then
                    noDanYinShuangYinStr = s                                                        '不是单引号和双引号字符
                    s = " " 
                End If 
                yinghaoLabel = s 
            'call echo("yinghaoLabel",yinghaoLabel)
            End If 
        ElseIf isValue = True And yinghaoLabel <> "" Then
            '为引号结束
            If yinghaoLabel = s Then
                isValue = False 
                If labelName = "a" And parentName = "href" And InStr(action, "|处理A链接|") > 0 Then
                    '处理
                    If InStr(valueStr, "?") > 0 Then
                        valueStr = Replace(valueStr, "?", "WenHao") & ".html" 
                    End If 
                    If InStr("|asp|php|aspx|jsp|", "|" & LCase(Mid(valueStr, InStrRev(valueStr, ".") + 1)) & "|") > 0 Then
                        valueStr = valueStr & ".html" 
                    End If 
                    valueStr = addToOrAddHttpUrl(addToHttpUrl, valueStr, "替换") 

                End If 
                'call echo("labelName",labelName)
                If yinghaoLabel = " " Then
                    c = c & """" & noDanYinShuangYinStr & valueStr & """ "                          '追加 不是单引号和双引号字符            补全
                Else
                    c = c & yinghaoLabel & valueStr & yinghaoLabel                                  '追加 不是单引号和双引号字符
                End If 
                yinghaoLabel = "" 
                noDanYinShuangYinStr = ""                                                       '不是单引号和双引号字符 清空
            Else
                valueStr = valueStr & s 
            End If 
        '为 分割
        ElseIf s = " " Then
            '暂存内容不为空
            If temp <> "" Then
                If Left(Trim(behindStr) & " ", 1) = "=" Then
                '后面一个字符等于=不处理
                Else
                    '为标签
                    If isValue = False Then
                        temp = LCase(temp) & " "                              '标签类型名称转小写
                    End If 
                    c = c & temp 
                    temp = "" 
                End If 
            End If 
        Else
            temp = temp & s 
        End If 

    Next 
    c = Trim(c) 
    handleHtmlAHref = c 
End Function 
'追加或替换网址(20150922) 配合上面   addToOrAddHttpUrl("http://127.0.0.1/aa/","http://127.0.0.1/4.asp","替换") = http://127.0.0.1/aa/4.asp
Function addToOrAddHttpUrl(httpurl, ByVal url, action)
    Dim s 
    action = "|" & action & "|" 
    If InStr(action, "|替换|") > 0 Then
        s = getwebsite(url) 
        If s <> "" Then
            url = Replace(url, s, "") 
        End If 
    End If 
    If InStr(url, httpurl) = False Then
        If Right(httpurl, 1) = "/" And(Left(url, 1) = "/" Or Left(url, 1) = "\") Then
            url = Mid(url, 2) 
        End If 
        url = httpurl & url 
    End If 

    addToOrAddHttpUrl = url 
End Function 

'获得HTML标签名 call rwend(getHtmlLableName("<img src><a href=>",0))    输入  img
Function getHtmlLableName(content, nThisLabel)
    Dim i, endStr, s, c, labelName, nLabelCount 
    nLabelCount = 0 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        endStr = Mid(content, i) 
        If s = "<" Then
            If InStr(endStr, ">") > 0 Then
                s = Mid(endStr, 1, InStr(endStr, ">")) 
                i = i + Len(s) - 1 
                s = Mid(s, 2, Len(s) - 2) 
                s = phptrim(s) 
                If Right(s, 1) = "/" Then
                    s = phptrim(Left(s, Len(s) - 1)) 
                End If 
                endStr = Right(endStr, Len(endStr) - Len(s) - 2)                                '最后字符减去当前标签  -2是因为它有<>二个字符
                '注意之前放在labelName下面
                labelName = Mid(s, 1, InStr(s & " ", " ") - 1) 
                labelName = LCase(labelName) 
                If nThisLabel = nLabelCount Then
                    Exit For 
                End If 
                nLabelCount = nLabelCount + 1 
            End If 
        End If 
        c = c & s 
    Next 
    getHtmlLableName = labelName 
End Function 

'删除html里空行 最笨的方法 删除空行
Function removeBlankLines(content)
    Dim s, c, splStr 
    splStr = Split(content, vbCrLf) 
    For Each s In splStr
        If Replace(Replace(s, vbTab, ""), " ", "") <> "" Then
            If c <> "" Then c = c & vbCrLf 
            c = c & s 
        End If 
    Next 
    removeBlankLines = c 
End Function 


'call echo("webtitle",getHtmlValue(content,"webtitle"))
'call echo("webdescription",getHtmlValue(content,"webdescription"))
'call echo("webkeywords",getHtmlValue(content,"webkeywords"))
'获得html里面指定值20160520 call echo("webtitle",getHtmlValue(content,"webtitle"))
Function getHtmlValue(content, sType)
    Dim i, endStr, s, labelName, startLabel, endLabel, LCaseEndStr, paramName 
    startLabel = "<" 
    endLabel = ">" 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        endStr = Mid(content, i) 
        If s = "<" Then
            If InStr(endStr, ">") > 0 Then
                s = Mid(endStr, 1, InStr(endStr, ">")) 
                i = i + Len(s) - 1 
                s = Mid(s, 2, Len(s) - 2) 
                s = phptrim(s) 
                If Right(s, 1) = "/" Then
                    s = phptrim(Left(s, Len(s) - 1)) 
                End If 
                endStr = Right(endStr, Len(endStr) - Len(s) - 2)                                '最后字符减去当前标签  -2是因为它有<>二个字符
                '注意之前放在labelName下面
                labelName = Mid(s, 1, InStr(s & " ", " ") - 1) 
                labelName = LCase(labelName) 

                If labelName = "title" And sType = "webtitle" Then
                    LCaseEndStr = LCase(endStr) 
					if InStr(LCaseEndStr, "</title>")>0 then
						s = Mid(endStr, 1, InStr(LCaseEndStr, "</title>") - 1) 
					else
						s=""
					end if
					getHtmlValue = s 
                    Exit Function 
                ElseIf labelName = "meta" And(sType = "webkeywords" Or sType = "webdescription") Then
                    LCaseEndStr = LCase(endStr) 
                    paramName = phptrim(LCase(getParamValue(s, "name"))) 
                    If "web" & paramName = sType Then
                        getHtmlValue = getParamValue(s, "content") 
                        Exit Function 
                    End If 


                End If 

            End If 
        End If 
    Next 
    getHtmlValue = "" 
End Function 

'获得参数值20160520  call rwend(getParamValue("meta name=""keywords"" content=""{$web_keywords$}""","name"))
Function getParamValue(content, paramName)
    Dim LCaseContent, s, splStart, splEnd, i, startStr, endStr 
    LCaseContent = LCase(content) 

    splStart = Array("=""", "='", "=") 
    splEnd = Array("""", "'", ">") 
    For i = 0 To UBound(splStart)
        startStr = paramName & splStart(i) 
        endStr = splEnd(i) 
        If InStr(LCaseContent, startStr) > 0 And InStr(LCaseContent, endStr) > 0 Then
            s = strCut(content, startStr, endStr, 2) 
            If s <> "" Then
                getParamValue = s 
                Exit Function 
            End If 
        End If 
    Next 
End Function 

%>    



