<% 
'FSO  VBscript.RegExp组件操作大全 (2013,9,29)

'检查字符是否存在,大小写不分
Function checkStr(ByVal content, ByVal SerchValue)
    On Error Resume Next 
    Dim Rep, search 
    checkStr = True 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.Global = True                                                               '设置全程性质。
        Rep.IgnoreCase = True                                                           '设置是否区分字母的大小写。True不区别，Falsh区分
        Rep.pattern = SerchValue 
        Set search = Rep.Execute(content)
            If search.Count = 0 Then checkStr = False 
        Set search = Nothing 
    Set Rep = Nothing 
    If Err Then Call doError(Err.Description, "CheckStr 检查字符是否存在 函数出错，SerchValue=" & SerchValue) 
End Function 
'检查字符是否存在 (辅助)
Function regExp_CheckStr(content, SerchValue)
    regExp_CheckStr = checkStr(content, SerchValue) 
End Function 
'检查字符是否存在 (辅助)
Function existsStr(content, SerchValue)
    existsStr = checkStr(content, SerchValue) 
End Function 
'检查字符是否存在 (辅助)
Function regExp_Instr(content, SerchValue)
    regExp_Instr = checkStr(content, SerchValue) 
End Function 

'替换内容,不区分大小写
Function regExp_Replace(ByVal content, ByVal SetPattern, ByVal TStr)
    On Error Resume Next 
    Dim Rep 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.pattern = SetPattern                                                        '设置模式。
        Rep.IgnoreCase = True                                                           '设置是否区分字母的大小写。
        Rep.Global = True                                                               '设置全程性质。
        regExp_Replace = Rep.Replace(content, TStr) 
    Set Rep = Nothing 
    If Err Then Call doError(Err.Description, "RegExp_Replace 替换内容 函数出错，SetPattern=" & SetPattern & " |Str=" & TStr) 
End Function 
'替换内容自动加颜色,不区分大小写
Function replaceColor(content, SetPattern, TStr, color)
    If color <> "" Then TStr = "<font color=" & color & ">" & TStr & "</font>" 
    replaceColor = regExp_Replace(content, SetPattern, TStr) 
End Function
'删除HTML标签
Function delHtml(ByVal strHtml)
    On Error Resume Next 
    Dim Rep, strOutput
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.IgnoreCase = True 
        Rep.Global = True 
        Rep.pattern = "(<[a-zA-Z].*?>)|(<[\/][a-zA-Z].*?>)" 
        strOutput = Rep.Replace(strHtml, "") 
        delHtml = strOutput 
    Set Rep = Nothing 
    If Err Then Call doError(Err.Description, "DelHtml 删除HTML标签 函数出错，") 
End Function

'删除HTML标签 (辅助)
Function regExp_DelHtml(strHtml)
    regExp_DelHtml = delHtml(strHtml)
End Function 
'获得内容中链接与标题列表  getAUrlTitleList(Content,"网址")
Function getAUrlTitleList(ByVal content, ByVal SType)
    On Error Resume Next 
    Dim s, c, SetExecute, SetValue, Rep 
    c = ""
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.IgnoreCase = True 
        Rep.Global = True 
        Rep.pattern = "<a [^>]*href *= *[""']([^""'> ]+)[^>]*>([^<]+)" 
        Set SetExecute = Rep.Execute(content)
            For Each SetValue In SetExecute
                If SType = "链接" Or SType = "网址" Or SType = "0" Then
                    s = SetValue.SubMatches(0) 
                ElseIf SType = "标题" Or SType = "1" Then
                    s = SetValue.SubMatches(1) 
                Else
                    s = SetValue.SubMatches(0) & "$Array$" & SetValue.SubMatches(1) 
                End If 
                If InStr(vbCrLf & c & vbCrLf, vbCrLf & s & vbCrLf) = False Then
                    c = c & s & vbCrLf 
                End If
            Next 
        Set SetExecute = Nothing 
    Set Rep = Nothing
    If c <> "" Then c = Left(c, Len(c) - 2)
    getAUrlTitleList = c 
    If Err Then Call doError(Err.Description, "GetHrefUrlTitleList 获得内容中链接与标题列表 函数出错，C=" & c) 
End Function 
'获得链接中网址 (辅助)
Function getAURL(content)
    getAURL = getAUrlTitleList(content, 0) 
End Function 
'获得链接中标题 (辅助)
Function getATitle(content)
    getATitle = getAUrlTitleList(content, 1) 
End Function 
'获得链接中网址和标题 (辅助)
Function getAURLTitle(content)
    getAURLTitle = getAUrlTitleList(content, "链接和标题") 
End Function 


'获得内容中链接与标题列表  GetAUrlTitleList(Content,"网址")
Function getContentAUrlList(ByVal content)
    On Error Resume Next 
    Dim s, c, SetExecute, SetValue, Rep 
    c = "" 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.IgnoreCase = True 
        Rep.Global = True 
        Rep.pattern = "<a [^>]*href *= *[""']([^""'> ]+)[^>]*>" 
        Set SetExecute = Rep.Execute(content)
            For Each SetValue In SetExecute
                s = SetValue.SubMatches(0) 
                If InStr(vbCrLf & c & vbCrLf, vbCrLf & s & vbCrLf) = False Then
                    c = c & s & vbCrLf 
                End If 
            Next 
        Set SetExecute = Nothing 
    Set Rep = Nothing 
    If c <> "" Then c = Left(c, Len(c) - 2) 
    getContentAUrlList = c 
    If Err Then Call doError(Err.Description, "GetHrefUrlTitleList 获得内容中链接与标题列表 函数出错，C=" & c) 
End Function 


'获得图片中网址
Function regExp_GetImgUrlList(ByVal content, ByVal SType)
    On Error Resume Next 
    Dim SetExecute, SetValue, Rep, ImgType, s, c 
    c = "" 
    SType = LCase(SType) 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.IgnoreCase = True 
        Rep.Global = True 
        Rep.pattern = "<img *src *= *[""|'| ]*([^""'> ]+)" 
        Set SetExecute = Rep.Execute(content)
            For Each SetValue In SetExecute
                s = SetValue.SubMatches(0) 
                ImgType = Right(s, Len(s) - InStrRev(s, ".")) 
                If SType = ImgType Or SType = "全部" Or SType = "" Then
                    If InStr(vbCrLf & c & vbCrLf, vbCrLf & s & vbCrLf) = False Then
                        c = c & s & vbCrLf 
                    End If 
                End If 
            Next 
        Set SetExecute = Nothing 
    Set Rep = Nothing 
    If c <> "" Then c = Left(c, Len(c) - 2) 
    regExp_GetImgUrlList = c 
    If Err Then Call doError(Err.Description, "GetIMGUrlList 获得图片中网址 函数出错，c=" & c) 
End Function 
'获得图片中网址 (辅助)
Function getIMG(content)
    getIMG = regExp_GetImgUrlList(content, "全部") 
End Function 
'获得图片中JPG网址 (辅助)
Function getIMGJpg(content)
    getIMGJpg = regExp_GetImgUrlList(content, "JPG") 
End Function 
'获得图片中Gif网址 (辅助)
Function getIMGGif(content)
    getIMGGif = regExp_GetImgUrlList(content, "GIF") 
End Function 
'获得图片中Bmp网址 (辅助)
Function getIMGBmp(content)
    getIMGBmp = regExp_GetImgUrlList(content, "Bmp") 
End Function 
'获得ASP自定函数列表，配合获得ASP中自定函数(2013,9,30
Function getASPFunctionList(ByVal content)
    On Error Resume Next 
    Dim SetExecute, SetValue, Rep, ImgType, s, c 
    c = "" 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.IgnoreCase = True 
        Rep.Global = True 
        Rep.pattern = "(sub|function) .*?\)" 
        'Rep.Pattern = "\s(sub|function)( .*?)(\(.*)|\s(sub|function)( .*?[ |    ])"            '更精准20150722
        Set SetExecute = Rep.Execute(content)
            For Each SetValue In SetExecute
                c = c & SetValue & vbCrLf 
            Next 
        Set SetExecute = Nothing 
    Set Rep = Nothing 
    If c <> "" Then c = Left(c, Len(c) - 2) 
    getASPFunctionList = c 
    If Err Then Call doError(Err.Description, "GetASPFunctionList 获得ASP函数列表 函数出错，c=" & c) 
End Function 
'获得ASP自定函数列表，配合获得ASP中自定函数(2013,9,30
Function getASPDimFunction(ByVal content, FunValue, FunName, FunDim)
    'On Error Resume Next
    content = Replace(content, "()", "( )")                                         '为防止截取不到函数
    Dim SetExecute, SetValue, Rep 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.IgnoreCase = True 
        Rep.Global = True 
        'Rep.Pattern = "(sub|function) .*?\)"
        content = regExp_Replace(content, "ByVal ", "") 
        Rep.pattern = "(sub|function) *([^\(]+) *\(([^\)]+)\)" 
        Set SetExecute = Rep.Execute(content)
            For Each SetValue In SetExecute
                '{字符可以判断是JS函数
                If InStr(FunValue, "{") = False Then
                    FunValue = FunValue & Trim(SetValue) & vbCrLf 
                    FunName = FunName & Trim(SetValue.SubMatches(1)) & vbCrLf 
                    FunDim = FunDim & Replace(SetValue.SubMatches(2), " ", "") & vbCrLf 
                End If 
            Next 
        Set SetExecute = Nothing 
    Set Rep = Nothing 
    'If FunDim <> "" Then FunDim = Left(FunDim, Len(FunDim)-2)                '不让用太怪了
    getASPDimFunction = FunValue 
    If Err Then Call doError(Err.Description, "GetASPFunctionList 获得ASP函数列表 函数出错，") 
End Function 
'Response.Write(StrLength("中国121aa"))
'正则表达式获得字符长度 中文二个字符
Function strLength(str)
    Dim Rep, lens, i 
    lens = 0 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.Global = True 
        Rep.IgnoreCase = True 
        Rep.pattern = "[\u4E00-\u9FA5\uF900-\uFA2D]" 
        For Each i In Rep.Execute(str)
            lens = lens + 1 
        Next 
    Set Rep = Nothing 
    lens = lens + Len(str) 
    strLength = lens 
End Function 
'正则表达式获得字符长度 中文二个字符
Function regExpGetLenght(str)
    regExpGetLenght = strLength(str) 
End Function 
'正则表达式获得字符长度 中文二个字符
Function stringLength(str)
    stringLength = strLength(str) 
End Function 
'获得HTML中Css
Function getCssHref(ByVal content)
    On Error Resume Next 
    Dim SetExecute, SetValue, Rep, ImgType, s, c 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.IgnoreCase = True 
        Rep.Global = True 
        Rep.pattern = "<link .*?href *= *[""|'| ]*([^""'>]+)" 
        Set SetExecute = Rep.Execute(content)
            For Each SetValue In SetExecute
                If InStr(LCase(SetValue), "stylesheet") > 0 Then
                    s = SetValue.SubMatches(0) 
                    c = c & s & vbCrLf 
                End If 
            Next 
        Set SetExecute = Nothing 
    Set Rep = Nothing 
    If c <> "" Then c = Left(c, Len(c) - 2) 
    getCssHref = c 
    If Err Then Call doError(Err.Description, "GetCssHref 获得HTML中Css 函数出错，C =" & c) 
End Function 
'获得HTML中Js
Function getJsSrc(ByVal content)
    On Error Resume Next 
    Dim SetExecute, SetValue, Rep, ImgType, s, c 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.IgnoreCase = True 
        Rep.Global = True 
        Rep.pattern = "<script .*?src *= *[""|'| ]*([^""'>]+)" 
        Set SetExecute = Rep.Execute(content)
            For Each SetValue In SetExecute
                s = SetValue.SubMatches(0) 
                c = c & s & vbCrLf 
            Next 
        Set SetExecute = Nothing 
    Set Rep = Nothing 
    If c <> "" Then c = Left(c, Len(c) - 2) 
    getJsSrc = c 
    If Err Then Call doError(Err.Description, "GetJsSrc 获得HTML中Js 函数出错，C =" & c) 
End Function 
'删除HTML中<scirpt
Function delHTMLScript(ByVal strHtml)
    On Error Resume Next 
    Dim Rep, strOutput 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.IgnoreCase = True 
        Rep.Global = True 
        Rep.pattern = "<script.*</script>" 
        strOutput = Rep.Replace(strHtml, "") 
        delHTMLScript = strOutput 
    Set Rep = Nothing 
    If Err Then Call doError(Err.Description, "删除HTML中<scirpt 函数出错，") 
End Function 
'获得HTML中keywords
Function getMeta(ByVal content, SType, nOK)
    On Error Resume Next 
    Dim SetExecute, SetValue, Rep, ImgType, s, c 
    c = "" 
    SType = LCase(SType) 
    nOK = 0 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.IgnoreCase = True 
        Rep.Global = True 
        Rep.pattern = "<meta .*?content *= *[""|'| ]*([^""'>]+)" 
        Set SetExecute = Rep.Execute(content)
            For Each SetValue In SetExecute
                If InStr(LCase(SetValue), SType) > 0 Then
                    s = SetValue.SubMatches(0) 
                    c = c & s & vbCrLf 
                    nOK = nOK + 1 
                End If 
            Next 
        Set SetExecute = Nothing 
    Set Rep = Nothing 
    If c <> "" Then c = Left(c, Len(c) - 2) 
    getMeta = c 
    If Err Then Call doError(Err.Description, "GetKeyWords 获得HTML中keywords 函数出错，C =" & c) 
End Function 
'去除HTML注释
Function delHtmlNote(ContentStr)
    Dim ClsTempLoseStr, regex 
    ClsTempLoseStr = CStr(ContentStr) 
    Set regex = CreateObject("VBscript.RegExp")
        regex.pattern = "<!--\/*[^<>]*-->" 
        regex.IgnoreCase = True 
        regex.Global = True 
        ClsTempLoseStr = regex.Replace(ClsTempLoseStr, "") 
        delHtmlNote = ClsTempLoseStr 
    Set regex = Nothing 
End Function 
'删除自己定义的HTML注释
Function delHtmlMyNote(ContentStr)
    Dim ClsTempLoseStr, regex 
    ClsTempLoseStr = CStr(ContentStr) 
    Set regex = CreateObject("VBscript.RegExp")
        regex.pattern = "<!--#\/*[^<>]*#-->" 
        regex.IgnoreCase = True 
        regex.Global = True 
        ClsTempLoseStr = regex.Replace(ClsTempLoseStr, "") 
        delHtmlMyNote = ClsTempLoseStr 
    Set regex = Nothing 
End Function 
'获得内容中图片列表
Function getStrImgList(ByVal content)
    content = Replace(content, "'", """") 
    Dim result, Result1, m, n, regex 
    result = "" : Result1 = "" 
    Set regex = CreateObject("VBscript.RegExp")
        'Set regEx = CreateObject("VBscript.RegExp")
        regex.IgnoreCase = True 
        regex.Global = True 
        regex.pattern = "<img [^>]*src=""([^"">]+)[^>]+>" 
        Set m = regex.Execute(content)
            For Each n In m
                result = result & n & "|" 
                Result1 = Result1 & n.SubMatches(0) & "|" 
            Next 
        Set m = Nothing 
    Set regex = Nothing 
    If result <> "" Then
        result = Left(result, Len(result) - 1) 
        Result1 = Left(Result1, Len(Result1) - 1) 
    End If 
    'result=split(result,"|") '存储<img>
    'result1=split(result1,"|") '存储图像地址
    getStrImgList = Result1 
End Function 
'获得邮箱列表
Function getEMailList(content)                                          'patrn:需要查找的字符 strng:被查找的字符串
    Dim regex, Match, Matches, RetStr                                               '创建变量。
    Set regex = CreateObject("VBscript.RegExp")                                     '创建正则表达式。
        regex.pattern = "\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"    '设置模式。"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"
        regex.IgnoreCase = True                                                         '设置是否区分大小写。
        regex.Global = True                                                             '设置全程匹配。
        RetStr = "" 
        Set Matches = regex.Execute(content)                                            '执行搜索。
            For Each Match In Matches                                      '循环遍历Matches集合。
                If InStr("," & RetStr & ",", "," & Match.Value & ",") = 0 Then
                    RetStr = RetStr & Match.Value & "," 
                End If 
            Next 
        Set Matches = Nothing 
    Set regex = Nothing 
    If RetStr <> "" Then RetStr = Left(RetStr, Len(RetStr) - 1) 
    getEMailList = RetStr 
End Function 
'将html内容代码里的标签大写转换成小写
Function oflink_lcasetag(content)
    Dim regex, Match, Matches                                                       '建立变量。
    Set regex = CreateObject("VBscript.RegExp")                                     '建立正则表达式。
        regex.pattern = "<.+?\>"                                                     '设置模式。
        regex.IgnoreCase = True                                                         '设置是否区分字符大小写。
        regex.Global = True                                                             '设置全局可用性。
        Set Matches = regex.Execute(content)                                            '执行搜索。
            content = content 
            For Each Match In Matches                                      '遍历匹配集合。
                content = Replace(content, Match.Value, LCase(Match.Value)) 
            Next 
            oflink_lcasetag = content 
        Set Matches = Nothing 
    Set regex = Nothing 
End Function 
'==================================================
'函数名：ScriptHtml
'作  用：过滤html标记
'参  数：ConStr ------ 要过滤的字符串   1为删除开始标记  3为删除前后两标记 2为测试无效果
'==================================================
Function scriptHtml(ByVal ConStr, TagName, FType)
    Dim re 
    Set re = CreateObject("VBscript.RegExp")
        re.IgnoreCase = True 
        re.Global = True 
        Select Case FType
            Case 1
                re.pattern = "<" & TagName & "([^>])*>" 
                ConStr = re.Replace(ConStr, "") 
            Case 2
                re.pattern = "<" & TagName & "([^>])*>.*?</" & TagName & "([^>])*>" 
                ConStr = re.Replace(ConStr, "") 
            Case 3
                re.pattern = "<" & TagName & "([^>])*>" 
                ConStr = re.Replace(ConStr, "") 
                re.pattern = "</" & TagName & "([^>])*>" 
                ConStr = re.Replace(ConStr, "") 
        End Select
        scriptHtml = ConStr 
    Set re = Nothing 
End Function 


'截取内容，可指定3个匹配参数 [$Array$]为行 [$Split$]为列    2014 12 10
'例子一
'Content = " tablename='ListMenu' bigclassname='系统管理' smallclassname='Robots在线修改' sort='3' url='Robots.Asp' display='1' "
'Call Rw(GetRegExp(" (.*?)\=.*?'(.*?)'", Content, 0,1,""))'
Function getRegExp(Patrn, strng, nSplit1, nSplit2, nSplit3)
    Dim regex, Match, Matches, c, SplitYes, s 
    c = "" 
    If nSplit1 <> "" Then nSplit1 = CInt(nSplit1)                                   '当前分割数
    If nSplit2 <> "" Then nSplit2 = CInt(nSplit2)                                   '当前分割数
    If nSplit3 <> "" Then nSplit3 = CInt(nSplit3)                                   '当前分割数
    SplitYes = False                                                                '分割是否为
    If nSplit1 <> "" Or nSplit2 <> "" Or nSplit3 <> "" Then SplitYes = True 
    Set regex = CreateObject("VBscript.RegExp")                                     '建立正则表达式。
        regex.pattern = Patrn                                                           '设置模式。
        regex.IgnoreCase = True                                                         '设置是否区分字符大小写。
        regex.Global = True                                                             '设置全局可用性。
        Set Matches = regex.Execute(strng)                                              '执行搜索。
            For Each Match In Matches                                      '遍历匹配集合。
                If SplitYes = True Then
                    s = "" 
                    If nSplit1 <> "" Then
                        s = s & Match.SubMatches(nSplit1) 
                    End If 
                    If nSplit2 <> "" Then
                        If s <> "" Then s = s & "[$Split$]" 
                        s = s & Match.SubMatches(nSplit2) 
                    End If 
                    If nSplit3 <> "" Then
                        If s <> "" Then s = s & "[$Split$]" 
                        s = s & Match.SubMatches(nSplit3) 
                    End If 
                    If s <> "" Then c = c & s & "[$Array$]" 
                Else
                    c = c & Match & "[$Array$]" 
                End If 
            Next 
            If c <> "" Then c = Left(c, Len(c) - 9) 
            getRegExp = c 
End Function



'获得内容里Img与Js中链接  SType=1为不重复   20150126
Function getImgJsUrl(content, SType)
    Dim s, c, splStr, url, UrlList 
    SType = LCase(SType)                                                            '类型转字小写并且是字符类型
    c = RegExpGetStr("<[img|script][^<>]*src[^<>]*[\/]?>", content, 2) 
    splStr = Split(c, "[$Array$]") 
    UrlList = "" 
    For Each s In splStr
        url = RegExpGetStr("src=[""|']?([^""' ]*)([""|']?).*[\/]?>", s, 1) 
        If SType = "1" Or SType = "不重复" Then
            If InStr(vbCrLf & UrlList & vbCrLf, vbCrLf & url & vbCrLf) Then
                url = "" 
            End If 
        End If 
        If url <> "" Then
            UrlList = UrlList & url & vbCrLf 
        End If 
    Next 
    If UrlList <> "" Then UrlList = Left(UrlList, Len(UrlList) - 2) 
    getImgJsUrl = UrlList 
End Function 
'正则表达式处理20150126
'提取网站中标题、关键词、描述
'WebTitle = RegExpGetStr("<TITLE>(.*)</TITLE>", Content, 1)                    '第一种
'WebTitle = RegExpGetStr("<TITLE>([^<>]*)</TITLE>", Content, 1)                    '第二种 完整
'WebDes = RegExpGetStr("<meta[^<>]*description[^<>]*[\/]?>", Content, 0)
'WebDes = RegExpGetStr("content=[""|']?([^""' ]*)([""|']?).*[\/]?>", WebDes, 1)
'WebKey = RegExpGetStr("<meta[^<>]*keywords[^<>]*[\/]?>", Content, 0)
'WebKey = RegExpGetStr("content=[""|']?([^""' ]*)([""|']?).*[\/]?>", WebKey, 1)
Function regExpGetStr(patternStr, content, SType)
    Dim re, RemoteFileurl, RemoteFile, c
	SType=cstr(SType)
    Set re = CreateObject("VBscript.RegExp")
        re.IgnoreCase = True 
        re.Global = True 
        re.pattern = patternStr 
        Set RemoteFile = re.Execute(content)
            For Each RemoteFileurl In RemoteFile
                If SType = "1" Then
                    RemoteFileurl = re.Replace(RemoteFileurl, "$1") 
                ElseIf SType = "2" Then
                    c = c & RemoteFileurl & "[$Array$]" 
                End If 
                regExpGetStr = RemoteFileurl 
            Next 
        Set RemoteFile = Nothing 
    Set re = Nothing 
    If SType = "2" Then
        If c <> "" Then c = Left(c, Len(c) - 9) 
        regExpGetStr = c 
    End If 
End Function 

'获得IP地址
Function newGetPingIP(HttpUrl)
    Dim CmdPath, content, splStr 
    Dim RemoteUrl, url, startStr, endStr 
    newGetPingIP = "" 
    CmdPath = "/备份/cmd.exe" 
    'HttpUrl="www.intcochina.com.cn"
    content = GetPingInfo(CmdPath, HttpUrl) 
    splStr = Split(content, "Request timed out.") 
    'Call Rw(Content & "," & Ubound(SplStr))
    If UBound(splStr) >= 4 Then
        url = Trim(Replace(Replace(Replace(HttpUrl, "http://", ""), "www.", ""), "/", "")) 
        RemoteUrl = "http://whois.www.net.cn/whois/api_webinfo?host=" & url & "&_=" & getRnd(9) 
        content = GetHttpPage(RemoteUrl, "utf-8") 
        startStr = """ip"":""" : endStr = """" 
        If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
            newGetPingIP = StrCut(content, startStr, endStr, 2) 
        End If 
    Else
        startStr = " [" : endStr = "]" 
        If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
            newGetPingIP = StrCut(content, startStr, endStr, 2) 
        Else
            newGetPingIP = "NO" 
        'Call Eerr("S",S)
        End If 
    End If 
End Function 
'用Cmd的Ping获得网址IP 用法：Response.Write(GetPingIP("/cmd.exe", "www.shuyate.cn"))
Function getPingIP(CmdPath, HttpUrl)
    Dim s, startStr, endStr 
    'StartStr = HttpUrl & " [" : EndStr="]"
    startStr = " [" : endStr = "]" 
    s = GetPingInfo(CmdPath, HttpUrl) 
    If InStr(s, startStr) > 0 And InStr(s, endStr) > 0 Then
        getPingIP = StrCut(s, startStr, endStr, 2) 
    Else
        getPingIP = "NO" 
    'Call Eerr("S",S)
    End If 
End Function 
'获得ping信息  getPingInfo("C:\Windows\winsxs\wow64_microsoft-windows-commandprompt_31bf3856ad364e35_6.1.7601.17514_none_f387767e655cd5ab\cmd.exe", " 10.10.10.2")
Function getPingInfo(CmdPath, HttpUrl)
    Dim wS, dD 
    Set wS = CreateObject("WScript.Shell")
        Call HandlePath(CmdPath)                                                        '获得完整路径
        Set dD = wS.exec(CmdPath & " /c ping " & HttpUrl)
            getPingInfo = dD.stdout.readall 
End Function
'显示真与假 待用
Function regExpTest(str, s_Pattern)
    Dim re 
    Set re = CreateObject("VBscript.RegExp")
        re.IgnoreCase = True 
        re.Global = True 
        re.pattern = s_Pattern 
        regExpTest = re.Test(str) 
    Set re = Nothing 
End Function 

'正则表达式获得内容 给asp转php里用
Function getRegExpStr(ByVal content, ByVal PatternStr, ByVal ReplaceStr)
    On Error Resume Next 
    Dim s, c, SetExecute, SetValue, Rep, newStr 
    c = "" 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.IgnoreCase = True 
        Rep.Global = True 
        Rep.pattern = PatternStr 
        Set SetExecute = Rep.Execute(content)
            For Each SetValue In SetExecute
                newStr = Replace(ReplaceStr, "\S1\", SetValue.SubMatches(0)) 
                newStr = Replace(newStr, "\S2\", SetValue.SubMatches(1)) 
                newStr = Replace(newStr, "\S3\", SetValue.SubMatches(2)) 
                newStr = Replace(newStr, "\S4\", SetValue.SubMatches(3)) 
                newStr = Replace(newStr, "\S5\", SetValue.SubMatches(4)) 
                content = Replace(content, SetValue, newStr) 
            Next 
        Set SetExecute = Nothing 
    Set Rep = Nothing 
    getRegExpStr = content 
End Function 



'替换CSS目录 （20150814）  call rwend(regExpReplaceCssDir("<link rel=""stylesheet"" href=""Images/jquery.mobile-1.4.5.min.css"">", "css/"))
Function regExpReplaceCssDir(ByVal content, addToDir)
    Dim setExecute, setValue, rep, replaceStr 
    Dim filePath, fileName 
    Set rep = CreateObject("VBscript.RegExp")
        rep.ignoreCase = True 
        rep.global = True 
        rep.pattern = "<link .*?href *= *[""|'| ]*([^""'>]+)" 
        Set setExecute = rep.Execute(content)
            For Each setValue In setExecute
                If InStr(LCase(setValue), "stylesheet") > 0 Then
                    filePath = setValue.subMatches(0) 
                    fileName = getStrFileName(filePath) 
                    replaceStr = Replace(setValue, filePath, addToDir & fileName) 
                    content = Replace(content, setValue, replaceStr) 
                End If 
            Next 
        Set setExecute = Nothing 
    Set rep = Nothing 
    regExpReplaceCssDir = content 
End Function 


'删除空行
Function deleteNullRow(ByVal content)
    deleteNullRow = regExp_Replace(content, "\r\n\s*\r\n", "")                      '删除空行
End Function 

'空链接加默认名称 20160405  待测试
function nullLinkAddDefaultName(content)
	dim rep
	Set rep = CreateObject("VBscript.RegExp")
	rep.IgnoreCase=true 
	rep.Global=True 
	rep.Pattern="<a [^>]*href *= *[""'][""'][^>]*>" 
	content=rep.Replace(content,"<a href='index.html'>") 
	Set rep=Nothing 
	nullLinkAddDefaultName=content
end function
%>  

