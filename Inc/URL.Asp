<% 
'URL 网址处理 (2013,9,27)

'  byref PubAHrefList, byref PubATitleList   为为什么要加上byref  是因为在转PHP时需要这个判断，请明它是地址传递 是关联变量

'NOVBNet start
'使用手册
'call echo("'获得当前网址第一种（getUrl）",getUrl())
'call echo("'获得当前网址第二种（getThisUrl）",getThisUrl())
'call echo("'获得当前网址无参数（getThisUrlNoParam）",getThisUrlNoParam())
'call echo("'获得来访网址（getGoToUrl）",getGoToUrl())
'call echo("'获得来访网址无参数（getGoToUrlNoParam）",getGoToUrlNoParam())
'call echo("'获得来访网址无文件名称（getGoToUrlNoFileName）",getGoToUrlNoFileName())
'call echo("'域名（webDoMain）",webDoMain())
'call echo("'主机（host）",host())
'call echo("'获得当前网址加参数（getUrlAddToParam）",getUrlAddToParam(GetUrl(), "PageSize=20", "replace"))


'获得当前网址第一种（getUrl）：http://127.0.0.1/atemp.asp?act=1
'获得当前网址第二种（getThisUrl）：http://127.0.0.1/atemp.asp?act=1
'获得当前网址无参数（getThisUrlNoParam）：http://127.0.0.1/atemp.asp
'获得来访网址（getGoToUrl）：http://127.0.0.1/5.asp?act=5&aa=aa&bb=bb
'获得来访网址无参数（getGoToUrlNoParam）：http://127.0.0.1/5.asp
'获得来访网址无文件名称（getGoToUrlNoFileName）：http://127.0.0.1/
'域名（webDoMain）：http://127.0.0.1
'主机（host）：http://127.0.0.1/
'获得当前网址加参数（GetUrlAddToParam）：http://127.0.0.1/atemp.asp?PageSize=20&act=1
'getThisUrlFileName()    : 4.asp
'getThisUrlFileParam()    : 4.asp?act=11

'Url = getUrlAddToParam("http://www.baidu.com/?a=1&b=2&c=3","?a=11&b=22&c=333","")        'http://www.baidu.com/?a=1&b=2&c=3
'Url = getUrlAddToParam("http://www.baidu.com/?a=1&b=2&c=3","?a=11&b=22&c=333","replace")        'http://www.baidu.com/?a=11&b=22&c=333

'获得当前网后的文件名与参数
Function getThisUrlFileParam()
    Dim url 
    url = getUrl() 
    url = Mid(url, InStrRev(url, "/") + 1) 
    getThisUrlFileParam = url 
End Function 

'获得当前网址无参数
Function getThisUrlNoParam()
    Dim httpType 
    If LCase(Request.ServerVariables("HTTPS")) = "off" Then
        httpType = "http://" 
    Else
        httpType = "https://" 
    End If 
    getThisUrlNoParam = httpType & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME") 
End Function 

'获得传过来网址
Function getGoToUrl()
    getGoToUrl = Request.ServerVariables("HTTP_REFERER") 
End Function

'获得传过来网址 无参数
Function getGoToUrlNoParam()
    Dim url 
    url = getGoToUrl() 
    If InStr(url, "?") > 0 Then
        url = Mid(url, 1, InStr(url, "?") - 1) 
    End If 
    getGoToUrlNoParam = url 
End Function 

'获得传过来网址 无文件名称
Function getGoToUrlNoFileName()
    Dim url 
    url = getGoToUrl() 
    If Right(url, 1) <> "/" Then
        If InStrRev(url, "/") > 0 Then
            url = Mid(getGoToUrlNoParam(), 1, InStrRev(getGoToUrlNoParam(), "/")) 
        End If 
    End If 
    If Right(url, 1) <> "/" Then url = url & "/" 
    getGoToUrlNoFileName = url 
End Function
'移除网址文件名部分
function remoteUrlFileName(url)
    If Right(url, 1) <> "/" Then
        If InStrRev(url, "/") > 0 Then
            url = Mid(url, 1, InStrRev(url, "/")) 
        End If 
    End If 
	remoteUrlFileName=url
end function
'移除网址参数部分
function remoteUrlParam(url)
    If Right(url, 1) <> "?" Then
        If InStrRev(url, "?") > 0 Then
            url = Mid(url, 1, InStrRev(url, "?")-1) 
        End If 
    End If 
	remoteUrlParam=url
end function
'获得网址目录部分
function getUrlDir(url)
	url=trim(url)
	if left(url,7)="http://" then
		url = mid(url,8)
	end if
	if instr(url,"/")>0 then
		url=Mid(url,instr(url,"/")+1)
	end if
	getUrlDir=url
end function


'获取客户端IP地址第二种
Function getIP2()
    On Error Resume Next 
    Dim x, y, addr 
    x = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
    y = Request.ServerVariables("REMOTE_ADDR") 
    addr = IIF(isNul(x) Or LCase(x) = "unknown", y, x) 
    If InStr(addr, ".") = 0 Then addr = "0.0.0.0" 
    getIP2 = addr 
End Function 

'获取IP地址 别人写得好像很专业一样 很全
Function getIP()
    On Error Resume Next 
    Dim strIPAddr 
    If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" Or InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then
        strIPAddr = Request.ServerVariables("REMOTE_ADDR") 
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then
        strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") - 1) 
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then
        strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") - 1) 
    Else
        strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
    End If 
    getIP = Trim(Mid(strIPAddr, 1, 30)) 
End Function 

'获得服务器IP
Function getServicerIP()
    getServicerIP = Request.ServerVariables("LOCAL_ADDR") 
End Function 

'获得服务器IP   辅助
Function getRemoteIP()
    getRemoteIP = getServicerIP() 
End Function 


'NOVBNet end

'处理网址\转/  '待完善
Function handleHttpUrl(httpUrl)
    Dim headStr, url 
    If isNul(httpUrl) Then Exit Function                                            '为空则退出
    httpUrl = Replace(Trim(httpUrl), "\", "/") 
    headStr = Mid(httpUrl, 1, InStr(httpUrl, ":") + 2) 
    httpUrl = Mid(httpUrl, InStr(httpUrl, ":") + 3) 
    httpUrl = Replace(httpUrl, "http://", "【|http|】")
    httpUrl = Replace(httpUrl, "https://", "【|https|】")
    httpUrl = Replace(httpUrl, "ftp://", "【|ftp|】")

    While InStr(httpUrl, "//") > 0
        httpUrl = Replace(httpUrl, "//", "/") 
    Wend 
    httpUrl = Replace(httpUrl, "【|http|】", "http://") 
    httpUrl = Replace(httpUrl, "【|https|】", "https://") 
    httpUrl = Replace(httpUrl, "【|ftp|】", "ftp://") 

    url = headStr & httpUrl 
    '/"http://www.qibosoft.com/images/qibosoft/loading.gif/"
    If Left(url, 2) = "/""" Then
        url = Mid(url, 3) 
    End If 
    If Right(url, 2) = "/""" Then
        url = Mid(url, 1, Len(url) - 2) 
    End If 
    handleHttpUrl = url 
End Function 

'处理文件/转\   使用了While判断，再完善
Function handleFileUrl(fileUrl)
    fileUrl = Replace(fileUrl, "/", "\") 
    Dim i 
    For i = 1 To 99
        fileUrl = Replace(fileUrl, "\\", "\") 

        If InStr(fileUrl, "\\") = False Then
            Exit For 
        End If 
    Next 
    handleFileUrl = fileUrl 
End Function 

'处理网址完善性
Function handleUrlComplete(httpUrl)
    Dim lastStr 
    handleUrlComplete = httpUrl 
    If InStr(httpUrl, "?") > 0 Then Exit Function                                   '有?符号则退出
    '网址最后没有/  判断如果为域名 则在最后加上/退出
    If Right(httpUrl, 1) <> "/" Then
        If httpUrl & "/" = getWebSite(httpUrl) Then
            handleUrlComplete = httpUrl & "/" 
            Exit Function 
        End If 
    End If 
    lastStr = Mid(httpUrl, InStrRev(httpUrl, "/") + 1) 
    If lastStr <> "" And InStr(lastStr, ".") = False Then
        handleUrlComplete = httpUrl & "/" 
    End If 
End Function 

'给网址添加域名
Function urlAddHttpUrl(httpUrl, url)
    httpUrl = Replace(httpUrl, "\", "/") 
    url = handleHttpUrl(url) 
    If InStr(LCase(url), "http://") = 0 And InStr(LCase(url), "www.") = 0 Then
        If Right(httpUrl, 1) = "/" And Left(url, 1) = "/" Then
            url = httpUrl & Mid(url, 2) 
        ElseIf Right(httpUrl, 1) <> "/" And Left(url, 1) <> "/" Then
            url = httpUrl & "/" & url 
        Else
            url = httpUrl & url 
        End If 
    End If 
    urlAddHttpUrl = url 
End Function 
'获得主机端口号
Function getPort()
    Dim port 
    port = CStr(Request.ServerVariables("SERVER_PORT")) 
    If port <> "80" And port <> "8080" Then
        port = ":" & port 
    Else
        port = "" 
    End If 
    getPort = port 
End Function 

'获得域名
Function webDoMain()
    webDoMain = "http://" & Request.ServerVariables("SERVER_NAME") & getPort() 
End Function 

'获得当前域名
Function host()
    host = "http://" & Request.ServerVariables("HTTP_HOST") & "/" 
End Function 

'获得当前域名 (辅助)
Function httpHost()
    httpHost = host() 
End Function 

'获得当前域名 (辅助)
Function getHost()
    getHost = host() 
End Function 

'网得当前网址
Function getUrl()
    getUrl = getThisUrl() 
'GetUrl = WebDoMain() & Request.ServerVariables("SCRIPT_NAME") & Request.ServerVariables("QUERY_STRING")
End Function 

'获得当前带参数网址
Function getThisUrl()
    Dim url 
    'vbdel start
    url =Request.ServerVariables("server_name") 
    'PHP版上面直接获得端口
    If InStr(url, ":") = False Then
        url = url & getPort() 
    End If 
    url = url & Request.ServerVariables("script_name") 
    If Request.ServerVariables("QUERY_STRING") <> "" Then url = url & "?" & Request.ServerVariables("QUERY_STRING") 
    'vbdel end
    getThisUrl =  "http://" & url 
End Function 
'获得当前无文件名称网址
Function getThisUrlNoFileName()
    Dim url 
    url = getThisUrl() 
    If Right(url, 1) <> "/" Then
        If InStrRev(url, "/") > 0 Then
            url = Mid(url, 1, InStrRev(url, "/")) 
        End If 
    End If 
    getThisUrlNoFileName = url 
End Function 

'获得网址域名名称 http://www.aaa.bb.mywebname.com/   mywebname
Function getWebSiteName(httpUrl)
    Dim url, splStr, s, domainName 
    url = getWebsite(httpUrl) 
    url = Replace(url, "://", "://.") 
    splStr = Split(url, ".") 
    For Each s In splStr
        If InStr(s, "/") = False And s <> "" Then
            If Len(s) >= 4 Then
                domainName = s 
            End If 
        End If 
    Next 
    getWebSiteName = domainName 
End Function 

'分析网址 获得域名
Function getWebSite(ByVal httpUrl)
    'On Error Resume Next
    '新版获得域名方法
    Dim url, tempHttpUrl, is_WebSite,httpHead
    tempHttpUrl = httpUrl 
    url = Trim(LCase(Replace(httpUrl, "?", "/"))) 
    url = Replace(Replace(url, "\", "/"), "http://", "") 
    If InStr(url, "/") > 0 Then url = Mid(url, 1, InStr(url, "/") - 1) 
    url = "http://" & url & "/" 
    Dim splStr, s, c 
    httpUrl = Replace(LCase(httpUrl), "http://", "")
	httpHead="http://"
	'增加了https://这种安全请求方式  20160526
	if instr(LCase(httpUrl), "https://")>0 then
		httpUrl = Replace(LCase(httpUrl), "https://", "")
		httpHead="https://"
	end if
	'删除/后台的值20160526
	if instr(httpUrl,"/")>0 then
		httpUrl=mid(httpUrl,1,instr(httpUrl,"/")-1)
	end if
    If InStr(httpUrl, "?") > 0 Then httpUrl = Mid(httpUrl, 1, InStr(httpUrl, "?") - 1) 
    If Left(httpUrl, 9) = "localhost" Then
        If InStr(httpUrl, "/") > 0 Then
            httpUrl = Mid(httpUrl, 1, InStr(httpUrl, "/") - 1) 
        Else
            httpUrl = "localhost" 
        End If 
    ElseIf Left(httpUrl, 8) = "192.168." Or Left(httpUrl, 9) = "127.0.0.1" Then
        httpUrl = httpUrl & "/" 
        httpUrl = "http://" & Mid(httpUrl, 1, InStr(httpUrl, "/") - 1) & "/" 
        getWebSite = httpUrl : Exit Function 
    Else
        splStr = Split(httpUrl, ".") 
        If(InStr(httpUrl, "www.") > 0 And UBound(splStr) >= 2) Or UBound(splStr) >= 1 Then
            If InStr(httpUrl, "/") > 0 Then
                s = Mid(httpUrl, 1, InStr(httpUrl, "/") - 1) 
                If s = getDianNumb(s) Then
                    httpUrl = s 
                End If 
            End If 
        Else
            httpUrl = ""                                                                    '没有则为空
        End If 
    End If 
    is_WebSite = False                                                              '是域名为假
    c = ".com.hk|.sh.cn|.com.cn|.net.cn|.org.cn" 
    c = c & "|.com|.net|.org|.tv|.cc|.info|.cn|.tw|:81|:99|.biz|.mobi|.hk|.us|.la|.gl|.in" 
    splStr = Split(c, "|") 
    For Each s In splStr
        If s <> "" Then
            If InStr(httpUrl, s) > 0 Then
                httpUrl = httpHead & Left(httpUrl, InStr(httpUrl, s) + Len(s) - 1) & "/" : is_WebSite = True : Exit For 
            End If 
        End If 
    Next 
    getWebSite = Left(httpUrl, 255)                                                 '域名不存在，则只截取255个字符
    'GetWebSite = ""                        '域名不存在则为空 20150104
    If getWebSite = "http:///" Then getWebSite = ""                                 '没有找到域名
    If is_WebSite = False Then
        getWebSite = "" 
    End If 

End Function 
'检测是否为网址
Function checkUrl(url)
    checkUrl = IIF(getWebSite(url) = "", False, True) 
End Function 
'检测是否为网址
Function checkHttpUrl(url)
    checkHttpUrl = checkUrl(url) 
End Function 
'检测是否为网址
Function isUrl(url)
    isUrl = checkUrl(url) 
End Function 
'检测是否为网址
Function isHttpUrl(url)
    isHttpUrl = checkUrl(url) 
End Function 

'完整网址
Function fullHttpUrl(ByVal httpUrl, ByVal url)
    'On Error Resume Next
    Err.Clear                                                                       '清除错误 要不然会报错
    Dim rootUrl, thisUrl, splStr, i, s, c, parentUrl, parentParentUrl, parentParentParentUrl, rootWebSite, thisWebSite, handleYes 
    httpUrl = pHPTrim(httpUrl)                                                      '清除两边空格
    url = pHPTrim(url)                                                              '清除两边空格

    If url = "" Then fullHttpUrl = "" : Exit Function                               '网址为空退出(20150805)
    If Trim(httpUrl) = "" Then fullHttpUrl = url : Exit Function                    '主网址为空退出 返回网址
    httpUrl = Replace(httpUrl, "\", "/")                                            '把网址\转/符号
    url = Replace(url, "\", "/")                                                    '把网址\转/符号

    '网址前两个字符为//则退出
    If Left(url, 2) = "//" Then
        fullHttpUrl = "http:" & url 
        Exit Function 
    End If 


    rootUrl = getWebSite(httpUrl)                                                   '主域名，也就是主网址
    rootWebSite = rootUrl 
    thisWebSite = getWebSite(url) 
    If Right(rootUrl, 1) = "/" Then rootUrl = Left(rootUrl, Len(rootUrl) - 1) 
    thisUrl = Left(httpUrl, InStrRev(httpUrl, "/"))                                 '当前网址
    splStr = Split(httpUrl, "/") 
    For i = 0 To UBound(splStr)
        If i + 1 = UBound(splStr) Then parentUrl = c 
        If i + 2 = UBound(splStr) Then parentParentUrl = c 
        If i + 3 = UBound(splStr) Then parentParentParentUrl = c 
        s = splStr(i) 
        c = c & s & "/" 
    Next 
    url = Trim(url)                                                                 '去除网址左右空格
    handleYes = False                                                               '操作为假
    If url <> "" And InStr(Left(url, 10), "www.") = False And InStr(Left(url, 10), "http://") = False And InStr(Left(url, 10), "https://") = False Then
        handleYes = True 
        If rootWebSite <> thisWebSite Then
            If rootWebSite = Replace(thisWebSite, "http://", "http://www.") Then
                handleYes = False 
                If InStr(LCase(url), "http://") > 0 Then
                    url = "http://www." & Right(url, Len(url) - 7) 
                Else
                    url = "http://www." & url 
                End If 
            End If 
        End If 
    End If 
    '操作是否为真
    If handleYes = True Then
        If Left(url, 1) = "/" Then
            url = rootUrl & url 
        ElseIf Left(url, 9) = "../../../" Then
            url = parentParentParentUrl & Right(url, Len(url) - 9) 
        ElseIf Left(url, 6) = "../../" Then
            url = parentParentUrl & Right(url, Len(url) - 6) 
        ElseIf Left(url, 3) = "../" Then
            url = parentUrl & Right(url, Len(url) - 3) 
        ElseIf Left(url, 2) = "./" Then
            url = thisUrl & Mid(url, 3) 
        Else
            url = thisUrl & url 
        End If 
    End If 
    If InStr(LCase(url), "http://") = False And InStr(LCase(url), "https://") = False Then
        If InStr(LCase(httpUrl), "http://") > 0 And InStr(LCase(url), "http://") = False Then
            url = "http://" & url 
        ElseIf InStr(LCase(httpUrl), "https://") > 0 And InStr(LCase(url), "https://") = False Then
            url = "https://" & url 
        End If 
    End If 
    fullHttpUrl = url 
    If Err Then doError Err.Description, "FullHttpUrl 获得域名 完整网址，HttpUrl=" & httpUrl & "<hr>Url=" & url 
End Function 

'批量处理网址完整20150728
Function batchFullHttpUrl(webSite, urlList)
    Dim splStr, url, c 
    splStr = Split(urlList, vbCrLf) 
    For Each url In splStr
        If Len(url) > 3 Then
            If c <> "" Then c = c & vbCrLf 
            c = c & fullHttpUrl(webSite, url) 
        End If 
    Next 
    batchFullHttpUrl = c 
End Function 


'网址特殊字符 简洁处理
Function uRLJianJieHandle(ByVal url)
    url = Replace(url, "&amp;", "&") 
    uRLJianJieHandle = url 
End Function 

'URL加密 待完善中。。。。
Function urlToAsc(url)
    Dim i 
    For i = 1 To Len(url)
        urlToAsc = urlToAsc & "%" & Hex(AscW(Mid(url, i, 1))) 
    Next 
End Function 


'获得网站标题
Function getWebTitle(content)
    getWebTitle = getStrCut(content, "<title>", "</title>", 0) 
End Function 

'获得容中网址列表 (缺陷是网址全部小写了20150728)
Function getContentAHref(httpUrl, content, byref PubAHrefList, byref PubATitleList)
    Dim i, s, TempS, LalType, nLen, LalStr, c 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        If s = "<" Then
            TempS = LCase(Mid(content, i)) 
            LalType = LCase(Mid(TempS, 1, InStr(TempS, " "))) 
            If LalType = "<a " Then
                LalStr = Mid(TempS, 1, InStr(TempS, "</") + 2) 
                nLen = Len(LalStr) - 1 
                c = c & handleLink(httpUrl, LalStr, "href", "", "url", PubAHrefList, PubATitleList) & vbCrLf 
                i = i + nLen 
            End If 
        End If 
        doEvents 
    Next 
    If c <> "" Then c = Left(c, Len(c) - 2) 
    getContentAHref = c 
End Function 

'获得内容中图片列表
Function getContentImgSrc(httpUrl, content, byref PubAHrefList, byref PubATitleList)
    Dim i, s, TempS, LalType, nLen, LalStr, c 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        If s = "<" Then
            TempS = LCase(Mid(content, i)) 
            LalType = LCase(Mid(TempS, 1, InStr(TempS, " "))) 
            If LalType = "<img " Then
                LalStr = Mid(TempS, 1, InStr(TempS, ">")) 
                nLen = Len(LalStr) - 1 
                'Call Echo(I,LalStr)
                c = c & handleLink(httpUrl, LalStr, "src", "", "url", PubAHrefList, PubATitleList) & vbCrLf 
                i = i + nLen 
            End If 
        End If 
        doEvents 
    Next 
    If c <> "" Then c = Left(c, Len(c) - 2) 
    getContentImgSrc = c 
End Function 

'让内容中网址完整 sType=|*|link|img|a|script|embed|param|meta|
Function handleConentUrl(httpUrl, content, sType, byref PubAHrefList, byref PubATitleList)
    Dim i, s, YuanStr, TempS, LalType, nLen, LalStr, c 
    sType = "|" & sType & "|" 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        If s = "<" Then
            YuanStr = Mid(content, i) 
            TempS = LCase(YuanStr) 
            TempS = Replace(Replace(TempS, Chr(10), " " & vbCrLf), Chr(13), " " & vbCrLf) '让处理图片素材更完整  比如  <img换行src=""  也可以获得 20150714
            LalStr = Mid(YuanStr, 1, InStr(YuanStr, ">")) 
            LalType = LCase(Mid(TempS, 1, InStr(TempS, " "))) 
            If LalType = "<link " And(InStr(sType, "|link|") > 0 Or InStr(sType, "|*|") > 0) Then
                nLen = Len(LalStr) - 1 
                c = c & handleLink(httpUrl, LalStr, "href", "", "", PubAHrefList, PubATitleList) 
                i = i + nLen 
            ElseIf LalType = "<img " And(InStr(sType, "|img|") > 0 Or InStr(sType, "|*|") > 0) Then
                nLen = Len(LalStr) - 1 
                c = c & handleLink(httpUrl, LalStr, "src", "", "", PubAHrefList, PubATitleList) 
                i = i + nLen 
            ElseIf LalType = "<a " And(InStr(sType, "|a|") > 0 Or InStr(sType, "|*|") > 0) Then
                nLen = Len(LalStr) - 1 
                '没有javascript就运行，但是还是有不足之处
                If InStr(LCase(LalStr), "javascript:") = 0 Then
                    c = c & handleLink(httpUrl, LalStr, "href", "", "", PubAHrefList, PubATitleList) 
                Else
                    c = c & LalStr 
                End If 
                i = i + nLen 
            ElseIf LalType = "<script " And(InStr(sType, "|script|") > 0 Or InStr(sType, "|*|") > 0) Then
                nLen = Len(LalStr) - 1 
                If InStr(LCase(LalStr), "src") > 0 Then
                    c = c & handleLink(httpUrl, LalStr, "src", "", "", PubAHrefList, PubATitleList) 
                Else
                    c = c & LalStr 
                End If 
                i = i + nLen 
            ElseIf LalType = "<embed " And(InStr(sType, "|embed|") > 0 Or InStr(sType, "|*|") > 0) Then
                nLen = Len(LalStr) - 1 
                c = c & handleLink(httpUrl, LalStr, "src", "", "", PubAHrefList, PubATitleList) 
                i = i + nLen 
            ElseIf LalType = "<param " And(InStr(sType, "|param|") > 0 Or InStr(sType, "|*|") > 0) Then
                nLen = Len(LalStr) - 1 
                If InStr(LCase(LalStr), "movie") > 0 Then
                    c = c & handleLink(httpUrl, LalStr, "value", "", "", PubAHrefList, PubATitleList) 
                Else
                    c = c & LalStr 
                End If 
                i = i + nLen 
            ElseIf LalType = "<meta " And(InStr(sType, "|meta|") > 0 Or InStr(sType, "|*|") > 0) Then
                nLen = Len(LalStr) - 1 
                '替换关键词
                If InStr(LCase(LalStr), "keywords") > 0 Then
                    c = c & handleLink(httpUrl, LalStr, "content", WebKeywords, "", PubAHrefList, PubATitleList) 
                '替换网站描述
                ElseIf InStr(LCase(LalStr), "description") > 0 Then
                    c = c & handleLink(httpUrl, LalStr, "content", WebDescription, "", PubAHrefList, PubATitleList) 
                Else
                    c = c & LalStr 
                End If 
                i = i + nLen 
            Else
                c = c & s 
            End If 
        Else
            c = c & s 
        End If 
        doEvents 
    Next 
    handleConentUrl = c 
End Function 


'替换内容里全部Js目录 20150722  call rwend(handleConentUrl("/admin/js/", "<script src='aa/js.js' ><script src=""bb/js.js"" >","",""))
Function replaceContentJsDir(content, dirPath, byref PubAHrefList, byref PubATitleList)
    Dim splStr, s, c 
    splStr = Split(content, vbCrLf) 
    For Each s In splStr
        If c <> "" Then c = c & vbCrLf 
        If InStr(s, "<script ") > 0 And InStr(s, "</script>") > 0 Then
            s = handleLink(dirPath, s, "src", "", "replaceDir", PubAHrefList, PubATitleList) 
        End If 
        c = c & s 
    Next 
    replaceContentJsDir = c 
End Function 


'处理链接地址 HttpUrl=追加网址，Content=内容  SType=类型
'替换目录方法  call rw(HandleLink("Js/", "111111<script src=""js/Jquery.Min.js""></"& "script>","src", "newurl", "replaceDir","",""))
Function handleLink(httpUrl, content, sType, SetStr, UrlOrContent, byref PubAHrefList, byref PubATitleList)
    On Error Resume Next 
    Dim splStr, i, s, c, TempContent, FindUrl, HandleUrl, startStr, endStr, s1, s2, tempHttpUrl 
    tempHttpUrl = httpUrl 
    UrlOrContent = LCase(UrlOrContent) 
    content = Replace(Replace(content, "= ", "="), "= ", "=") 
    content = Replace(Replace(content, " =", "="), " =", "=") 
    TempContent = LCase(content) 
    '没有链接退出
    If InStr(TempContent, " href=") = 0 And InStr(TempContent, " src=") = 0 Then
        handleLink = "" 
        Exit Function 
    ElseIf InStr(TempContent, " href=\""") > 0 Then
        content = Replace(content, "\""", """") : TempContent = LCase(content) 
    End If 
    startStr = sType & "=""" 
    endStr = """" 
    If InStr(TempContent, startStr) > 0 And InStr(TempContent, endStr) > 0 Then
        'call echo("提示","1")
        FindUrl = strCut(content, startStr, endStr, 2) 
        If SetStr <> "" Then
            HandleUrl = SetStr 
        Else
            HandleUrl = fullHttpUrl(httpUrl, FindUrl)
            '替换目录
            If UrlOrContent = "replacedir" Then
                HandleUrl = tempHttpUrl & handleFilePathArray(HandleUrl)(2) 
            End If 
            PubAHrefList = PubAHrefList & HandleUrl & vbCrLf 
            '链接标题
            s1 = InStr(content, ">") 
            s2 = Right(content, Len(content) - s1) 
            s2 = Mid(s2, 1, InStrRev(s2, "</") - 1) 
            s2 = Replace(s2, vbCrLf, "【换行】") 
            PubATitleList = PubATitleList & s2 & vbCrLf 
        End If 
        If FindUrl <> HandleUrl Then
            '强强强旱替换
            s1 = InStr(content, startStr) - 1 + Len(startStr)                               '这里面用TempContent而不用Content因为有大小写在里面20140726
            s2 = Right(content, Len(content) - s1) 
            s2 = Mid(s2, InStr(s2, endStr)) 
            s1 = Left(content, s1) 
            content = s1 & HandleUrl & s2 
        End If 
        If UrlOrContent = "url" Then
            handleLink = HandleUrl 
        Else
            handleLink = content 
        End If 
        Exit Function 
    End If 
    startStr = sType & "='" 
    endStr = "'" 
    If InStr(TempContent, startStr) > 0 And InStr(TempContent, endStr) > 0 Then
        'call echo("提示","2")
        FindUrl = strCut(TempContent, startStr, endStr, 2) 
        If SetStr <> "" Then
            HandleUrl = SetStr 
        Else
            HandleUrl = fullHttpUrl(httpUrl, FindUrl) 
            '替换目录
            If UrlOrContent = "replacedir" Then
                HandleUrl = tempHttpUrl & handleFilePathArray(HandleUrl)(2) 
            End If 
            PubAHrefList = PubAHrefList & HandleUrl & vbCrLf 
            '链接标题
            s1 = InStr(content, ">") 
            s2 = Right(content, Len(content) - s1) 
            s2 = Mid(s2, 1, InStrRev(s2, "</") - 1) 
            s2 = Replace(s2, vbCrLf, "【换行】") 
            PubATitleList = PubATitleList & s2 & vbCrLf 
        End If 
        If FindUrl <> HandleUrl Then
            '强强强旱替换
            s1 = InStr(content, startStr) - 1 + Len(startStr) 
            s2 = Right(content, Len(content) - s1) 
            s2 = Mid(s2, InStr(s2, endStr)) 
            s1 = Left(content, s1) 
            content = s1 & HandleUrl & s2 
        End If 
        If UrlOrContent = "url" Then
            handleLink = HandleUrl 
        Else
            handleLink = content 
        End If 
        Exit Function 
    End If 
    startStr = sType & "=" 
    endStr = ">"                                                                    '这里面把之家的 空格换成>
    If InStr(TempContent, startStr) > 0 And InStr(TempContent, endStr) > 0 Then
        FindUrl = strCut(TempContent, startStr, endStr, 2) 

        If SetStr <> "" Then
            HandleUrl = SetStr 
        Else
            HandleUrl = fullHttpUrl(httpUrl, FindUrl) 
            '替换目录
            If UrlOrContent = "replacedir" Then
                HandleUrl = tempHttpUrl & handleFilePathArray(HandleUrl)(2) 
            End If 
            PubAHrefList = PubAHrefList & handleHttpUrl(HandleUrl) & vbCrLf 
            '链接标题
            s1 = InStr(content, ">") 
            s2 = Right(content, Len(content) - s1) 
            s2 = Mid(s2, 1, InStrRev(s2, "</") - 1) 
            s2 = Replace(s2, vbCrLf, "【换行】") 
            PubATitleList = PubATitleList & s2 & vbCrLf 
        End If 
        If FindUrl <> HandleUrl Then
            '强强强旱替换
            s1 = InStr(content, startStr) - 1 + Len(startStr) 
            s2 = Right(content, Len(content) - s1) 
            s2 = Mid(s2, InStr(s2, endStr)) 
            s1 = Left(content, s1) 
            content = s1 & HandleUrl & s2 
        End If 
        If UrlOrContent = "url" Then
            handleLink = HandleUrl 
        Else
            handleLink = content 
        End If 
        Exit Function 
    End If 



    If UrlOrContent <> "url" Then handleLink = content 
    Call createAddFile("出错内容列表.txt", httpUrl & vbCrLf & content & vbCrLf & sType & vbCrLf & SetStr & vbCrLf & UrlOrContent & vbCrLf & "----------------------" & vbCrLf) 
End Function 

'获得网站目录文件夹名称 \Templates\WeiZhanLue\  得到WeiZhanLue
Function getEndUrlHandleName(FileUrl)
    Dim url 
    url = Replace(Trim(FileUrl), "\", "/") 
    If Right(url, 1) = "/" Then url = Mid(url, 1, Len(url) - 1) 
    url = Mid(url, InStrRev(url, "/") + 1) 
    getEndUrlHandleName = url 
End Function 


'获得列表中不同域名列表
Function getUrlListInWebSiteList(content)
    Dim url, UrlList, splStr 
    splStr = Split(content, vbCrLf) 
    For Each url In splStr
        url = getWebSite(url) 
        If url <> "" And InStr(vbCrLf & UrlList & vbCrLf, vbCrLf & url & vbCrLf) = 0 Then
            UrlList = UrlList & url & vbCrLf 
        End If 
        doevents 
    Next 
    getUrlListInWebSiteList = UrlList 
End Function 

'获得当前网址的文件名称
Function getThisUrlFileName()
    Dim url 
    url = Request.ServerVariables("SCRIPT_NAME") 
    If Left(url, 1) = "/" Then url = Right(url, Len(url) - 1) 
    getThisUrlFileName = url 
End Function 

'处理网站HTML中Img    写得不是特别的完善好  Content = HandleWebHtmlImg("/aa/bb/",Content)
Function handleWebHtmlImg(RootPath, content, byref PubAHrefList, byref PubATitleList)
    Dim ImgList, splStr, imgUrl, NewImgUrl 
    Dim startStr, endStr 
    ImgList = getContentImgSrc("", content, PubAHrefList, PubATitleList) 
    splStr = Split(ImgList, vbCrLf) 
    For Each imgUrl In splStr
        If imgUrl <> "" Then
            NewImgUrl = handleHttpUrl(imgUrl) 
            If InStr(NewImgUrl, "/") > 0 Then
                NewImgUrl = Mid(NewImgUrl, InStrRev(NewImgUrl, "/") + 1) 
            End If 
            NewImgUrl = RootPath & NewImgUrl 
            'Call Echo(NewImgUrl,ImgUrl)
            startStr = "src=""" : endStr = """" 
            If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
                content = regExp_Replace(content, startStr & imgUrl & endStr, startStr & NewImgUrl & endStr) 
            End If 
            startStr = "src='" : endStr = "'" 
            If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
                content = regExp_Replace(content, startStr & imgUrl & endStr, startStr & NewImgUrl & endStr) 
            End If 
        End If 
    Next 
    handleWebHtmlImg = content 
End Function 

'处理网站Css中Img    写得不是特别的完善好  Content = HandleWebHtmlImg("/aa/bb/",Content)
Function handleWebCssImg(RootPath, content)
    Dim startStr, endStr, ImgList, splStr, c, imgUrl, NewImgUrl 
    startStr = "url\(" 
    endStr = "\)" 
    ImgList = getArray(content, startStr, endStr, False, False) 
    'Call RwEnd(ImgList)
    splStr = Split(ImgList, "$Array$") 
    For Each imgUrl In splStr
        If imgUrl <> "" Then
            NewImgUrl = handleHttpUrl(imgUrl) 
            If InStr(NewImgUrl, "/") > 0 Then
                NewImgUrl = Mid(NewImgUrl, InStrRev(NewImgUrl, "/") + 1) 
            End If 
            NewImgUrl = RootPath & NewImgUrl 
            startStr = "url(" 
            endStr = ")" 
            If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
                'call echo(StartStr,"StartStr")
                content = regExp_Replace(content, startStr & imgUrl & endStr, startStr & NewImgUrl & endStr) 
            End If 
        End If 
    Next 
    handleWebCssImg = content 
End Function 

'批量处理网址完整性
Function batchHandleUrlIntegrity(httpUrl, UrlList)
    Dim splUrl, url, c 
    splUrl = Split(UrlList, vbCrLf) 
    For Each url In splUrl
        If url <> "" Then
            url = fullHttpUrl(httpUrl, url) 
            If InStr(vbCrLf & c & vbCrLf, vbCrLf & url & vbCrLf) = False Then
                c = c & url & vbCrLf 
            End If 
        End If 
    Next 
    batchHandleUrlIntegrity = c 
End Function 

'处理内容中链接图片路径 目录是显示图片
Function replaceContentImagePath(RootFolder, content)
    Dim ImageFile, ToImageFile, ImageList, splxx 
    RootFolder = handleHttpUrl(RootFolder) 
    If Right(RootFolder, 1) <> "/" Then RootFolder = RootFolder & "/" 
    ImageList = getDirFileNameList(RootFolder,"") 
    splxx = Split(ImageList, vbCrLf) 
    For Each ImageFile In splxx
        If ImageFile <> "" Then
            ToImageFile = "file:///" & RootFolder & ImageFile 
            'html中图片路径替换
            content = Replace(content, """" & ImageFile & """", """" & ToImageFile & """") 
            content = Replace(content, "'" & ImageFile & "'", """" & ToImageFile & """") 
            content = Replace(content, "=" & ImageFile & " ", """" & ToImageFile & """") 
            content = Replace(content, "=" & ImageFile & ">", """" & ToImageFile & """") 
            'Css中图片路径替换
            content = Replace(content, "(" & ImageFile & ")", "(" & ToImageFile & ")") 
            content = Replace(content, "(" & ImageFile & ";", "(" & ToImageFile & ";") 
        End If 
    Next 
    replaceContentImagePath = content 
End Function 

'获得Css链接列表 是名称
Function getCssLinkList(ByVal content)
    Dim startStr, endStr, splStr, s, c, fileName 
    startStr = "<link" : endStr = "/>" 
    content = getArray(content, startStr, endStr, False, False) 
    splStr = Split(content, "$Array$") 
    For Each s In splStr
        If InStr(LCase(s), "stylesheet") > 0 Then
            fileName = strCut(s, "href=""", """", 2) 
            Call echo(fileName, s) 
            c = c & fileName & vbCrLf 
        End If 
    Next 
    getCssLinkList = c 
End Function 

'获得Html中图片地址列表
Function getHtmlBackGroundUrlList(content)
    Dim i, s, YuanStr, TempS, LalType, nLen, LalStr, c, startStr, endStr 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        If s = "<" Then
            YuanStr = Mid(content, i) 
            TempS = LCase(YuanStr) 
            LalStr = Mid(YuanStr, 1, InStr(YuanStr, ">")) 
            LalType = LCase(Mid(TempS, 1, InStr(TempS, " "))) 
            If InStr(LalStr, "url(") > 0 Then
                startStr = "url(" : endStr = ")" 
                c = c & strCut(LalStr, startStr, endStr, 2) & vbCrLf 
                i = i + nLen 
            End If 
        End If 
        doEvents 
    Next 
    getHtmlBackGroundUrlList = c 
End Function 

'获得图片地址列表
Function getImgUrlList(content)
    Dim i, s, YuanStr, TempS, LalType, nLen, LalStr, c 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        If s = "<" Then
            YuanStr = Mid(content, i) 
            TempS = LCase(YuanStr) 
            LalStr = Mid(YuanStr, 1, InStr(YuanStr, ">")) 
            LalType = LCase(Mid(TempS, 1, InStr(TempS, " "))) 
            If LalType = "<img " Then
                c = c & getLinkUrl(LalStr, "src") & vbCrLf 
                i = i + nLen 

            End If 
        End If 
        doEvents 
    Next 
    getImgUrlList = c 
End Function 

'获得img或a地址 如GetLinkUrl(LalStr, "src")
Function getLinkUrl(ByVal LinkStr, LinkType)
    Dim TempLinkStr, startStr, endStr, LinkUrl 
    LinkStr = Replace(Replace(LinkStr, "= ", "="), "= ", "=") 
    LinkStr = Replace(Replace(LinkStr, " =", "="), " =", "=") 
    TempLinkStr = LCase(LinkStr) 

    startStr = LinkType & "=""" 
    endStr = """" 
    If InStr(TempLinkStr, startStr) > 0 And InStr(TempLinkStr, endStr) > 0 Then
        LinkUrl = strCut(TempLinkStr, startStr, endStr, 2) 
        getLinkUrl = LinkUrl 
    End If 
End Function 

'是本地网址或远程网址        20141120
Function localUrlOrRemoteUrl(filePath, UrlYes)
    Dim httpUrl 
    UrlYes = False 
    filePath = Trim(filePath) 
    httpUrl = Replace(filePath, "\", "/") 
    If Left(LCase(httpUrl), 7) = "http://" Or Left(LCase(httpUrl), 4) = "www." Or Left(LCase(httpUrl), 4) = "[网址]" Then
        UrlYes = True 
    End If 
    If UrlYes = False Then
        filePath = setFileName(filePath) 
    End If 
    localUrlOrRemoteUrl = filePath 
End Function 

'检测是否为远程网址
Function checkRemoteUrl(url)
    Dim httpUrl 
    url = Trim(url) 
    httpUrl = Replace(unSetFileName(url), "\", "/") 
    checkRemoteUrl = False 
    If Left(LCase(httpUrl), 7) = "http://" Or Left(LCase(httpUrl), 4) = "www." Or Left(LCase(httpUrl), 4) = "[网址]" Then
        If Left(LCase(httpUrl), 4) = "[网址]" Then
            url = Mid(url, InStr(url, "[网址]") + 1) 
        End If 
        checkRemoteUrl = True 
    End If 
End Function 



'函数名称：AsaiLinkAdd   20141217引用别人
'函数作用：正则自动添加链接
'例：Response.Write AsaiLinkAdd("http://www.wzl99.com/小孙、小孙http://www.wzl99.com/小孙、小孙http://www.wzl99.com//小孙、小孙http://www.wzl99.com/小孙、小孙http://www.wzl99.com/小孙、小孙www.wzl99.com/小孙、小孙fengying789@126.com小孙、小孙")
Function asaiLinkAdd(str)
    Dim re 
    str = " " & str 
    Set re = CreateObject("VBscript.RegExp")                                        '建立正则表达式对象regular expression
        re.IgnoreCase = True                                                            '忽略大小写
        re.Global = True                                                                '搜索匹配字串针对全部文本
        '查找带协议前缀的链接
        re.pattern = "(http://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?)" 
        str = re.Replace(str, "<a href='$1' target='_blank'>$1</a>") 
        '查找不带协议前缀的链接
        re.pattern = "([^(http:\/\/)])(www\.([\w-]+\.)+[\w]+(\/[\w-]+)*[\/]?([\w-]+\.[\w]+)?(\?[\w]+=[\w]+(&[\w]+=[\w]+)*)?)" 
        str = re.Replace(str, "$1<a href='http://$2' target='_blank'>$2</a>") 
        '查找邮件链接
        re.pattern = "(mailto:)?([\w]+@([\w-]+\.)+[\w]+)" 
        str = re.Replace(str, "<a href='mailto:$2'>$1$2</a>") 
    Set re = Nothing 
    asaiLinkAdd = Replace(Trim(str), " </a>", "") 
    asaiLinkAdd = Replace(Trim(str), " ;</a>", "") 
    asaiLinkAdd = Replace(Trim(str), " ；</a>", "") 
End Function 

'函数名称：AsaiLinkDel Asai(浅井)   20141217引用别人
'函数作用：正则自动去除链接
'例：Response.Write AsaiLinkDel("<a href='http://www.wzl99.com/' target='_blank'>http://www.wzl99.com/</a>小孙、小孙<a href='http://www.wzl99.com/' target='_blank'>http://www.wzl99.com/</a>小孙、")
Function asaiLinkDel(htmlStr)
    Dim ra 
    Set ra = CreateObject("VBscript.RegExp")
        ra.IgnoreCase = True 
        ra.Global = True 
        ra.pattern = "<a[^>]+>(.+?)<\/a>" 
        asaiLinkDel = ra.Replace(htmlStr, "$1") 
End Function

'判断域名是否合法  暂存
Function iswww(strng)
    iswww = False 
    Dim regex, Match 
    Set regex = CreateObject("VBscript.RegExp")
        regex.pattern = "^\w+((-\w+)|(\.\w+))*[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z]+$" 
        regex.IgnoreCase = True 
        Set Match = regex.Execute(strng)
            If Match.Count Then iswww = True 
End Function


'加强于20150220
'追加网址参数20150121  getUrlAddToParam("aa","&a=b","replace")   addto replace    SType为追加还是替换
'Url = getUrlAddToParam("http://www.baidu.com/?a=1&b=2&c=3","?a=11&b=22&c=333","")        'http://www.baidu.com/?a=1&b=2&c=3
'Url = getUrlAddToParam("http://www.baidu.com/?a=1&b=2&c=3","?a=11&b=22&c=333","replace")        'http://www.baidu.com/?a=11&b=22&c=333
'Url = getUrlAddToParam(GetUrl(),"id=" & Rs("Id"),"replace")
'Call Echo(Url,getUrlAddToParam(Url,"id=1&aa=1&bb=2","delete"))          批量删除参数
Function getUrlAddToParam(ByVal url, ByVal AddToUrl, sType)
    Dim content, splStr, splxx, s, c, httpUrl, urlFileName, webSite 
    Dim urlParam                                                                    '网址参数 是获得网址后台参数值
    Dim paramName                                                                   '参数名称
    Dim paramValue                                                                  '参数值
    Dim paramNameList                                                               '参数名称列表，防止重复
    Dim handleYes                                                                   '处理为真

    AddToUrl = handleHttpUrl(AddToUrl) 

    '处理网址
    url = Trim(url) 
    '当前网址最后一个字符为?或&给删除掉 无用
    If Right(url, 1) = "?" Or Right(url, 1) = "&" Then
        url = Trim(Mid(url, 1, Len(url) - 1)) 
    End If 
    '处理追加网址
    AddToUrl = Trim(AddToUrl) 
    '追加网址最后一个字符为?或&给删除掉 无用
    If Left(AddToUrl, 1) = "?" Or Left(AddToUrl, 1) = "&" Then
        AddToUrl = Trim(Mid(AddToUrl, 2)) 
    End If 
    '网址为空则返回追加网址 并退出
    If url = "" Then getUrlAddToParam = "?" & AddToUrl : Exit Function 

    httpUrl = url 
    If InStr(url, "?") > 0 Then
        httpUrl = Mid(url, 1, InStr(url, "?") - 1)                                      '获得当前路径网址
        webSite = getWebSite(url)                                                       '处理域名
        urlParam = Mid(url, InStr(url, "?") + 1) 

        httpUrl = handleHttpUrl(httpUrl) 
        If Right(httpUrl, 1) <> "/" Then
            urlFileName = Mid(httpUrl, InStrRev(httpUrl, "/") + 1) 
            httpUrl = Left(httpUrl, Len(httpUrl) - Len(urlFileName)) 
        'Call Echo(HttpUrl,UrlFileName)
        End If 
    End If 

    '类型选择  追加 不是替换
    If LCase(sType) = "replace" Or sType = "替换" Then
        content = AddToUrl & "&" & urlParam 
        'Call echo("Content",Content)
        If InStr(AddToUrl, "?") > 0 Then
            urlFileName = Mid(AddToUrl, 1, InStr(AddToUrl, "?") - 1) 
        'Call Echo(AddToUrl,UrlFileName)
        End If 
    ElseIf(LCase(sType) <> "delete" Or LCase(sType) <> "del") Then
        content = urlParam & "&" & AddToUrl 
    End If 

    content = Replace(content, "?", "&") 

    '处理删除参数 20150210
    If LCase(sType) = "delete" Or LCase(sType) = "del" Then
        splStr = Split(LCase(AddToUrl), "&") : AddToUrl = "&" 
        For Each s In splStr
            If InStr(s, "=") Then
                s = Mid(s, 1, InStr(s, "=") - 1) 
            End If 
            If s <> "" Then
                AddToUrl = AddToUrl & s & "&" 
            End If 
        Next 
    'Call Eerr("AddToUrl",AddToUrl)
    End If 

    'Call Echo("Content",Content)
    splStr = Split(content, "&") 
    For Each s In splStr
        If InStr(s, "=") > 0 Then
            splxx = Split(s, "=") 
            paramName = splxx(0) 
            paramValue = splxx(1) 

            handleYes = True 

            If LCase(sType) = "delete" Or LCase(sType) = "del" Then
                If InStr("&" & AddToUrl & "&", "&" & LCase(paramName) & "&") > 0 Then
                    handleYes = False 
                End If 
            End If 

            If InStr("|" & paramNameList & "|", "|" & LCase(paramName) & "|") = False And handleYes = True Then
                paramNameList = paramNameList & LCase(paramName) & "|" 
                c = c & IIF(c = "", "?", "&") 
                c = c & paramName & "=" & paramValue 
            End If 
        End If 
    Next 



    c = urlFileName & c 
    If getWebSite(c) = "" Then
        If Left(AddToUrl, 1) = "/" Then
            c = webSite & c 
        Else
            c = httpUrl & c 
        End If 

    End If 

	c=replace(c,"\","/")			'20160313
    getUrlAddToParam = c 
End Function 




'组合网址 20150706 call echo("",groupUrl("www.baidu.com//","/1.asp"))
Function groupUrl(url1, url2)
    Dim urlType, i 
    urlType = "/" 
    url1 = Replace(url1, IIF(urlType = "/", "\", "/"), urlType) 
    url2 = Replace(url2, IIF(urlType = "/", "\", "/"), urlType) 
    url1 = phptrim(url1) 
    url2 = phptrim(url2) 
    For i = 0 To 99
        If Right(url1, 1) = urlType Then
            url1 = Mid(url1, 1, Len(url1) - 1) 
        Else
            Exit For 
        End If 
    Next 
    For i = 0 To 99
        If Left(url2, 1) = urlType Then
            url2 = Mid(url2, 2) 
        Else
            Exit For 
        End If 
    Next 
    groupUrl = url1 & urlType & url2 
End Function 



'处理POST或Cookes发送方式的参数处理
Function handlePostCookiesParame(Parame, sType)
    Dim splStr, s, c, leftC, rightC 
    splStr = Split(Parame, "&") 
    For Each s In splStr
        If InStr(s, "=") > 0 Then
            leftC = Mid(s, 1, InStr(s, "=")) 
            rightC = Mid(s, InStr(s, "=") + 1) 
            If LCase(sType) = "post" Or sType = "" Then
                If c <> "" Then c = c & "&" 
                rightC = escape(rightC) 
            ElseIf LCase(sType) = "cookies" Or LCase(sType) = "cookie" Then
                If c <> "" Then c = c & ";" 
                rightC = URLEncoding(rightC) 
            End If 
            c = c & leftC & rightC 
        'call echo(leftC,RightC)
        End If 
    Next 
    handlePostCookiesParame = c 
End Function 


'移除网址中参数20150724
Function remoteHttpUrlParameter(httpUrl)
    Dim splStr, s, c, leftC, rightC 
    '没有?号退出
    If InStr(httpUrl, "?") = False Then
        remoteHttpUrlParameter = httpUrl 
        Exit Function 
    End If 
    splStr = Split(httpUrl, "&") 
    For Each s In splStr
        If InStr(s, "=") > 0 Then
            leftC = Mid(s, 1, InStr(s, "=")) 
            rightC = Mid(s, InStr(s, "=") + 1) 
            If c <> "" Then c = c & "&" 
            c = c & leftC 
        End If 
    Next 
    remoteHttpUrlParameter = c 
End Function 


'检测当前网址里文件名称是否存在(20150909)
'用法 call echo("checkUrlName",checkUrlName("1|2|3|"))      <%=IIF(checkUrlName("1|2|3|"),"111","222")% >
Function checkUrlName(searchUrlName)
    Dim splStr, urlName, url 
    searchUrlName = LCase(searchUrlName)                                            '搜索网址名称转小写
    url = LCase(Request.ServerVariables("script_name")) 
    splStr = Split(searchUrlName, "|") 
    For Each urlName In splStr
        If urlName <> "" Then
            If InStr(url, urlName) > 0 Then
                checkUrlName = True 
                Exit Function 
            End If 
        End If 
    Next 
    checkUrlName = False 
End Function 





'0 url：http://127.0.0.1/aa/4.asp?act=sdf&v=sdf
'1 urlDir：http://127.0.0.1/
'2 fileName：4.asp
'3 FileType：asp
'4 fileStr：4.asp?act=sdf&v=sdf
'5 HttpAgreement：http
'6 webSite：http://127.0.0.1/
'7 folderDir：aa/

'网址处理成数组20150124  数组  0原文件路径 1为文件路径   2为文件名称  3为去除文件类型文件名称   4为文件类型后缀名
Function handleHttpUrlArray(ByVal url)
    Dim urlDir, fileName, FileType, fileStr, httpAgreement, webSite, folderDir 
    url = handleHttpUrl(url) 

    urlDir = Mid(url, 1, InStrRev(url, "/")) 
    fileStr = Mid(url, InStrRev(url, "/") + 1) : fileName = fileStr 
    If InStr(fileStr, "?") > 0 Then
        fileName = Mid(fileStr, 1, InStr(fileStr, "?") - 1) 
    End If 
    FileType = Mid(fileName, InStrRev(fileName, ".") + 1) 
    httpAgreement = Mid(url, 1, InStr(url, ":") - 1) 
    webSite = getWebSite(url) 
    Call echo("url", url) 
    folderDir = Mid(urlDir, Len(webSite)) 
    'HandleHttpUrlArray = Array(url, urlDir, fileName, fileType, fileStr, HttpAgreement, webSite, folderDir)
    Dim arrayData 
    arrayData = Split(url & vbCrLf & urlDir & vbCrLf & fileName & vbCrLf & FileType & vbCrLf & fileStr & vbCrLf & httpAgreement & vbCrLf & webSite & vbCrLf & folderDir, vbCrLf) 
    handleHttpUrlArray = arrayData 
End Function 


'移除jsCss后的参数Param (20151019)
Function remoteJsCssParam(content, PubAHrefList)
    remoteJsCssParam = handleRemoteJsCssParam(content, PubAHrefList, "|替换内容|替换网址") 
End Function 

'处理移除jsCss后的参数Param (20151019)
Function handleRemoteJsCssParam(content, urlList, sType)
    Dim splStr, c, url, dataArray, fileName, fileType, fileStr, replaceStr 
    sType = "|" & sType & "|" 
    splStr = Split(urlList, vbCrLf) 
    For Each url In splStr
        If url <> "" Then
            If c <> "" Then c = c & vbCrLf 
            dataArray = handleHttpUrlArray(url) 
            fileName = dataArray(2) 
            fileType = LCase(dataArray(3)) 
            fileStr = LCase(dataArray(4)) 
            If(fileType = "js" Or fileType = "css") And InStr(fileStr, "?") > 0 Then
                replaceStr = Mid(url, 1, InStr(url, "?") - 1) 
                'call echo(replaceStr,fileStr)
                '这种替换方法还是不精准，待改进
                If InStr(sType, "|替换内容|") > 0 Then
                    content = Replace(content, url, replaceStr) 
                End If 
                If InStr(sType, "|替换网址|") > 0 Then
                    urlList = Replace(urlList, url, replaceStr) 
                End If 
            End If 
        End If 
    Next 
End Function 


'批量处理网址完整(20151022)
Function batchHandleHttpUrlComplete(httpUrl, content)
    Dim webSite, splStr, url, lCaseUrl, c 
    webSite = getwebsite(httpUrl) 
    splStr = Split(content, vbCrLf) 
    For Each url In splStr
        url = phptrim(url) 
        lCaseUrl = LCase(url) 
        If lCaseUrl <> "#" And Left(lCaseUrl, 11) <> "javascript:" Then
            If InStr(vbCrLf & LCase(c) & vbCrLf, vbCrLf & lCaseUrl & vbCrLf) = False Then
                If c <> "" Then c = c & vbCrLf 
                c = c & urlAddHttpUrl(webSite, url) 
            End If 
        End If 
    Next 
    batchHandleHttpUrlComplete = c 
End Function 


'检测同域名(20151023)
Function isWebSite(ByVal url1, ByVal url2)
    isWebSite = handleIsWebSite(url1, url2, "") 
End Function 
'检测同子域名(20151023)
Function isSonWebSite(ByVal url1, ByVal url2)
    isSonWebSite = handleIsWebSite(url1, url2, "子域名") 
End Function 

'处理两网址是否域名同等(20151023)
Function handleIsWebSite(ByVal url1, ByVal url2, sType)
    url1 = getwebsite(url1) 
    url2 = getwebsite(url2) 
    If InStr(url1, "://") > 0 Then
        url1 = Mid(url1, InStr(url1, "://") + 3) 
    End If 
    If Left(url1, 4) = "www." Then
        url1 = Mid(url1, 5) 
    End If 
    If InStr(url2, "://") > 0 Then
        url2 = Mid(url2, InStr(url2, "://") + 3) 
    End If 
    If Left(url2, 4) = "www." Then
        url2 = Mid(url2, 5) 
    End If 

    If sType = "子域名" Then
        Dim splStr, s, c 
        c = ".com.hk|.sh.cn|.com.cn|.net.cn|.org.cn" 
        c = c & "|.com|.net|.org|.tv|.cc|.info|.cn|.tw|:81|:99|.biz|.mobi|.hk|.us|.la|.gl|.in" 
        splStr = Split(c, "|") 
        For Each s In splStr
            If s <> "" Then
                url1 = Replace(url1, s, "") 
                url2 = Replace(url2, s, "") 
            End If 
        Next 

        If InStr(url1, ".") Then
            url1 = Mid(url1, InStr(url1, ".") + 1) 
        End If 
        If InStr(url2, ".") Then
            url2 = Mid(url2, InStr(url2, ".") + 1) 
        End If 


    End If 
    handleIsWebSite = False 
    If url1 = url2 Then
        handleIsWebSite = True 
    End If 
End Function 


'获得内容里网址列表(20161025)
Function getContentUrlList(httpUrl, ByVal content)
    getContentUrlList = handleGetContentUrlList(httpUrl, content, "|*|内链|") 
End Function 
'处理获得内容里网址列表(20161025)
Function handleGetContentUrlList(httpUrl, ByVal content, sType)
    Dim i, s, nextS, endSLCase, endS, urlStr, nLen, urlList, url, urlLCase, webSite, labelType, isHandle, valueLabel 
    sType = "|" & LCase(Trim(sType)) & "|" 
    webSite = getwebsite(httpUrl) 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        nextS = Mid(content & " ", i + 1, 1) 
        endS = Mid(content, i + 1) : endSLCase = LCase(endS) 
        If s = "<" Then
            url = "" 
            labelType = "" 
            isHandle = False 
            If Left(endSLCase, 2) = "a " Then
                labelType = "a" 
                valueLabel = "href" 
                isHandle = True 
            ElseIf Left(endSLCase, 5) = "link " Then
                labelType = "link" 
                valueLabel = "href" 
                isHandle = True 
            ElseIf Left(endSLCase, 4) = "img " Then
                labelType = "img" 
                valueLabel = "src" 
                isHandle = True 
            ElseIf Left(endSLCase, 7) = "script " Then
                labelType = "script" 
                valueLabel = "src" 
                isHandle = True 
            End If 
            If isHandle = True Then
                If InStr(sType, "|" & labelType & "|") > 0 Or InStr(sType, "|*|") > 0 Then
                    nLen = InStr(endS, ">") 
                    urlStr = Mid(endS, 1, nLen) 
                    url = RParam(urlStr, valueLabel) 
                    i = i + nLen 
                End If 
            End If 
            If url <> "" Then
                urlLCase = LCase(url) 
                If urlLCase <> "#" And Left(urlLCase, 11) <> "javascript:" Then
                    If InStr(vbCrLf & urlList & vbCrLf, vbCrLf & url & vbCrLf) = False Then
                        url = fullHttpUrl(httpUrl, url) 
                        isHandle = isSonWebSite(url, httpUrl) 
                        If InStr(sType, "|内链|") > 0 Then
                            If isHandle = True Then
                                urlList = urlList & url & vbCrLf 
                            End If 
                        ElseIf InStr(sType, "|外链|") > 0 Then
                            If isHandle = False Then
                                urlList = urlList & url & vbCrLf 
                            End If 
                        Else
                            urlList = urlList & url & vbCrLf 
                        End If 
                    End If 
                End If 
            End If 
        End If 
    Next 
    If urlList <> "" Then urlList = Left(urlList, Len(urlList) - 2) 
    handleGetContentUrlList = urlList 
End Function 
'获得网址中内链与外链列表
Function getInChain(httpUrl, urlList)
    Dim splStr, url, c, urlLCase, isHandle 
    splStr = Split(urlList, vbCrLf) 
    urlList = "" 
    For Each url In splStr
        If url <> "" Then
            urlLCase = LCase(url)
            If Left(urlLCase, 1) <> "#" And Left(urlLCase, 11) <> "javascript:" Then
                If InStr(vbCrLf & urlList & vbCrLf, vbCrLf & url & vbCrLf) = False Then
                    url = fullHttpUrl(httpUrl, url) 
                    isHandle = isSonWebSite(url, httpUrl) 
                    If isHandle = True Then
                        urlList = urlList & url & vbCrLf 
                    End If 
                End If
            End If
        End If 
    Next 
    If urlList <> "" Then urlList = Left(urlList, Len(urlList) - 2) 
    getInChain = urlList 
End Function


'处理扫描后网址列表 20160428
function handleScanUrlList(httpurl,urlList)
	dim splstr,url,c,lCaseUrl
	splstr=split(urlList,vbcrlf)
	for each url in splstr
		url=phptrim(url)
		lCaseUrl=lcase(url) 
		if url<>"" and left(url,10)<>"tencent://" and left(url,11)<>"javascript:" and left(url,1)<>"#" then		
			url = fullHttpUrl(httpurl,url)
			if instr(vbcrlf & c & vbcrlf,vbcrlf & url & vbcrlf)=false then
				c=c & url & vbcrlf			
			end if
		end if
	next
	handleScanUrlList=c
end function

'处理相同域名列表 20160501
function handleWithWebSiteList(httpurl,urllist)
	dim website,splstr,url,c,urlWebsite,s
	website=lcase(getwebsite(httpurl))
	splstr=split(urllist,vbcrlf)
	for each url in splstr
		if url <>"" then
			if right(url,1)<>"/" and instr(url,"?")=false then
				s=mid(url,instrrev(url,"/"))
				'call echo("s",s)
				if (instr(s,".")=false or instr(s,".com")>0 or instr(s,".cn")>0 or instr(s,".net")>0) and instr(s,"@")=false then
					url = url & "/"
				end if
				'call echo("url",url)
			end if
			urlWebsite=lcase(getwebsite(url))
			if website=urlWebsite and instr(vbcrlf & c & vbcrlf,vbcrlf & url & vbcrlf)=false then
				if c<>"" then 
					c=c & vbcrlf
				end if
				c=c & url
			end if
		end if
	next
	handleWithWebSiteList=c
end function
'处理不同域名列表 20160501
function handleDifferenceWebSiteList(httpurl,urllist)
	dim website,splstr,url,c,urlWebsite,websiteList
	website=lcase(getwebsite(httpurl))
	splstr=split(urllist,vbcrlf)
	for each url in splstr
		urlWebsite=lcase(getwebsite(url))
		if urlWebsite <>"" and website<>urlWebsite and instr(vbcrlf & websiteList & vbcrlf,vbcrlf & urlWebsite & vbcrlf)=false then
			websiteList=websiteList & urlWebsite
		end if
	next
	handleDifferenceWebSiteList=websiteList
end function


 




'检测域名存在 20160511   例：checkDomainName('http://www.baidu.com/a/b/sdf')
function checkDomainName(httpurl)
	dim url,url2,s,s2
	url=getwebsite(httpurl)
	url2=url & "/a/1/b/2/cdefg/"
	call echo(url,url2)
	s=getHttpUrl(url,"")
	s2=getHttpUrl(url2,"")
	if s=s2 then
		checkDomainName=false
	else
		checkDomainName=true
	end if
end function

'获得url状态说明
function getHttpUrlStateAbout(nState)
	dim s,c
	s=cstr(trim(nState))
	select case s
		case "100" : c="继续"
		case "101" : c="开关协议"
		case "200" : c="成功"
		case "201" : c="创建"
		case "202" : c="接受"
		case "203" : c="非权威信息"
		case "204" : c="不含内容"
		case "205" : c="重置内容"
		case "206" : c="部分内容"
		case "300" : c="多项选择"
		case "301" : c="移动永久"
		case "302" : c="暂时移动"
		case "303" : c="看其他的"
		case "304" : c="未修改"
		case "305" : c="使用代理"
		case "307" : c="临时重定向"
		
		case "400" : c="坏的要求"
		case "401" : c="未经授权"
		case "402" : c="付款要求"
		case "403" : c="禁止"
		case "404" : c="未找到"
		case "405" : c="不允许的方法"
		case "406" : c="不可接受"
		case "407" : c="代理验证所需"
		case "408" : c="请求超时"
		case "409" : c="冲突"
		case "410" : c="消失"
		case "411" : c="所需长度"
		case "412" : c="先决条件"
		case "413" : c="请求实体过大"
		case "414" : c="的请求URI太长"
		case "415" : c="不支持的媒体类型"
		case "416" : c="的请求范围不满足"
		case "417" : c="期望失败"
		
		case "500" : c="内部服务器错误"
		case "501" : c="未实施"
		case "502" : c="坏网关"
		case "503" : c="服务不可用"
		case "504" : c="网关超时"
		case "505" : c="的HTTP版本不支持"
		case "509" : c="带宽限制超过"
	end select
	getHttpUrlStateAbout=c
end function
 
%>