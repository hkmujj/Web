<% 
'XML Microsoft.XMLHTTP组件操作大全 (2013,9,27)
'下载远程文件  [出错时 待完善] Call SaveRemoteFile(RemoteFileUrl, LocalFileName)
Function saveRemoteFile(ByVal remoteFileurl, ByVal localFileName)
    On Error Resume Next 
    saveRemoteFile = 0 
    Dim ads, xmlHttp, getRemoteData 
    Call handlePath(localFileName)                                                  '获得完整路径
    Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
        With xmlHttp
            .Open "Get", remoteFileurl, False, "", ""                                       'Flase同步,True异步
            .send 
            If Err Then
                saveRemoteFile = 9999 : Exit Function                         '[网址错误]
            End If 
            getRemoteData =.responseBody 
            saveRemoteFile =.Status 
        End With 
    Set xmlHttp = Nothing 
    Set ads = CreateObject("ADODB.Stream")
        With ads
            .Type = 1 
            .Open 
            .Write getRemoteData 
            'Response.Write("<br><b>["&LocalFileName&"]</b><br>")
            .SaveToFile localFileName, 2 
            .Cancel 
            .Close 
        End With 
    Set ads = Nothing 
'    If Err Then doError Err.Description, "SaveRemoteFile 下载远程文件 函数出错，RemoteFileUrl=" & remoteFileurl & "<hr>LocalFileName=" & localFileName 
End Function 


'下载远程文件 (辅助)
Function downFile(remoteFileurl, localFileName)
    downFile = saveRemoteFile(remoteFileurl, localFileName) 
End Function 


'下载文件程序
Function popupDownFile(path)
    Dim OSM, SZ 
    Call handlePath(path)                                                           '获得完整路径
    Response.Clear 
    Set OSM = CreateObject("ADODB.Stream")
        OSM.Open 
        OSM.Type = 1 
        OSM.loadFromFile path 
        SZ = InStrRev(path, "\") + 1 
        Response.AddHeader "Content-Disposition", "attachment; filename=" & Mid(path, SZ) 
        Response.AddHeader "Content-Length", OSM.Size 
        Response.Charset = "UTF-8" 
        Response.ContentType = "application/octet-stream" 
        Response.BinaryWrite OSM.read 
        Response.flush 
        Response.Write("") 
        OSM.Close 
    Set OSM = Nothing 
End Function 

'用xml获得网页状态
Function XMLGetStatus(ByVal url)
    On Error Resume Next 
    Dim http 
    Set http = CreateObject("Microsoft.XMLHTTP")
        With http
            .Open "Get", url, False 
            .setRequestHeader "cache-control", "no-cache" 
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
            .send 
            XMLGetStatus =.Status 
        End With 
    Set http = Nothing 
    'responseBody： 结果返回为无符号整数数组。
    'responseStream： 结果返回为IStream流。
    'responseText ： 结果返回为字符串。
'responseXML： 结果返回为XML格式数据。
End Function 



'Post发送以Gb2312方式
'例20150105 Call Echo(Url,XMLPost("http://127.0.0.1/5.asp", "title='"& escape("标题"& vbCrlf &"1") & "'&content='内容'&UpDateTime='时间'"))
Function XMLPost(url, PostStr)
    XMLPost = handleXMLPost(url, PostStr, "gb2312")(0) 
End Function 


'Post发送以Utf-8方式
Function UTFXMLPost(url, PostStr)
    UTFXMLPost = handleXMLPost(url, PostStr, "utf-8")(0) 
End Function 


'用xml传网页表单数据, 发送中文耍处理(如 桌=%D7%C0) (已完善)
Function handleXMLPost(ByVal url, ByVal PostStr, sSetChar)
    Dim http, dataArray(6), content 
    sSetChar = handleStrCharSet(sSetChar) 
    Set http = CreateObject("Microsoft.XMLHTTP")
        Call http.Open("POST", url, False) 
        Call http.setRequestHeader("cache-control", "no-cache") 
        Call http.setRequestHeader("Content-Type", "application/x-www-form-urlencoded") 
        Call http.setRequestHeader("Connection", "close") 
        Call http.setRequestHeader("Content-Length", Len(PostStr))                      '可以不需要
		
        Call http.send(PostStr) 
        If http.readyState <> 4 Then
            content = "error" 
        Else
            content = bytesToBstr(http.responseBody, sSetChar) 
        'content = bytes2BSTR(Http.responseBody)    '这个要比上面那个好用   有时也不好用
        End If 
        dataArray(0) = content 
        dataArray(1) = http.Status 
        dataArray(2) = http.responseBody                                                '字节
        handleXMLPost = dataArray 
    Set http = Nothing 
End Function 


'用xml传网页表单数据, 发送中文耍处理(如 桌=%D7%C0) (已完善)
Function xmlGet(ByVal url)
    xmlGet = handleXmlGet(url, "gb2312")(0) 
End Function 


'处理get方法获得内容
Function handleXmlGet(ByVal url, sSetChar)
    'on error resume next
    Dim http, dataArray(1), content 
    'call echo("sSetChar",sSetChar)
    sSetChar = handleStrCharSet(sSetChar) 
    Set http = CreateObject("Microsoft.XMLHTTP")
        Call http.Open("Get", url, False) 
        Call http.setRequestHeader("cache-control", "no-cache") 
        Call http.setRequestHeader("Content-Type", "application/x-www-form-urlencoded") 
        'Call http.SetRequestHeader("Cookie", "Admin%5FPassWord=admin; Admin%5FUserName=admin; CNZZDATA5865366=cnzz_eid%3D1338721452-1438760869-%26ntime%3D1438760869; style=default; sswbr=1; Hm_lvt_c39492dd0d3beab26d196a53cbd72d54=1440725802,1442725700; BD_UPN=112451; ASPSESSIONIDSCDQDRAC=CNOHMAHANIPHOFLHKCIKFIGH") 
        Call http.send 
        If http.readyState <> 4 Then
            content = "error" 
        Else
            'call eerr("sSetChar",sSetChar)
            content = bytesToBstr(http.responseBody, sSetChar) 
        'content = bytes2BSTR(Http.responseBody)        '这个要比上面那个好用我了   有时也不好用
        End If 
        dataArray(0) = content 
        dataArray(1) = http.Status 
    Set http = Nothing 
    If Err Then
        dataArray(0) = "" 
        dataArray(1) = -1 
    End If 
    handleXmlGet = dataArray 
End Function 


'Cookies提交 (已完善)
Function cookiesPost(PostUrl, PostCok)
    cookiesPost = handleCookiesPost(PostUrl, PostCok, "gb2312") 
End Function 


'处理Cookies提交
Function handleCookiesPost(PostUrl, PostCok, sSetChar)
    Dim http, dataArray(1), content 
    sSetChar = handleStrCharSet(sSetChar) 
    PostCok = "ASPSESSIONIDAQACTAQB=HKFHJOPDOMAIKGMPGBJJDKLJ;" & PostCok 
    Set http = CreateObject("msxml2.serverXMLHTTP")
        Call http.Open("POST", PostUrl, False) 
        '.SetRequestHeader "Content-Length", Len(PostStr)        '可以不需要
        Call http.setRequestHeader("Content-Type", "application/x-www-form-urlencoded") 
        '.SetRequestHeader "Referer", PostRef
        Call http.setRequestHeader("Cookie", PostCok) 
        Call http.send 
        If http.readyState <> 4 Then
            content = "error" 
        Else
            content = bytesToBstr(http.responseBody, sSetChar) 
        'content = bytes2BSTR(Http.responseBody)    '这个要比上面那个好用我了   有时也不好用
        End If 
        dataArray(0) = content 
        dataArray(1) = http.Status 
    Set http = Nothing 
    handleCookiesPost = dataArray 
End Function 


'功能：ASP获取远程文件大小   待测试
'参数：url 远程文件地址
'返回：远程文件大小(单位：字节)
Function getRemoteFileSize(url)
    Dim xmlHttp 
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
        With xmlHttp
            .Open "get", url, False 
            Call.setRequestHeader("range", "bytes=-1") 
            .send 
            getRemoteFileSize = Split(.getResponseHeader("Content-Range"), "/")(1) 
        End With 
    Set xmlHttp = Nothing 
End Function 

'下载远程文件
Function downRemoteFile(path)
    Dim OSM, SZ 
    Call handlePath(path)                                                           '获得完整路径
    Response.Clear 
    Set OSM = CreateObject("ADODB.Stream")
        With OSM
            .Open 
            .Type = 1 
            .loadFromFile path 
            SZ = InStrRev(path, "\") + 1 
            Response.AddHeader "Content-Disposition", "attachment; filename=" & Mid(path, SZ) 
            Response.AddHeader "Content-Length",.Size 
            Response.Charset = "UTF-8" 
            Response.ContentType = "application/octet-stream" 
            Response.BinaryWrite.read 
            Response.flush 
            Response.Write("") 
            .Close 
        End With 
    Set OSM = Nothing 
End Function


'图片转Base64 20150722引用网上 call echo("images/error.gif",ImagesToBase64("images/error.gif"))
Function imagesToBase64(fileName)
    Dim xml 
    Dim root 
    Dim fs 
    Dim filePath 
    Dim objStream 
    Dim objXMLDoc 
    Dim base64 
    Dim getFileExt 
    '定义变量完结
    '创建对像
    filePath = handlePath(fileName) 
    getFileExt = Mid(filePath, InStrRev(filePath, ".") + 1) 
    Set objXMLDoc = CreateObject("msxml2.FreeThreadedDOMDocument")
        '设定生成XML文档的根为 Base64Data
        Call objXMLDoc.loadxml("<?xml version='1.0'?><Base64Data />") 
        Set fs = CreateObject("Scripting.FileSystemObject")                             ''服务器需要FSO组件
            If fs.fileExists(filePath) Then                                                 '判断File文件是否存在
                '用 stream 来读取数据
                Set objStream = CreateObject("ADODB.Stream")
                    objStream.Type = 1 
                    objStream.Open 
                    Call objStream.loadFromFile(filePath) 


                    objXMLDoc.documentElement.DataType = "bin.base64" 
                    objXMLDoc.documentElement.nodeTypedValue = objStream.read 
                    '数据流读取结束.得到了值 objXMLDoc
                    '创建XML文件
                    Set xml = CreateObject("msxml2.FreeThreadedDOMDocument")
                        Call xml.Load(objXMLDoc) 
                        If xml.readyState > 2 Then
                            Set root = xml.getElementsByTagName("Base64Data")
                                'Base64="<img src=""data:image/"&getFileExt&";base64,"&vbcrlf&root.Item(0).Text&""">"
                                'Base64 = "data:image/" & getFileExt & ";base64," & vbCrLf & root.item(0).text
                                base64 = root.Item(0).text 
                        Else
                            base64 = "" 
                        End If 
                            Set xml = Nothing
                    Set objStream = Nothing
            Else
                base64 = "" 
            End If 

                Set fs = Nothing
        Set objXMLDoc = Nothing


        imagesToBase64 = base64 
End Function

'Base64转图片 20150722 引用网上   call Base64ToImages("1.jpg",Base64Data)
Function base64ToImages(saveImagePath, Base64Data)
    Dim xmlStr 
    saveImagePath = handlePath(saveImagePath) 
    xmlStr = "<data>" & Base64Data & "</data>" 
    Dim xml : Set xml = CreateObject("MSXML2.DOMDocument")
        Dim stm : Set stm = CreateObject("ADODB.Stream")
            xml.resolveExternals = False 
            Call xml.loadxml(xmlStr) 
            Call xml.documentElement.setAttribute("xmlns:dt", "urn:schemas-microsoft-com:datatypes") 
            xml.documentElement.DataType = "bin.base64" 
            stm.Type = 1                                                                    'adTypeBinary
            stm.Open 
            Call stm.Write(xml.documentElement.nodeTypedValue) 
            Call stm.SaveToFile(saveImagePath) 
            stm.Close 
        Set xml = Nothing 
    Set stm = Nothing 
End Function 

'SQL注入部分

'获得SQL注入内容
Function getSqlInContent(httpUrl, canShu, methodType, selectWebShowType)
    getSqlInContent = handleSqlIn(httpUrl, canShu, methodType, selectWebShowType, "gb2312")(0) 
End Function 


'处理获得SQL注入信息 如 call rwend( handleSqlIn("http://127.0.0.1/4.asp?", "&url=aaa3222333&url2=sss333333", "cookies", 1, "utf-8")(0) )
Function handleSqlIn(httpUrl, canShu, methodType, selectWebShowType, txtCharSet)
    Dim content, dataArray                                                   '这个(9)不能要，要不然就会报错，不能给数组赋值
    methodType = LCase(methodType) 
    'Post不行  待从家里弄
    If methodType = "post" Then
        dataArray = handleXMLPost(httpUrl, handlePostCookiesParame(canShu, "post"), txtCharSet) 
    'Get这样可以
    ElseIf methodType = "get" Then
        canShu = handlUrlCanShu(canShu)                                                 '参数处理
        dataArray = handleXmlGet(httpUrl & canShu, txtCharSet) 

    'Cookies可以用
    ElseIf methodType = "cookies" Then
        dataArray = handleCookiesPost(httpUrl, handlePostCookiesParame(canShu, "cookies"), txtCharSet) 

    End If 
    If CStr(selectWebShowType) = "0" Then
        dataArray(0) = Replace(dataArray(0), "<", "&lt;") 
    ElseIf CStr(selectWebShowType) = "1" Then
        dataArray(0) = Replace(dataArray(0), "<br>", vbCrLf) 
        dataArray(0) = delHtml(dataArray(0)) 
    End If 
    handleSqlIn = dataArray 
End Function 



'判断是否可以注入
Function sqlInUrl(ByVal httpUrl)
    Dim XMLObject, IsTrue1, IsTrue2 
    sqlInUrl = False 
    IsTrue1 = False : IsTrue2 = False 
    Set XMLObject = CreateObject("Microsoft.XMLHTTP")
        With XMLObject
            .Open "GET", httpUrl, False                                                     '测试页面是否存在
            .send 
            If XMLObject.Status <> 200 Then
                Exit Function 
            End If 
            Call.Open("GET", httpUrl & " And 1=1", False)                                   '测试页面是否存在
            .send 
            IsTrue1 =.Status 

            Call.Open("GET", httpUrl & " And 1=2", False)                                   '测试页面是否存在
            .send "" 
            IsTrue2 =.Status 
        End With 
    Set XMLObject = Nothing 
    If IsTrue1 = 200 And IsTrue2 = 500 Then sqlInUrl = True 
End Function 

'处理注入网址，配置获得网站注意网址
Function handlUrlCanShu(httpUrl)
    Dim url, splStr, i, s, s1, s2 
    splStr = Split(httpUrl, "=") 
    For i = 0 To UBound(splStr)
        s = splStr(i) 
        If url <> "" Then url = url & "=" 
        If InStr(s, "&") Then
            s1 = Mid(s, 1, InStr(s, "&") - 1) 
            s2 = Mid(s, InStr(s, "&")) 
            s = escape(s1) & s2 
        ElseIf i = UBound(splStr) Then
            s = escape(s) 
        End If 
        url = url & s 
    Next 
    handlUrlCanShu = url 
End Function 


'处理字符编码 20150723
Function handleStrCharSet(sSetChar)
    If sSetChar = "1" Or UCase(sSetChar) = "GB2312" Or sSetChar = "" Then
        sSetChar = "GB2312" 
    ElseIf sSetChar = "0" Or UCase(sSetChar) = "UTF-8" Then
        sSetChar = "UTF-8" 
    ElseIf sSetChar = "2" Or UCase(sSetChar) = "UNICODE" Then
        sSetChar = "UNICODE" 
    End If 
    handleStrCharSet = sSetChar 
End Function 


'URL加密  自己写得不再使用，UTF8_URLEncoding没有它好用
Function URLEncoding(str)
    '追加可不要
    str = Replace(str, "1", "%31") 
    str = Replace(str, "0", "%30") 
    str = Replace(str, "A", "%41") 
    str = Replace(str, "我", "%CE%D2") 
    str = Replace(str, "#", "%23") 
    '原始
    str = Replace(str, "a", "%61") 
    str = Replace(str, "n", "%6E") 
    str = Replace(str, "d", "%64") 
    str = Replace(str, " ", "%20") 
    str = Replace(str, "=", "%3D") 
    str = Replace(str, "e", "%65") 
    str = Replace(str, "x", "%78") 
    str = Replace(str, "i", "%69") 
    str = Replace(str, "s", "%73") 
    str = Replace(str, "t", "%74") 
    str = Replace(str, "(", "%28") 
    str = Replace(str, ")", "%29") 
    str = Replace(str, "l", "%6C") 
    str = Replace(str, "c", "%63") 
    str = Replace(str, "*", "%2A") 
    str = Replace(str, "f", "%66") 
    str = Replace(str, "r", "%72") 
    str = Replace(str, "o", "%6F") 
    str = Replace(str, "m", "%6D") 
    str = Replace(str, "w", "%77") 
    str = Replace(str, "h", "%68") 
    URLEncoding = str 
End Function 






'伪造来访网址20150922  call rw(getWinHttp("http://127.0.0.1/4.asp","","Admin%5FPassWord=admin; Admin%5FUserName=admin; CNZZDATA5865366=cnzz_eid%3D1338721452-1438760869-%26ntime%3D1438760869; style=default; sswbr=1; Hm_lvt_c39492dd0d3beab26d196a53cbd72d54=1440725802,1442725700; BD_UPN=112451; ASPSESSIONIDSCDQDRAC=CNOHMAHANIPHOFLHKCIKFIGH",""))
Function getWinHttp(httpUrl, refererUrl, cookiesStr, postStr, sSetChar)
    Dim winHttp, content 
    Set winHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
        With winHttp
            '来路网址为空 则为当前网址域名
            If refererUrl = "" Then
                refererUrl = getwebsite(httpUrl) 
            End If 
            '设置cookies字符串不能为空
            If cookiesStr = "" Then
                cookiesStr = "ASPSESSIONIDAQACTAQB=HKFHJOPDOMAIKGMPGBJJDKLJ;" 
            End If 
            '字符编号
            sSetChar = Trim(sSetChar) 
            If sSetChar = "" Or sSetChar = "1" Or sSetChar = "gb2312" Then
                sSetChar = "gb2312" 
            Else
                sSetChar = "utf-8" 
            End If 

            '设置参数
            .setTimeOuts 60000, 60000, 60000, 3000                                          '设置操作超时时间
            '.SetTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
            'resolveTimeout = 10000 '解析 DNS 名字的超时时间，10000 毫秒。
            'connectTimeout = 10000 '建立 Winsock 连接的超时时间，10000 毫秒。
            'sendTimeout = 120000 '发送数据的超时时间，120000 毫秒。
            'receiveTimeout = 60000 '接收 response 的超时时间，60000 毫秒。
            .Option(4) = 13056                                                              '忽略错误标志
            .Option(6) = False                                                              '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态。
            '.Open "GET", "http://www.baidu.com/", False 'GET 或 POST, Url, False 同步方式；True 异步方式
            .Open "GET", httpUrl, False 
            '组成 HTTP 头信息
            .SetRequestHeader "Accept", "*/*"                                               '接受数据类型
            .SetRequestHeader "Accept-Language", "zh-cn,zh"                                 '用户系统语言
            .SetRequestHeader "User-Agent", "Mozilla/6.0"                                   '用户浏览器信息
            .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"           '编码方式
            .SetRequestHeader "Referer", refererUrl                                         '来路
            .SetRequestHeader "Connection", "Close"                                         'Close = 不保持连接，Keep-Alive = 保持连接(持久连接)
            '.SetRequestHeader "Accept-Encoding", "gzip, deflate" '如果发送，会返回 gzip, deflate 压缩过的编码
            '.SetRequestHeader "Content-Length", Len(Str) '内容长度，Post 方式用的。
            .SetRequestHeader "Cookie", cookiesStr                                          '设置 Cookie   直接把cookies字符串放进去就可以了，  获得方法:IE》F12》网络

            postStr = handlePostCookiesParame(postStr, "post") 

            '发送数据
            .send(postStr)                                                                  'Post 方式：.Send (参数)
            .WaitForResponse                                                                '等待返回请求，XMLHTTP中也可以使用
            '输出结果
            'Response.Write .Status '当前 HTTP 状态
            'Response.Write .ResponseText '文本数据
            content = bytesToBstr(.responseBody, sSetChar) 

        'Response.BinaryWrite .ResponseBody '二进制数据流数据   好方法
        End With 
        getWinHttp = content 
End Function


%>     


