<% 
'XML Microsoft.XMLHTTP���������ȫ (2013,9,27)
'����Զ���ļ�  [����ʱ ������] Call SaveRemoteFile(RemoteFileUrl, LocalFileName)
Function saveRemoteFile(ByVal remoteFileurl, ByVal localFileName)
    On Error Resume Next 
    saveRemoteFile = 0 
    Dim ads, xmlHttp, getRemoteData 
    Call handlePath(localFileName)                                                  '�������·��
    Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
        With xmlHttp
            .Open "Get", remoteFileurl, False, "", ""                                       'Flaseͬ��,True�첽
            .send 
            If Err Then
                saveRemoteFile = 9999 : Exit Function                         '[��ַ����]
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
    If Err Then doError Err.Description, "SaveRemoteFile ����Զ���ļ� ����������RemoteFileUrl=" & remoteFileurl & "<hr>LocalFileName=" & localFileName 
End Function 


'����Զ���ļ� (����)
Function downFile(remoteFileurl, localFileName)
    downFile = saveRemoteFile(remoteFileurl, localFileName) 
End Function 


'�����ļ�����
Function popupDownFile(path)
    Dim OSM, SZ 
    Call handlePath(path)                                                           '�������·��
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

'��xml�����ҳ״̬
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
    'responseBody�� �������Ϊ�޷����������顣
    'responseStream�� �������ΪIStream����
    'responseText �� �������Ϊ�ַ�����
'responseXML�� �������ΪXML��ʽ���ݡ�
End Function 



'Post������Gb2312��ʽ
'��20150105 Call Echo(Url,XMLPost("http://127.0.0.1/5.asp", "title='"& escape("����"& vbCrlf &"1") & "'&content='����'&UpDateTime='ʱ��'"))
Function XMLPost(url, PostStr)
    XMLPost = handleXMLPost(url, PostStr, "gb2312")(0) 
End Function 


'Post������Utf-8��ʽ
Function UTFXMLPost(url, PostStr)
    UTFXMLPost = handleXMLPost(url, PostStr, "utf-8")(0) 
End Function 


'��xml����ҳ��������, ��������ˣ����(�� ��=%D7%C0) (������)
Function handleXMLPost(ByVal url, ByVal PostStr, sSetChar)
    Dim http, dataArray(6), content 
    sSetChar = handleStrCharSet(sSetChar) 
    Set http = CreateObject("Microsoft.XMLHTTP")
        Call http.Open("POST", url, False) 
        Call http.setRequestHeader("cache-control", "no-cache") 
        Call http.setRequestHeader("Content-Type", "application/x-www-form-urlencoded") 
        Call http.setRequestHeader("Connection", "close") 
        Call http.setRequestHeader("Content-Length", Len(PostStr))                      '���Բ���Ҫ
		
        Call http.send(PostStr) 
        If http.readyState <> 4 Then
            content = "error" 
        Else
            content = bytesToBstr(http.responseBody, sSetChar) 
        'content = bytes2BSTR(Http.responseBody)    '���Ҫ�������Ǹ�����   ��ʱҲ������
        End If 
        dataArray(0) = content 
        dataArray(1) = http.Status 
        dataArray(2) = http.responseBody                                                '�ֽ�
        handleXMLPost = dataArray 
    Set http = Nothing 
End Function 


'��xml����ҳ��������, ��������ˣ����(�� ��=%D7%C0) (������)
Function xmlGet(ByVal url)
    xmlGet = handleXmlGet(url, "gb2312")(0) 
End Function 


'����get�����������
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
        'content = bytes2BSTR(Http.responseBody)        '���Ҫ�������Ǹ���������   ��ʱҲ������
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


'Cookies�ύ (������)
Function cookiesPost(PostUrl, PostCok)
    cookiesPost = handleCookiesPost(PostUrl, PostCok, "gb2312") 
End Function 


'����Cookies�ύ
Function handleCookiesPost(PostUrl, PostCok, sSetChar)
    Dim http, dataArray(1), content 
    sSetChar = handleStrCharSet(sSetChar) 
    PostCok = "ASPSESSIONIDAQACTAQB=HKFHJOPDOMAIKGMPGBJJDKLJ;" & PostCok 
    Set http = CreateObject("msxml2.serverXMLHTTP")
        Call http.Open("POST", PostUrl, False) 
        '.SetRequestHeader "Content-Length", Len(PostStr)        '���Բ���Ҫ
        Call http.setRequestHeader("Content-Type", "application/x-www-form-urlencoded") 
        '.SetRequestHeader "Referer", PostRef
        Call http.setRequestHeader("Cookie", PostCok) 
        Call http.send 
        If http.readyState <> 4 Then
            content = "error" 
        Else
            content = bytesToBstr(http.responseBody, sSetChar) 
        'content = bytes2BSTR(Http.responseBody)    '���Ҫ�������Ǹ���������   ��ʱҲ������
        End If 
        dataArray(0) = content 
        dataArray(1) = http.Status 
    Set http = Nothing 
    handleCookiesPost = dataArray 
End Function 


'���ܣ�ASP��ȡԶ���ļ���С   ������
'������url Զ���ļ���ַ
'���أ�Զ���ļ���С(��λ���ֽ�)
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

'����Զ���ļ�
Function downRemoteFile(path)
    Dim OSM, SZ 
    Call handlePath(path)                                                           '�������·��
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


'ͼƬתBase64 20150722�������� call echo("images/error.gif",ImagesToBase64("images/error.gif"))
Function imagesToBase64(fileName)
    Dim xml 
    Dim root 
    Dim fs 
    Dim filePath 
    Dim objStream 
    Dim objXMLDoc 
    Dim base64 
    Dim getFileExt 
    '����������
    '��������
    filePath = handlePath(fileName) 
    getFileExt = Mid(filePath, InStrRev(filePath, ".") + 1) 
    Set objXMLDoc = CreateObject("msxml2.FreeThreadedDOMDocument")
        '�趨����XML�ĵ��ĸ�Ϊ Base64Data
        Call objXMLDoc.loadxml("<?xml version='1.0'?><Base64Data />") 
        Set fs = CreateObject("Scripting.FileSystemObject")                             ''��������ҪFSO���
            If fs.fileExists(filePath) Then                                                 '�ж�File�ļ��Ƿ����
                '�� stream ����ȡ����
                Set objStream = CreateObject("ADODB.Stream")
                    objStream.Type = 1 
                    objStream.Open 
                    Call objStream.loadFromFile(filePath) 


                    objXMLDoc.documentElement.DataType = "bin.base64" 
                    objXMLDoc.documentElement.nodeTypedValue = objStream.read 
                    '��������ȡ����.�õ���ֵ objXMLDoc
                    '����XML�ļ�
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

'Base64תͼƬ 20150722 ��������   call Base64ToImages("1.jpg",Base64Data)
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

'SQLע�벿��

'���SQLע������
Function getSqlInContent(httpUrl, canShu, methodType, selectWebShowType)
    getSqlInContent = handleSqlIn(httpUrl, canShu, methodType, selectWebShowType, "gb2312")(0) 
End Function 


'�������SQLע����Ϣ �� call rwend( handleSqlIn("http://127.0.0.1/4.asp?", "&url=aaa3222333&url2=sss333333", "cookies", 1, "utf-8")(0) )
Function handleSqlIn(httpUrl, canShu, methodType, selectWebShowType, txtCharSet)
    Dim content, dataArray                                                   '���(9)����Ҫ��Ҫ��Ȼ�ͻᱨ�������ܸ����鸳ֵ
    methodType = LCase(methodType) 
    'Post����  ���Ӽ���Ū
    If methodType = "post" Then
        dataArray = handleXMLPost(httpUrl, handlePostCookiesParame(canShu, "post"), txtCharSet) 
    'Get��������
    ElseIf methodType = "get" Then
        canShu = handlUrlCanShu(canShu)                                                 '��������
        dataArray = handleXmlGet(httpUrl & canShu, txtCharSet) 

    'Cookies������
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



'�ж��Ƿ����ע��
Function sqlInUrl(ByVal httpUrl)
    Dim XMLObject, IsTrue1, IsTrue2 
    sqlInUrl = False 
    IsTrue1 = False : IsTrue2 = False 
    Set XMLObject = CreateObject("Microsoft.XMLHTTP")
        With XMLObject
            .Open "GET", httpUrl, False                                                     '����ҳ���Ƿ����
            .send 
            If XMLObject.Status <> 200 Then
                Exit Function 
            End If 
            Call.Open("GET", httpUrl & " And 1=1", False)                                   '����ҳ���Ƿ����
            .send 
            IsTrue1 =.Status 

            Call.Open("GET", httpUrl & " And 1=2", False)                                   '����ҳ���Ƿ����
            .send "" 
            IsTrue2 =.Status 
        End With 
    Set XMLObject = Nothing 
    If IsTrue1 = 200 And IsTrue2 = 500 Then sqlInUrl = True 
End Function 

'����ע����ַ�����û����վע����ַ
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


'�����ַ����� 20150723
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


'URL����  �Լ�д�ò���ʹ�ã�UTF8_URLEncodingû��������
Function URLEncoding(str)
    '׷�ӿɲ�Ҫ
    str = Replace(str, "1", "%31") 
    str = Replace(str, "0", "%30") 
    str = Replace(str, "A", "%41") 
    str = Replace(str, "��", "%CE%D2") 
    str = Replace(str, "#", "%23") 
    'ԭʼ
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






'α��������ַ20150922  call rw(getWinHttp("http://127.0.0.1/4.asp","","Admin%5FPassWord=admin; Admin%5FUserName=admin; CNZZDATA5865366=cnzz_eid%3D1338721452-1438760869-%26ntime%3D1438760869; style=default; sswbr=1; Hm_lvt_c39492dd0d3beab26d196a53cbd72d54=1440725802,1442725700; BD_UPN=112451; ASPSESSIONIDSCDQDRAC=CNOHMAHANIPHOFLHKCIKFIGH",""))
Function getWinHttp(httpUrl, refererUrl, cookiesStr, postStr, sSetChar)
    Dim winHttp, content 
    Set winHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
        With winHttp
            '��·��ַΪ�� ��Ϊ��ǰ��ַ����
            If refererUrl = "" Then
                refererUrl = getwebsite(httpUrl) 
            End If 
            '����cookies�ַ�������Ϊ��
            If cookiesStr = "" Then
                cookiesStr = "ASPSESSIONIDAQACTAQB=HKFHJOPDOMAIKGMPGBJJDKLJ;" 
            End If 
            '�ַ����
            sSetChar = Trim(sSetChar) 
            If sSetChar = "" Or sSetChar = "1" Or sSetChar = "gb2312" Then
                sSetChar = "gb2312" 
            Else
                sSetChar = "utf-8" 
            End If 

            '���ò���
            .setTimeOuts 60000, 60000, 60000, 3000                                          '���ò�����ʱʱ��
            '.SetTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
            'resolveTimeout = 10000 '���� DNS ���ֵĳ�ʱʱ�䣬10000 ���롣
            'connectTimeout = 10000 '���� Winsock ���ӵĳ�ʱʱ�䣬10000 ���롣
            'sendTimeout = 120000 '�������ݵĳ�ʱʱ�䣬120000 ���롣
            'receiveTimeout = 60000 '���� response �ĳ�ʱʱ�䣬60000 ���롣
            .Option(4) = 13056                                                              '���Դ����־
            .Option(6) = False                                                              'Ϊ True ʱ��������ҳ���ض�����תʱ�Զ���ת��False ���Զ���ת����ȡ����˷��ص�302״̬��
            '.Open "GET", "http://www.baidu.com/", False 'GET �� POST, Url, False ͬ����ʽ��True �첽��ʽ
            .Open "GET", httpUrl, False 
            '��� HTTP ͷ��Ϣ
            .SetRequestHeader "Accept", "*/*"                                               '������������
            .SetRequestHeader "Accept-Language", "zh-cn,zh"                                 '�û�ϵͳ����
            .SetRequestHeader "User-Agent", "Mozilla/6.0"                                   '�û��������Ϣ
            .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"           '���뷽ʽ
            .SetRequestHeader "Referer", refererUrl                                         '��·
            .SetRequestHeader "Connection", "Close"                                         'Close = ���������ӣ�Keep-Alive = ��������(�־�����)
            '.SetRequestHeader "Accept-Encoding", "gzip, deflate" '������ͣ��᷵�� gzip, deflate ѹ�����ı���
            '.SetRequestHeader "Content-Length", Len(Str) '���ݳ��ȣ�Post ��ʽ�õġ�
            .SetRequestHeader "Cookie", cookiesStr                                          '���� Cookie   ֱ�Ӱ�cookies�ַ����Ž�ȥ�Ϳ����ˣ�  ��÷���:IE��F12������

            postStr = handlePostCookiesParame(postStr, "post") 

            '��������
            .send(postStr)                                                                  'Post ��ʽ��.Send (����)
            .WaitForResponse                                                                '�ȴ���������XMLHTTP��Ҳ����ʹ��
            '������
            'Response.Write .Status '��ǰ HTTP ״̬
            'Response.Write .ResponseText '�ı�����
            content = bytesToBstr(.responseBody, sSetChar) 

        'Response.BinaryWrite .ResponseBody '����������������   �÷���
        End With 
        getWinHttp = content 
End Function


%>     

