<% 
'FSO  VBscript.RegExp���������ȫ (2013,9,29)

'����ַ��Ƿ����,��Сд����
Function checkStr(ByVal content, ByVal SerchValue)
    On Error Resume Next 
    Dim Rep, search 
    checkStr = True 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.Global = True                                                               '����ȫ�����ʡ�
        Rep.IgnoreCase = True                                                           '�����Ƿ�������ĸ�Ĵ�Сд��True������Falsh����
        Rep.pattern = SerchValue 
        Set search = Rep.Execute(content)
            If search.Count = 0 Then checkStr = False 
        Set search = Nothing 
    Set Rep = Nothing 
    If Err Then Call doError(Err.Description, "CheckStr ����ַ��Ƿ���� ����������SerchValue=" & SerchValue) 
End Function 
'����ַ��Ƿ���� (����)
Function regExp_CheckStr(content, SerchValue)
    regExp_CheckStr = checkStr(content, SerchValue) 
End Function 
'����ַ��Ƿ���� (����)
Function existsStr(content, SerchValue)
    existsStr = checkStr(content, SerchValue) 
End Function 
'����ַ��Ƿ���� (����)
Function regExp_Instr(content, SerchValue)
    regExp_Instr = checkStr(content, SerchValue) 
End Function 

'�滻����,�����ִ�Сд
Function regExp_Replace(ByVal content, ByVal SetPattern, ByVal TStr)
    On Error Resume Next 
    Dim Rep 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.pattern = SetPattern                                                        '����ģʽ��
        Rep.IgnoreCase = True                                                           '�����Ƿ�������ĸ�Ĵ�Сд��
        Rep.Global = True                                                               '����ȫ�����ʡ�
        regExp_Replace = Rep.Replace(content, TStr) 
    Set Rep = Nothing 
    If Err Then Call doError(Err.Description, "RegExp_Replace �滻���� ����������SetPattern=" & SetPattern & " |Str=" & TStr) 
End Function 
'�滻�����Զ�����ɫ,�����ִ�Сд
Function replaceColor(content, SetPattern, TStr, color)
    If color <> "" Then TStr = "<font color=" & color & ">" & TStr & "</font>" 
    replaceColor = regExp_Replace(content, SetPattern, TStr) 
End Function
'ɾ��HTML��ǩ
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
    If Err Then Call doError(Err.Description, "DelHtml ɾ��HTML��ǩ ����������") 
End Function 
'ɾ��HTML��ǩ (����)
Function regExp_DelHtml(strHtml)
    regExp_DelHtml = delHtml(strHtml)
End Function 
'�������������������б�  getAUrlTitleList(Content,"��ַ")
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
                If SType = "����" Or SType = "��ַ" Or SType = "0" Then
                    s = SetValue.SubMatches(0) 
                ElseIf SType = "����" Or SType = "1" Then
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
    If Err Then Call doError(Err.Description, "GetHrefUrlTitleList �������������������б� ����������C=" & c) 
End Function 
'�����������ַ (����)
Function getAURL(content)
    getAURL = getAUrlTitleList(content, 0) 
End Function 
'��������б��� (����)
Function getATitle(content)
    getATitle = getAUrlTitleList(content, 1) 
End Function 
'�����������ַ�ͱ��� (����)
Function getAURLTitle(content)
    getAURLTitle = getAUrlTitleList(content, "���Ӻͱ���") 
End Function 


'�������������������б�  GetAUrlTitleList(Content,"��ַ")
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
    If Err Then Call doError(Err.Description, "GetHrefUrlTitleList �������������������б� ����������C=" & c) 
End Function 


'���ͼƬ����ַ
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
                If SType = ImgType Or SType = "ȫ��" Or SType = "" Then
                    If InStr(vbCrLf & c & vbCrLf, vbCrLf & s & vbCrLf) = False Then
                        c = c & s & vbCrLf 
                    End If 
                End If 
            Next 
        Set SetExecute = Nothing 
    Set Rep = Nothing 
    If c <> "" Then c = Left(c, Len(c) - 2) 
    regExp_GetImgUrlList = c 
    If Err Then Call doError(Err.Description, "GetIMGUrlList ���ͼƬ����ַ ����������c=" & c) 
End Function 
'���ͼƬ����ַ (����)
Function getIMG(content)
    getIMG = regExp_GetImgUrlList(content, "ȫ��") 
End Function 
'���ͼƬ��JPG��ַ (����)
Function getIMGJpg(content)
    getIMGJpg = regExp_GetImgUrlList(content, "JPG") 
End Function 
'���ͼƬ��Gif��ַ (����)
Function getIMGGif(content)
    getIMGGif = regExp_GetImgUrlList(content, "GIF") 
End Function 
'���ͼƬ��Bmp��ַ (����)
Function getIMGBmp(content)
    getIMGBmp = regExp_GetImgUrlList(content, "Bmp") 
End Function 
'���ASP�Զ������б�����ϻ��ASP���Զ�����(2013,9,30
Function getASPFunctionList(ByVal content)
    On Error Resume Next 
    Dim SetExecute, SetValue, Rep, ImgType, s, c 
    c = "" 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.IgnoreCase = True 
        Rep.Global = True 
        Rep.pattern = "(sub|function) .*?\)" 
        'Rep.Pattern = "\s(sub|function)( .*?)(\(.*)|\s(sub|function)( .*?[ |    ])"            '����׼20150722
        Set SetExecute = Rep.Execute(content)
            For Each SetValue In SetExecute
                c = c & SetValue & vbCrLf 
            Next 
        Set SetExecute = Nothing 
    Set Rep = Nothing 
    If c <> "" Then c = Left(c, Len(c) - 2) 
    getASPFunctionList = c 
    If Err Then Call doError(Err.Description, "GetASPFunctionList ���ASP�����б� ����������c=" & c) 
End Function 
'���ASP�Զ������б�����ϻ��ASP���Զ�����(2013,9,30
Function getASPDimFunction(ByVal content, FunValue, FunName, FunDim)
    'On Error Resume Next
    content = Replace(content, "()", "( )")                                         'Ϊ��ֹ��ȡ��������
    Dim SetExecute, SetValue, Rep 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.IgnoreCase = True 
        Rep.Global = True 
        'Rep.Pattern = "(sub|function) .*?\)"
        content = regExp_Replace(content, "ByVal ", "") 
        Rep.pattern = "(sub|function) *([^\(]+) *\(([^\)]+)\)" 
        Set SetExecute = Rep.Execute(content)
            For Each SetValue In SetExecute
                '{�ַ������ж���JS����
                If InStr(FunValue, "{") = False Then
                    FunValue = FunValue & Trim(SetValue) & vbCrLf 
                    FunName = FunName & Trim(SetValue.SubMatches(1)) & vbCrLf 
                    FunDim = FunDim & Replace(SetValue.SubMatches(2), " ", "") & vbCrLf 
                End If 
            Next 
        Set SetExecute = Nothing 
    Set Rep = Nothing 
    'If FunDim <> "" Then FunDim = Left(FunDim, Len(FunDim)-2)                '������̫����
    getASPDimFunction = FunValue 
    If Err Then Call doError(Err.Description, "GetASPFunctionList ���ASP�����б� ����������") 
End Function 
'Response.Write(StrLength("�й�121aa"))
'�������ʽ����ַ����� ���Ķ����ַ�
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
'�������ʽ����ַ����� ���Ķ����ַ�
Function regExpGetLenght(str)
    regExpGetLenght = strLength(str) 
End Function 
'�������ʽ����ַ����� ���Ķ����ַ�
Function stringLength(str)
    stringLength = strLength(str) 
End Function 
'���HTML��Css
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
    If Err Then Call doError(Err.Description, "GetCssHref ���HTML��Css ����������C =" & c) 
End Function 
'���HTML��Js
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
    If Err Then Call doError(Err.Description, "GetJsSrc ���HTML��Js ����������C =" & c) 
End Function 
'ɾ��HTML��<scirpt
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
    If Err Then Call doError(Err.Description, "ɾ��HTML��<scirpt ����������") 
End Function 
'���HTML��keywords
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
    If Err Then Call doError(Err.Description, "GetKeyWords ���HTML��keywords ����������C =" & c) 
End Function 
'ȥ��HTMLע��
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
'ɾ���Լ������HTMLע��
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
'���������ͼƬ�б�
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
    'result=split(result,"|") '�洢<img>
    'result1=split(result1,"|") '�洢ͼ���ַ
    getStrImgList = Result1 
End Function 
'��������б�
Function getEMailList(content)                                          'patrn:��Ҫ���ҵ��ַ� strng:�����ҵ��ַ���
    Dim regex, Match, Matches, RetStr                                               '����������
    Set regex = CreateObject("VBscript.RegExp")                                     '�����������ʽ��
        regex.pattern = "\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"    '����ģʽ��"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"
        regex.IgnoreCase = True                                                         '�����Ƿ����ִ�Сд��
        regex.Global = True                                                             '����ȫ��ƥ�䡣
        RetStr = "" 
        Set Matches = regex.Execute(content)                                            'ִ��������
            For Each Match In Matches                                      'ѭ������Matches���ϡ�
                If InStr("," & RetStr & ",", "," & Match.Value & ",") = 0 Then
                    RetStr = RetStr & Match.Value & "," 
                End If 
            Next 
        Set Matches = Nothing 
    Set regex = Nothing 
    If RetStr <> "" Then RetStr = Left(RetStr, Len(RetStr) - 1) 
    getEMailList = RetStr 
End Function 
'��html���ݴ�����ı�ǩ��дת����Сд
Function oflink_lcasetag(content)
    Dim regex, Match, Matches                                                       '����������
    Set regex = CreateObject("VBscript.RegExp")                                     '�����������ʽ��
        regex.pattern = "<.+?\>"                                                     '����ģʽ��
        regex.IgnoreCase = True                                                         '�����Ƿ������ַ���Сд��
        regex.Global = True                                                             '����ȫ�ֿ����ԡ�
        Set Matches = regex.Execute(content)                                            'ִ��������
            content = content 
            For Each Match In Matches                                      '����ƥ�伯�ϡ�
                content = Replace(content, Match.Value, LCase(Match.Value)) 
            Next 
            oflink_lcasetag = content 
        Set Matches = Nothing 
    Set regex = Nothing 
End Function 
'==================================================
'��������ScriptHtml
'��  �ã�����html���
'��  ����ConStr ------ Ҫ���˵��ַ���   1Ϊɾ����ʼ���  3Ϊɾ��ǰ������� 2Ϊ������Ч��
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


'��ȡ���ݣ���ָ��3��ƥ����� [$Array$]Ϊ�� [$Split$]Ϊ��    2014 12 10
'����һ
'Content = " tablename='ListMenu' bigclassname='ϵͳ����' smallclassname='Robots�����޸�' sort='3' url='Robots.Asp' display='1' "
'Call Rw(GetRegExp(" (.*?)\=.*?'(.*?)'", Content, 0,1,""))'
Function getRegExp(Patrn, strng, nSplit1, nSplit2, nSplit3)
    Dim regex, Match, Matches, c, SplitYes, s 
    c = "" 
    If nSplit1 <> "" Then nSplit1 = CInt(nSplit1)                                   '��ǰ�ָ���
    If nSplit2 <> "" Then nSplit2 = CInt(nSplit2)                                   '��ǰ�ָ���
    If nSplit3 <> "" Then nSplit3 = CInt(nSplit3)                                   '��ǰ�ָ���
    SplitYes = False                                                                '�ָ��Ƿ�Ϊ
    If nSplit1 <> "" Or nSplit2 <> "" Or nSplit3 <> "" Then SplitYes = True 
    Set regex = CreateObject("VBscript.RegExp")                                     '�����������ʽ��
        regex.pattern = Patrn                                                           '����ģʽ��
        regex.IgnoreCase = True                                                         '�����Ƿ������ַ���Сд��
        regex.Global = True                                                             '����ȫ�ֿ����ԡ�
        Set Matches = regex.Execute(strng)                                              'ִ��������
            For Each Match In Matches                                      '����ƥ�伯�ϡ�
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



'���������Img��Js������  SType=1Ϊ���ظ�   20150126
Function getImgJsUrl(content, SType)
    Dim s, c, splStr, url, UrlList 
    SType = LCase(SType)                                                            '����ת��Сд�������ַ�����
    c = RegExpGetStr("<[img|script][^<>]*src[^<>]*[\/]?>", content, 2) 
    splStr = Split(c, "[$Array$]") 
    UrlList = "" 
    For Each s In splStr
        url = RegExpGetStr("src=[""|']?([^""' ]*)([""|']?).*[\/]?>", s, 1) 
        If SType = "1" Or SType = "���ظ�" Then
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
'�������ʽ����20150126
'��ȡ��վ�б��⡢�ؼ��ʡ�����
'WebTitle = RegExpGetStr("<TITLE>(.*)</TITLE>", Content, 1)                    '��һ��
'WebTitle = RegExpGetStr("<TITLE>([^<>]*)</TITLE>", Content, 1)                    '�ڶ��� ����
'WebDes = RegExpGetStr("<meta[^<>]*description[^<>]*[\/]?>", Content, 0)
'WebDes = RegExpGetStr("content=[""|']?([^""' ]*)([""|']?).*[\/]?>", WebDes, 1)
'WebKey = RegExpGetStr("<meta[^<>]*keywords[^<>]*[\/]?>", Content, 0)
'WebKey = RegExpGetStr("content=[""|']?([^""' ]*)([""|']?).*[\/]?>", WebKey, 1)
Function regExpGetStr(title, content, SType)
    Dim re, RemoteFileurl, RemoteFile, c 
    c = "" 
    regExpGetStr = "" 
    Set re = CreateObject("VBscript.RegExp")
        re.IgnoreCase = True 
        re.Global = True 
        re.pattern = title 
        Set RemoteFile = re.Execute(content)
            For Each RemoteFileurl In RemoteFile
                If SType = 1 Then
                    RemoteFileurl = re.Replace(RemoteFileurl, "$1") 
                ElseIf SType = 2 Then
                    c = c & RemoteFileurl & "[$Array$]" 
                End If 
                regExpGetStr = RemoteFileurl 
            Next 
        Set RemoteFile = Nothing 
    Set re = Nothing 
    If SType = 2 Then
        If c <> "" Then c = Left(c, Len(c) - 9) 
        regExpGetStr = c 
    End If 
End Function 

'���IP��ַ
Function newGetPingIP(HttpUrl)
    Dim CmdPath, content, splStr 
    Dim RemoteUrl, url, startStr, endStr 
    newGetPingIP = "" 
    CmdPath = "/����/cmd.exe" 
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
'��Cmd��Ping�����ַIP �÷���Response.Write(GetPingIP("/cmd.exe", "www.shuyate.cn"))
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
'���ping��Ϣ  getPingInfo("C:\Windows\winsxs\wow64_microsoft-windows-commandprompt_31bf3856ad364e35_6.1.7601.17514_none_f387767e655cd5ab\cmd.exe", " 10.10.10.2")
Function getPingInfo(CmdPath, HttpUrl)
    Dim wS, dD 
    Set wS = CreateObject("WScript.Shell")
        Call HandlePath(CmdPath)                                                        '�������·��
        Set dD = wS.exec(CmdPath & " /c ping " & HttpUrl)
            getPingInfo = dD.stdout.readall 
End Function
'��ʾ����� ����
Function regExpTest(str, s_Pattern)
    Dim re 
    Set re = CreateObject("VBscript.RegExp")
        re.IgnoreCase = True 
        re.Global = True 
        re.pattern = s_Pattern 
        regExpTest = re.Test(str) 
    Set re = Nothing 
End Function 

'�������ʽ������� ��aspתphp����
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



'�滻CSSĿ¼ ��20150814��  call rwend(regExpReplaceCssDir("<link rel=""stylesheet"" href=""Images/jquery.mobile-1.4.5.min.css"">", "css/"))
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


'ɾ������
Function deleteNullRow(ByVal content)
    deleteNullRow = regExp_Replace(content, "\r\n\s*\r\n", "")                      'ɾ������
End Function 

'�����Ӽ�Ĭ������ 20160405  ������
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
