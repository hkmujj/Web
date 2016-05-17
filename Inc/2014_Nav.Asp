<% 
'����
'������վ����
Sub handleWebConfig()
    Dim did, sql 
    did = IIF(PubNavDid <> "", PubNavDid, PubProDid) 
    '�����վ��������
    Call GetNavConfig(did) 
    If PubId <> "" Then
        sql = "Select * From [Product] Where Id=" & PubId 
        rs.Open sql, conn, 1, 1 
        If Not rs.EOF Then
            '�����վ��������
            Call GetNavConfig(rs("BigClassName")) 
            MainNav = rs("BigClassName")                                                    '��������
        End If : rs.Close 
        Exit Sub 
    End If 
End Sub 
'�����վ��������
Function getNavConfig(did)
    RsTemp.Open "Select * From [NavBigClass] Where BigClassName='" & did & "'", conn, 1, 1 
    If Not RsTemp.EOF Then
        MainNav = RsTemp("BigClassName")                                                '��������
        PubNavType = RsTemp("NavType")                                                  '��վ��������
        PubNavDid = MainNav 
    End If : RsTemp.Close 
End Function 
'�����վ��������
Function getNavType(did)
    RsTemp.Open "Select * From [NavBigClass] Where BigClassName='" & did & "'", conn, 1, 1 
    If Not RsTemp.EOF Then
        getNavType = RsTemp("NavType")                                                  '��վ��������
    End If : RsTemp.Close 
End Function 

'VB������ Start
'������CSS���� ����VB����ʹ�ã�����վ�������� 20150617
Function readNavCSS(id, styleValue)
End Function 
'VB������ End

'������վ����  ��ǰʹ�ð汾  20141215�������޸Ŀ���
Function handleNavigation(action, addSql)
    Dim i, c, s, url, fileName, bigFolder, did, sid, showDid, target, sql, dropDownMenu, focusType 
    Dim smallMenuStr, j, showSid 
    Dim isConcise                                                                   '�����ʾ20150212
    Dim styleId, styleValue                                                         '��ʽID����ʽ����
    Dim cssNameAddId 



    styleId = PHPTrim(RParam(action, "styleID")) 
    styleValue = PHPTrim(RParam(action, "styleValue")) 
    If styleId <> "" Then
        Call ReadNavCSS(styleId, styleValue) 
    End If 

    'Ϊ�������� ���Զ���ȡ��ʽ����  20150615
    If checkStrIsNumberType(styleValue) Then
        cssNameAddId = "_" & styleValue                                                 'Css����׷��Id���
    End If 

    '����SQLΪ����  ����������
    If addSql = "" Then
        addSql = RParam(action, "AddSql") 
    End If 
    '׷��SQLΪ�� ����Ĭ��
    If addSql = "" Then
        addSql = " Where NavTop<>0 Order By Sort Asc" 
    End If 


    sql = "Select * From [NavBigClass]" 
    If UserId <> "" Then sql = GetWhereAnd(sql, " And UserId=" & UserId)            'ָ���û�
    sql = GetWhereAnd(sql, addSql) 
    If CheckSql(sql) = False Then Call eerr("Sql", sql) 
    dropDownMenu = LCase(RParam(action, "DropDownMenu")) 
    focusType = LCase(RParam(action, "FocusType")) 
    isConcise = LCase(RParam(action, "isConcise")) 
    If isConcise = "true" Then
        isConcise = False 
    Else
        isConcise = True 
    End If 


    rs.Open sql, conn, 1, 1 
    If isConcise = True Then c = c & CopyStr(" ", 4) & "<li class=left></li>" & vbCrLf 
    For i = 1 To rs.RecordCount
        fileName = Trim(rs("FileName")) 

        If fileName <> "" And Right(fileName, 1) <> "/" And Right(fileName, 1) <> "\" Then fileName = fileName & ".html" 
        bigFolder = "/" & GetWebFolderName() & rs("FolderName") & "/" 
        did = rs("BigClassName") 

        url = rs("FileName") 
        If CheckMakeHtmlFile(url) = True Then
            url = bigFolder & fileName 
            url = Replace(Replace(url, "//", "/"), "//", "/") 
            url = FullHttpUrl(host, url) 
        End If 
        Call WebDebug(url, "act=Nav&NavDid=" & did)                                 '����
        If InStr("|��Ʒ|����|����|", "|" & rs("NavType") & "|") > 0 Then
            Call WebDebug(url, "act=Nav&NavDid=" & did & "&Page=1")                 '����
        End If 

        showDid = FontColorFontB(did, rs("FontB"), rs("FontColor")) 
        target = rs("Target") 
        If PubNavDid = rs("BigClassName") Then
            If focusType = "a" Then
                s = CopyStr(" ", 8) & "<li class=focus><a " & AHref(url, showDid, target) & ">" & showDid & "</a>" 
            Else
                s = CopyStr(" ", 8) & "<li class=focus>" & showDid 
            End If 
        Else
            'Call Eerr("Url",Url)
            s = CopyStr(" ", 8) & "<li><a " & AHref(url, showDid, target) & ">" & showDid & "</a>" 
        End If 
        'S = DisplayOnlineEditDialog(""& adminDir &"NavManage.Asp?act=EditNavBig&Id=" & Rs("Id") & "&n=" & GetRnd(11), S)
        s = DisplayOnlineED2(WEB_ADMINURL &"NavManage.Asp?act=EditNavBig&Id=" & rs("Id") & "&n=" & getRnd(11), s, "<li") 

        c = c & s 

        smallMenuStr = ""                                                               'С��˵��������
        '����С���ı��˵�
        If rs("NavType") = "�ı�" And dropDownMenu = "true" Then
            rss.Open "Select * From [NavSmallClass] Where BigClassName='" & rs("BigClassName") & "'", conn, 1, 1 
            If Not rss.EOF Then
                'C=C & vbCrlf & CopyStr(" ",8) & "<ul>"& vbCrlf & CopyStr(" ",12) & "<li>" & vbCrlf
                c = c & vbCrLf & CopyStr(" ", 8) & "<ul>" & vbCrLf & CopyStr(" ", 12) 
                For j = 1 To rss.RecordCount
                    fileName = Trim(rss("FileName")) 

                    If fileName <> "" And Right(fileName, 1) <> "/" And Right(fileName, 1) <> "\" Then fileName = fileName & ".html" 

                    sid = rss("SmallClassName") 

                    url = rss("FileName") 
                    If CheckMakeHtmlFile(url) = True Then
                        url = bigFolder & fileName 
                        url = bigFolder & "/" & rss("FolderName") & "/" & fileName 
                        url = Replace(Replace(url, "//", "/"), "//", "/") 
                        Call WebDebug(url, "act=Nav&NavDid=" & did & "&NavSid=" & sid)          '����
                    End If 

                    showSid = FontColorFontB(sid, rss("FontB"), rss("FontColor")) 
                    target = rss("Target") : target = "_blank" 
                    s = "<a " & AHref(url, showSid, target) & ">" & showSid & "</a>" 
                    'Call eerr(ShowSid,Rss("Id"))
                    s = DisplayOnlineEditDialog(""& adminDir &"NavManage.Asp?act=EditNavSmall&Id=" & rss("Id") & "&n=" & getRnd(11), s) 
                    smallMenuStr = smallMenuStr & CopyStr(" ", 8) & "<li>" & s & "</li>" & vbCrLf 

                rss.MoveNext : Next 
            End If : rss.Close 
        '����С���Ʒ�˵�
        ElseIf rs("NavType") = "��Ʒ" And dropDownMenu = "true" Then
            rss.Open "Select * From [SmallClass] Where BigClassName='" & rs("BigClassName") & "'", conn, 1, 1 
            If Not rss.EOF Then
                c = c & vbCrLf & CopyStr(" ", 8) & "<ul>" & vbCrLf & CopyStr(" ", 12) 
                For j = 1 To rss.RecordCount
                    fileName = Trim(rss("FileName")) 

                    If fileName <> "" And Right(fileName, 1) <> "/" And Right(fileName, 1) <> "\" Then fileName = fileName & ".html" 
                    sid = rss("SmallClassName") 

                    url = rss("FileName") 
                    If CheckMakeHtmlFile(url) = True Then
                        url = bigFolder & fileName 
                        url = bigFolder & "/" & rss("FolderName") & "/" & fileName 
                        url = Replace(Replace(url, "//", "/"), "//", "/") 
                        Call WebDebug(url, "act=CreateClass&ProDid=" & did & "&ProSid=" & sid & "&Page=1") '����
                    End If 
                    'Call eerr(ShowSid,Rss("Id"))
                    showSid = FontColorFontB(sid, rss("FontB"), rss("FontColor")) 
                    target = rss("Target") : target = "_blank" 
                    s = "<a " & AHref(url, showSid, target) & ">" & showSid & "</a>" 
                    s = DisplayOnlineEditDialog(""& adminDir &"ProductClassManage.Asp?act=ShowEditSmallClass&Id=" & rss("Id") & "&n=" & getRnd(11), s) 
                    smallMenuStr = PHPRTrim(smallMenuStr) & vbCrLf & CopyStr(" ", 12) & "<li>" & s & "</li>" & vbCrLf 
                rss.MoveNext : Next 
            End If : rss.Close 
        End If 
        If smallMenuStr <> "" Then c = c & smallMenuStr & "</ul>" & vbCrLf 


        c = c & CopyStr(" ", 8) & "</li>" & vbCrLf 

        If isConcise = True Then c = c & CopyStr(" ", 8) & "<li class=line></li>" & vbCrLf 
    rs.MoveNext : Next : rs.Close 
    If isConcise = True Then c = c & CopyStr(" ", 8) & "<li class=right></li>" & vbCrLf 

    If styleId <> "" Then
        c = "<ul class='nav" & styleId & cssNameAddId & "'>" & vbCrLf & c & vbCrLf & "</ul>" & vbCrLf 
    End If 

    handleNavigation = c 
End Function 
'Div����վ����
Function divTopNavigation(NavDid, addSql)
    Dim i, c, s, url, fileName, bigFolder, did, showDid, target, sql 
    sql = "Select * From [NavBigClass]" 
    If UserId <> "" Then sql = GetWhereAnd(sql, " And UserId=" & UserId)            'ָ���û�
    sql = GetWhereAnd(sql, addSql) 
    If CheckSql(sql) = False Then Call eerr("Sql", sql) 
    rs.Open sql, conn, 1, 1 
    c = c & "<li class=left></li>" & vbCrLf 
    For i = 1 To rs.RecordCount
        fileName = rs("FileName") 
        If fileName <> "" And Right(fileName, 1) <> "/" Then fileName = fileName & ".html" 
        bigFolder = "/" & GetWebFolderName() & rs("FolderName") & "/" 
        did = rs("BigClassName") 

        url = rs("FileName") 
        If CheckRemoteUrl(url) = False Then
            url = bigFolder & fileName 
            url = Replace(Replace(url, "//", "/"), "//", "/") 
            url = FullHttpUrl(host, url) 
            Call WebDebug(url, "act=Nav&NavDid=" & did)                                 '����
        End If 

        showDid = FontColorFontB(did, rs("FontB"), rs("FontColor")) 
        target = rs("Target") 
        If NavDid = rs("BigClassName") Then
            c = c & "<li class=focus>" & showDid & "</li>" & vbCrLf 
        Else
            c = c & "<li><a " & AHref(url, showDid, target) & ">" & showDid & "</a></li>" & vbCrLf 
        End If 
        c = c & "<li class=line></li>" & vbCrLf 
    rs.MoveNext : Next : rs.Close 
    c = c & "<li class=right></li>" & vbCrLf 
    divTopNavigation = c 
End Function 
'��վ�ײ�����            '��ʱ����
Function webFootNavigation(addSql)
    Dim i, c, url, fileName, bigFolder, did, showDid, target, sql 
    sql = "Select * From [NavBigClass] " 
    If UserId <> "" Then sql = GetWhereAnd(sql, " And UserId=" & UserId)            'ָ���û�
    sql = GetWhereAnd(sql, addSql) 
    rs.Open sql, conn, 1, 1 
    For i = 1 To rs.RecordCount
        fileName = rs("FileName") 
        If fileName <> "" And Right(fileName, 1) <> "/" Then fileName = fileName & ".html" 

        bigFolder = "/" & rs("FolderName") & "/" 
        showDid = FontColorFontB(did, rs("FontB"), rs("FontColor")) 
        target = rs("Target") 
        did = rs("BigClassName") 

        url = rs("FileName") 
        If CheckRemoteUrl(url) = False Then
            url = bigFolder & fileName 
            url = Replace(Replace(url, "//", "/"), "//", "/") 
            Call WebDebug(url, "act=Nav&NavDid=" & did)                                 '����
        End If 

        c = c & "<a " & AHref(url, showDid, target) & ">" & showDid & "</a>" 
        If i <> rs.RecordCount Then c = c & " &nbsp;| &nbsp; " 
    rs.MoveNext : Next : rs.Close 
    webFootNavigation = c 
End Function 
'======================== ������ ==========================
'��õ������Ͷ�Ӧ�ĵ�������  ��
Function getNavTypeInName(NavType)
    If LCase(NavType) = "home" Then NavType = "��ҳ"                                '׷����20150310
    getNavTypeInName = GetNavParam("NavBigClass", "BigClassName", "Where NavType='" & NavType & "'") 
End Function 
'��õ�������
Function getNavParam(ByVal TableName, ByVal FieldName, ByVal addSql)
    Dim sql, fileName, url, bigFolder 
    sql = "Select * From [" & TableName & "] " & addSql 
    TempRs.Open sql, conn, 1, 1 
    If Not TempRs.EOF Then
        '�������ƻ����ַ
        If FieldName = "[url]" Then
            fileName = TempRs("FileName") 
            If fileName <> "" And Right(fileName, 1) <> "/" Then fileName = fileName & ".html" 
            bigFolder = "/" & GetWebFolderName() & TempRs("FolderName") & "/" 

            url = TempRs("FileName") 
            If CheckRemoteUrl(url) = False Then
                url = bigFolder & fileName 
                url = Replace(Replace(url, "//", "/"), "//", "/") 
                url = FullHttpUrl(host, url) 
                Call WebDebug(url, "act=Nav&NavDid=" & rs("BigClassName"))                  '����         Did������20150617
            End If 

            getNavParam = url 
        ElseIf FieldName <> "" Then
            getNavParam = TempRs(FieldName) 
        End If 
    End If : TempRs.Close 
End Function 





'��ʾ�����б� 201507 10 (�򵥰�)  ������
Function XY_CustomNavList(action)
    Dim did, sid, tid, title, topNumb, cutStrNumb, addSql 
    Call HandleFunParameter(action, did, sid, tid, title, topNumb, cutStrNumb, addSql) '�����ô��������
    Call XY_AutoAddHandle(action)                                                   '����������Ϣ����
    Dim defaultContent, i, j, s, c, startStr, endStr, url, aUrl, aStr, aImg, showTitle, dateTime, imgFile, defaultImage, rndShow, idList, sql 
    Dim DescriptionLabel                                                            '��ǩ
    Dim WebDescription                                                              '��վ����
    Dim Price                                                                       '�۸�
    Dim content                                                                     '����
    Dim noFollow                                                                    '��׷�� 20141222
    defaultContent = GetDefaultValue(action)                                        '���Ĭ������
    Dim AutoAddDid, AutoAddSid, AutoAddTid, AutoAdd 
    Dim ArticleDescription                                                          '��������(��Ҫ���)
    Dim modI                                                                        '��ѭ��20150112
    Dim HtmlCode                                                                    'HTML����
    Dim splStr, splxx 
    Dim bigFolder, fileName 

    noFollow = Trim(LCase(RParam(action, "noFollow")))                              '��׷��
    sql = "Select * From [NavBigClass] Where NavTop<>0 Order By Sort Asc" 
    'Call Eerr("Sql=",Sql)
    rs.Open sql, conn, 1, 1 
    'Call eerr("Sql��"& Rs.RecordCount &"��", Sql)
    For i = 1 To rs.RecordCount
        If rs.EOF Then Exit For 
        If rs("BigClassName") = PubNavDid Then
            startStr = "[list-focus]" : endStr = "[/list-focus]" 
        Else
            startStr = "[list-" & i & "]" : endStr = "[/list-" & i & "]" 
        End If 
        '��[list-mod2]  [/list-mod2]    20150112
        For modI = 6 To 2 Step - 1
            If InStr(defaultContent, startStr) = False And i Mod modI = 0 Then
                startStr = "[list-mod" & modI & "]" : endStr = "[/list-mod" & modI & "]" 
                If InStr(defaultContent, startStr) > 0 Then
                    'Call Echo("ModI=" & ModI & ",I=" & I,StartStr)
                    Exit For 
                End If 
            End If 
        Next 

        'û������Ĭ��
        If InStr(defaultContent, startStr) = False Then
            startStr = "[list]" : endStr = "[/list]" 
        End If 



        If InStr(defaultContent, startStr) > 0 And InStr(defaultContent, endStr) > 0 Then
            s = StrCut(defaultContent, startStr, endStr, 2) 


            fileName = Trim(rs("FileName")) 

            If fileName <> "" And Right(fileName, 1) <> "/" And Right(fileName, 1) <> "\" Then fileName = fileName & ".html" 
            bigFolder = "/" & GetWebFolderName() & rs("FolderName") & "/" 
            did = rs("BigClassName") 

            url = rs("FileName") 
            If CheckMakeHtmlFile(url) = True Then
                url = bigFolder & fileName 
                url = Replace(Replace(url, "//", "/"), "//", "/") 
                url = FullHttpUrl(host, url) 
            End If 
            '20151013�Ľ�
            If rs("CustomAUrl") <> "" Then
                url = rs("CustomAUrl") 
            Else
                Call WebDebug(url, "act=Nav&NavDid=" & did)                                 '����
            End If 

            If InStr("|��Ʒ|����|����|", "|" & rs("NavType") & "|") > 0 Then
                Call WebDebug(url, "act=Nav&NavDid=" & did & "&Page=1")                 '����
            End If 

            aUrl = AHref(url, rs("BigClassName"), rs("Target")) 

            aStr = "<a " & AHref(url, rs("BigClassName"), rs("Target")) & SetHtmlParam(action, "aclass") & IIF(noFollow = "true", " rel='nofollow'", "") & ">" & showTitle & "</a>" & vbCrLf 


            WebDescription = rs("WebDescription") 

            showTitle = CutStr(rs("BigClassName"), cutStrNumb, "... ") 
            showTitle = FontColorFontB(showTitle, rs("FontB"), rs("FontColor")) 



            For j = 1 To 3
                s = ReplaceValueParam(s, "ni", i)                                               '����Ϊi����Ϊi����imgurl��ͻ [$i$]
                s = ReplaceValueParam(s, "���-1", i - 1)                                       '����Ϊi����Ϊi����imgurl��ͻ [$i$]
                s = ReplaceValueParam(s, "���", i)                                             '����Ϊi����Ϊi����imgurl��ͻ [$i$]

                'S = ReplaceValueParam(S,"���-1%2",Fix(I/2)-1)                    '����Ϊi����Ϊi����imgurl��ͻ [$i$]




                s = ReplaceValueParam(s, "bigclassname", rs("BigClassName")) 
                s = ReplaceValueParam(s, "title", rs("BigClassName")) 
                s = ReplaceValueParam(s, "showtitle", showTitle) 
                s = ReplaceValueParam(s, "url", url) 
                s = ReplaceValueParam(s, "aurl", aUrl) 
                s = ReplaceValueParam(s, "astr", aStr) 
            Next 

            url = ""& adminDir &"Product.Asp?act=ShowEditProduct&Id=" & rs("Id") & "&n=" & getRnd(11) 
            'S = DisplayOnlineEditDialog(Url, S)

            s = DisplayOnlineEditDialog(""& adminDir &"NavManage.Asp?act=EditNavBig&Id=" & rs("Id") & "&n=" & getRnd(11), s) 


            '��ʼλ�ü�Dialog����
            startStr = "[list-" & i & " startdialog]" : endStr = "[/list-" & i & " startdialog]" 
            If InStr(defaultContent, startStr) > 0 And InStr(defaultContent, endStr) > 0 Then
                s = StrCut(defaultContent, startStr, endStr, 2) & s 
            End If 
            '����λ�ü�Dialog����
            startStr = "[list-" & i & " enddialog]" : endStr = "[/list-" & i & " enddialog]" 
            If InStr(defaultContent, startStr) > 0 And InStr(defaultContent, endStr) > 0 Then
                s = s & StrCut(defaultContent, startStr, endStr, 2) 
            End If 
            c = c & s 
        End If 
    rs.MoveNext : Next : rs.Close 

    '��ʼ���ݼ�Dialog����
    startStr = "[dialog start]" : endStr = "[/dialog start]" 
    If InStr(defaultContent, startStr) > 0 And InStr(defaultContent, endStr) > 0 Then
        c = StrCut(defaultContent, startStr, endStr, 2) & c 
    End If 
    '�������ݼ�Dialog����
    startStr = "[dialog end]" : endStr = "[/dialog end]" 
    If InStr(defaultContent, startStr) > 0 And InStr(defaultContent, endStr) > 0 Then
        c = c & StrCut(defaultContent, startStr, endStr, 2) 
    End If 
    XY_CustomNavList = c 
End Function 













'��õ������ģ������
Function getNavTemplateText(NavDid, NavSid, NavType)
    Dim addSql, TemplateFileName, fileName, NavBigFolderName, NavSmallFolder 
    If NavSid <> "" Then
        addSql = " Where BigClassName='" & NavDid & "' And SmallClassName='" & NavSid & "'" 
    ElseIf NavDid <> "" Then
        addSql = " Where BigClassName='" & NavDid & "'" 
    Else
        addSql = " Where NavType='" & NavType & "'" 
    End If 
    '��������
    If NavDid <> "" Then
        tempRs.Open "Select * From [NavBigClass] Where BigClassName='" & NavDid & "'", conn, 1, 1 
        If Not tempRs.EOF Then
            NavBigFolderName = "/" & tempRs("FolderName") & "/" 
            WebNavType = tempRs("NavType")                                                  '�����վ��������
            MainTitle = tempRs("BigClassName")                                              '��Ϣ��������
        End If : tempRs.Close 
    End If 
    '����С��
    If NavSid <> "" Then
        tempRs.Open "Select * From [NavSmallClass]  " & addSql, conn, 1, 1 
        If Not tempRs.EOF Then
            '�Զ���ģ��
            If Trim(tempRs("TemplatePath")) <> "" Then
                TemplateFileName = tempRs("TemplatePath") 
            Else
                TemplateFileName = "Main_Model.Html" 
            End If 
            fileName = GetHandleFileName(tempRs("FileName")) 

            If tempRs("FolderName") = "" Then
                NavSmallFolder = NavBigFolderName 
            Else
                NavSmallFolder = Trim(Replace(tempRs("FolderName"), "\", "/")) 
                If Left(NavSmallFolder, 1) <> "/" Then
                    NavSmallFolder = NavBigFolderName & "/" & NavSmallFolder 
                End If 
            End If 
            CreateHTMLPath = LocalTestFolder & NavSmallFolder & fileName 


            '����ת��ַ��Ϊ��(20151022)
            If Trim(tempRs("CustomAUrl")) <> "" Then
                CreateHTMLPath = "" 
            End If 
            'call echo("LocalTestFolder", LocalTestFolder)
            'call EErr("CreateHTMLPath",CreateHTMLPath)

            'CreateHTMLPath = GetHandleUrl(NavSmallFolder)    '����URL   �Ѿ��������ˣ������ٴ�����

            'Call eerr(NavBigFolderName & ",CreateHTMLPath",CreateHTMLPath)

            MainContent = tempRs("Content")                                                 '����ı�������
            '���������޸�
            MainContent = DisplayOnlineEditDialog("" & adminDir & "NavManage.Asp?act=SmallNavInfo&Id=" & tempRs("Id") & "&n=" & getRnd(11), MainContent) 



            '��ʾ��ƪ���µ���վ����ؼ�������
            WebTitle = IIf(tempRs("WebTitle") <> "", tempRs("WebTitle"), WebTitle) 
            WebKeywords = IIf(tempRs("WebKeywords") <> "", tempRs("WebKeywords"), WebKeywords) 
            WebDescription = IIf(tempRs("WebDescription") <> "", tempRs("WebDescription"), WebDescription) 

            MainTitle = tempRs("SmallClassName")                                            '��Ϣ��������(20151013)   ����д�÷ǳ������ˣ�����ȸĽ�
        End If : tempRs.Close 
        getNavTemplateText = TemplateFileName : Exit Function 
    End If 
    '��������
    rs.Open "Select * From [NavBigClass] " & addSql, conn, 1, 1 
    If Not rs.EOF Then
        '�Զ���ģ��
        If Trim(rs("TemplatePath")) <> "" Then
            TemplateFileName = rs("TemplatePath") 
        Else
            Select Case rs("NavType")
                Case "��ҳ" : TemplateFileName = "Index_Model.Html"
                Case "�ı�"
                    TemplateFileName = "Main_Model.Html" : MainContent = rs("Content")              '����ı�������
                    '���������޸�
                    MainContent = DisplayOnlineEditDialog("" & adminDir & "NavManage.Asp?act=EditInformation&Id=" & rs("Id") & "&n=" & getRnd(11), MainContent) 

                Case Else : TemplateFileName = "Main_Model.Html"                                'Ĭ��Ϊ���ģ��
            End Select
        End If 
        MainTitle = rs("BigClassName")                                                  '��Ϣ��������
        '����ļ�����
        fileName = GetHandleFileName(rs("FileName")) 
        '����һ���ļ���(�����ļ���·��)
        Call CreateDirFolder("/" & LocalTestFolder & rs("FolderName")) 
        CreateHTMLPath = "/" & LocalTestFolder & rs("FolderName") & "/" & fileName 
        CreateHTMLPath = GetHandleUrl(CreateHTMLPath)                                   '����URL

        '����ת��ַ��Ϊ��(20151022)
        If Trim(rs("CustomAUrl")) <> "" Then
            CreateHTMLPath = "" 
        End If 


        '��ʾ��ƪ���µ���վ����ؼ�������
        WebTitle = IIf(rs("WebTitle") <> "", rs("WebTitle"), WebTitle) 
        WebKeywords = IIf(rs("WebKeywords") <> "", rs("WebKeywords"), WebKeywords) 
        WebDescription = IIf(rs("WebDescription") <> "", rs("WebDescription"), WebDescription) 
    Else
        'û�д˵��� ��ΪĬ��ģ��
        TemplateFileName = "Index_Model.Html" 
        Call echo("û�е������ࡾ" & NavDid & "���˵���ģ��ΪĬ��", TemplateFileName) 
    End If : rs.Close 
    getNavTemplateText = TemplateFileName 
End Function 
'�����Ϣ���ģ������
Function getClassTemplateText(ProDid, ProSid, ProTid)
    Dim sql, TemplateFileName, fileName, RootFolder, dirNameStr 
    If ProTid <> "" Then
        sql = "Select * From [ThreeClass] Where BigClassName='" & ProDid & "' And SmallClassName='" & ProSid & "' And ThreeClassName='" & ProTid & "'" 
    ElseIf ProSid <> "" Then
        sql = "Select * From [SmallClass] Where BigClassName='" & ProDid & "' And SmallClassName='" & ProSid & "'" 
    Else
        sql = "Select * From [BigClass] Where BigClassName='" & ProDid & "'" 
    End If 

    '������
    If ProDid <> "" Then
        tempRs.Open "Select * From [BigClass] Where BigClassName='" & ProDid & "'", conn, 1, 1 
        If Not tempRs.EOF Then
            RootFolder = "/" & tempRs("FolderName") & "/" 
        End If : tempRs.Close 
    End If 
    '���С��
    If ProDid <> "" Then
        tempRs.Open "Select * From [SmallClass] Where BigClassName='" & ProDid & "' And SmallClassName='" & ProSid & "'", conn, 1, 1 
        If Not tempRs.EOF Then
            dirNameStr = Trim(Replace(tempRs("FolderName"), "\", "/")) 
            If Left(dirNameStr, 1) <> "/" Then
                RootFolder = RootFolder & "/" & dirNameStr & "/" 
            Else
                RootFolder = "/" & tempRs("FolderName") & "/" 
            End If 

        End If : tempRs.Close 
    End If 
    '�������
    If ProTid <> "" Then
        tempRs.Open "Select * From [ThreeClass] Where BigClassName='" & ProDid & "' And SmallClassName='" & ProSid & "' And ThreeClassName='" & ProTid & "'", conn, 1, 1 
        If Not tempRs.EOF Then

            dirNameStr = Trim(Replace(tempRs("FolderName"), "\", "/")) 
            If Left(dirNameStr, 1) <> "/" Then
                RootFolder = RootFolder & "/" & dirNameStr & "/" 
            Else
                RootFolder = "/" & tempRs("FolderName") & "/" 
            End If 

        End If : tempRs.Close 
    End If 
    If ProDid <> "" And ProSid = "" And ProTid = "" Then
        tempRs.Open "Select * From[NavBigClass] Where (NavType='��Ʒ' Or NavType='����' Or NavType='����') And BigClassName='" & ProDid & "'", conn, 1, 1 

        '֮ǰΪ��ʡ�£����������CheckMakeHtmlFile(TempRs("FileName"))�ж� �����ˣ��������Ա�ˣ������� ��Ӧ��ѽ��20141223
        If Not tempRs.EOF Then
            If CheckMakeHtmlFile(tempRs("FileName")) = True Then
                RootFolder = "/" & tempRs("FolderName") & "/" 
            Else
                tempRs.Close 
                tempRs.Open sql, conn, 1, 1 
            End If 
        Else
            tempRs.Close 
            tempRs.Open sql, conn, 1, 1 
        End If 
    Else
        tempRs.Open sql, conn, 1, 1 
    End If 

    If Not tempRs.EOF Then
        '�Զ���ģ��
        If Trim(tempRs("TemplatePath")) <> "" Then
            TemplateFileName = tempRs("TemplatePath") 
        Else
            TemplateFileName = "Main_Model.Html" 
        End If 
        fileName = GetHandleFileName(tempRs("FileName")) 
        '����һ���ļ���(�����ļ���·��)
        Call CreateDirFolder("/" & LocalTestFolder & RootFolder) 
        CreateHTMLPath = "/" & LocalTestFolder & RootFolder & "/" & fileName 
        CreateHTMLPath = GetHandleUrl(CreateHTMLPath)                                   '����URL

        'MainContent = TempRs("Content")        '����ı�������
        If tempRs("WebTitle") <> "" Then WebTitle = tempRs("WebTitle") 
        If tempRs("WebKeywords") <> "" Then WebKeywords = tempRs("WebKeywords") 
        If tempRs("WebDescription") <> "" Then WebDescription = tempRs("WebDescription") 
    End If : tempRs.Close 

    getClassTemplateText = TemplateFileName 
End Function 



'================================== 2015�� ==================================


'���������ϸҳ��ʽ 20150114
Function getArticleInfoStyle(did, sid, tid, id)
    Dim rs, addSql 
    Set rs = CreateObject("Adodb.RecordSet")
        '����
        addSql = " Where BigClassName='" & did & "' And ArticleInfoStyle<>'' And ArticleInfoStyle<>'0'" 
        rs.Open "Select * From [BigClass]" & addSql, conn, 1, 1 
        If Not rs.EOF Then
            getArticleInfoStyle = rs("ArticleInfoStyle") 
        End If : rs.Close 
        'С��
        If sid <> "" Then
            addSql = addSql & " And SmallClassName='" & sid & "'" 
            rs.Open "Select * From [SmallClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                getArticleInfoStyle = rs("ArticleInfoStyle") 
            End If : rs.Close 
        End If 
        '����
        If sid <> "" Then
            addSql = addSql & " And ThreeClassName='" & tid & "'" 
            rs.Open "Select * From [ThreeClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                getArticleInfoStyle = rs("ArticleInfoStyle") 
            End If : rs.Close 
        End If 
        '��Ʒ
        If id <> "" Then
            rs.Open "Select * From [Product] Where Id=" & id, conn, 1, 1 
            If Not rs.EOF Then
                If rs("ArticleInfoStyle") <> "" And rs("ArticleInfoStyle") <> "0" Then
                    getArticleInfoStyle = rs("ArticleInfoStyle") 
                End If 
            End If : rs.Close 
        End If 
        '���Ϊ������Ĭ��չʾ��ʽ
        If getArticleInfoStyle = "" Then
            getArticleInfoStyle = "Ĭ��������ϸ.html" 
        End If 
End Function
'���������ϸҳ�� 20150114
Function getArticleInfoPhotoWidth(did, sid, tid, id)
    Dim rs, addSql 
    Set rs = CreateObject("Adodb.RecordSet")
        '����
        addSql = " Where BigClassName='" & did & "' And ArticleInfoPhotoWidth<>'' And ArticleInfoPhotoWidth<>'0'" 
        rs.Open "Select * From [BigClass]" & addSql, conn, 1, 1 
        If Not rs.EOF Then
            If CInt(rs("ArticleInfoPhotoWidth")) > 0 Then
                getArticleInfoPhotoWidth = rs("ArticleInfoPhotoWidth") 
            End If 
        End If : rs.Close 
        'С��
        If sid <> "" Then
            addSql = addSql & " And SmallClassName='" & sid & "'" 
            rs.Open "Select * From [SmallClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                If CInt(rs("ArticleInfoPhotoWidth")) > 0 Then
                    getArticleInfoPhotoWidth = rs("ArticleInfoPhotoWidth") 
                End If 
            End If : rs.Close 
        End If 
        '����
        If sid <> "" Then
            addSql = addSql & " And ThreeClassName='" & tid & "'" 
            rs.Open "Select * From [ThreeClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                If CInt(rs("ArticleInfoPhotoWidth")) > 0 Then
                    getArticleInfoPhotoWidth = rs("ArticleInfoPhotoWidth") 
                End If 
            End If : rs.Close 
        End If 
        '��Ʒ
        If id <> "" Then
            rs.Open "Select * From [Product] Where Id=" & id, conn, 1, 1 
            If Not rs.EOF Then
                If(rs("ArticleInfoPhotoWidth") <> "" And rs("ArticleInfoPhotoWidth") <> "0") Or rs("ArticleInfoPhotoWidth") = "-1" Then
                    getArticleInfoPhotoWidth = rs("ArticleInfoPhotoWidth") 
                End If 
            End If : rs.Close 
        End If 
        If getArticleInfoPhotoWidth = "0" Or getArticleInfoPhotoWidth = "" Then getArticleInfoPhotoWidth = 600 
End Function


'�������չʾ�б�ÿҳ��ʾ���� 20150114
Function getArticlePageSize(did, sid, tid)
    Dim rs, addSql 
    Set rs = CreateObject("Adodb.RecordSet")
        '����
        addSql = " Where BigClassName='" & did & "' And ArticleListStyle<>''" 
        rs.Open "Select * From [BigClass]" & addSql, conn, 1, 1 
        If Not rs.EOF Then
            getArticlePageSize = rs("ArticlePageSize") 
        End If : rs.Close 
        'С��
        If sid <> "" Then
            addSql = addSql & " And SmallClassName='" & sid & "'" 
            rs.Open "Select * From [SmallClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                getArticlePageSize = rs("ArticlePageSize") 
            End If : rs.Close 
        End If 
        '����
        If sid <> "" Then
            addSql = addSql & " And ThreeClassName='" & tid & "'" 
            rs.Open "Select * From [ThreeClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                getArticlePageSize = rs("ArticlePageSize") 
            End If : rs.Close 
        End If 
        '��ǰÿҳ��ʾ��Ϊ�ջ�0ʱ��ΪĬ����ʾ9��
        If getArticlePageSize = "" Or getArticlePageSize = "0" Then
            getArticlePageSize = 9 
        End If 
End Function


'�������չʾ�б���ʽ 20150114   ���� S = ReadTemplateFileSource("��Ʒ��ϸ") ʹ��
Function getArticleListStyle(did, sid, tid)
    Dim rs, addSql 
    Set rs = CreateObject("Adodb.RecordSet")
        '�������� Ĭ�ϴ�������
        rs.Open "Select * From [NavBigClass] Where BigClassName='" & did & "'", conn, 1, 1 
        If Not rs.EOF Then
            If InStr("|��Ʒ|����|����|", "|" & rs("NavType") & "|") > 0 Then
                getArticleListStyle = "Ĭ��" & rs("NavType") & "�б�.html" 
            End If 
        End If : rs.Close 
        '����
        addSql = " Where BigClassName='" & did & "' And ArticleListStyle<>''" 
        rs.Open "Select * From [BigClass]" & addSql, conn, 1, 1 
        If Not rs.EOF Then
            getArticleListStyle = rs("ArticleListStyle") 
        End If : rs.Close 
        'С��
        If sid <> "" Then
            addSql = addSql & " And SmallClassName='" & sid & "'" 
            rs.Open "Select * From [SmallClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                getArticleListStyle = rs("ArticleListStyle") 
            End If : rs.Close 
        End If 
        '����
        If sid <> "" Then
            addSql = addSql & " And ThreeClassName='" & tid & "'" 
            rs.Open "Select * From [ThreeClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                getArticleListStyle = rs("ArticleListStyle") 
            End If : rs.Close 
        End If 
        '���Ϊ������Ĭ��չʾ��ʽ
        If getArticleListStyle = "" Then
            getArticleListStyle = "Ĭ�������б�.html" 
        End If 
End Function

%> 