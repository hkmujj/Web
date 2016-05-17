<% 
'导航
'处理网站配置
Sub handleWebConfig()
    Dim did, sql 
    did = IIF(PubNavDid <> "", PubNavDid, PubProDid) 
    '获得网站导航配置
    Call GetNavConfig(did) 
    If PubId <> "" Then
        sql = "Select * From [Product] Where Id=" & PubId 
        rs.Open sql, conn, 1, 1 
        If Not rs.EOF Then
            '获得网站导航配置
            Call GetNavConfig(rs("BigClassName")) 
            MainNav = rs("BigClassName")                                                    '导航大类
        End If : rs.Close 
        Exit Sub 
    End If 
End Sub 
'获得网站导航配置
Function getNavConfig(did)
    RsTemp.Open "Select * From [NavBigClass] Where BigClassName='" & did & "'", conn, 1, 1 
    If Not RsTemp.EOF Then
        MainNav = RsTemp("BigClassName")                                                '导航大类
        PubNavType = RsTemp("NavType")                                                  '网站导航类型
        PubNavDid = MainNav 
    End If : RsTemp.Close 
End Function 
'获得网站导航类型
Function getNavType(did)
    RsTemp.Open "Select * From [NavBigClass] Where BigClassName='" & did & "'", conn, 1, 1 
    If Not RsTemp.EOF Then
        getNavType = RsTemp("NavType")                                                  '网站导航类型
    End If : RsTemp.Close 
End Function 

'VB不引用 Start
'读导航CSS内容 配置VB软件使用，在网站里无意义 20150617
Function readNavCSS(id, styleValue)
End Function 
'VB不引用 End

'处理网站导航  当前使用版本  20141215加在线修改控制
Function handleNavigation(action, addSql)
    Dim i, c, s, url, fileName, bigFolder, did, sid, showDid, target, sql, dropDownMenu, focusType 
    Dim smallMenuStr, j, showSid 
    Dim isConcise                                                                   '简洁显示20150212
    Dim styleId, styleValue                                                         '样式ID与样式内容
    Dim cssNameAddId 



    styleId = PHPTrim(RParam(action, "styleID")) 
    styleValue = PHPTrim(RParam(action, "styleValue")) 
    If styleId <> "" Then
        Call ReadNavCSS(styleId, styleValue) 
    End If 

    '为数字类型 则自动提取样式内容  20150615
    If checkStrIsNumberType(styleValue) Then
        cssNameAddId = "_" & styleValue                                                 'Css名称追加Id编号
    End If 

    '参数SQL为空则  从配置里获得
    If addSql = "" Then
        addSql = RParam(action, "AddSql") 
    End If 
    '追加SQL为空 则用默认
    If addSql = "" Then
        addSql = " Where NavTop<>0 Order By Sort Asc" 
    End If 


    sql = "Select * From [NavBigClass]" 
    If UserId <> "" Then sql = GetWhereAnd(sql, " And UserId=" & UserId)            '指定用户
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
        Call WebDebug(url, "act=Nav&NavDid=" & did)                                 '调试
        If InStr("|产品|新闻|下载|", "|" & rs("NavType") & "|") > 0 Then
            Call WebDebug(url, "act=Nav&NavDid=" & did & "&Page=1")                 '调试
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

        smallMenuStr = ""                                                               '小类菜单内容清空
        '导航小类文本菜单
        If rs("NavType") = "文本" And dropDownMenu = "true" Then
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
                        Call WebDebug(url, "act=Nav&NavDid=" & did & "&NavSid=" & sid)          '调试
                    End If 

                    showSid = FontColorFontB(sid, rss("FontB"), rss("FontColor")) 
                    target = rss("Target") : target = "_blank" 
                    s = "<a " & AHref(url, showSid, target) & ">" & showSid & "</a>" 
                    'Call eerr(ShowSid,Rss("Id"))
                    s = DisplayOnlineEditDialog(""& adminDir &"NavManage.Asp?act=EditNavSmall&Id=" & rss("Id") & "&n=" & getRnd(11), s) 
                    smallMenuStr = smallMenuStr & CopyStr(" ", 8) & "<li>" & s & "</li>" & vbCrLf 

                rss.MoveNext : Next 
            End If : rss.Close 
        '导航小类产品菜单
        ElseIf rs("NavType") = "产品" And dropDownMenu = "true" Then
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
                        Call WebDebug(url, "act=CreateClass&ProDid=" & did & "&ProSid=" & sid & "&Page=1") '调试
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
'Div版网站导航
Function divTopNavigation(NavDid, addSql)
    Dim i, c, s, url, fileName, bigFolder, did, showDid, target, sql 
    sql = "Select * From [NavBigClass]" 
    If UserId <> "" Then sql = GetWhereAnd(sql, " And UserId=" & UserId)            '指定用户
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
            Call WebDebug(url, "act=Nav&NavDid=" & did)                                 '调试
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
'网站底部导航            '暂时无用
Function webFootNavigation(addSql)
    Dim i, c, url, fileName, bigFolder, did, showDid, target, sql 
    sql = "Select * From [NavBigClass] " 
    If UserId <> "" Then sql = GetWhereAnd(sql, " And UserId=" & UserId)            '指定用户
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
            Call WebDebug(url, "act=Nav&NavDid=" & did)                                 '调试
        End If 

        c = c & "<a " & AHref(url, showDid, target) & ">" & showDid & "</a>" 
        If i <> rs.RecordCount Then c = c & " &nbsp;| &nbsp; " 
    rs.MoveNext : Next : rs.Close 
    webFootNavigation = c 
End Function 
'======================== 辅助区 ==========================
'获得导航类型对应的导航名称  旧
Function getNavTypeInName(NavType)
    If LCase(NavType) = "home" Then NavType = "首页"                                '追加于20150310
    getNavTypeInName = GetNavParam("NavBigClass", "BigClassName", "Where NavType='" & NavType & "'") 
End Function 
'获得导航参数
Function getNavParam(ByVal TableName, ByVal FieldName, ByVal addSql)
    Dim sql, fileName, url, bigFolder 
    sql = "Select * From [" & TableName & "] " & addSql 
    TempRs.Open sql, conn, 1, 1 
    If Not TempRs.EOF Then
        '根据名称获得网址
        If FieldName = "[url]" Then
            fileName = TempRs("FileName") 
            If fileName <> "" And Right(fileName, 1) <> "/" Then fileName = fileName & ".html" 
            bigFolder = "/" & GetWebFolderName() & TempRs("FolderName") & "/" 

            url = TempRs("FileName") 
            If CheckRemoteUrl(url) = False Then
                url = bigFolder & fileName 
                url = Replace(Replace(url, "//", "/"), "//", "/") 
                url = FullHttpUrl(host, url) 
                Call WebDebug(url, "act=Nav&NavDid=" & rs("BigClassName"))                  '调试         Did被改了20150617
            End If 

            getNavParam = url 
        ElseIf FieldName <> "" Then
            getNavParam = TempRs(FieldName) 
        End If 
    End If : TempRs.Close 
End Function 





'显示导航列表 201507 10 (简单版)  待完善
Function XY_CustomNavList(action)
    Dim did, sid, tid, title, topNumb, cutStrNumb, addSql 
    Call HandleFunParameter(action, did, sid, tid, title, topNumb, cutStrNumb, addSql) '获得这么函数参数
    Call XY_AutoAddHandle(action)                                                   '处理添加信息处理
    Dim defaultContent, i, j, s, c, startStr, endStr, url, aUrl, aStr, aImg, showTitle, dateTime, imgFile, defaultImage, rndShow, idList, sql 
    Dim DescriptionLabel                                                            '标签
    Dim WebDescription                                                              '网站描述
    Dim Price                                                                       '价格
    Dim content                                                                     '内容
    Dim noFollow                                                                    '不追踪 20141222
    defaultContent = GetDefaultValue(action)                                        '获得默认内容
    Dim AutoAddDid, AutoAddSid, AutoAddTid, AutoAdd 
    Dim ArticleDescription                                                          '文章描述(简要简介)
    Dim modI                                                                        '余循环20150112
    Dim HtmlCode                                                                    'HTML代码
    Dim splStr, splxx 
    Dim bigFolder, fileName 

    noFollow = Trim(LCase(RParam(action, "noFollow")))                              '不追踪
    sql = "Select * From [NavBigClass] Where NavTop<>0 Order By Sort Asc" 
    'Call Eerr("Sql=",Sql)
    rs.Open sql, conn, 1, 1 
    'Call eerr("Sql【"& Rs.RecordCount &"】", Sql)
    For i = 1 To rs.RecordCount
        If rs.EOF Then Exit For 
        If rs("BigClassName") = PubNavDid Then
            startStr = "[list-focus]" : endStr = "[/list-focus]" 
        Else
            startStr = "[list-" & i & "]" : endStr = "[/list-" & i & "]" 
        End If 
        '例[list-mod2]  [/list-mod2]    20150112
        For modI = 6 To 2 Step - 1
            If InStr(defaultContent, startStr) = False And i Mod modI = 0 Then
                startStr = "[list-mod" & modI & "]" : endStr = "[/list-mod" & modI & "]" 
                If InStr(defaultContent, startStr) > 0 Then
                    'Call Echo("ModI=" & ModI & ",I=" & I,StartStr)
                    Exit For 
                End If 
            End If 
        Next 

        '没有则用默认
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
            '20151013改进
            If rs("CustomAUrl") <> "" Then
                url = rs("CustomAUrl") 
            Else
                Call WebDebug(url, "act=Nav&NavDid=" & did)                                 '调试
            End If 

            If InStr("|产品|新闻|下载|", "|" & rs("NavType") & "|") > 0 Then
                Call WebDebug(url, "act=Nav&NavDid=" & did & "&Page=1")                 '调试
            End If 

            aUrl = AHref(url, rs("BigClassName"), rs("Target")) 

            aStr = "<a " & AHref(url, rs("BigClassName"), rs("Target")) & SetHtmlParam(action, "aclass") & IIF(noFollow = "true", " rel='nofollow'", "") & ">" & showTitle & "</a>" & vbCrLf 


            WebDescription = rs("WebDescription") 

            showTitle = CutStr(rs("BigClassName"), cutStrNumb, "... ") 
            showTitle = FontColorFontB(showTitle, rs("FontB"), rs("FontColor")) 



            For j = 1 To 3
                s = ReplaceValueParam(s, "ni", i)                                               '不对为i，因为i会与imgurl冲突 [$i$]
                s = ReplaceValueParam(s, "编号-1", i - 1)                                       '不对为i，因为i会与imgurl冲突 [$i$]
                s = ReplaceValueParam(s, "编号", i)                                             '不对为i，因为i会与imgurl冲突 [$i$]

                'S = ReplaceValueParam(S,"编号-1%2",Fix(I/2)-1)                    '不对为i，因为i会与imgurl冲突 [$i$]




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


            '开始位置加Dialog内容
            startStr = "[list-" & i & " startdialog]" : endStr = "[/list-" & i & " startdialog]" 
            If InStr(defaultContent, startStr) > 0 And InStr(defaultContent, endStr) > 0 Then
                s = StrCut(defaultContent, startStr, endStr, 2) & s 
            End If 
            '结束位置加Dialog内容
            startStr = "[list-" & i & " enddialog]" : endStr = "[/list-" & i & " enddialog]" 
            If InStr(defaultContent, startStr) > 0 And InStr(defaultContent, endStr) > 0 Then
                s = s & StrCut(defaultContent, startStr, endStr, 2) 
            End If 
            c = c & s 
        End If 
    rs.MoveNext : Next : rs.Close 

    '开始内容加Dialog内容
    startStr = "[dialog start]" : endStr = "[/dialog start]" 
    If InStr(defaultContent, startStr) > 0 And InStr(defaultContent, endStr) > 0 Then
        c = StrCut(defaultContent, startStr, endStr, 2) & c 
    End If 
    '结束内容加Dialog内容
    startStr = "[dialog end]" : endStr = "[/dialog end]" 
    If InStr(defaultContent, startStr) > 0 And InStr(defaultContent, endStr) > 0 Then
        c = c & StrCut(defaultContent, startStr, endStr, 2) 
    End If 
    XY_CustomNavList = c 
End Function 













'获得导航类别模板内容
Function getNavTemplateText(NavDid, NavSid, NavType)
    Dim addSql, TemplateFileName, fileName, NavBigFolderName, NavSmallFolder 
    If NavSid <> "" Then
        addSql = " Where BigClassName='" & NavDid & "' And SmallClassName='" & NavSid & "'" 
    ElseIf NavDid <> "" Then
        addSql = " Where BigClassName='" & NavDid & "'" 
    Else
        addSql = " Where NavType='" & NavType & "'" 
    End If 
    '导航大类
    If NavDid <> "" Then
        tempRs.Open "Select * From [NavBigClass] Where BigClassName='" & NavDid & "'", conn, 1, 1 
        If Not tempRs.EOF Then
            NavBigFolderName = "/" & tempRs("FolderName") & "/" 
            WebNavType = tempRs("NavType")                                                  '获得网站导航类型
            MainTitle = tempRs("BigClassName")                                              '信息标题名称
        End If : tempRs.Close 
    End If 
    '导航小类
    If NavSid <> "" Then
        tempRs.Open "Select * From [NavSmallClass]  " & addSql, conn, 1, 1 
        If Not tempRs.EOF Then
            '自定义模板
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


            '有跳转网址则为空(20151022)
            If Trim(tempRs("CustomAUrl")) <> "" Then
                CreateHTMLPath = "" 
            End If 
            'call echo("LocalTestFolder", LocalTestFolder)
            'call EErr("CreateHTMLPath",CreateHTMLPath)

            'CreateHTMLPath = GetHandleUrl(NavSmallFolder)    '处理URL   已经处理过了，无需再处理了

            'Call eerr(NavBigFolderName & ",CreateHTMLPath",CreateHTMLPath)

            MainContent = tempRs("Content")                                                 '获得文本中内容
            '添加在线修改
            MainContent = DisplayOnlineEditDialog("" & adminDir & "NavManage.Asp?act=SmallNavInfo&Id=" & tempRs("Id") & "&n=" & getRnd(11), MainContent) 



            '显示这篇文章的网站标题关键词描述
            WebTitle = IIf(tempRs("WebTitle") <> "", tempRs("WebTitle"), WebTitle) 
            WebKeywords = IIf(tempRs("WebKeywords") <> "", tempRs("WebKeywords"), WebKeywords) 
            WebDescription = IIf(tempRs("WebDescription") <> "", tempRs("WebDescription"), WebDescription) 

            MainTitle = tempRs("SmallClassName")                                            '信息标题名称(20151013)   程序写得非常的乱了，等深度改进
        End If : tempRs.Close 
        getNavTemplateText = TemplateFileName : Exit Function 
    End If 
    '导航大类
    rs.Open "Select * From [NavBigClass] " & addSql, conn, 1, 1 
    If Not rs.EOF Then
        '自定义模板
        If Trim(rs("TemplatePath")) <> "" Then
            TemplateFileName = rs("TemplatePath") 
        Else
            Select Case rs("NavType")
                Case "首页" : TemplateFileName = "Index_Model.Html"
                Case "文本"
                    TemplateFileName = "Main_Model.Html" : MainContent = rs("Content")              '获得文本中内容
                    '添加在线修改
                    MainContent = DisplayOnlineEditDialog("" & adminDir & "NavManage.Asp?act=EditInformation&Id=" & rs("Id") & "&n=" & getRnd(11), MainContent) 

                Case Else : TemplateFileName = "Main_Model.Html"                                '默认为这个模板
            End Select
        End If 
        MainTitle = rs("BigClassName")                                                  '信息标题名称
        '获得文件名称
        fileName = GetHandleFileName(rs("FileName")) 
        '创建一组文件夹(导航文件夹路径)
        Call CreateDirFolder("/" & LocalTestFolder & rs("FolderName")) 
        CreateHTMLPath = "/" & LocalTestFolder & rs("FolderName") & "/" & fileName 
        CreateHTMLPath = GetHandleUrl(CreateHTMLPath)                                   '处理URL

        '有跳转网址则为空(20151022)
        If Trim(rs("CustomAUrl")) <> "" Then
            CreateHTMLPath = "" 
        End If 


        '显示这篇文章的网站标题关键词描述
        WebTitle = IIf(rs("WebTitle") <> "", rs("WebTitle"), WebTitle) 
        WebKeywords = IIf(rs("WebKeywords") <> "", rs("WebKeywords"), WebKeywords) 
        WebDescription = IIf(rs("WebDescription") <> "", rs("WebDescription"), WebDescription) 
    Else
        '没有此导航 则为默认模板
        TemplateFileName = "Index_Model.Html" 
        Call echo("没有导航大类【" & NavDid & "】此导航模板为默认", TemplateFileName) 
    End If : rs.Close 
    getNavTemplateText = TemplateFileName 
End Function 
'获得信息类别模板内容
Function getClassTemplateText(ProDid, ProSid, ProTid)
    Dim sql, TemplateFileName, fileName, RootFolder, dirNameStr 
    If ProTid <> "" Then
        sql = "Select * From [ThreeClass] Where BigClassName='" & ProDid & "' And SmallClassName='" & ProSid & "' And ThreeClassName='" & ProTid & "'" 
    ElseIf ProSid <> "" Then
        sql = "Select * From [SmallClass] Where BigClassName='" & ProDid & "' And SmallClassName='" & ProSid & "'" 
    Else
        sql = "Select * From [BigClass] Where BigClassName='" & ProDid & "'" 
    End If 

    '类别大类
    If ProDid <> "" Then
        tempRs.Open "Select * From [BigClass] Where BigClassName='" & ProDid & "'", conn, 1, 1 
        If Not tempRs.EOF Then
            RootFolder = "/" & tempRs("FolderName") & "/" 
        End If : tempRs.Close 
    End If 
    '类别小类
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
    '类别子类
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
        tempRs.Open "Select * From[NavBigClass] Where (NavType='产品' Or NavType='新闻' Or NavType='下载') And BigClassName='" & ProDid & "'", conn, 1, 1 

        '之前为了省事，这里面加了CheckMakeHtmlFile(TempRs("FileName"))判断 出错了，多年程序员了，还这样 不应该呀。20141223
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
        '自定义模板
        If Trim(tempRs("TemplatePath")) <> "" Then
            TemplateFileName = tempRs("TemplatePath") 
        Else
            TemplateFileName = "Main_Model.Html" 
        End If 
        fileName = GetHandleFileName(tempRs("FileName")) 
        '创建一组文件夹(导航文件夹路径)
        Call CreateDirFolder("/" & LocalTestFolder & RootFolder) 
        CreateHTMLPath = "/" & LocalTestFolder & RootFolder & "/" & fileName 
        CreateHTMLPath = GetHandleUrl(CreateHTMLPath)                                   '处理URL

        'MainContent = TempRs("Content")        '获得文本中内容
        If tempRs("WebTitle") <> "" Then WebTitle = tempRs("WebTitle") 
        If tempRs("WebKeywords") <> "" Then WebKeywords = tempRs("WebKeywords") 
        If tempRs("WebDescription") <> "" Then WebDescription = tempRs("WebDescription") 
    End If : tempRs.Close 

    getClassTemplateText = TemplateFileName 
End Function 



'================================== 2015年 ==================================


'获得文章详细页样式 20150114
Function getArticleInfoStyle(did, sid, tid, id)
    Dim rs, addSql 
    Set rs = CreateObject("Adodb.RecordSet")
        '大类
        addSql = " Where BigClassName='" & did & "' And ArticleInfoStyle<>'' And ArticleInfoStyle<>'0'" 
        rs.Open "Select * From [BigClass]" & addSql, conn, 1, 1 
        If Not rs.EOF Then
            getArticleInfoStyle = rs("ArticleInfoStyle") 
        End If : rs.Close 
        '小类
        If sid <> "" Then
            addSql = addSql & " And SmallClassName='" & sid & "'" 
            rs.Open "Select * From [SmallClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                getArticleInfoStyle = rs("ArticleInfoStyle") 
            End If : rs.Close 
        End If 
        '子类
        If sid <> "" Then
            addSql = addSql & " And ThreeClassName='" & tid & "'" 
            rs.Open "Select * From [ThreeClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                getArticleInfoStyle = rs("ArticleInfoStyle") 
            End If : rs.Close 
        End If 
        '产品
        If id <> "" Then
            rs.Open "Select * From [Product] Where Id=" & id, conn, 1, 1 
            If Not rs.EOF Then
                If rs("ArticleInfoStyle") <> "" And rs("ArticleInfoStyle") <> "0" Then
                    getArticleInfoStyle = rs("ArticleInfoStyle") 
                End If 
            End If : rs.Close 
        End If 
        '最后为空则用默认展示样式
        If getArticleInfoStyle = "" Then
            getArticleInfoStyle = "默认新闻详细.html" 
        End If 
End Function
'获得文章详细页宽 20150114
Function getArticleInfoPhotoWidth(did, sid, tid, id)
    Dim rs, addSql 
    Set rs = CreateObject("Adodb.RecordSet")
        '大类
        addSql = " Where BigClassName='" & did & "' And ArticleInfoPhotoWidth<>'' And ArticleInfoPhotoWidth<>'0'" 
        rs.Open "Select * From [BigClass]" & addSql, conn, 1, 1 
        If Not rs.EOF Then
            If CInt(rs("ArticleInfoPhotoWidth")) > 0 Then
                getArticleInfoPhotoWidth = rs("ArticleInfoPhotoWidth") 
            End If 
        End If : rs.Close 
        '小类
        If sid <> "" Then
            addSql = addSql & " And SmallClassName='" & sid & "'" 
            rs.Open "Select * From [SmallClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                If CInt(rs("ArticleInfoPhotoWidth")) > 0 Then
                    getArticleInfoPhotoWidth = rs("ArticleInfoPhotoWidth") 
                End If 
            End If : rs.Close 
        End If 
        '子类
        If sid <> "" Then
            addSql = addSql & " And ThreeClassName='" & tid & "'" 
            rs.Open "Select * From [ThreeClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                If CInt(rs("ArticleInfoPhotoWidth")) > 0 Then
                    getArticleInfoPhotoWidth = rs("ArticleInfoPhotoWidth") 
                End If 
            End If : rs.Close 
        End If 
        '产品
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


'获得文章展示列表每页显示条数 20150114
Function getArticlePageSize(did, sid, tid)
    Dim rs, addSql 
    Set rs = CreateObject("Adodb.RecordSet")
        '大类
        addSql = " Where BigClassName='" & did & "' And ArticleListStyle<>''" 
        rs.Open "Select * From [BigClass]" & addSql, conn, 1, 1 
        If Not rs.EOF Then
            getArticlePageSize = rs("ArticlePageSize") 
        End If : rs.Close 
        '小类
        If sid <> "" Then
            addSql = addSql & " And SmallClassName='" & sid & "'" 
            rs.Open "Select * From [SmallClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                getArticlePageSize = rs("ArticlePageSize") 
            End If : rs.Close 
        End If 
        '子类
        If sid <> "" Then
            addSql = addSql & " And ThreeClassName='" & tid & "'" 
            rs.Open "Select * From [ThreeClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                getArticlePageSize = rs("ArticlePageSize") 
            End If : rs.Close 
        End If 
        '当前每页显示数为空或0时则为默认显示9条
        If getArticlePageSize = "" Or getArticlePageSize = "0" Then
            getArticlePageSize = 9 
        End If 
End Function


'获得文章展示列表样式 20150114   配置 S = ReadTemplateFileSource("产品详细") 使用
Function getArticleListStyle(did, sid, tid)
    Dim rs, addSql 
    Set rs = CreateObject("Adodb.RecordSet")
        '导航大类 默认大类类型
        rs.Open "Select * From [NavBigClass] Where BigClassName='" & did & "'", conn, 1, 1 
        If Not rs.EOF Then
            If InStr("|产品|新闻|下载|", "|" & rs("NavType") & "|") > 0 Then
                getArticleListStyle = "默认" & rs("NavType") & "列表.html" 
            End If 
        End If : rs.Close 
        '大类
        addSql = " Where BigClassName='" & did & "' And ArticleListStyle<>''" 
        rs.Open "Select * From [BigClass]" & addSql, conn, 1, 1 
        If Not rs.EOF Then
            getArticleListStyle = rs("ArticleListStyle") 
        End If : rs.Close 
        '小类
        If sid <> "" Then
            addSql = addSql & " And SmallClassName='" & sid & "'" 
            rs.Open "Select * From [SmallClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                getArticleListStyle = rs("ArticleListStyle") 
            End If : rs.Close 
        End If 
        '子类
        If sid <> "" Then
            addSql = addSql & " And ThreeClassName='" & tid & "'" 
            rs.Open "Select * From [ThreeClass]" & addSql, conn, 1, 1 
            If Not rs.EOF Then
                getArticleListStyle = rs("ArticleListStyle") 
            End If : rs.Close 
        End If 
        '最后为空则用默认展示样式
        If getArticleListStyle = "" Then
            getArticleListStyle = "默认新闻列表.html" 
        End If 
End Function

%> 
