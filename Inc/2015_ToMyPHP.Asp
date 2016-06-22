<% 
'与php通用   我的后台

'获得网站底部内容aa
Function XY_AP_WebSiteBottom(action)
    If InStr(cfg_websitebottom, "[$aoutadd$]") > 0 Then
        cfg_websitebottom = getDefaultValue(action)                                     '获得默认内容
        conn.Execute("update " & db_PREFIX & "website set websitebottom='" & ADSql(cfg_websitebottom) & "'") 
    End If 
    XY_AP_WebSiteBottom = cfg_websitebottom 
End Function 

'asp与php版本
function XY_EDITORTYPE(action)
	dim aspValue,phpValue,s
	aspValue = LCase(RParam(action, "asp")) 
	phpValue = LCase(RParam(action, "php")) 
	if EDITORTYPE="asp" then
		s=aspValue
	else
		s=phpValue
	end if
	XY_EDITORTYPE=s
end function


'加载文件
Function XY_Include(action)
    Dim templateFilePath, Block, startStr, endStr, content 
    templateFilePath = LCase(RParam(action, "File")) 
    Block = LCase(RParam(action, "Block")) 

    Dim findstr, replaceStr                                                         '查找字符，替换字符
    findstr = moduleFindContent(action, "findstr")                                  '先找块
    replaceStr = moduleFindContent(action, "replacestr")                            '先找块

    templateFilePath = handleFileUrl(templateFilePath)                              '处理文件路径
    If checkFile(templateFilePath) = False Then
        templateFilePath = webTemplate & templateFilePath 
    End If 
    content = getFText(templateFilePath) 
    If Block <> "" Then
        startStr = "<!--#" & Block & " start#-->" 
        endStr = "<!--#" & Block & " end#-->" 
        If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
            content = strCut(content, startStr, endStr, 2) 
        End If 
    End If 
    '替换读出来的内容
    If findstr <> "" Then
        content = Replace(content, findstr, replaceStr) 
    End If 

    XY_Include = content 
End Function 

'栏目菜单
Function XY_AP_ColumnMenu(action)
    Dim defaultStr, thisId, parentid 
    parentid = Trim(RParam(action, "parentid")) 
    If parentid = "" Then parentid = -1 

    thisId = glb_columnId 
    If thisId = "" Then thisId = -1 
    defaultStr = getDefaultValue(action)                                            '获得默认内容
    XY_AP_ColumnMenu = showColumnList(parentid, thisId, 0, defaultStr) 
End Function 



'显示栏目列表
Function XY_AP_ColumnList(action)
    Dim sql, flags, addSql, columnname 
    sql = RParam(action, "sql") 
    flags = RParam(action, "flags") 
    addSql = RParam(action, "addSql") 
    columnname = RParam(action, "columnname") 
    If flags <> "" Then
        sql = " where flags like'%" & flags & "%'" 
    End If 
    If columnname <> "" Then
        sql = getWhereAnd(sql, "where parentid=" & getColumnId(columnname)) 
    'call echo(sql,columnName)
    End If 
    '追加sql
    If addSql <> "" Then
        sql = getWhereAnd(sql, addSql) 
    End If 
    XY_AP_ColumnList = XY_AP_GeneralList(action, "WebColumn", sql) 

End Function 

'显示文章列表
Function XY_AP_ArticleList(action)
    Dim sql, addSql, columnName, columnId, topNumb, idRand, splStr, s, columnIdList 
    action = replaceGlobleVariable(action)                                          '处理下替换标签
    sql = RParam(action, "sql") 
    topNumb = RParam(action, "topNumb") 


    'id随机
    idRand = LCase(RParam(action, "rand")) 
    If idRand = "true" Or idRand = "1" Then
        sql = sql & " where id in(" & getRandArticleId("", topNumb) & ")" 
    End If 

    '栏目名称 对栏目数组处理如 模板分享下载[Array]CSS3[Array]HTML5
    s = RParam(action, "columnName") 
    If s = "" Then
        s = RParam(action, "did") 
    End If 
    If s <> "" Then
        splStr = Split(s, "[Array]") 
        For Each columnName In splStr
            columnId = getColumnId(columnName) 
            If columnId <> "" Then
                If columnIdList <> "" Then
                    columnIdList = columnIdList & "," 
                End If 
                columnIdList = columnIdList & columnId 
            End If 
        Next 
    End If 
    If columnIdList <> "" Then
        sql = getWhereAnd(sql, "where parentId in(" & columnIdList & ")") 
    End If 
    '追加sql
    addSql = RParam(action, "addSql") 
    If addSql <> "" Then
        sql = getWhereAnd(sql, addSql) 
    End If 
    sql = replaceGlobleVariable(sql) 
    'call echo(RParam(action, "columnName") ,sql)
    XY_AP_ArticleList = XY_AP_GeneralList(action, "ArticleDetail", sql) 
End Function 

'显示评论列表
Function XY_AP_CommentList(action)
    Dim itemID, sql, addSql 
    addSql = RParam(action, "addsql") 
    itemID = RParam(action, "itemID") 
    itemID = replaceGlobleVariable(itemID) 

    If itemID <> "" Then
        sql = " where itemID=" & itemID 
    End If 
    '追加sql
    If addSql <> "" Then
        sql = getWhereAnd(sql, addSql) 
    End If 
    XY_AP_CommentList = XY_AP_GeneralList(action, "TableComment", sql) 
End Function 

'显示搜索统计
Function XY_AP_SearchStatList(action)
    Dim addSql 
    addSql = RParam(action, "addSql") 
    XY_AP_SearchStatList = XY_AP_GeneralList(action, "SearchStat", addSql) 
End Function 
'显示友情链接
Function XY_AP_Links(action)
    Dim addSql 
    addSql = RParam(action, "addSql") 
    XY_AP_Links = XY_AP_GeneralList(action, "FriendLink", addSql) 
End Function 
 
'通用信息列表
Function XY_AP_GeneralList(action, tableName, addSql)
    Dim title, topNumb, nTop, isB, sql 
    Dim columnName, columnEnName, aboutcontent, bodyContent, showTitle 
    Dim bannerImage, smallImage, bigImage, id 
    Dim defaultStr, i, j, s, c, startStr, endStr, url 
    Dim noFollow                                                                    '不追踪 20141222
    defaultStr = getDefaultValue(action)                                            '获得默认内容
    Dim modI                                                                        '余循环20150112
    noFollow = Trim(LCase(RParam(action, "noFollow")))                              '不追踪
    Dim lableTitle                                                                  '标题标题
    Dim target                                                                      'a链接打开目标方式
    Dim adddatetime                                                                 '添加时间
    Dim isFocus 
    Dim fieldNameList                                                               '字段列表
    Dim abcolorStr                                                                  'A加粗和颜色
    Dim atargetStr                                                                  'A链接打开方式
    Dim atitleStr                                                                   'A链接的title20160407
    Dim anofollowStr                                                                'A链接的nofollow
    Dim splFieldName, fieldName, replaceStr, k 

    fieldNameList = getHandleFieldList(db_PREFIX & tableName, "字段列表") 
    splFieldName = Split(fieldNameList, ",") 

    topNumb = RParam(action, "topNumb") : nTop = topNumb 
    If nTop <> "" Then
        nTop = CInt(nTop) 
    Else
        nTop = 999 
    End If 
    If sql = "" Then
        If topNumb <> "" Then
            topNumb = " top " & topNumb & " " 
        End If 
        sql = "Select " & topNumb & "* From " & db_PREFIX & tableName 
    End If 
    '追加sql
    If addSql <> "" Then
        sql = getWhereAnd(sql, addSql) 
    End If 
    sql = replaceGlobleVariable(sql)                                                '替换全局变量

    '检测SQL
    If checksql(sql) = False Then
        Call errorLog("出错提示：<br>action=" & action & "<hr>sql=" & sql & "<br>") 
        Exit Function 
    End If 
    rs.Open sql, conn, 1, 1 
    For i = 1 To rs.RecordCount
        'call echo(sql,i & "," & nTop)
        If i > nTop Then
            Exit For 
        End If 
        '【PHP】$rs=mysql_fetch_array($rsObj);                                            //给PHP用，因为在 asptophp转换不完善
        isFocus = False                                                                 '交点为假
        id = rs("id") 
        '【导航】
        If tableName = "WebColumn" Then
            If isMakeHtml = True Then
                url = getRsUrl(rs("fileName"), rs("customAUrl"), "/nav" & rs("id")) 
            Else
                url = handleWebUrl("?act=nav&columnName=" & rs("columnname"))               '会追加gl等参数
                If rs("customAUrl") <> "" Then
                    url = rs("customAUrl") 
                    url = replaceGlobleVariable(url) 
                End If 
            End If 
            '全局栏目名称为空则为自动定位首页 追加(20160128)
            If glb_columnName = "" And rs("columnType") = "首页" Then
                glb_columnName = rs("columnName") 
            End If 
            If rs("columnName") = glb_columnName Then
                isFocus = True 
            End If 
        '【文章】
        ElseIf tableName = "ArticleDetail" Then
            If isMakeHtml = True Then
                url = getRsUrl(rs("fileName"), rs("customAUrl"), "detail/detail" & rs("id")) 
            Else
                url = handleWebUrl("?act=detail&id=" & rs("id"))                            '会追加gl等参数
                If rs("customAUrl") <> "" Then
                    url = rs("customAUrl") 
                End If 
            End If 
        '评论
        ElseIf tableName = "TableComment" Then

        End If 

        'A链接添加颜色
        abcolorStr = "" 
        If InStr(fieldNameList, ",titlecolor,") > 0 Then
            'A链接颜色
            If rs("titlecolor") <> "" Then
                abcolorStr = "color:" & rs("titlecolor") & ";" 
            End If 
        End If 
        If InStr(fieldNameList, ",flags,") > 0 Then
            'A链接加粗
            If InStr(rs("flags"), "|b|") > 0 Then
                abcolorStr = abcolorStr & "font-weight:bold;" 
            End If 
        End If 
        If abcolorStr <> "" Then
            abcolorStr = " style=""" & abcolorStr & """" 
        End If 

        '打开方式2016
        If InStr(fieldNameList, ",target,") > 0 Then
            atargetStr = IIF(rs("target") <> "", " target=""" & rs("target") & """", "") 
        End If 

        'A的title
        If InStr(fieldNameList, ",title,") > 0 Then
            atitleStr = IIF(rs("title") <> "", " title=""" & rs("title") & """", "") 
        End If 

        'A的nofollow
        If InStr(fieldNameList, ",nofollow,") > 0 Then
            anofollowStr = IIF(rs("nofollow") <> 0, " rel=""nofollow""", "") 
        End If 

        '交点判断(给栏目导航用的)
        If isFocus = True Then
            startStr = "[list-focus]" : endStr = "[/list-focus]" 
        Else
            startStr = "[list-" & i & "]" : endStr = "[/list-" & i & "]" 
        End If 

        '在最后时排序当前交点20160202
        If i = topNumb And isFocus = False Then
            startStr = "[list-end]" : endStr = "[/list-end]" 
        End If 

        '例[list-mod2]  [/list-mod2]    20150112
        For modI = 6 To 2 Step - 1
            If InStr(defaultStr, startStr) = False And i Mod modI = 0 Then
                startStr = "[list-mod" & modI & "]" : endStr = "[/list-mod" & modI & "]" 
                If InStr(defaultStr, startStr) > 0 Then
                    Exit For 
                End If 
            End If 
        Next 

        '没有则用默认
        If InStr(defaultStr, startStr) = False Then
            startStr = "[list]" : endStr = "[/list]" 
        End If 


        If InStr(defaultStr, startStr) > 0 And InStr(defaultStr, endStr) > 0 Then
            s = strCut(defaultStr, startStr, endStr, 2) 

            s = replaceValueParam(s, "i", i)                                                '循环编号
            s = replaceValueParam(s, "编号", i)                                             '循环编号
            s = replaceValueParam(s, "id", rs("id"))                                        'id编号 因为获得字段他不获得id
            s = replaceValueParam(s, "url", url)                                            '网址
            s = replaceValueParam(s, "aurl", "href=""" & url & """")                        '网址
            s = replaceValueParam(s, "abcolor", abcolorStr)                                 'A链接加颜色与加粗
            s = replaceValueParam(s, "atitle", atitleStr)                                   'A链接title
            s = replaceValueParam(s, "anofollow", anofollowStr)                             'A链接nofollow
            s = replaceValueParam(s, "atarget", atargetStr)                                 'A链接打开方式



            For k = 0 To UBound(splFieldName)
                If splFieldName(k) <> "" Then
                    fieldName = splFieldName(k) 
                    replaceStr = rs(fieldName) & "" 
                    s = replaceValueParam(s, fieldName, replaceStr) 
                End If 
            Next 


            '开始位置加Dialog内容
            startStr = "[list-" & i & " startdialog]" : endStr = "[/list-" & i & " startdialog]" 
            If InStr(defaultStr, startStr) > 0 And InStr(defaultStr, endStr) > 0 Then
                s = strCut(defaultStr, startStr, endStr, 2) & s 
            End If 
            '结束位置加Dialog内容
            startStr = "[list-" & i & " enddialog]" : endStr = "[/list-" & i & " enddialog]" 
            If InStr(defaultStr, startStr) > 0 And InStr(defaultStr, endStr) > 0 Then
                s = s & strCut(defaultStr, startStr, endStr, 2) 
            End If 

            '加控制
            '【导航】
            If tableName = "WebColumn" Then
                url = WEB_ADMINURL & "?act=addEditHandle&actionType=WebColumn&lableTitle=网站栏目&nPageSize=10&page=&id=" & rs("id") & "&n=" & getRnd(11) 
            '【文章】
            ElseIf tableName = "ArticleDetail" Then
                url = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page=&parentid=&id=" & rs("id") & "&n=" & getRnd(11) 

                s = replaceValueParam(s, "columnurl", getColumnUrl(rs("parentid"), ""))         '文章对应栏目URL 20160304
                s = replaceValueParam(s, "columnname", getColumnName(rs("parentid")))           '文章对应栏目名称 20160304

            End If 
            s = handleDisplayOnlineEditDialog(url, s, "", "div|li|span")                    '处理是否添加在线修改管理器
            c = c & s 
        End If 
    rs.MoveNext : Next : rs.Close 

    '开始内容加Dialog内容
    startStr = "[dialog start]" : endStr = "[/dialog start]" 
    If InStr(defaultStr, startStr) > 0 And InStr(defaultStr, endStr) > 0 Then
        c = strCut(defaultStr, startStr, endStr, 2) & c 
    End If 
    '结束内容加Dialog内容
    startStr = "[dialog end]" : endStr = "[/dialog end]" 
    If InStr(defaultStr, startStr) > 0 And InStr(defaultStr, endStr) > 0 Then
        c = c & strCut(defaultStr, startStr, endStr, 2) 
    End If 
    XY_AP_GeneralList = c 
End Function 


'处理获得表内容
Function XY_handleGetTableBody(action, tableName, fieldParamName, defaultFileName, adminUrl)
    Dim url, content, id, sql, addSql, fieldName, fieldParamValue, fieldNameList, nLen, delHtmlYes, trimYes 
    fieldName = RParam(action, "fieldname")                                         '字段名称

    fieldNameList = getHandleFieldList(db_PREFIX & tableName, "字段列表") 
    '字段名称不为空，并且要在表字段里
    If fieldName = "" Or InStr(fieldNameList, "," & fieldName & ",") = False Then
        fieldName = defaultFileName 
    End If 
    fieldName = LCase(fieldName)                                                    '转为小写，因为在PHP里是全小写的

    fieldParamValue = RParam(action, fieldParamName)                                '截取字段内容
    id = handleNumber(RParam(action, "id"))                                         '获得ID
    addSql = " where " & fieldParamName & "='" & fieldParamValue & "'" 
    If id <> "" Then
        addSql = " where id=" & id 
    End If 

    content = getDefaultValue(action)                                               '获得默认内容
    sql = "select * from " & db_PREFIX & tableName & addSql 
    rs.Open sql, conn, 1, 1 
    If rs.EOF Then
        '自动添加 20160113
        If RParam(action, "autoadd") = "true" Then
            conn.Execute("insert into " & db_PREFIX & tableName & " (" & fieldParamName & "," & fieldName & ") values('" & fieldParamValue & "','" & ADSql(content) & "')") 
        End If 
    Else
        id = rs("id") 
        content = rs(fieldName) 
    End If : rs.Close 

    '删除Html
    delHtmlYes = RParam(action, "delHtml")                                          '是否删除Html
    If delHtmlYes = "true" Then content = Replace(delHtml(content), "<", "&lt;") 'HTML处理
    '删除两边空格
    trimYes = RParam(action, "trim")                                                '是否删除两边空格
    If trimYes = "true" Then content = trimVbCrlf(content) 

    '截取字符处理
    nLen = RParam(action, "len")                                                    '字符长度值
    nLen = handleNumber(nLen) 
    'If nLen<>"" Then ReplaceStr = CutStr(ReplaceStr,nLen,"null")' Left(ReplaceStr,nLen)
    If nLen <> "" Then content = cutStr(content, nLen, "...")                       'Left(ReplaceStr,nLen)


    If id = "" Then
        id = XY_AP_GetFieldValue("", sql, "id") 
    End If 
    url = adminUrl & "&id=" & id & "&n=" & getRnd(11) 
    If Request("gl") = "edit" Then
        content = "<span>" & content & "</span>" 
    End If 

    'call echo(sql,url)
    content = handleDisplayOnlineEditDialog(url, content, "", "span") 
    XY_handleGetTableBody = content 

End Function 

'获得单页内容
Function XY_AP_GetOnePageBody(action)
    Dim adminUrl 
    adminUrl = WEB_ADMINURL & "?act=addEditHandle&actionType=OnePage&lableTitle=单页管理&nPageSize=10&page=&switchId=2" 
    XY_AP_GetOnePageBody = XY_handleGetTableBody(action, "onepage", "title", "bodycontent", adminUrl) 
End Function 

'获得导航内容
Function XY_AP_GetColumnBody(action)
    Dim adminUrl 
    adminUrl = WEB_ADMINURL & "?act=addEditHandle&actionType=WebColumn&lableTitle=网站栏目&nPageSize=10&page=&switchId=2" 
    XY_AP_GetColumnBody = XY_handleGetTableBody(action, "webcolumn", "columnname", "bodycontent", adminUrl) 
End Function 

'显示文章内容
Function XY_AP_GetArticleBody(action)
    Dim adminUrl 
    adminUrl = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page=&switchId=2" 
    XY_AP_GetArticleBody = XY_handleGetTableBody(action, "articledetail", "title", "bodycontent", adminUrl) 
End Function 


'获得栏目URL
Function XY_GetColumnUrl(action)
    Dim columnName, url 
    columnName = RParam(action, "columnName") 
    url = getColumnUrl(columnName, "name") 
    'call echo(columnName, url)
    If Request("gl") <> "" Then
        url = url & "&gl=" & Request("gl") 
    End If 
    XY_GetColumnUrl = url 

End Function 

'获得文章URL
Function XY_GetArticleUrl(action)
    Dim title, url 
    title = RParam(action, "title") 
    url = getArticleUrl(title) 
    If Request("gl") <> "" Then
        url = url & "&gl=" & Request("gl") 
    End If 
    XY_GetArticleUrl = url 
End Function 

'获得单页URL
Function XY_GetOnePageUrl(action)
    Dim title, url 
    title = RParam(action, "title") 
    url = getOnePageUrl(title) 
    If Request("gl") <> "" Then
        url = url & "&gl=" & Request("gl") 
    End If 
    XY_GetOnePageUrl = url 
End Function 


'获得单个字段内容
Function XY_AP_GetFieldValue(action, sql, fieldName)
    Dim title, content 
    rs.Open sql, conn, 1, 1 
    If Not rs.EOF Then
        content = rs(fieldName) 
    End If : rs.Close 
    XY_AP_GetFieldValue = content 
End Function 


'Js版网站统计
Function XY_JsWebStat(action)
    Dim s, fileName, sType 
    sType = RParam(action, "stype") 
    fileName = Trim(RParam(action, "fileName")) 
    If fileName = "" Then
        fileName = "[$WEB_VIEWURL$]?act=webstat&stype=" & sType 
    End If 
    fileName = Replace(fileName, "/", "\/") 
    s = "<script>document.writeln(""<script src=\'" & fileName & "&GoToUrl=""" 
    s = s & "+escape(document.referrer)+""&ThisUrl=""+escape(window.location.href)+""&screen=""+escape(window.screen.width+""x""+window.screen.height)" 
    s = s & "+""&co=""+escape(document.cookie)"                                 '收集cookie 不需要则屏蔽掉
    s = s & "+"" \'><\/script>"");</script>" 
    XY_JsWebStat = s 
End Function 



'普通链接A
Function XY_HrefA(action)
    Dim content, Href, c, AContent, AType, url, title 
    action = handleInModule(action, "start") 
    content = RParam(action, "Content") 
    AType = RParam(action, "Type") 
    If AType = "收藏" Then
        '第一种方法
        'Url = "window.external.addFavorite('"& WebUrl &"','"& WebTitle &"')"
        url = "shoucang(document.title,window.location)" 
        c = "<a href='javascript:;' onClick=""" & url & """ " & setHtmlParam(action, "target|title|alt|id|class|style") & ">" & content & "</a>" 
    ElseIf AType = "设为首页" Then
        '第一种方法
        'Url = "var strHref=window.location.href;this.style.behavior='url(#default#homepage)';this.setHomePage('"& WebUrl &"');"
        url = "SetHome(this,window.location)" 
        c = "<a href='javascript:;' onClick=""" & url & """" & setHtmlParam(action, "target|title|alt|id|class|style") & ">" & content & "</a>" 
    Else
        content = RParam(action, "Title") 
    End If 

    content = handleInModule(content, "end") 
    If c = "" Then c = "<a" & setHtmlParam(action, "href|target|title|alt|id|class|rel|style") & ">" & content & "</a>" 

    XY_HrefA = c 
End Function 



'布局20151231
Function XY_Layout(action)
    Dim layoutName, s, c, sourceStr, replaceStr, splSource, splReplace, i 

    layoutName = RParam(action, "layoutname") 
    rs.Open "select * from " & db_PREFIX & "weblayout where layoutname='" & layoutName & "'", conn, 1, 1 
    If Not rs.EOF Then
        c = rs("bodycontent") 

        sourceStr = rs("sourceStr")                                                     '源内容 被替换内容
        replaceStr = rs("replaceStr")                                                   '替换内容
        splSource = Split(sourceStr, "[Array]")                                    '源内容数组
        splReplace = Split(replaceStr, "[Array]")                                  '替换内容数组

        For i = 0 To UBound(splSource)
            sourceStr = splSource(i) 
            replaceStr = splReplace(i) 
            If sourceStr <> "" Then
                c = Replace(c, sourceStr, replaceStr) 
                'call echo(sourceStr,replaceStr)
            'call echo(c,instr(c,sourcestr))
            End If 
        Next 
    'call rwend(c)
    End If : rs.Close 
    XY_Layout = c 
End Function 

'模块20151231
Function XY_Module(action)
    Dim moduleName, s, c, sourceStr, replaceStr, splSource, splReplace, i 
    moduleName = RParam(action, "modulename") 
    rs.Open "select * from " & db_PREFIX & "webmodule where modulename='" & moduleName & "'", conn, 1, 1 
    If Not rs.EOF Then
        c = rs("bodycontent") 

        sourceStr = RParam(action, "sourceStr")                                         '源内容 被替换内容
        replaceStr = RParam(action, "replaceStr")                                       '替换内容

        splSource = Split(sourceStr, "[Array]")                                    '源内容数组
        splReplace = Split(replaceStr, "[Array]")                                  '替换内容数组

        For i = 0 To UBound(splSource)
            sourceStr = splSource(i) 
            replaceStr = splReplace(i) 
            If sourceStr <> "" Then
                c = Replace(c, sourceStr, replaceStr) 
                'call echo(sourceStr,replaceStr)
            'call echo(c,instr(c,sourcestr))
            End If 
        Next 
    'call rwend(c)
    End If : rs.Close 
    XY_Module = c 
End Function 

'显示包裹块20160127
Function XY_DisplayWrap(ByVal action)
    Dim content 
    content = getDefaultValue(action) 
    XY_DisplayWrap = content 
End Function 




'嵌套标题 测试
Function XY_getLableValue(action)
    Dim title, content, c 
    'call echo("Action",Action)
    title = RParam(action, "title") 
    content = RParam(action, "content") 
    c = c & "title=" & getContentRunStr(title) & "<hr>" 
    c = c & "content=" & getContentRunStr(content) & "<hr>" 
    XY_getLableValue = c 
    Call echo("title", title) 
    XY_getLableValue = "【title=】【" & title & "】" 
End Function 
'标题在搜索引擎中搜索列表
Function XY_TitleInSearchEngineList(action)
    Dim title, sType 
    title = RParam(action, "title") 
    sType = RParam(action, "sType") 
    Dim c 
    c = c & "<strong>更多关于《" & title & "》在各大搜索引擎中相关信息</strong><br>" & vbCrLf 
    c = c & "<ul class=""list""> " & vbCrLf 
    c = c & "<li><a href=""https://www.baidu.com/s?ie=gb2312&word=" & title & """ rel=""nofollow"" target=""_blank"">【baidu搜索】在百度里搜索(" & title & ")</a></li>" & vbCrLf 
    c = c & "<li><a href=""http://www.haosou.com/s?ie=gb2312&q=" & title & """ rel=""nofollow"" target=""_blank"">【haosou搜索】在好搜里搜索(" & title & ")</a></li>" & vbCrLf 
    c = c & "<li><a href=""https://search.yahoo.com/search;_ylt=A86.JmbkJatWH5YARmebvZx4?toggle=1&cop=mss&ei=gb2312&fr=yfp-t-901&fp=1&p=" & title & """ rel=""nofollow"" target=""_blank"">【yahoo搜索】在雅虎里搜索(" & title & ")</a></li>" & vbCrLf 

    c = c & "<li><a href=""https://www.sogou.com/sogou?ie=utf8&query=" & GBtoUTF8(title) & """ rel=""nofollow"" target=""_blank"">【sogou搜索】在搜狗里搜索(" & title & ")</a></li>" & vbCrLf 
    c = c & "<li><a href=""http://www.youdao.com/search?ue=utf8&q=" & GBtoUTF8(title) & """ rel=""nofollow"" target=""_blank"">【youdao搜索】在有道里搜索(" & title & ")</a></li>" & vbCrLf 
    c = c & "<li><a href=""http://search.yam.com/Search/Web/DefaultKSA.aspx?SearchType=web&l=0&p=0&k=" & GBtoUTF8(title) & """ rel=""nofollow"" target=""_blank"">【yam搜索(google提供技术)】在蕃薯藤里搜索(" & title & ")</a></li>" & vbCrLf 
	
	
    c = c & "<li><a href=""http://cn.bing.com/search?q=" & GBtoUTF8(title) & """ rel=""nofollow"" target=""_blank"">【bing搜索】在必应里搜索(" & title & ")</a></li>" & vbCrLf 
    c = c & "</ul>" & vbCrLf 

    XY_TitleInSearchEngineList = c 
End Function 

'URL加密
Function XY_escape(action)
    Dim content 
    content = RParam(action, "content") 
    XY_escape = escape(content) 
End Function 
'URL解密
Function XY_unescape(action)
    Dim content 
    content = RParam(action, "content") 
    XY_unescape = escape(content) 
End Function 

%>           


