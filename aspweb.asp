<!--#Include File = "Inc/_Config.Asp"-->          
<!--#Include File = "web/function.asp"-->    
<% 
'asp服务器

Call openconn() 
'=========

Dim code                                                                        'html代码
Dim templateName                                                                '模板名称
Dim cfg_webSiteUrl, cfg_webTemplate, cfg_webImages, cfg_webCss, cfg_webJs, cfg_webTitle, cfg_webKeywords, cfg_webDescription, cfg_webSiteBottom, cfg_flags 
Dim glb_columnName, glb_columnId, glb_id, glb_columnType, glb_columnENType, glb_table, glb_detailTitle, glb_flags 
Dim webTemplate                                                                 '网站模板路径
Dim glb_url, glb_filePath                                                       '当前链接网址,和文件路径
Dim glb_isonhtml                                                                '是否生成静态网页
Dim glb_locationType                                                            '位置类型

Dim glb_bodyContent                                                             '主体内容
Dim glb_artitleAuthor                                                           '文章作者
Dim glb_artitleAdddatetime                                                      '文章添加时间
Dim glb_upArticle                                                               '上一篇文章
Dim glb_downArticle                                                             '下一篇文章
Dim glb_aritcleRelatedTags                                                      '文章标签组
Dim glb_aritcleSmallImage, glb_aritcleBigImage                                  '文章小图与文章大图
Dim glb_searchKeyWord                                                           '搜索关键词
dim cacheHtmlFilePath															'缓冲html文件路径

Dim isMakeHtml                                                                  '是否生成网页
'处理动作   ReplaceValueParam为控制字符显示方式
Function handleAction(content)
    Dim startStr, endStr, ActionList, splStr, action, s, HandYes 
    startStr = "{\$" : endStr = "\$}" 
    ActionList = getArray(content, startStr, endStr, True, True) 
    'Call echo("ActionList ", ActionList)
    splStr = Split(ActionList, "$Array$") 
    For Each s In splStr
        action = Trim(s) 
        action = handleInModule(action, "start")                                        '处理\'替换掉
        If action <> "" Then
            action = Trim(Mid(action, 3, Len(action) - 4)) & " " 
            'call echo("s",s)
            HandYes = True                                                                  '处理为真
            '{VB #} 这种是放在图片路径里，目的是为了在VB里不处理这个路径
            If checkFunValue(action, "# ") = True Then
                action = "" 
            '测试
            ElseIf checkFunValue(action, "GetLableValue ") = True Then
                action = XY_getLableValue(action) 
            '标题在搜索引擎里列表
            ElseIf checkFunValue(action, "TitleInSearchEngineList ") = True Then
                action = XY_TitleInSearchEngineList(action) 

            '加载文件
            ElseIf checkFunValue(action, "Include ") = True Then
                action = XY_Include(action) 
            '栏目列表
            ElseIf checkFunValue(action, "ColumnList ") = True Then
                action = XY_AP_ColumnList(action) 
            '文章列表
            ElseIf checkFunValue(action, "ArticleList ") = True Or checkFunValue(action, "CustomInfoList ") = True Then
                action = XY_AP_ArticleList(action) 
            '评论列表
            ElseIf checkFunValue(action, "CommentList ") = True Then
                action = XY_AP_CommentList(action) 
            '搜索统计列表
            ElseIf checkFunValue(action, "SearchStatList ") = True Then
                action = XY_AP_SearchStatList(action) 
            '友情链接列表
            ElseIf checkFunValue(action, "Links ") = True Then
                action = XY_AP_Links(action)

            '显示单页内容
            ElseIf checkFunValue(action, "GetOnePageBody ") = True Or checkFunValue(action, "MainInfo ") = True Then
                action = XY_AP_GetOnePageBody(action) 
            '显示文章内容
            ElseIf checkFunValue(action, "GetArticleBody ") = True Then
                action = XY_AP_GetArticleBody(action) 
            '显示栏目内容
            ElseIf checkFunValue(action, "GetColumnBody ") = True Then
                action = XY_AP_GetColumnBody(action) 

            '获得栏目URL
            ElseIf checkFunValue(action, "GetColumnUrl ") = True Then
                action = XY_GetColumnUrl(action) 
            '获得文章URL
            ElseIf checkFunValue(action, "GetArticleUrl ") = True Then
                action = XY_GetArticleUrl(action) 
            '获得单页URL
            ElseIf checkFunValue(action, "GetOnePageUrl ") = True Then
                action = XY_GetOnePageUrl(action) 

            '显示包裹块 作用不大
            ElseIf checkFunValue(action, "DisplayWrap ") = True Then
                action = XY_DisplayWrap(action) 
            '显示布局
            ElseIf checkFunValue(action, "Layout ") = True Then
                action = XY_Layout(action) 
            '显示模块
            ElseIf checkFunValue(action, "Module ") = True Then
                action = XY_Module(action) 
            '获得内容模块 20150108
            ElseIf checkFunValue(action, "GetContentModule ") = True Then
                action = XY_ReadTemplateModule(action) 
            '读模板样式并设置标题与内容   软件里有个栏目Style进行设置
            ElseIf checkFunValue(action, "ReadColumeSetTitle ") = True Then
                action = XY_ReadColumeSetTitle(action) 

            '显示JS渲染ASP/PHP/VB等程序的编辑器
            ElseIf checkFunValue(action, "displayEditor ") = True Then
                action = displayEditor(action) 

            'Js版网站统计
            ElseIf checkFunValue(action, "JsWebStat ") = True Then
                action = XY_JsWebStat(action) 

                '------------------- 链接区 -----------------------
            '普通链接A
            ElseIf checkFunValue(action, "HrefA ") = True Then
                action = XY_HrefA(action) 

            '栏目菜单(引用后台栏目程序)
            ElseIf checkFunValue(action, "ColumnMenu ") = True Then
                action = XY_AP_ColumnMenu(action) 

            '网站底部
            ElseIf checkFunValue(action, "WebSiteBottom ") = True Then
                action = XY_AP_WebSiteBottom(action) 
            '显示网站栏目 20160331
            ElseIf checkFunValue(action, "DisplayWebColumn ") = True Then
                action = XY_DisplayWebColumn(action) 
            'URL加密
            ElseIf checkFunValue(action, "escape ") = True Then
                action = XY_escape(action) 
            'URL解密
            ElseIf checkFunValue(action, "unescape ") = True Then
                action = XY_unescape(action) 
            'asp与php版本
            ElseIf checkFunValue(action, "EDITORTYPE ") = True Then
                action = XY_EDITORTYPE(action)


            '暂时不屏蔽
            ElseIf checkFunValue(action, "copyTemplateMaterial ") = True Then
                action = "" 
            ElseIf checkFunValue(action, "clearCache ") = True Then
                action = "" 
            Else
                HandYes = False                                                                 '处理为假
            End If 
            '注意这样，有的则不显示 晕 And IsNul(action)=False
            If isNul(action) = True Then action = "" 
            If HandYes = True Then
                content = Replace(content, s, action) 
            End If 
        End If 
    Next 
    handleAction = content 
End Function 

'显示网站栏目 新版 把之前网站 导航程序改进过来的
Function XY_DisplayWebColumn(action)
    Dim i, c, s, url, sql, dropDownMenu, focusType, addSql 
    Dim isConcise                                                                   '简洁显示20150212
    Dim styleId, styleValue                                                         '样式ID与样式内容
    Dim cssNameAddId 
    Dim shopnavidwrap                                                               '是否显示栏目ID包

    styleId = PHPTrim(RParam(action, "styleID"))
    styleValue = PHPTrim(RParam(action, "styleValue"))
    addSql = PHPTrim(RParam(action, "addSql"))
    shopnavidwrap = PHPTrim(RParam(action, "shopnavidwrap"))
    'If styleId <> "" Then
    'Call ReadNavCSS(styleId, styleValue)
    'End If

    '为数字类型 则自动提取样式内容  20150615
    If checkStrIsNumberType(styleValue) Then
        cssNameAddId = "_" & styleValue                                                 'Css名称追加Id编号
    End If 
    sql = "select * from " & db_PREFIX & "webcolumn" 
    '追加sql
    If addSql <> "" Then
        sql = getWhereAnd(sql, addSql) 
    End If 
    If checkSql(sql) = False Then Call eerr("Sql", sql) 
    rs.Open sql, conn, 1, 1 
    dropDownMenu = LCase(RParam(action, "DropDownMenu")) 
    focusType = LCase(RParam(action, "FocusType")) 
    isConcise = LCase(RParam(action, "isConcise")) 
    If isConcise = "true" Then
        isConcise = False 
    Else
        isConcise = True 
    End If 
    If isConcise = True Then c = c & copyStr(" ", 4) & "<li class=left></li>" & vbCrLf 
    For i = 1 To rs.RecordCount

        '【PHP】$rs=mysql_fetch_array($rsObj);                                            //给PHP用，因为在 asptophp转换不完善
        url = getColumnUrl(rs("columnname"), "name") 
        If rs("columnName") = glb_columnName Then
            If focusType = "a" Then
                s = copyStr(" ", 8) & "<li class=focus><a href=""" & url & """>" & rs("columname") & "</a>" 
            Else
                s = copyStr(" ", 8) & "<li class=focus>" & rs("columnname") 
            End If 
        Else
            s = copyStr(" ", 8) & "<li><a href=""" & url & """>" & rs("columnname") & "</a>" 
        End If 
        url = WEB_ADMINURL & "?act=addEditHandle&actionType=WebColumn&lableTitle=网站栏目&nPageSize=10&page=&id=" & rs("id") & "&n=" & getRnd(11) 
        s = handleDisplayOnlineEditDialog(url, s, "", "div|li|span")                    '处理是否添加在线修改管理器

        c = c & s 

        '小类

        c = c & copyStr(" ", 8) & "</li>" & vbCrLf 

        If isConcise = True Then c = c & copyStr(" ", 8) & "<li class=line></li>" & vbCrLf 
    rs.MoveNext : Next : rs.Close 
    If isConcise = True Then c = c & copyStr(" ", 8) & "<li class=right></li>" & vbCrLf 

    If styleId <> "" Then
        c = "<ul class='nav" & styleId & cssNameAddId & "'>" & vbCrLf & c & vbCrLf & "</ul>" & vbCrLf 
    End If 
    If shopnavidwrap = "1" Or shopnavidwrap = "true" Then
        c = "<div id='nav" & styleId & cssNameAddId & "'>" & vbCrLf & c & vbCrLf & "</div>" & vbCrLf 
    End If 

    XY_DisplayWebColumn = c 
End Function 

'替换全局变量 {$cfg_websiteurl$}
Function replaceGlobleVariable(ByVal content)
    content = handleRGV(content, "{$cfg_webSiteUrl$}", cfg_webSiteUrl)              '网址
    content = handleRGV(content, "{$cfg_webTemplate$}", cfg_webTemplate)            '模板
    content = handleRGV(content, "{$cfg_webImages$}", cfg_webImages)                '图片路径
    content = handleRGV(content, "{$cfg_webCss$}", cfg_webCss)                      'css路径
    content = handleRGV(content, "{$cfg_webJs$}", cfg_webJs)                        'js路径
    content = handleRGV(content, "{$cfg_webTitle$}", cfg_webTitle)                  '网站标题
    content = handleRGV(content, "{$cfg_webKeywords$}", cfg_webKeywords)            '网站关键词
    content = handleRGV(content, "{$cfg_webDescription$}", cfg_webDescription)      '网站描述
    content = handleRGV(content, "{$cfg_webSiteBottom$}", cfg_webSiteBottom)        '网站描述

    content = handleRGV(content, "{$glb_columnId$}", glb_columnId)                  '栏目Id
    content = handleRGV(content, "{$glb_columnName$}", glb_columnName)              '栏目名称
    content = handleRGV(content, "{$glb_columnType$}", glb_columnType)              '栏目类型
    content = handleRGV(content, "{$glb_columnENType$}", glb_columnENType)          '栏目英文类型

    content = handleRGV(content, "{$glb_Table$}", glb_table)                        '表
    content = handleRGV(content, "{$glb_Id$}", glb_id)                              'id


    '兼容旧版本 渐渐把它去掉
    content = handleRGV(content, "{$WebImages$}", cfg_webImages)                    '图片路径
    content = handleRGV(content, "{$WebCss$}", cfg_webCss)                          'css路径
    content = handleRGV(content, "{$WebJs$}", cfg_webJs)                            'js路径
    content = handleRGV(content, "{$Web_Title$}", cfg_webTitle) 
    content = handleRGV(content, "{$Web_KeyWords$}", cfg_webKeywords) 
    content = handleRGV(content, "{$Web_Description$}", cfg_webDescription) 


    content = handleRGV(content, "{$EDITORTYPE$}", EDITORTYPE)                      '后缀
    content = handleRGV(content, "{$WEB_VIEWURL$}", WEB_VIEWURL)                    '首页显示网址
    '文章用到
    content = handleRGV(content, "{$glb_artitleAuthor$}", glb_artitleAuthor)        '文章作者
    content = handleRGV(content, "{$glb_artitleAdddatetime$}", glb_artitleAdddatetime) '文章添加时间
    content = handleRGV(content, "{$glb_upArticle$}", glb_upArticle)                '上一篇文章
    content = handleRGV(content, "{$glb_downArticle$}", glb_downArticle)            '下一篇文章
    content = handleRGV(content, "{$glb_aritcleRelatedTags$}", glb_aritcleRelatedTags) '文章标签组
    content = handleRGV(content, "{$glb_aritcleBigImage$}", glb_aritcleBigImage)    '文章大图
    content = handleRGV(content, "{$glb_aritcleSmallImage$}", glb_aritcleSmallImage) '文章小图
    content = handleRGV(content, "{$glb_searchKeyWord$}", glb_searchKeyWord)        '首页显示网址


    replaceGlobleVariable = content 
End Function 

'处理替换
Function handleRGV(ByVal content, findStr, replaceStr)
    Dim lableName 
    '对[$$]处理
    lableName = Mid(findStr, 3, Len(findStr) - 4) & " " 
    lableName = Mid(lableName, 1, InStr(lableName, " ") - 1) 
    content = replaceValueParam(content, lableName, replaceStr) 
    content = replaceValueParam(content, LCase(lableName), replaceStr) 
    '直接替换{$$}这种方式，兼容之前网站
    content = Replace(content, findStr, replaceStr) 
    content = Replace(content, LCase(findStr), replaceStr) 
    handleRGV = content 
End Function 

'加载网址配置
Sub loadWebConfig()
    Dim templatedir 
    Call openconn() 
    rs.Open "select * from " & db_PREFIX & "website", conn, 1, 1 
    If Not rs.EOF Then
        cfg_webSiteUrl = phptrim(rs("webSiteUrl"))                                      '网址
        cfg_webTemplate = webDir & phptrim(rs("webTemplate"))                           '模板路径
        cfg_webImages = webDir & phptrim(rs("webImages"))                               '图片路径
        cfg_webCss = webDir & phptrim(rs("webCss"))                                     'css路径
        cfg_webJs = webDir & phptrim(rs("webJs"))                                       'js路径
        cfg_webTitle = rs("webTitle")                                                   '网址标题
        cfg_webKeywords = rs("webKeywords")                                             '网站关键词
        cfg_webDescription = rs("webDescription")                                       '网站描述
        cfg_webSiteBottom = rs("webSiteBottom")                                         '网站地底
        cfg_flags = rs("flags")                                                         '旗

        '改换模板20160202
        If Request("templatedir") <> "" Then
			'删除绝对目录前面的目录，不需要那个东西20160414
            templatedir = replace(handlePath(Request("templatedir")),handlePath("/"),"/")
			'call eerr("templatedir",templatedir)
			
            If(InStr(templatedir, ":") > 0 Or InStr(templatedir, "..") > 0) And getIP() <> "127.0.0.1" Then
                Call eerr("提示", "模板目录有非法字符") 
            End If 
            templatedir = handlehttpurl(Replace(templatedir, handlePath("/"), "/")) 

            cfg_webImages = Replace(cfg_webImages, cfg_webTemplate, templatedir) 
            cfg_webCss = Replace(cfg_webCss, cfg_webTemplate, templatedir) 
            cfg_webJs = Replace(cfg_webJs, cfg_webTemplate, templatedir) 
            cfg_webTemplate = templatedir 
        End If 
        webTemplate = cfg_webTemplate 
    End If : rs.Close 
End Sub 

'网站位置 待完善
Function thisPosition(content)
    Dim c 
    c = "<a href=""/"">首页</a>" 
    If glb_columnName <> "" Then
        c = c & " >> <a href=""" & getColumnUrl(glb_columnName, "name") & """>" & glb_columnName & "</a>" 
    End If 
    '20160330
    If glb_locationType = "detail" Then
        c = c & " >> 查看内容" 
    End If 
    'call echo("glb_locationType",glb_locationType)

    content = Replace(content, "[$detailPosition$]", c) 
    content = Replace(content, "[$detailTitle$]", glb_detailTitle) 
    content = Replace(content, "[$detailContent$]", glb_bodyContent) 

    thisPosition = content 
End Function 

'显示管理列表
Function getDetailList(action, content, actionName, lableTitle, ByVal fieldNameList, nPageSize, nPage, addSql)
    Call openconn() 
    Dim defaultList, i, s, c, tableName, j, splxx, sql 
    Dim x, url, nCount 
    Dim pageInfo 

    Dim fieldName                                                                   '字段名称
    Dim splFieldName                                                                '分割字段

    Dim replaceStr                                                                  '替换字符
    tableName = LCase(actionName)                                                   '表名称
    Dim listFileName                                                                '列表文件名称
    listFileName = RParam(action, "listFileName") 
    Dim abcolorStr                                                                  'A加粗和颜色
    Dim atargetStr                                                                  'A链接打开方式
    Dim atitleStr                                                                   'A链接的title20160407
    Dim anofollowStr                                                                'A链接的nofollow

    Dim id 
    id = rq("id") 
    Call checkIDSQL(Request("id")) 

    If fieldNameList = "*" Then
        fieldNameList = getHandleFieldList(db_PREFIX & tableName, "字段列表") 
    End If 

    fieldNameList = specialStrReplace(fieldNameList)                                '特殊字符处理
    splFieldName = Split(fieldNameList, ",")                                        '字段分割成数组


    defaultList = getStrCut(content, "[list]", "[/list]", 2) 
    pageInfo = getStrCut(content, "[page]", "[/page]", 1) 
    If pageInfo <> "" Then
        content = Replace(content, pageInfo, "") 
    End If 

    sql = "select * from " & db_PREFIX & tableName & " " & addSql 
    '检测SQL
    If checksql(sql) = False Then
        Call errorLog("出错提示：<br>sql=" & sql & "<br>") 
        Exit Function 
    End If 
    rs.Open sql, conn, 1, 1 
    nCount = rs.RecordCount 

    '为动态翻页网址
    If isMakeHtml = True Then
        url = "" 
        If Len(listFileName) > 5 Then
            url = Mid(listFileName, 1, Len(listFileName) - 5) & "[id].html" 
            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
        End If 
    Else
        url = getUrlAddToParam(getUrl(), "?page=[id]", "replace") 
    End If 
    content = Replace(content, "[$pageInfo$]", webPageControl(nCount, nPageSize, nPage, url, pageInfo)) 

    If EDITORTYPE = "asp" Then
        x = getRsPageNumber(rs, nCount, nPageSize, nPage)                               '获得Rs页数                                                  '记录总数
    Else
        If nPage <> "" Then
            nPage = nPage - 1 
        End If 
        sql = "select * from " & db_PREFIX & "" & tableName & " " & addSql & " limit " & nPageSize * nPage & "," & nPageSize 
        rs.Open sql, conn, 1, 1 
        x = rs.RecordCount 
    End If 
    For i = 1 To x
        '【PHP】$rs=mysql_fetch_array($rsObj);                                            //给PHP用，因为在 asptophp转换不完善

        s = defaultList 
        s = Replace(s, "[$id$]", rs("id")) 
        For j = 0 To UBound(splFieldName)
            If splFieldName(j) <> "" Then
                splxx = Split(splFieldName(j) & "|||", "|") 
                fieldName = splxx(0) 
                replaceStr = rs(fieldName) & ""
                s = replaceValueParam(s, fieldName, replaceStr)
            End If 

            If isMakeHtml = True Then
                url = getHandleRsUrl(rs("fileName"), rs("customAUrl"), "/detail/detail" & rs("id")) 
            Else
                url = handleWebUrl("?act=detail&id=" & rs("id")) 
                If rs("customAUrl") <> "" Then
                    url = rs("customAUrl") 					 
                End If 
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



            s = replaceValueParam(s, "url", url) 
            s = replaceValueParam(s, "abcolor", abcolorStr)                                 'A链接加颜色与加粗
            s = replaceValueParam(s, "atitle", atitleStr)                                   'A链接title
            s = replaceValueParam(s, "anofollow", anofollowStr)                             'A链接nofollow
            s = replaceValueParam(s, "atarget", atargetStr)                                 'A链接打开方式


        Next 
        '文章列表加在线编辑
        url = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page=&parentid=&id=" & rs("id") & "&n=" & getRnd(11) 
        s = handleDisplayOnlineEditDialog(url, s, "", "div|li|span") 

        c = c & s 
    rs.MoveNext : Next : rs.Close 
    content = Replace(content, "[list]" & defaultList & "[/list]", c) 

    If isMakeHtml = True Then
        url = "" 
        If Len(listFileName) > 5 Then
            url = Mid(listFileName, 1, Len(listFileName) - 5) & "[id].html" 
            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
        End If 
    Else
        url = getUrlAddToParam(getUrl(), "?page=[id]", "replace") 
    End If 

    getDetailList = content 
End Function 


'****************************************************
'默认列表模板
Function defaultListTemplate()
    Dim c, templateHtml, listTemplate, lableName, startStr, endStr 

    templateHtml = getFText(cfg_webTemplate & "/" & templateName) 

    lableName = "list" 
    startStr = "<!--#" & lableName & " start#-->" 
    endStr = "<!--#" & lableName & " end#-->" 
    'call rwend(templateHtml)
    If InStr(templateHtml, startStr) > 0 And InStr(templateHtml, endStr) > 0 Then
        listTemplate = strCut(templateHtml, startStr, endStr, 2) 
    Else
        startStr = "<!--#" & lableName 
        endStr = "#-->" 
        If InStr(templateHtml, startStr) > 0 And InStr(templateHtml, endStr) > 0 Then
            listTemplate = strCut(templateHtml, startStr, endStr, 2) 
        End If 
    End If 
    If listTemplate = "" Then
        c = "<ul class=""list"">" & vbCrLf 
        c = c & "[list]    <li><a href=""[$url$]""[$atitle$][$atarget$][$abcolor$][$anofollow$]>[$title$]</a><span class=""time"">[$adddatetime format_time='7'$]</span></li>" & vbCrLf 
        c = c & "[/list]" & vbCrLf 
        c = c & "</ul>" & vbCrLf 
        c = c & "<div class=""clear10""></div>" & vbCrLf 
        c = c & "<div>[$pageInfo$]</div>" & vbCrLf 
        listTemplate = c 
    End If 

    defaultListTemplate = listTemplate 
End Function 

'缓冲处理20160622
cacheHtmlFilePath="/cache/html/" & setFileName(getThisUrlFileParam()) & ".html"
'启用缓冲
if request("cache")<>"false" and onCacheHtml=true then
	if checkFile(cacheHtmlFilePath)=true then
		'call echo("读取缓冲文件","OK")
		call rwend(getftext(cacheHtmlFilePath))
	end if
end if

'记录表前缀
If Request("db_PREFIX") <> "" Then
    db_PREFIX = Request("db_PREFIX") 
ElseIf Session("db_PREFIX") <> "" Then
    db_PREFIX = Session("db_PREFIX") 
End If 
'加载网址配置
Call loadWebConfig() 
isMakeHtml = False                                                              '默认生成HTML为关闭
If Request("isMakeHtml") = "1" Or Request("isMakeHtml") = "true" Then
    isMakeHtml = True 
End If 
templateName = Request("templateName")                                          '模板名称

'保存数据处理页
Select Case Request("act")
    Case "savedata" : saveData(Request("stype")) : Response.End()                   '保存数据
    ''站长统计 | 今日IP[653] | 今日PV[9865] | 当前在线[65]')
    Case "webstat" : webStat(adminDir & "/Data/Stat/"):Response.End()    '网站统计
	
    Case "saveSiteMap" : isMakeHtml=true:saveSiteMap() :response.End()                                             '保存sitemap.xml
	
	case "handleAction" 
	if request("ishtml")="1" then
		isMakeHtml = True
	end if
	rwend(handleAction(request("content")))		'处理动作
End Select


'生成html
If Request("act") = "makehtml" Then
    Call echo("makehtml", "makehtml") 
    isMakeHtml = True 
    Call makeWebHtml(" action actionType='" & Request("act") & "' columnName='" & Request("columnName") & "' id='" & Request("id") & "' ") 
    Call createFileGBK("index.html", code) 

'复制Html到网站
ElseIf Request("act") = "copyHtmlToWeb" Then
    Call copyHtmlToWeb() 
'全部生成
ElseIf Request("act") = "makeallhtml" Then
    Call makeAllHtml("", "", Request("id")) 

'生成当前页面
ElseIf Request("isMakeHtml") <> "" And Request("isSave") <> "" Then

    Call handlePower("生成当前HTML页面")                                            '管理权限处理
    Call writeSystemLog("", "生成当前HTML页面")                                     '系统日志

    isMakeHtml = True 


    Call checkIDSQL(Request("id")) 
    Call rw(makeWebHtml(" action actionType='" & Request("act") & "' columnName='" & Request("columnName") & "' columnType='" & Request("columnType") & "' id='" & Request("id") & "' npage='" & Request("page") & "' ")) 
    glb_filePath = Replace(glb_url, cfg_webSiteUrl, "") 
    If Right(glb_filePath, 1) = "/" Then
        glb_filePath = glb_filePath & "index.html" 
    ElseIf glb_filePath = "" And glb_columnType = "首页" Then
        glb_filePath = "index.html" 
    End If 
    '文件不为空  并且开启生成html
    If glb_filePath <> "" And glb_isonhtml = True Then
        Call createDirFolder(getFileAttr(glb_filePath, "1")) 
        Call createFileGBK(glb_filePath, code) 
        If Request("act") = "detail" Then
            conn.Execute("update " & db_PREFIX & "ArticleDetail set ishtml=true where id=" & Request("id")) 
        ElseIf Request("act") = "nav" Then
            If Request("id") <> "" Then
                conn.Execute("update " & db_PREFIX & "WebColumn set ishtml=true where id=" & Request("id")) 
            Else
                conn.Execute("update " & db_PREFIX & "WebColumn set ishtml=true where columnname='" & Request("columnName") & "'") 
            End If 
        End If 
        Call echo("生成文件路径", "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>") 

        '新闻则批量生成 20160216
        If glb_columnType = "新闻" Then
            Call makeAllHtml("", "", glb_columnId) 
        End If 

    End If 

'全部生成
ElseIf Request("act") = "Search" Then
    Call rw(makeWebHtml("actionType='Search' npage='1' ")) 
Else
    If LCase(Request("issave")) = "1" Then
        Call makeAllHtml(Request("columnType"), Request("columnName"), Request("columnId")) 
    Else
        Call checkIDSQL(Request("id")) 
        Call rw(makeWebHtml(" action actionType='" & Request("act") & "' columnName='" & Request("columnName") & "' columnType='" & Request("columnType") & "' id='" & Request("id") & "' npage='" & Request("page") & "' ")) 
    End If 
End If
'开启缓冲html 
if onCacheHtml=true then
	call createFile(cacheHtmlFilePath,code)		'保存到缓冲文件里20160622
end if
'检测ID是否SQL安全
Function checkIDSQL(id)
    If checkNumber(id) = False And id <> "" Then
        Call eerr("提示", "id中有非法字符") 
    End If 
End Function 




'http://127.0.0.1/aspweb.asp?act=nav&columnName=ASP
'http://127.0.0.1/aspweb.asp?act=detail&id=75
'生成html静态页
Function makeWebHtml(action)
    Dim actionType, npagesize, npage, url, addSql, sortSql 
    actionType = RParam(action, "actionType") 
    npage = RParam(action, "npage") 
    npage = getnumber(npage) 
    If npage = "" Then
        npage = 1 
    Else
        npage = CInt(npage) 
    End If 
    '导航
    If actionType = "nav" Then
        glb_columnType = RParam(action, "columnType") 
        glb_columnName = RParam(action, "columnName") 
        glb_columnId = RParam(action, "columnId") 
        If glb_columnId = "" Then
            glb_columnId = RParam(action, "id") 
        End If 
        If glb_columnType <> "" Then
            addSql = "where columnType='" & glb_columnType & "'" 
        End If 
        If glb_columnName <> "" Then
            addSql = getWhereAnd(addSql, "where columnName='" & glb_columnName & "'") 
        End If 
        If glb_columnId <> "" Then
            addSql = getWhereAnd(addSql, "where id=" & glb_columnId & "") 
        End If 
        'call echo("addsql",addsql)
        rs.Open "Select * from " & db_PREFIX & "webcolumn " & addSql, conn, 1, 1 
        If Not rs.EOF Then
            glb_columnId = rs("id") 
            glb_columnName = rs("columnname") 
            glb_columnType = rs("columntype") 
            glb_bodyContent = rs("bodycontent") 
            glb_detailTitle = glb_columnName 
            glb_flags = rs("flags") 
            npagesize = rs("npagesize")                                                     '每页显示条数
            glb_isonhtml = rs("isonhtml")                                                   '是否生成静态网页
            sortSql = " " & rs("sortsql")                                                   '排序SQL

            If rs("webTitle") <> "" Then
                cfg_webTitle = rs("webTitle")                                                   '网址标题
            End If 
            If rs("webKeywords") <> "" Then
                cfg_webKeywords = rs("webKeywords")                                             '网站关键词
            End If 
            If rs("webDescription") <> "" Then
                cfg_webDescription = rs("webDescription")                                       '网站描述
            End If 
            If templateName = "" Then
                If Trim(rs("templatePath")) <> "" Then
                    templateName = rs("templatePath") 
                ElseIf rs("columntype") <> "首页" Then
                    templateName = getDateilTemplate(rs("id"), "List") 
                End If 
            End If 
        End If : rs.Close 
        glb_columnENType = handleColumnType(glb_columnType) 
        glb_url = getColumnUrl(glb_columnName, "name") 

        '文章类列表
        If InStr("|产品|新闻|视频|下载|案例|", "|" & glb_columnType & "|") > 0 Then
            glb_bodyContent = getDetailList(action, defaultListTemplate(), "ArticleDetail", "栏目列表", "*", npagesize, npage, "where parentid=" & glb_columnId & sortSql) 
        '留言类列表
        ElseIf InStr("|留言|", "|" & glb_columnType & "|") > 0 Then
            glb_bodyContent = getDetailList(action, defaultListTemplate(), "GuestBook", "留言列表", "*", npagesize, npage, " where isthrough<>0 " & sortSql) 
        ElseIf glb_columnType = "文本" Then
            '航行栏目加管理
            If Request("gl") = "edit" Then
                glb_bodyContent = "<span>" & glb_bodyContent & "</span>" 
            End If 
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=WebColumn&lableTitle=网站栏目&nPageSize=10&page=&id=" & glb_columnId & "&n=" & getRnd(11) 
            glb_bodyContent = handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span") 

        End If 
    '细节
    ElseIf actionType = "detail" Then
        glb_locationType = "detail" 
        rs.Open "Select * from " & db_PREFIX & "articledetail where id=" & RParam(action, "id"), conn, 1, 1 
        If Not rs.EOF Then
            glb_columnName = getColumnName(rs("parentid")) 
            glb_detailTitle = rs("title") 
            glb_flags = rs("flags") 
            glb_isonhtml = rs("isonhtml")                                                   '是否生成静态网页
            glb_id = rs("id")                                                               '文章ID
            If isMakeHtml = True Then
                glb_url = getHandleRsUrl(rs("fileName"), rs("customAUrl"), "/detail/detail" & rs("id")) 
            Else
                glb_url = handleWebUrl("?act=detail&id=" & rs("id")) 
            End If 

            If rs("webTitle") <> "" Then
                cfg_webTitle = rs("webTitle")                                                   '网址标题
            End If 
            If rs("webKeywords") <> "" Then
                cfg_webKeywords = rs("webKeywords")                                             '网站关键词
            End If 
            If rs("webDescription") <> "" Then
                cfg_webDescription = rs("webDescription")                                       '网站描述
            End If 

            glb_artitleAuthor = rs("author") 
            glb_artitleAdddatetime = rs("adddatetime") 
            glb_upArticle = upArticle(rs("parentid"), "sortrank", rs("sortrank")) 
            glb_downArticle = downArticle(rs("parentid"), "sortrank", rs("sortrank")) 
            glb_aritcleRelatedTags = aritcleRelatedTags(rs("relatedtags")) 
            glb_aritcleSmallImage = rs("smallimage") 
            glb_aritcleBigImage = rs("bigimage") 

            '文章内容
            'glb_bodyContent = "<div class=""articleinfowrap"">[$articleinfowrap$]</div>" & rs("bodycontent") & "[$relatedtags$]<ul class=""updownarticlewrap"">[$updownArticle$]</ul>"
            '上一篇文章，下一篇文章
            'glb_bodyContent = Replace(glb_bodyContent, "[$updownArticle$]", upArticle(rs("parentid"), "sortrank", rs("sortrank")) & downArticle(rs("parentid"), "sortrank", rs("sortrank")))
            'glb_bodyContent = Replace(glb_bodyContent, "[$articleinfowrap$]", "来源：" & rs("author") & " &nbsp; 发布时间：" & format_Time(rs("adddatetime"), 1))
            'glb_bodyContent = Replace(glb_bodyContent, "[$relatedtags$]", aritcleRelatedTags(rs("relatedtags")))

            glb_bodyContent = rs("bodycontent") 

            '文章详细加控制
            If Request("gl") = "edit" Then
                glb_bodyContent = "<span>" & glb_bodyContent & "</span>" 
            End If 
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page=&parentid=&id=" & RParam(action, "id") & "&n=" & getRnd(11) 
            glb_bodyContent = handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span") 

            If templateName = "" Then
                If Trim(rs("templatePath")) <> "" Then
                    templateName = rs("templatePath") 
                Else
                    templateName = getDateilTemplate(rs("parentid"), "Detail") 
                End If 
            End If 

        End If : rs.Close 

    '单页
    ElseIf actionType = "onepage" Then
        rs.Open "Select * from " & db_PREFIX & "onepage where id=" & RParam(action, "id"), conn, 1, 1 
        If Not rs.EOF Then
            glb_detailTitle = rs("title") 
            glb_isonhtml = rs("isonhtml")                                                   '是否生成静态网页
            If isMakeHtml = True Then
                glb_url = getHandleRsUrl(rs("fileName"), rs("customAUrl"), "/page/page" & rs("id")) 
            Else
                glb_url = handleWebUrl("?act=detail&id=" & rs("id")) 
            End If 

            If rs("webTitle") <> "" Then
                cfg_webTitle = rs("webTitle")                                                   '网址标题
            End If 
            If rs("webKeywords") <> "" Then
                cfg_webKeywords = rs("webKeywords")                                             '网站关键词
            End If 
            If rs("webDescription") <> "" Then
                cfg_webDescription = rs("webDescription")                                       '网站描述
            End If 
            '内容
            glb_bodyContent = rs("bodycontent") 


            '文章详细加控制
            If Request("gl") = "edit" Then
                glb_bodyContent = "<span>" & glb_bodyContent & "</span>" 
            End If 
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page=&parentid=&id=" & RParam(action, "id") & "&n=" & getRnd(11) 
            glb_bodyContent = handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span") 


            If templateName = "" Then
                If Trim(rs("templatePath")) <> "" Then
                    templateName = rs("templatePath") 
                Else
                    templateName = "Main_Model.html" 
                'call echo(templateName,"templateName")
                End If 
            End If 

        End If : rs.Close 

    '搜索
    ElseIf actionType = "Search" Then
        templateName = "Main_Model.html" 
        glb_searchKeyWord = Request("wd") 
        addSql = " where title like '%" & glb_searchKeyWord & "%'" 
        npagesize = 20 
        'call echo(npagesize, npage)
        glb_bodyContent = getDetailList(action, defaultListTemplate(), "ArticleDetail", "网站栏目", "*", npagesize, npage, addSql) 

    '加载等待
    ElseIf actionType = "loading" Then
        Call rwend("页面正在加载中。。。") 
    End If 
    '模板为空，则用默认首页模板
    If templateName = "" Then
        templateName = "Index_Model.html"                                               '默认模板
    End If 
    '检测当前路径是否有模板
    If InStr(templateName, "/") = False Then
        templateName = cfg_webTemplate & "/" & templateName 
    End If 
    'call echo("templateName",templateName)
    code = getftext(templateName) 


    code = handleAction(code)                                                       '处理动作
    code = thisPosition(code)                                                       '位置
    code = replaceGlobleVariable(code)                                              '替换全局标签
    code = handleAction(code)                                                       '处理动作    '再来一次，处理数据内容里动作

    code = handleAction(code)                                                       '处理动作
    code = handleAction(code)                                                       '处理动作
    code = thisPosition(code)                                                       '位置
    code = replaceGlobleVariable(code)                                              '替换全局标签
    code = delTemplateMyNote(code)                                                  '删除无用内容

    '格式化HTML
    If InStr(cfg_flags, "|formattinghtml|") > 0 Then
        'code = HtmlFormatting(code)        '简单
        code = handleHtmlFormatting(code, False, 0, "删除空行")                         '自定义
    '格式化HTML第二种
    ElseIf InStr(cfg_flags, "|formattinghtmltow|") > 0 Then
        code = htmlFormatting(code)                                                     '简单
        code = handleHtmlFormatting(code, False, 0, "删除空行")                         '自定义
    '压缩HTML
    ElseIf InStr(cfg_flags, "|ziphtml|") > 0 Then
        code = ziphtml(code) 

    End If 
    '闭合标签
    If InStr(cfg_flags, "|labelclose|") > 0 Then
        code = handleCloseHtml(code, True, "")                                          '图片自动加alt  "|*|",
    End If 

    '在线编辑20160127
    If rq("gl") = "edit" Then
        If InStr(code, "</head>") > 0 Then
            If InStr(code, "jquery.Min.js") = False Then
                code = Replace(code, "</head>", "<script src=""/Jquery/jquery.Min.js""></script></head>") 
            End If 
            code = Replace(code, "</head>", "<script src=""/Jquery/Callcontext_menu.js""></script></head>") 
        End If 
        If InStr(code, "<body>") > 0 Then
        'Code = Replace(Code,"<body>", "<body onLoad=""ContextMenu.intializeContextMenu()"">")
        End If 
    End If 
    'call echo(templateName,templateName)
    makeWebHtml = code 
End Function 

'获得默认细节模板页
Function getDateilTemplate(parentid, templateType)
    Dim templateName 
    templateName = "Main_Model.html" 
    rsx.Open "select * from " & db_PREFIX & "webcolumn where id=" & parentid, conn, 1, 1 
    If Not rsx.EOF Then
        'call echo("columntype",rsx("columntype"))
        If rsx("columntype") = "新闻" Then
            '新闻细节页
            If checkFile(cfg_webTemplate & "/News_" & templateType & ".html") = True Then
                templateName = "News_" & templateType & ".html" 
            End If 
        ElseIf rsx("columntype") = "产品" Then
            '产品细节页
            If checkFile(cfg_webTemplate & "/Product_" & templateType & ".html") = True Then
                templateName = "Product_" & templateType & ".html" 
            End If 
        ElseIf rsx("columntype") = "下载" Then
            '下载细节页
            If checkFile(cfg_webTemplate & "/Down_" & templateType & ".html") = True Then
                templateName = "Down_" & templateType & ".html" 
            End If 

        ElseIf rsx("columntype") = "视频" Then
            '视频细节页
            If checkFile(cfg_webTemplate & "/Video_" & templateType & ".html") = True Then
                templateName = "Video_" & templateType & ".html" 
            End If 
        ElseIf rsx("columntype") = "留言" Then
            '视频细节页
            If checkFile(cfg_webTemplate & "/GuestBook_" & templateType & ".html") = True Then
                templateName = "Video_" & templateType & ".html" 
            End If 
        ElseIf rsx("columntype") = "文本" Then
            '视频细节页
            If checkFile(cfg_webTemplate & "/Page_" & templateType & ".html") = True Then
                templateName = "Page_" & templateType & ".html" 
            End If 
        End If 
    End If : rsx.Close 
    'call echo(templateType,templateName)
    getDateilTemplate = templateName 

End Function 


'生成全部html页面
Sub makeAllHtml(columnType, columnName, columnId)
    Dim action, s, i, nPageSize, nCountSize, nPage, addSql, url, articleSql 
    Call handlePower("生成全部HTML页面")                                            '管理权限处理
    Call writeSystemLog("", "生成全部HTML页面")                                     '系统日志

    isMakeHtml = True 
    '栏目
    Call echo("栏目", "") 
    If columnType <> "" Then
        addSql = "where columnType='" & columnType & "'" 
    End If 
    If columnName <> "" Then
        addSql = getWhereAnd(addSql, "where columnName='" & columnName & "'") 
    End If 
    If columnId <> "" Then
        addSql = getWhereAnd(addSql, "where id in(" & columnId & ")") 
    End If 
    rss.Open "select * from " & db_PREFIX & "webcolumn " & addSql & " order by sortrank asc", conn, 1, 1 
    While Not rss.EOF
        glb_columnName = "" 
        '开启生成html
        If rss("isonhtml") = True Then
            If InStr("|产品|新闻|视频|下载|案例|留言|反馈|招聘|订单|", "|" & rss("columntype") & "|") > 0 Then
                If rss("columntype") = "留言" Then
                    nCountSize = getRecordCount(db_PREFIX & "guestbook", "")                        '记录数
                Else
                    nCountSize = getRecordCount(db_PREFIX & "articledetail", " where parentid=" & rss("id")) '记录数
                End If 
                nPageSize = rss("npagesize") 
                nPage = getPageNumb(CInt(nCountSize), CInt(nPageSize)) 
                If nPage <= 0 Then
                    nPage = 1 
                End If 
                For i = 1 To nPage
                    url = getHandleRsUrl(rss("fileName"), rss("customAUrl"), "/nav" & rss("id")) 
                    glb_filePath = Replace(url, cfg_webSiteUrl, "") 
                    If Right(glb_filePath, 1) = "/" Or glb_filePath = "" Then
                        glb_filePath = glb_filePath & "index.html" 
                    End If 
                    'call echo("glb_filePath",glb_filePath)
                    action = " action actionType='nav' columnName='" & rss("columnname") & "' npage='" & i & "' listfilename='" & glb_filePath & "' " 
                    'call echo("action",action)
                    Call makeWebHtml(action) 
                    If i > 1 Then
                        glb_filePath = Mid(glb_filePath, 1, Len(glb_filePath) - 5) & i & ".html" 
                    End If 
                    s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
                    Call echo(action, s) 
                    If glb_filePath <> "" Then
                        Call createDirFolder(getFileAttr(glb_filePath, "1")) 
                        Call createFileGBK(glb_filePath, code) 
                    End If 
                    doevents() 
                    templateName = ""                                                               '清空模板文件名称
                Next 
            Else
                action = " action actionType='nav' columnName='" & rss("columnname") & "'" 
                Call makeWebHtml(action) 
                glb_filePath = Replace(getColumnUrl(rss("columnname"), "name"), cfg_webSiteUrl, "") 
                If Right(glb_filePath, 1) = "/" Or glb_filePath = "" Then
                    glb_filePath = glb_filePath & "index.html" 
                End If 
                s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
                Call echo(action, s) 
                If glb_filePath <> "" Then
                    Call createDirFolder(getFileAttr(glb_filePath, "1")) 
                    Call createFileGBK(glb_filePath, code) 
                End If 
                doevents() 
                templateName = "" 
            End If 
            conn.Execute("update " & db_PREFIX & "WebColumn set ishtml=true where id=" & rss("id")) '更新导航为生成状态
        End If 
    rss.MoveNext : Wend : rss.Close 

    '单独处理指定栏目对应文章
    If columnId <> "" Then
        articleSql = "select * from " & db_PREFIX & "articledetail where parentid=" & columnId & " order by sortrank asc" 
    '批量处理文章
    ElseIf addSql = "" Then
        articleSql = "select * from " & db_PREFIX & "articledetail order by sortrank asc" 
    End If 
    If articleSql <> "" Then
        '文章
        Call echo("文章", "") 
        rss.Open articleSql, conn, 1, 1 
        While Not rss.EOF
            glb_columnName = "" 
            action = " action actionType='detail' columnName='" & rss("parentid") & "' id='" & rss("id") & "'" 
            'call echo("action",action)
            Call makeWebHtml(action) 
            glb_filePath = Replace(glb_url, cfg_webSiteUrl, "") 
            If Right(glb_filePath, 1) = "/" Then
                glb_filePath = glb_filePath & "index.html" 
            End If 
            s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
            Call echo(action, s) 
            '文件不为空  并且开启生成html
            If glb_filePath <> "" And rss("isonhtml") = True Then
                Call createDirFolder(getFileAttr(glb_filePath, "1")) 
                Call createFileGBK(glb_filePath, code) 
                conn.Execute("update " & db_PREFIX & "ArticleDetail set ishtml=true where id=" & rss("id")) '更新文章为生成状态
            End If 
            templateName = ""                                                               '清空模板文件名称
        rss.MoveNext : Wend : rss.Close 
    End If 

    If addSql = "" Then
        '单页
        Call echo("单页", "") 
        rss.Open "select * from " & db_PREFIX & "onepage order by sortrank asc", conn, 1, 1 
        While Not rss.EOF
            glb_columnName = "" 
            action = " action actionType='onepage' id='" & rss("id") & "'" 
            'call echo("action",action)
            Call makeWebHtml(action) 
            glb_filePath = Replace(glb_url, cfg_webSiteUrl, "") 
            If Right(glb_filePath, 1) = "/" Then
                glb_filePath = glb_filePath & "index.html" 
            End If 
            s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
            Call echo(action, s) 
            '文件不为空  并且开启生成html
            If glb_filePath <> "" And rss("isonhtml") = True Then
                Call createDirFolder(getFileAttr(glb_filePath, "1")) 
                Call createFileGBK(glb_filePath, code) 
                conn.Execute("update " & db_PREFIX & "onepage set ishtml=true where id=" & rss("id")) '更新单页为生成状态
            End If 
            templateName = ""                                                               '清空模板文件名称
        rss.MoveNext : Wend : rss.Close 

    End If 


End Sub 

'复制html到网站
Sub copyHtmlToWeb()
    Dim webDir,toWebDir, toFilePath, filePath, fileName, fileList, splStr, content, s, s1, c, webImages, webCss, webJs, splJs 
    Dim webFolderName, jsFileList, setFileCode, nErrLevel, jsFilePath

    setFileCode = Request("setcode")                                                '设置文件保存编码

    Call handlePower("复制生成HTML页面")                                            '管理权限处理
    Call writeSystemLog("", "复制生成HTML页面")                                     '系统日志

    webFolderName = cfg_webTemplate 
    If Left(webFolderName, 1) = "/" Then
        webFolderName = Mid(webFolderName, 2) 
    End If 
    If Right(webFolderName, 1) = "/" Then
        webFolderName = Mid(webFolderName, 1, Len(webFolderName) - 1) 
    End If 
    If InStr(webFolderName, "/") > 0 Then
        webFolderName = Mid(webFolderName, InStr(webFolderName, "/") + 1) 
    End If 
    webDir = "/htmladmin/" & webFolderName & "/"
	toWebDir="/htmlw" & "eb/viewweb/" 
	call createDirFolder(toWebDir)
	toWebDir = toWebDir & pinYin2(webFolderName) & "/" 
	
	call deleteFolder(toWebDir)				'删除
	call createFolder("/htmlweb/web")		'创建文件夹 防止web文件夹不存在20160504
    Call deleteFolder(webDir) 
    Call createDirFolder(webDir) 
    webImages = webDir & "Images/"
    webCss = webDir & "Css/" 
    webJs = webDir & "Js/"
    Call copyFolder(cfg_webImages, webImages) 
    Call copyFolder(cfg_webCss, webCss) 
    Call createFolder(webJs)                                                        '创建Js文件夹


    '处理Js文件夹
    splJs = Split(getDirJsList(webJs), vbCrLf) 
    For Each filePath In splJs
        If filePath <> "" Then
            toFilePath = webJs & getFileName(filePath) 
            Call echo("js", filePath) 
            Call moveFile(filePath, toFilePath) 
        End If 
    Next 
    '处理Css文件夹
    splStr = Split(getDirCssList(webCss), vbCrLf) 
    For Each filePath In splStr
        If filePath <> "" Then
            content = getftext(filePath) 
            content = Replace(content, cfg_webImages, "../images/") 
			
			content = deleteCssNote(content)			
			content=phptrim(content)
			'设置为utf-8编码 20160527
			if lcase(setFileCode)="utf-8" then
				content=replace(content,"gb2312","utf-8")
			end if
            Call writeToFile(filePath, content, setFileCode) 
            Call echo("css", cfg_webImages) 
        End If 
    Next 
    '复制栏目HTML
    isMakeHtml = True 
    rss.Open "select * from " & db_PREFIX & "webcolumn where isonhtml=true", conn, 1, 1 
    While Not rss.EOF
        glb_filePath = Replace(getColumnUrl(rss("columnname"), "name"), cfg_webSiteUrl, "") 
        If Right(glb_filePath, 1) = "/" Or Right(glb_filePath, 1) = "" Then
            glb_filePath = glb_filePath & "index.html" 
        End If 
        If Right(glb_filePath, 5) = ".html" Then
            If Right(glb_filePath, 11) = "/index.html" Then
                fileList = fileList & glb_filePath & vbCrLf 
            Else
                fileList = glb_filePath & vbCrLf & fileList 
            End If 
            fileName = Replace(glb_filePath, "/", "_") 
            toFilePath = webDir & fileName 
            Call copyfile(glb_filePath, toFilePath) 
            Call echo("导航", glb_filePath) 
        End If 
    rss.MoveNext : Wend : rss.Close 
    '复制文章HTML
    rss.Open "select * from " & db_PREFIX & "articledetail where isonhtml=true", conn, 1, 1 
    While Not rss.EOF
        glb_url = getHandleRsUrl(rss("fileName"), rss("customAUrl"), "/detail/detail" & rss("id")) 
        glb_filePath = Replace(glb_url, cfg_webSiteUrl, "") 
        If Right(glb_filePath, 1) = "/" Or Right(glb_filePath, 1) = "" Then
            glb_filePath = glb_filePath & "index.html" 
        End If 
        If Right(glb_filePath, 5) = ".html" Then
            If Right(glb_filePath, 11) = "/index.html" Then
                fileList = fileList & glb_filePath & vbCrLf 
            Else
                fileList = glb_filePath & vbCrLf & fileList 
            End If 
            fileName = Replace(glb_filePath, "/", "_") 
            toFilePath = webDir & fileName 
            Call copyfile(glb_filePath, toFilePath) 
            Call echo("文章" & rss("title"), glb_filePath) 
        End If 
    rss.MoveNext : Wend : rss.Close 
    '复制单面HTML
    rss.Open "select * from " & db_PREFIX & "onepage where isonhtml=true", conn, 1, 1 
    While Not rss.EOF
        glb_url = getHandleRsUrl(rss("fileName"), rss("customAUrl"), "/page/page" & rss("id")) 
        glb_filePath = Replace(glb_url, cfg_webSiteUrl, "") 
        If Right(glb_filePath, 1) = "/" Or Right(glb_filePath, 1) = "" Then
            glb_filePath = glb_filePath & "index.html" 
        End If 
        If Right(glb_filePath, 5) = ".html" Then
            If Right(glb_filePath, 11) = "/index.html" Then
                fileList = fileList & glb_filePath & vbCrLf 
            Else
                fileList = glb_filePath & vbCrLf & fileList 
            End If 
            fileName = Replace(glb_filePath, "/", "_") 
            toFilePath = webDir & fileName 
            Call copyfile(glb_filePath, toFilePath) 
            Call echo("单页" & rss("title"), glb_filePath) 
        End If 
    rss.MoveNext : Wend : rss.Close 
    '批量处理html文件列表
    'call echo(cfg_webSiteUrl,cfg_webTemplate)
    'call rwend(fileList)
    Dim sourceUrl, replaceUrl 
    splStr = Split(fileList, vbCrLf) 
    For Each filePath In splStr
        If filePath <> "" Then
            filePath = webDir & Replace(filePath, "/", "_") 
            Call echo("filePath", filePath) 
            content = getftext(filePath) 

            For Each s In splStr
                s1 = s 
                If Right(s1, 11) = "/index.html" Then
                    s1 = Left(s1, Len(s1) - 11) & "/" 
                End If 
                sourceUrl = cfg_webSiteUrl & s1 
                replaceUrl = cfg_webSiteUrl & Replace(s, "/", "_") 
                'Call echo(sourceUrl, replaceUrl) 							'屏蔽  否则大量显示20160613
                content = Replace(content, sourceUrl, replaceUrl) 
            Next 
            content = Replace(content, cfg_webSiteUrl, "")                                  '删除网址
            content = Replace(content, cfg_webTemplate, "")                                 '删除模板路径
            'content=nullLinkAddDefaultName(content)
            For Each s In splJs
                If s <> "" Then
                    fileName = getFileName(s) 
                    content = Replace(content, "Images/" & fileName, "js/" & fileName) 
                End If 
            Next 
            If InStr(content, "/Jquery/Jquery.Min.js") > 0 Then
                content = Replace(content, "/Jquery/Jquery.Min.js", "js/Jquery.Min.js") 
                Call copyfile("/Jquery/Jquery.Min.js", webJs & "/Jquery.Min.js") 
            End If 
            content = Replace(content, "<a href="""" ", "<a href=""index.html"" ")    '让首页加index.html
            Call createFileGBK(filePath, content) 
        End If 
    Next 

    '把复制网站夹下的images/文件夹下的js移到js/文件夹下  20160315
    Dim htmlFileList, splHtmlFile, splJsFile, htmlFilePath, jsFileName 
    jsFileList = getDirJsNameList(webImages) 
    htmlFileList = getDirHtmlList(webDir) 
    splHtmlFile = Split(htmlFileList, vbCrLf) 
    splJsFile = Split(jsFileList, vbCrLf) 
    For Each htmlFilePath In splHtmlFile
        content = getftext(htmlFilePath) 
        For Each jsFileName In splJsFile
            content = regExp_Replace(content, "Images/" & jsFileName, "js/" & jsFileName) 
        Next 

        nErrLevel = 0 
        content = handleHtmlFormatting(content, False, nErrLevel, "|删除空行|")         '|删除空行|
        content = handleCloseHtml(content, True, "")                                    '闭合标签
        nErrLevel = checkHtmlFormatting(content) 
        If checkHtmlFormatting(content) = False Then
            Call eerr(htmlFilePath & "(格式化错误)", nErrLevel) 		'注意
        End If 
		'设置为utf-8编码
		if lcase(setFileCode)="utf-8" then
			content=replace(content,"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />","<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />")
		end if
		content=phptrim(content)
        Call writeToFile(htmlFilePath, content, setFileCode)
    Next 
    'images下js移动到js下
    For Each jsFileName In splJsFile
		jsFilePath=webImages & jsFileName
        content = getftext(jsFilePath) 		
		content=phptrim(content)
        Call writeToFile(webJs & jsFileName, content, setFileCode) 
		call deleteFile(jsFilePath)
    Next 
	
	call copyFolder(webDir,toWebDir)
	'使htmlWeb文件夹用php压缩
	if request("isMakeZip")="1" then
	    Call makeHtmlWebToZip(webDir) 
	end if
	'使网站用xml打包20160612
	if request("isMakeXml")="1" then
		call makeHtmlWebToXmlZip("/htmladmin/", webFolderName)
	end if
End Sub 

'使htmlWeb文件夹用php压缩
Function makeHtmlWebToZip(webDir)
    Dim content, splStr, filePath, c, fileArray, fileName, fileType, isTrue 
    Dim webFolderName 
    Dim cleanFileList 
    splStr = Split(webDir, "/") 
    webFolderName = splStr(2) 
    'call eerr(webFolderName,webDir)
    content = getFileFolderList(webDir, True, "全部", "", "全部文件夹", "", "") 
    splStr = Split(content, vbCrLf) 
    For Each filePath In splStr
        If checkfolder(filePath) = False Then
            fileArray = handleFilePathArray(filePath) 
            fileName = LCase(fileArray(2)) 
            fileType = LCase(fileArray(4)) 
            fileName = remoteNumber(fileName) 
            isTrue = True 

            If InStr("|" & cleanFileList & "|", "|" & fileName & "|") > 0 And fileType = "html" Then
                isTrue = False 
            End If 
            If isTrue = True Then
                'call echo(fileType,fileName)
                If c <> "" Then c = c & "|" 
                c = c & Replace(filePath, handlePath("/"), "") 
                cleanFileList = cleanFileList & fileName & "|" 
            End If 
        End If 
    Next 
    Call rw(c) 
    c = c & "|||||" 
    Call createFileGBK("htmlweb/1.txt", c) 
    Call echo("<hr>cccccccccccc", c) 
    '先判断这个文件存在20160309
    If checkFile("/myZIP.php") = True Then
        Call echo("", XMLPost(getHost() & "/myZIP.php?webFolderName=" & webFolderName , "content=" & escape(c))) 
    End If 

End Function 
'使网站用xml打包20160612
function makeHtmlWebToXmlZip(newWebDir,rootDir)
        Dim xmlFileName,xmlSize
        xmlFileName = getIP() & "_update.xml" 

		'newWebDir="\Templates2015\"
		'rootDir="\sharembweb\"
	
        Dim objXmlZIP : Set objXmlZIP = new xmlZIP
            Call objXmlZIP.run(handlePath(newWebDir), handlePath(newWebDir & rootDir), False, xmlFileName)  
			call echo(handlePath(newWebDir), handlePath(newWebDir & rootDir))
        Set objXmlZIP = Nothing 
		doevents
		xmlSize=getFSize(xmlFileName)
		xmlSize=printSpaceValue(xmlSize)
        Call echo("下载xml打包文件", "<a href=/tools/downfile.asp?act=download&downfile=" & xorEnc("/" & xmlFileName, 31380) & " title='点击下载'>点击下载" & xmlFileName & "("& xmlSize &")</a>") 
end function


'生成更新sitemap.xml 20160118
Sub saveSiteMap()
    Dim isWebRunHtml                                                                '是否为html方式显示网站
    Dim changefreg                                                                  '更新频率
    Dim priority                                                                    '优先级
    Dim c, url 
    Call handlePower("修改生成SiteMap")                                             '管理权限处理

    changefreg = Request("changefreg") 
    priority = Request("priority") 
    Call loadWebConfig()                                                            '加载配置
    'call eerr("cfg_flags",cfg_flags)
    If InStr(cfg_flags, "|htmlrun|") > 0 Then
        isWebRunHtml = True 
    Else
        isWebRunHtml = False 
    End If 

    c = c & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf 
    c = c & vbTab & "<urlset xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"">" & vbCrLf 
	dim rsx:Set rsx = CreateObject("Adodb.RecordSet")
    '栏目
    rsx.Open "select * from " & db_PREFIX & "webcolumn order by sortrank asc", conn, 1, 1 
    While Not rsx.EOF
        If rsx("nofollow") = False Then
            c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
            If isWebRunHtml = True Then
                url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/nav" & rsx("id")) 
				url=handleAction(url)
            Else
                url = escape("?act=nav&columnName=" & rsx("columnname")) 
            End If 
            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
            'call echo(cfg_webSiteUrl,url)

            c = c & copystr(vbTab, 3) & "<loc>" & url & "</loc>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<lastmod>" & format_Time(rsx("updatetime"), 2) & "</lastmod>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<changefreq>" & changefreg & "</changefreq>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<priority>" & priority & "</priority>" & vbCrLf 
            c = c & copystr(vbTab, 2) & "</url>" & vbCrLf 
            Call echo("栏目", "<a href=""" & url & """ target='_blank'>" & url & "</a>") 
        End If 
    rsx.MoveNext : Wend : rsx.Close 

    '文章
    rsx.Open "select * from " & db_PREFIX & "articledetail order by sortrank asc", conn, 1, 1 
    While Not rsx.EOF
        If rsx("nofollow") = False Then
            c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
            If isWebRunHtml = True Then
                url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/detail/detail" & rsx("id"))
				url=handleAction(url)
            Else
                url = "?act=detail&id=" & rsx("id") 
            End If
            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
            'call echo(cfg_webSiteUrl,url)

            c = c & copystr(vbTab, 3) & "<loc>" & url & "</loc>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<lastmod>" & format_Time(rsx("updatetime"), 2) & "</lastmod>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<changefreq>" & changefreg & "</changefreq>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<priority>" & priority & "</priority>" & vbCrLf 
            c = c & copystr(vbTab, 2) & "</url>" & vbCrLf 
            Call echo("文章", "<a href=""" & url & """ target='_blank'>" & url & "</a>") 
        End If 
    rsx.MoveNext : Wend : rsx.Close 

    '单页
    rsx.Open "select * from " & db_PREFIX & "onepage order by sortrank asc", conn, 1, 1 
    While Not rsx.EOF
        If rsx("nofollow") = False Then
            c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
            If isWebRunHtml = True Then
                url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/page/detail" & rsx("id")) 
				url=handleAction(url)				
            Else
                url = "?act=onepage&id=" & rsx("id") 
            End If
            url = urlAddHttpUrl(cfg_webSiteUrl, url)
            'call echo(cfg_webSiteUrl,url)

            c = c & copystr(vbTab, 3) & "<loc>" & url & "</loc>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<lastmod>" & format_Time(rsx("updatetime"), 2) & "</lastmod>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<changefreq>" & changefreg & "</changefreq>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<priority>" & priority & "</priority>" & vbCrLf 
            c = c & copystr(vbTab, 2) & "</url>" & vbCrLf 
            Call echo("单页", "<a href=""" & url & """ target='_blank'>" & url & "</a>") 
        End If 
    rsx.MoveNext : Wend : rsx.Close 


    c = c & vbTab & "</urlset>" & vbCrLf 

    Call loadWebConfig() 
	
    Call createfile("sitemap.xml", c) 
    Call echo("生成sitemap.xml文件成功", "<a href='/sitemap.xml' target='_blank'>点击预览sitemap.xml</a>") 

    '判断是否生成sitemap.html
    If Request("issitemaphtml") = "1" Then
        c = "" 
        '第二种
        '栏目
        rsx.Open "select * from " & db_PREFIX & "webcolumn order by sortrank asc", conn, 1, 1 
        While Not rsx.EOF
            If rsx("nofollow") = False Then
                If isWebRunHtml = True Then
                    url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/nav" & rsx("id")) 
					url=handleAction(url)
                Else
                    url = escape("?act=nav&columnName=" & rsx("columnname")) 
                End If 
                url = urlAddHttpUrl(cfg_webSiteUrl, url) 

                c = c & "<li style=""width:20%;""><a href=""" & url & """>" & rsx("columnname") & "</a><ul>" & vbCrLf 

                '文章
                rss.Open "select * from " & db_PREFIX & "articledetail where parentId=" & rsx("id") & " order by sortrank asc", conn, 1, 1 
                While Not rss.EOF
                    If rss("nofollow") = False Then
                        If isWebRunHtml = True Then
                            url = getRsUrl(rss("fileName"), rss("customAUrl"), "/detail/detail" & rss("id")) 
							url=handleAction(url)
                        Else
                            url = "?act=detail&id=" & rss("id") 
                        End If 
                        url = urlAddHttpUrl(cfg_webSiteUrl, url) 
                        c = c & "<li style=""width:20%;""><a href=""" & url & """ target=""_blank"">" & rss("title") & "</a>" & vbCrLf 
                    End If 
                rss.MoveNext : Wend : rss.Close 
                c = c & "</ul></li>" & vbCrLf 


            End If 
        rsx.MoveNext : Wend : rsx.Close 

        '单面
        c = c & "<li style=""width:20%;""><a href=""javascript:;"">单面列表</a><ul>" & vbCrLf 
        rsx.Open "select * from " & db_PREFIX & "onepage order by sortrank asc", conn, 1, 1 
        While Not rsx.EOF
            If rsx("nofollow") = False Then
                c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
                If isWebRunHtml = True Then
                    url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/page/detail" & rsx("id")) 
					url=handleAction(url)
                Else
                    url = "?act=onepage&id=" & rsx("id") 
                End If 
                c = c & "<li style=""width:20%;""><a href=""" & url & """ target=""_blank"">" & rsx("title") & "</a>" & vbCrLf 
            End If 
        rsx.MoveNext : Wend : rsx.Close 
        c = c & "</ul></li>" & vbCrLf 

        Dim templateContent 
        templateContent = getftext(adminDir & "/template_SiteMap.html") 


        templateContent = Replace(templateContent, "{$content$}", c) 
        templateContent = Replace(templateContent, "{$Web_Title$}", cfg_webTitle) 
		 
		
        Call createfile("sitemap.html", templateContent) 
        Call echo("生成sitemap.html文件成功", "<a href='/sitemap.html' target='_blank'>点击预览sitemap.html</a>") 
    End If 
    Call writeSystemLog("", "保存sitemap.xml")                                      '系统日志
End Sub 
%>       