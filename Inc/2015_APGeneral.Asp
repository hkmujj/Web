<% 
'ASP和PHP通用函数


'文章相关标签 组  待改进
Function aritcleRelatedTags(relatedTags)
    Dim c, splStr, s, url 
    splStr = Split(relatedTags, ",") 
    For Each s In splStr
        If s <> "" Then
            If c <> "" Then
                c = c & "," 
            End If 
            url = getColumnUrl(s, "name") 
            c = c & "<a href=""" & url & """ rel=""category tag"" class=""ablue"">" & s & "</a>" & vbCrLf 
        End If 
    Next 

    c = "<footer class=""articlefooter"">" & vbCrLf & "标签： " & c & "</footer>" & vbCrLf 
    aritcleRelatedTags = c 
End Function 


'获得随机文章id列表
Function getRandArticleId(addSql, topNumb)
    Dim splStr, s, c, nIndex 
    rs.Open "select * from " & db_PREFIX & "articledetail " & addSql, conn, 1, 1 
    While Not rs.EOF
        If c <> "" Then c = c & "," 
        c = c & rs("id") 
    rs.MoveNext : Wend : rs.Close 
    getRandArticleId = randomShow(c, ",", 4) 
    splStr = Split(c, ",") : c = "" : nIndex = 0 
    For Each s In splStr
        If c <> "" Then c = c & "," 
        c = c & s 
        nIndex = nIndex + 1 
        If nIndex >= topNumb Then Exit For 
    Next 
    getRandArticleId = c 
End Function 

'上一篇文章 这里面的sortrank(排序)也可以改为id,在引用的时候就要用id
Function upArticle(parentid, lableName, lableValue)
    Dim sql 
    sql = "select * from " & db_PREFIX & "articledetail where parentid=" & parentid & " and " & lableName & "<" & lableValue & " order by " & lableName & " desc" 
    upArticle = handleUpDownArticle("上一篇：", sql) 
End Function 
'下一篇文章
Function downArticle(parentid, lableName, lableValue)
    Dim sql 
    sql = "select * from " & db_PREFIX & "articledetail where parentid=" & parentid & " and " & lableName & ">" & lableValue & " order by " & lableName & " asc" 
    downArticle = handleUpDownArticle("下一篇：", sql) 
End Function 
'处理上下页
Function handleUpDownArticle(lableTitle, sql)
    Dim c, url 
    'call echo("sql",sql)
    rsx.Open sql, conn, 1, 1 
    If Not rsx.EOF Then
        If isMakeHtml = True Then
            url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/detail/detail" & rsx("id")) 
        Else
            url = handleWebUrl("?act=detail&id=" & rsx("id")) 
        End If 
        c = "<a href=""" & url & """>" & lableTitle & rsx("title") & "</a>" 
    Else
        c = lableTitle & "没有" 
    End If : rsx.Close 
    handleUpDownArticle = c 
End Function 
'获得RS网址 配置上一页 下一页
Function getRsUrl(ByVal fileName, customAUrl, defaultFileName)
    Dim url 
    '用默认文件名称
    If fileName = "" Then
        fileName = defaultFileName 
    End If 
    '网址
    If fileName <> "" Then
        fileName = LCase(fileName)                                                      '让文件名称小写20160315
        url = fileName 
        If InStr(LCase(url), ".html") = False And Right(url, 1) <> "/" Then
            url = url & ".html" 
        End If 
    End If 
    If Trim(customAUrl) <> "" Then
        url = Trim(customAUrl) 
    End If 
    If InStr(cfg_flags, "|addwebsite|") > 0 Then
        'url = replaceGlobleVariable(url)   '替换全局变量
        If InStr(url, "$cfg_websiteurl$") = False And InStr(url, "{$GetColumnUrl ") = False And InStr(url, "{$GetArticleUrl ") = False And InStr(url, "{$GetOnePageUrl ") = False Then
            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
        End If 
    End If 
    getRsUrl = url 
End Function 
'获得处理后RS网址
Function getHandleRsUrl(fileName, customAUrl, defaultFileName)
    Dim url 
    url = getRsUrl(fileName, customAUrl, defaultFileName) 
    '因为URL如果为自定义的则需要处理下全局变量，这样程序运行又会变慢，不就可以使用生成HTML方法解决这个问题，20160308
    url = replaceGlobleVariable(url) 
    getHandleRsUrl = url 
End Function 

'获得单页url 20160114
Function getOnePageUrl(title)
    Dim url 
    rsx.Open "select * from " & db_PREFIX & "onepage where title='" & title & "'", conn, 1, 1 
    If Not rsx.EOF Then
        If isMakeHtml = True Then
            url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/page/page" & rsx("id")) 
        Else
            url = handleWebUrl("?act=onepage&id=" & rsx("id")) 
            If rsx("customAUrl") <> "" Then
                url = rsx("customAUrl") 
            End If 
        End If 
    End If : rsx.Close 

    getOnePageUrl = url 
End Function 
'获得文章
Function getArticleUrl(title)
    Dim url 
    rsx.Open "select * from " & db_PREFIX & "articledetail where title='" & title & "'", conn, 1, 1 
    If Not rsx.EOF Then
        If isMakeHtml = True Then
            url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/detail/" & rsx("id")) 
        Else
            url = handleWebUrl("?act=article&id=" & rsx("id")) 
            If rsx("customAUrl") <> "" Then
                url = rsx("customAUrl") 
            End If 
        End If 
    End If : rsx.Close 

    getArticleUrl = url 
End Function 
'获得栏目url 20160114
Function getColumnUrl(columnNameOrId, sType)
    Dim url, addSql 

    columnNameOrId = replaceGlobleVariable(columnNameOrId)                          '处理动作 <a href="{$GetColumnUrl columnname='[$glb_columnName$]' $}" >更多图片</a>

    If sType = "name" Then
        addSql = " where columnname='" & columnNameOrId & "'" 
    Else
        addSql = " where id=" & columnNameOrId & "" 
    End If 
    rsx.Open "select * from " & db_PREFIX & "webcolumn" & addSql, conn, 1, 1 
    If Not rsx.EOF Then
        If isMakeHtml = True Then
            url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/nav" & rsx("id")) 
        Else
            url = handleWebUrl("?act=nav&columnName=" & rsx("columnname")) 
            If rsx("customAUrl") <> "" Then
                url = rsx("customAUrl") 
            End If 
        End If 
    End If : rsx.Close 

    getColumnUrl = url 
End Function 

'获得文章标题对应的id
Function getArticleId(title)
    title = Replace(title, "'", "")                                                 '注意，这个不能留
    getArticleId = -1 
    rsx.Open "Select * from " & db_PREFIX & "ArticleDetail where title='" & title & "'", conn, 1, 1 
    If Not rsx.EOF Then
        getArticleId = rsx("id") 
    End If : rsx.Close 
End Function 

'获得栏目名称对应的id
Function getColumnId(columnName)
    columnName = Replace(columnName, "'", "")                                       '注意，这个不能留
    getColumnId = -1 
    rsx.Open "Select * from " & db_PREFIX & "webcolumn where columnName='" & columnName & "'", conn, 1, 1 
    If Not rsx.EOF Then
        getColumnId = rsx("id") 
    End If : rsx.Close 
End Function 


'获得栏目ID对应的名称
Function getColumnName(columnID)
    If columnID <> "" Then
        rsx.Open "Select * from " & db_PREFIX & "webcolumn where id=" & columnID, conn, 1, 1 
        If Not rsx.EOF Then
            getColumnName = rsx("columnName") 
        End If : rsx.Close 
    End If 
End Function 




'获得栏目ID对应的类型
Function getColumnType(columnID)
    rsx.Open "Select * from " & db_PREFIX & "webcolumn where id=" & columnID, conn, 1, 1 
    If Not rsx.EOF Then
        getColumnType = rsx("columnType") 
    End If : rsx.Close 
End Function 


'获得栏目ID对应的内容
Function getColumnBodyContent(columnID)
    rsx.Open "Select * from " & db_PREFIX & "webcolumn where id=" & columnID, conn, 1, 1 
    If Not rsx.EOF Then
        getColumnBodyContent = rsx("bodyContent") 
    End If : rsx.Close 
End Function 







'网站统计2014
Function webStat(folderPath)
    Dim dateTime, content, splStr 
    Dim thisUrl, goToUrl, caiShu, c, fileName, co, ie, xp 
    goToUrl = Request.ServerVariables("HTTP_REFERER") 
    thisUrl = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME") 
    caiShu = Request.ServerVariables("QUERY_STRING") 
    If caiShu <> "" Then
        thisUrl = thisUrl & "?" & caiShu 
    End If 
    goToUrl = Request("GoToUrl") 
    thisUrl = Request("ThisUrl") 
    co = Request.QueryString("co") 
    dateTime = Now() 
    content = Request.ServerVariables("HTTP_USER_AGENT") 
    content = Replace(content, "MSIE", "Internet Explorer") 
    content = Replace(content, "NT 5.0", "2000") 
    content = Replace(content, "NT 5.1", "XP") 
    content = Replace(content, "NT 5.2", "2003") 

    splStr = Split(content & ";;;;", ";") 
    ie = splStr(1) 
    xp = Trim(splStr(2)) 
    If Right(xp, 1) = ")" Then xp = Mid(xp, 1, Len(xp) - 1) 
    c = "来访" & goToUrl & vbCrLf 
    c = c & "当前：" & thisUrl & vbCrLf 
    c = c & "时间：" & dateTime & vbCrLf 
    c = c & "IP:" & getIP() & vbCrLf 
    c = c & "IE:" & getBrType("") & vbCrLf 
    c = c & "Cookies=" & co & vbCrLf 
    c = c & "XP=" & xp & vbCrLf 
    c = c & "Screen=" & Request("screen") & vbCrLf                                  '屏幕分辨率
    c = c & "用户信息=" & Request.ServerVariables("HTTP_USER_AGENT") & vbCrLf       '用户信息

    c = c & "-------------------------------------------------" & vbCrLf 
    'c=c & "CaiShu=" & CaiShu & vbcrlf
    fileName = folderPath & format_Time(Now(), 2) & ".txt" 
    Call createAddFile(fileName, c) 
    c = c & vbCrLf & fileName 
    c = Replace(c, vbCrLf, "\n") 
    c = Replace(c, """", "\""") 
    'Response.Write("eval(""var MyWebStat=\""" & C & "\"""")")

    Dim splxx, nIP, nPV, ipList, s, ip 
    '判断是否显示回显记录
    If Request("stype") = "display" Then
        content = getftext(fileName) 
        splxx = Split(content, vbCrLf & "-------------------------------------------------" & vbCrLf) 
        nIP = 0 
        nPV = 0 
        ipList = "" 
        For Each s In splxx
            If InStr(s, "当前：") > 0 Then
                s = vbCrLf & s & vbCrLf 
                ip = ADSql(getStrCut(s, vbCrLf & "IP:", vbCrLf, 0)) 
                nPV = nPV + 1 
                If InStr(vbCrLf & ipList & vbCrLf, vbCrLf & ip & vbCrLf) = False Then
                    ipList = ipList & ip & vbCrLf 
                    nIP = nIP + 1 
                End If 
            End If 
        Next 
        Call rw("document.write('网长统计 | 今日IP[" & nIP & "] | 今日PV[" & nPV & "] ')") 
    End If 
    webStat = c 
End Function 

'判断传值是否相等
Function checkFunValue(action, FunName)
    checkFunValue =(Left(action, Len(FunName)) = FunName) 
End Function 
'HTML标签参数自动添加(target|title|alt|id|class|style|)    辅助类
Function setHtmlParam(content, ParamList)
    Dim splStr, startStr, endStr, c, paramValue, ReplaceStartStr 
    endStr = "'" 
    splStr = Split(ParamList, "|") 
    For Each startStr In splStr
        startStr = Trim(startStr) 
        If startStr <> "" Then
            '替换开始字符   因为开始字符类型可变 不同
            ReplaceStartStr = startStr 
            If Left(ReplaceStartStr, 3) = "img" Then
                ReplaceStartStr = Mid(ReplaceStartStr, 4) 
            ElseIf Left(ReplaceStartStr, 1) = "a" Then
                ReplaceStartStr = Mid(ReplaceStartStr, 2) 
            ElseIf InStr("|ul|li|", "|" & Left(ReplaceStartStr, 2) & "|") > 0 Then
                ReplaceStartStr = Mid(ReplaceStartStr, 3) 
            End If 
            ReplaceStartStr = " " & ReplaceStartStr & "='" 

            startStr = " " & startStr & "='" 
            If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
                paramValue = strCut(content, startStr, endStr, 2) 
                paramValue = handleInModule(paramValue, "end")                       '处理内部模块
                c = c & ReplaceStartStr & paramValue & endStr 
            End If 
        End If 
    Next 
    setHtmlParam = c 
End Function 
%>      
