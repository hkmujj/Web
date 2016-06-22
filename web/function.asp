<% 

'调用function文件函数
Function callFunction()
    Select Case Request("stype")
    Case "updateWebsiteStat" : updateWebsiteStat()                                  '更新网站统计
    Case "clearWebsiteStat" : clearWebsiteStat()                                    '清空网站统计
    Case "updateTodayWebStat" : updateTodayWebStat()                                '更新网站今天统计
    Case "websiteDetail" : websiteDetail()                                          '详细网站统计
	case "displayAccessDomain" : displayAccessDomain()										'显示访问域名
		
        Case Else : Call eerr("function1页里没有动作", Request("stype"))
    End Select
End Function

'显示访问域名
function displayAccessDomain()
	dim visitWebSite,visitWebSiteList,urlList,nOK
    Call openconn() 							
	nOK=0		
    rs.Open "select * from " & db_PREFIX & "websitestat", conn, 1, 1
    While Not rs.EOF
		visitWebSite=lcase(getWebSite(rs("visiturl")))
		'call echo("visitWebSite",visitWebSite)
		if instr(vbcrlf & visitWebSiteList & vbcrlf,vbcrlf & visitWebSite & vbcrlf)=false then
			if visitWebSite<>lcase(getWebSite(webDoMain())) then
				visitWebSiteList=visitWebSiteList & visitWebSite & vbcrlf
				nOK=nOK+1
				urlList=urlList & nOK & "、<a href='" & rs("visiturl") & "' target='_blank'>" & rs("visiturl") & "</a><br>"
			end if
		end if
	rs.movenext:wend:rs.close
	call echo("显示访问域名","操作完成 <a href='javascript:history.go(-1)'>点击返回</a>")
	call rwend(visitWebSiteList & "<br><hr><br>" & urlList)
end function
'获得处理后表列表 20160313
Function getHandleTableList()
    Dim s, lableStr 
    lableStr = "表列表[" & Request("mdbpath") & "]" 
    If WEB_CACHEContent = "" Then
        WEB_CACHEContent = getftext(WEB_CACHEFile) 
    End If 
    s = getConfigContentBlock(WEB_CACHEContent, "#" & lableStr & "#") 
    If s = "" Then
        s = LCase(getTableList()) 
        s = "|" & Replace(s, vbCrLf, "|") & "|" 
        WEB_CACHEContent = setConfigFileBlock(WEB_CACHEFile, s, "#" & lableStr & "#") 
		if isCacheTip=true then
        	Call echo("缓冲", lableStr)
		end if 
    End If 
    getHandleTableList = s 
End Function 

'获得处理的字段列表
Function getHandleFieldList(tableName, sType)
    Dim s 
    If WEB_CACHEContent = "" Then
        WEB_CACHEContent = getftext(WEB_CACHEFile) 
    End If 
    s = getConfigContentBlock(WEB_CACHEContent, "#" & tableName & sType & "#") 

    If s = "" Then
        If sType = "字段配置列表" Then
            s = LCase(getFieldConfigList(tableName)) 
        Else
            s = LCase(getFieldList(tableName)) 
        End If 
        WEB_CACHEContent = setConfigFileBlock(WEB_CACHEFile, s, "#" & tableName & sType & "#")
		if isCacheTip=true then
        	Call echo("缓冲", tableName & sType) 
		end if
    End If 
    getHandleFieldList = s 
End Function 
'读模板内容 20160310
Function getTemplateContent(templateFileName)
    Call loadWebConfig() 
    '读模板
    Dim templateFile, customTemplateFile, c
    customTemplateFile = ROOT_PATH & "template/" & db_PREFIX & "/" & templateFileName 
    '为手机端
    If checkMobile() = True or request("m")="mobile" Then			
        templateFile = ROOT_PATH & "/Template/mobile/" & templateFileName 
    End If 
    '判断手机端文件是否存在20160330
    If checkFile(templateFile) = False Then
        If checkFile(customTemplateFile) = True Then
            templateFile = customTemplateFile 
        Else
            templateFile = ROOT_PATH & templateFileName 
        End If 
    End If 
    c = getFText(templateFile) 
    c = replaceLableContent(c) 
    getTemplateContent = c
End Function 
'替换标签内容
Function replaceLableContent(content)
    content = Replace(content, "{$webVersion$}", webVersion)                        '网站版本
    content = Replace(content, "{$Web_Title$}", cfg_webTitle)                       '网站标题
    content = Replace(content, "{$EDITORTYPE$}", EDITORTYPE)                        'ASP与PHP
    content = Replace(content, "{$adminDir$}", adminDir)                            '后台目录

    content = Replace(content, "[$adminId$]", Session("adminId"))              '管理员ID
    content = Replace(content, "{$adminusername$}", Session("adminusername"))       '管理账号名称
    content = Replace(content, "{$EDITORTYPE$}", EDITORTYPE)                        '程序类型
    content = Replace(content, "{$WEB_VIEWURL$}", WEB_VIEWURL)                      '前台
    content = Replace(content, "{$webVersion$}", webVersion)                        '版本
    content = Replace(content, "{$WebsiteStat$}", getConfigFileBlock(WEB_CACHEFile, "#访客信息#")) '最近访客信息


    content = Replace(content, "{$DB_PREFIX$}", db_PREFIX)                          '表前缀
    content = Replace(content, "{$adminflags$}", IIF(Session("adminflags") = "|*|", "超级管理员", "普通管理员")) '管理员类型
    content = Replace(content, "{$SERVER_SOFTWARE$}", Request.ServerVariables("SERVER_SOFTWARE")) '服务器版本
    content = Replace(content, "{$SERVER_NAME$}", Request.ServerVariables("SERVER_NAME")) '服务器网址
    content = Replace(content, "{$LOCAL_ADDR$}", Request.ServerVariables("LOCAL_ADDR")) '服务器IP
    content = Replace(content, "{$SERVER_PORT$}", Request.ServerVariables("SERVER_PORT")) '服务器端口
    content = replaceValueParam(content, "mdbpath", Request("mdbpath")) 
    content = replaceValueParam(content, "webDir", webDir) 
	
	'20160614
	if EDITORTYPE="php" then
		content = Replace(content, "{$EDITORTYPE_PHP$}", "php")                          '给phpinc/用
	end if
	content = Replace(content, "{$EDITORTYPE_PHP$}", "")                          '给phpinc/用 
	
    replaceLableContent = content 
End Function 

'文章列表旗
Function displayFlags(flags)
    Dim c 
    '头条[h]
    If InStr("|" & flags & "|", "|h|") > 0 Then
        c = c & "头 " 
    End If 
    '推荐[c]
    If InStr("|" & flags & "|", "|c|") > 0 Then
        c = c & "推 " 
    End If 
    '幻灯[f]
    If InStr("|" & flags & "|", "|f|") > 0 Then
        c = c & "幻 " 
    End If 
    '特荐[a]
    If InStr("|" & flags & "|", "|a|") > 0 Then
        c = c & "特 " 
    End If 
    '滚动[s]
    If InStr("|" & flags & "|", "|s|") > 0 Then
        c = c & "滚 " 
    End If 
    '加粗[b]
    If InStr("|" & flags & "|", "|b|") > 0 Then
        c = c & "粗 " 
    End If 
    If c <> "" Then c = "[<font color=""red"">" & c & "</font>]" 

    displayFlags = c 
End Function 


'栏目类别循环配置       showColumnList(-1, 0,defaultList)   nCount为深度值
Function showColumnList(ByVal parentid, ByVal tableName, fileName, ByVal thisPId, nCount, ByVal action)
    Dim i, s, c, selectcolumnname, selStr, url, isFocus, sql, addSql 
    Dim rs : Set rs = CreateObject("Adodb.RecordSet")
        Dim fieldNameList, splFieldName, k, fieldName, replaceStr, startStr, endStr, topNumb, modI 
        Dim subHeaderStr, subFooterStr 

        subHeaderStr = getStrCut(action, "[subheader]", "[/subheader]", 2) 
        subFooterStr = getStrCut(action, "[subfooter]", "[/subfooter]", 2) 

        fieldNameList = getHandleFieldList(db_PREFIX & tableName, "字段列表") 
        splFieldName = Split(fieldNameList, ",") 
        sql = "select * from " & db_PREFIX & tableName & " where parentid=" & parentid 
        '处理追加SQL
        startStr = "[sql-" & nCount & "]" : endStr = "[/sql-" & nCount & "]" 
        If InStr(action, startStr) = False And InStr(action, endStr) = False Then
            startStr = "[sql]" : endStr = "[/sql]" 
        End If 
        addSql = getStrCut(action, startStr, endStr, 2) 
        If addSql <> "" Then
            sql = getWhereAnd(sql, addSql) 
        End If 
        'call echo("addsql",addsql)
        rs.Open sql & " order by sortrank asc", conn, 1, 1 
        For i = 1 To rs.RecordCount
            If Not rs.EOF Then
                selStr = "" 
                isFocus = False 
                If CStr(rs("id")) = CStr(thisPId) Then
                    selStr = " selected " 
                    isFocus = True 
                End If 

                '网址判断
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
                    If InStr(action, startStr) = False And i Mod modI = 0 Then
                        startStr = "[list-mod" & modI & "]" : endStr = "[/list-mod" & modI & "]" 
                        If InStr(action, startStr) > 0 Then
                            Exit For 
                        End If 
                    End If 
                Next 

                '没有则用默认
                If InStr(action, startStr) = False Then
                    startStr = "[list]" : endStr = "[/list]" 
                End If 

                'call rwend(action)
                'call echo(startStr,endStr)
                If InStr(action, startStr) > 0 And InStr(action, endStr) > 0 Then
                    s = strCut(action, startStr, endStr, 2) 

                    s = replaceValueParam(s, "id", rs("id")) 
                    s = replaceValueParam(s, "selected", selStr) 
                    selectcolumnname = rs(fileName) 
                    If nCount >= 1 Then
                        selectcolumnname = copystr("&nbsp;&nbsp;", nCount) & "├─" & selectcolumnname 
                    End If 
                    s = replaceValueParam(s, "selectcolumnname", selectcolumnname) 


                    For k = 0 To UBound(splFieldName)
                        If splFieldName(k) <> "" Then
                            fieldName = splFieldName(k) 
                            replaceStr = rs(fieldName) & "" 

                            s = replaceValueParam(s, fieldName, replaceStr) 
                        End If 
                    Next 

                    'url = WEB_VIEWURL & "?act=nav&columnName=" & rs(fileName)             '以栏目名称显示列表
                    url = WEB_VIEWURL & "?act=nav&id=" & rs("id")                               '以栏目ID显示列表

                    '自定义网址
                    If Trim(rs("customaurl")) <> "" Then
                        url = Trim(rs("customaurl")) 
                    End If 
                    s = Replace(s, "[$viewWeb$]", url) 
                    s = replaceValueParam(s, "url", url) 

                    If EDITORTYPE = "php" Then
                        s = Replace(s, "[$phpArray$]", "[]") 
                    Else
                        s = Replace(s, "[$phpArray$]", "") 
                    End If 

                    's=copystr("",nCount) & rs("columnname") & "<hr>"
                    c = c & s & vbCrLf 
                    s = showColumnList(rs("id"), tableName, fileName, thisPId, nCount + 1, action) 
                    If s <> "" Then s = vbCrLf & subHeaderStr & s & subFooterStr 
                    c = c & s 
                End If 
            End If 
        rs.MoveNext : Next : rs.Close 
        showColumnList = c 
End Function 
'msg1  辅助
Function getMsg1(msgStr, url)
    Dim content 
    content = getFText(ROOT_PATH & "msg.html") 
    msgStr = msgStr & "<br>" & jsTiming(url, 5) 
    content = Replace(content, "[$msgStr$]", msgStr) 
    content = Replace(content, "[$url$]", url)
	
	
    content = replaceL(content, "提示信息")
    content = replaceL(content, "如果您的浏览器没有自动跳转，请点击这里")
    content = replaceL(content, "倒计时")
	
	
    getMsg1 = content 
End Function 

'检测权力
Function checkPower(powerName)
    If InStr("|" & Session("adminflags") & "|", "|" & powerName & "|") > 0 Or InStr("|" & Session("adminflags") & "|", "|*|") > 0 Then
        checkPower = True 
    Else
        checkPower = False 
    End If 
End Function 
'处理后台管理权限
Function handlePower(powerName)
    If checkPower(powerName) = False Then
        Call eerr("提示", "你没有【" & powerName & "】权限，<a href='javascript:history.go(-1);'>点击返回</a>") 
    End If 
End Function 
 



'显示管理列表
Function dispalyManage(actionName, lableTitle, ByVal nPageSize, addSql)
    Call handlePower("显示" & lableTitle)                                           '管理权限处理
    Call loadWebConfig() 
    Dim content, i, s, c, fieldNameList, sql, action 
    Dim x, url, nCount, nPage 
    Dim idInputName 

    Dim tableName, j, splxx 
    Dim fieldName                                                                   '字段名称
    Dim splFieldName                                                                '分割字段
    Dim searchfield, keyWord                                                        '搜索字段，搜索关键词
    Dim parentid                                                                    '栏目id

    Dim replaceStr                                                                  '替换字符
    tableName = LCase(actionName)                                                   '表名称

    searchfield = Request("searchfield")                                            '获得搜索字段值
    keyWord = Request("keyword")                                                    '获得搜索关键词值
    If Request.Form("parentid") <> "" Then
        parentid = Request.Form("parentid") 
    Else
        parentid = Request.QueryString("parentid") 
    End If 

    Dim id 
    id = rq("id") 

    fieldNameList = getHandleFieldList(db_PREFIX & tableName, "字段列表") 

    fieldNameList = specialStrReplace(fieldNameList)                                '特殊字符处理
    splFieldName = Split(fieldNameList, ",")                                        '字段分割成数组

    '读模板
    content = getTemplateContent("manage_" & tableName & ".html") 

    action = getStrCut(content, "[list]", "[/list]", 2) 
    '网站栏目单独处理      栏目不一样20160301
    If actionName = "WebColumn" Then
        action = getStrCut(content, "[action]", "[/action]", 1) 
        content = Replace(content, action, showColumnList( -1, "WebColumn", "columnname", "", 0, action)) 
    ElseIf actionName = "ListMenu" Then
        action = getStrCut(content, "[action]", "[/action]", 1) 
        content = Replace(content, action, showColumnList( -1, "listmenu", "title", "", 0, action)) 
    Else
        If keyWord <> "" And searchfield <> "" Then
            addSql = getWhereAnd(" where " & searchfield & " like '%" & keyWord & "%' ", addSql) 
        End If 
        If parentid <> "" Then
            addSql = getWhereAnd(" where parentid=" & parentid & " ", addSql) 
        End If 
        'call echo(tableName,addsql)
        sql = "select * from " & db_PREFIX & tableName & " " & addSql 
        '检测SQL
        If checksql(sql) = False Then
            Call errorLog("出错提示：<br>action=" & action & "<hr>sql=" & sql & "<br>") 
            Exit Function 
        End If 
        rs.Open sql, conn, 1, 1 
        nCount = rs.RecordCount 
        nPage = Request("page") 
        content = Replace(content, "[$pageInfo$]", webPageControl(nCount, nPageSize, nPage, url, "")) 
        content = Replace(content, "[$accessSql$]", sql) 

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
            s = Replace(action, "[$id$]", rs("id")) 
            For j = 0 To UBound(splFieldName)
                If splFieldName(j) <> "" Then
                    splxx = Split(splFieldName(j) & "|||", "|") 
                    fieldName = splxx(0) 
                    replaceStr = rs(fieldName) & "" 
                    '对文章旗处理
                    If fieldName = "flags" Then
                        replaceStr = displayFlags(replaceStr) 
                    End If 
                    'call echo("fieldname",fieldname)
                    's = Replace(s, "[$" & fieldName & "$]", replaceStr)
                    s = replaceValueParam(s, fieldName, replaceStr) 

                End If 
            Next 

            idInputName = "id" 
            s = Replace(s, "[$selectid$]", "<input type='checkbox' name='" & idInputName & "' id='" & idInputName & "' value='" & rs("id") & "' >") 
            s = Replace(s, "[$phpArray$]", "") 
            url = "【NO】" 
            If actionName = "ArticleDetail" Then
                url = WEB_VIEWURL & "?act=detail&id=" & rs("id") 
            ElseIf actionName = "OnePage" Then
                url = WEB_VIEWURL & "?act=onepage&id=" & rs("id") 
            '给评论加预览=文章  20160129
            ElseIf actionName = "TableComment" Then
                url = WEB_VIEWURL & "?act=detail&id=" & rs("itemid") 
            End If 
            '必需有自定义字段
            If InStr(fieldNameList, "customaurl") > 0 Then
                '自定义网址
                If Trim(rs("customaurl")) <> "" Then
                    url = Trim(rs("customaurl")) 
                End If 
            End If 
            s = Replace(s, "[$viewWeb$]", url) 
            s = replaceValueParam(s, "cfg_websiteurl", cfg_webSiteUrl) 

            c = c & s 
        rs.MoveNext : Next : rs.Close 
        content = Replace(content, "[list]" & action & "[/list]", c) 
        '表单提交处理，parentid(栏目ID) searchfield(搜索字段) keyword(关键词) addsql(排序)
        url = "?page=[id]&addsql=" & Request("addsql") & "&keyword=" & Request("keyword") & "&searchfield=" & Request("searchfield") & "&parentid=" & Request("parentid") 
        url = getUrlAddToParam(getUrl(), url, "replace") 
        'call echo("url",url)
        content = Replace(content, "[list]" & action & "[/list]", c) 

    End If 

    If InStr(content, "[$input_parentid$]") > 0 Then
        action = "[list]<option value=""[$id$]""[$selected$]>[$selectcolumnname$]</option>[/list]" 
        c = "<select name=""parentid"" id=""parentid""><option value="""">≡ 选择栏目 ≡</option>" & showColumnList( -1, "webcolumn", "columnname", parentid, 0, action) & vbCrLf & "</select>" 
        content = Replace(content, "[$input_parentid$]", c)                        '上级栏目
    End If 

    content = replaceValueParam(content, "searchfield", Request("searchfield"))     '搜索字段
    content = replaceValueParam(content, "keyword", Request("keyword"))             '搜索关键词
    content = replaceValueParam(content, "nPageSize", Request("nPageSize"))         '每页显示条数
    content = replaceValueParam(content, "addsql", Request("addsql"))               '追加sql值条数
    content = replaceValueParam(content, "tableName", tableName)                    '表名称
    content = replaceValueParam(content, "actionType", Request("actionType"))       '动作类型
    content = replaceValueParam(content, "lableTitle", Request("lableTitle"))       '动作标题
    content = replaceValueParam(content, "id", id)                                  'id
    content = replaceValueParam(content, "page", Request("page"))                   '页

    content = replaceValueParam(content, "parentid", Request("parentid"))           '栏目id


    url = getUrlAddToParam(getThisUrl(), "?parentid=&keyword=&searchfield=&page=", "delete") 
    content = replaceValueParam(content, "position", "系统管理 > <a href='" & url & "'>" & lableTitle & "列表</a>") 'position位置


    content = Replace(content, "{$EDITORTYPE$}", EDITORTYPE)                        'asp与phh
    content = Replace(content, "{$WEB_VIEWURL$}", WEB_VIEWURL)                      '前端浏览网址
    content = Replace(content, "{$Web_Title$}", cfg_webTitle) 
 
    content = content & stat2016(True) 
	
	content=handleDisplayLanguage(content,"handleDisplayLanguage")			'语言处理
	
    Call rw(content) 
End Function  

'添加修改界面
Function addEditDisplay(actionName, lableTitle, ByVal fieldNameList)
    Dim content, addOrEdit, splxx, i, j, s, c, tableName, url, aStr 
    Dim fieldName                                                                   '字段名称
    Dim splFieldName                                                                '分割字段
    Dim fieldSetType                                                                '字段设置类型
    Dim fieldValue                                                                  '字段值
    Dim sql                                                                         'sql语句
    Dim defaultList                                                                 '默认列表
    Dim flagsInputName                                                              '旗input名称给ArticleDetail用
    Dim titlecolor                                                                  '标题颜色
    Dim flags                                                                       '旗
    Dim splStr, fieldConfig, defaultFieldValue, postUrl 
    Dim subTableName, subFileName                                                   '子列表的表名称，子列表字段名称


    Dim id 
    id = rq("id") 
    addOrEdit = "添加" 
    If id <> "" Then
        addOrEdit = "修改" 
    End If 

    If InStr(",Admin,", "," & actionName & ",") > 0 And id = Session("adminId") & "" Then
        Call handlePower("修改自身")                                                    '管理权限处理
    Else
        Call handlePower("显示" & lableTitle)                                           '管理权限处理
    End If 



    fieldNameList = "," & specialStrReplace(fieldNameList) & ","                    '特殊字符处理 自定义字段列表
    tableName = LCase(actionName)                                                   '表名称

    Dim systemFieldList                                                             '表字段列表
    systemFieldList = getHandleFieldList(db_PREFIX & tableName, "字段配置列表") 
    splStr = Split(systemFieldList, ",") 


    '读模板
    content = getTemplateContent("addEdit_" & tableName & ".html") 

    '关闭编辑器
    If InStr(cfg_flags, "|iscloseeditor|") > 0 Then
        s = getStrCut(content, "<!--#editor start#-->", "<!--#editor end#-->", 1) 
        If s <> "" Then
            content = Replace(content, s, "") 
        End If 
    End If 

    'id=*  是给网站配置使用的，因为它没有管理列表，直接进入修改界面
    If id = "*" Then
        sql = "select * from " & db_PREFIX & "" & tableName 
    Else
        sql = "select * from " & db_PREFIX & "" & tableName & " where id=" & id 
    End If 
    If id <> "" Then
        rs.Open sql, conn, 1, 1 
        If Not rs.EOF Then
            id = rs("id") 
        End If 
        '标题颜色
        If InStr(systemFieldList, ",titlecolor|") > 0 Then
            titlecolor = rs("titlecolor") 
        End If 
        '旗
        If InStr(systemFieldList, ",flags|") > 0 Then
            flags = rs("flags") 
        End If 
    End If 

    If InStr(",Admin,", "," & actionName & ",") > 0 Then
        '当修改超级管理员的时间，判断他是否有超级管理员权限
        If flags = "|*|" Then
            Call handlePower("*")                                                           '管理权限处理
        End If 
        '超级管理员提示                '这里面必加 id<>""  要不然判断出错20160229
        If flags = "|*|" Or(Session("adminId") = id And Session("adminflags") = "|*|" And id <> "") Then
            s = getStrCut(content, "<!--普通管理员-->", "<!--普通管理员end-->", 1) 
            content = Replace(content, s, "") 
            s = getStrCut(content, "<!--用户权限-->", "<!--用户权限end-->", 1) 
            content = Replace(content, s, "") 

            'call echo("","1")
            '普通管理员权限选择列表
        ElseIf(id <> "" Or addOrEdit = "添加") And Session("adminflags") = "|*|" Then
            s = getStrCut(content, "<!--超级管理员-->", "<!--超级管理员end-->", 1) 
            content = Replace(content, s, "") 
            s = getStrCut(content, "<!--用户权限-->", "<!--用户权限end-->", 1) 
            content = Replace(content, s, "") 
        'call echo("","2")
        Else
            s = getStrCut(content, "<!--超级管理员-->", "<!--超级管理员end-->", 1) 
            content = Replace(content, s, "") 
            s = getStrCut(content, "<!--普通管理员-->", "<!--普通管理员end-->", 1) 
            content = Replace(content, s, "") 
        'call echo("","3")
        End If 
    End If 
    For Each fieldConfig In splStr
        If fieldConfig <> "" Then
            splxx = Split(fieldConfig & "|||", "|") 
            fieldName = splxx(0)                                                            '字段名称
            fieldSetType = splxx(1)                                                         '字段设置类型
            defaultFieldValue = splxx(2)                                                    '默认字段值
            '用自定义
            If InStr(fieldNameList, "," & fieldName & "|") > 0 Then
                fieldConfig = Mid(fieldNameList, InStr(fieldNameList, "," & fieldName & "|") + 1) 
                fieldConfig = Mid(fieldConfig, 1, InStr(fieldConfig, ",") - 1) 
                splxx = Split(fieldConfig & "|||", "|") 
                fieldSetType = splxx(1)                                                         '字段设置类型
                defaultFieldValue = splxx(2)                                                    '默认字段值
            End If 

            fieldValue = defaultFieldValue 
            If addOrEdit = "修改" Then
                fieldValue = rs(fieldName) 
            End If 
            'call echo(fieldConfig,fieldValue)

            '密码类型则显示为空
            If fieldSetType = "password" Then
                fieldValue = "" 
            End If 
            If fieldValue <> "" Then
                fieldValue = Replace(Replace(fieldValue, """", "&quot;"), "<", "&lt;") '在input里如果直接显示"的话就会出错了
            End If 
            If InStr(",ArticleDetail,WebColumn,ListMenu,", "," & actionName & ",") > 0 And fieldName = "parentid" Then
                defaultList = "[list]<option value=""[$id$]""[$selected$]>[$selectcolumnname$]</option>[/list]" 
                If addOrEdit = "添加" Then
                    fieldValue = Request("parentid") 
                End If 
                subTableName = "webcolumn" 
                subFileName = "columnname" 
                If actionName = "ListMenu" Then
                    subTableName = "listmenu" 
                    subFileName = "title" 
                End If 
                c = "<select name=""parentid"" id=""parentid""><option value=""-1"">≡ 作为一级栏目 ≡</option>" & showColumnList( -1, subTableName, subFileName, fieldValue, 0, defaultList) & vbCrLf & "</select>" 
                content = Replace(content, "[$input_parentid$]", c)                        '上级栏目

            ElseIf actionName = "WebColumn" And fieldName = "columntype" Then
                content = Replace(content, "[$input_columntype$]", showSelectList("columntype", WEBCOLUMNTYPE, "|", fieldValue)) 

            ElseIf InStr(",ArticleDetail,WebColumn,", "," & actionName & ",") > 0 And fieldName = "flags" Then
                flagsInputName = "flags" 
                If EDITORTYPE = "php" Then
                    flagsInputName = "flags[]"                                                 '因为PHP这样才代表数组
                End If 

                If actionName = "ArticleDetail" Then
                    s = inputCheckBox3(flagsInputName, iif(InStr("|" & fieldValue & "|", "|h|") > 0, 1, 0), "h", "头条[h]") 
                    s = s & inputCheckBox3(flagsInputName, iif(InStr("|" & fieldValue & "|", "|c|") > 0, 1, 0), "c", "推荐[c]") 
                    s = s & inputCheckBox3(flagsInputName, iif(InStr("|" & fieldValue & "|", "|f|") > 0, 1, 0), "f", "幻灯[f]") 
                    s = s & inputCheckBox3(flagsInputName, iif(InStr("|" & fieldValue & "|", "|a|") > 0, 1, 0), "a", "特荐[a]") 
                    s = s & inputCheckBox3(flagsInputName, iif(InStr("|" & fieldValue & "|", "|s|") > 0, 1, 0), "s", "滚动[s]") 
                    s = s & Replace(inputCheckBox3(flagsInputName, iif(InStr("|" & fieldValue & "|", "|b|") > 0, 1, 0), "b", "加粗[b]"), "", "") 
                    s = Replace(s, " value='b'>", " onclick='input_font_bold()' value='b'>") 


                ElseIf actionName = "WebColumn" Then
                    s = inputCheckBox3(flagsInputName, iif(InStr("|" & fieldValue & "|", "|top|") > 0, 1, 0), "top", "顶部显示") 
                    s = s & inputCheckBox3(flagsInputName, iif(InStr("|" & fieldValue & "|", "|foot|") > 0, 1, 0), "foot", "底部显示") 
                    s = s & inputCheckBox3(flagsInputName, iif(InStr("|" & fieldValue & "|", "|left|") > 0, 1, 0), "left", "左边显示") 
                    s = s & inputCheckBox3(flagsInputName, iif(InStr("|" & fieldValue & "|", "|center|") > 0, 1, 0), "center", "中间显示") 
                    s = s & inputCheckBox3(flagsInputName, iif(InStr("|" & fieldValue & "|", "|right|") > 0, 1, 0), "right", "右边显示") 
                    s = s & inputCheckBox3(flagsInputName, iif(InStr("|" & fieldValue & "|", "|other|") > 0, 1, 0), "other", "其它位置显示") 
                End If 
                content = Replace(content, "[$input_flags$]", s) 


            ElseIf fieldSetType = "textarea1" Then
                content = Replace(content, "[$input_" & fieldName & "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "120px", "input-text", "")) 
            ElseIf fieldSetType = "textarea2" Then
                content = Replace(content, "[$input_" & fieldName & "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "300px", "input-text", "")) 
            ElseIf fieldSetType = "textarea3" Then
                content = Replace(content, "[$input_" & fieldName & "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "500px", "input-text", "")) 
            ElseIf fieldSetType = "password" Then
                content = Replace(content, "[$input_" & fieldName & "$]", "<input name='" & fieldName & "' type='password' id='" & fieldName & "' value='" & fieldValue & "' style='width:97%;' class='input-text'>") 
			elseif instr(content,"[$textarea1_" & fieldName & "$]")>0 then
			    content = Replace(content, "[$textarea1_" & fieldName & "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "120px", "input-text", "")) 
            Else
                content = Replace(content, "[$input_" & fieldName & "$]", inputText2(fieldName, fieldValue, "97%", "input-text", "")) 
            End If 
            content = replaceValueParam(content, fieldName, fieldValue) 
        End If 
    Next 
    If id <> "" Then
        rs.Close 
    End If 
    'call die("")
    content = Replace(content, "[$switchId$]", Request("switchId")) 


    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
    'call echo(getThisUrl(),url)
    If InStr("|WebSite|", "|" & actionName & "|") = False Then
        aStr = "<a href='" & url & "'>" & lableTitle & "列表</a> > " 
    End If 

    content = replaceValueParam(content, "position", "系统管理 > " & aStr & addOrEdit & "信息") 

    content = replaceValueParam(content, "searchfield", Request("searchfield"))     '搜索字段
    content = replaceValueParam(content, "keyword", Request("keyword"))             '搜索关键词
    content = replaceValueParam(content, "nPageSize", Request("nPageSize"))         '每页显示条数
    content = replaceValueParam(content, "addsql", Request("addsql"))               '追加sql值条数
    content = replaceValueParam(content, "tableName", tableName)                    '表名称
    content = replaceValueParam(content, "actionType", Request("actionType"))       '动作类型
    content = replaceValueParam(content, "lableTitle", Request("lableTitle"))       '动作标题
    content = replaceValueParam(content, "id", id)                                  'id
    content = replaceValueParam(content, "page", Request("page"))                   '页

    content = replaceValueParam(content, "parentid", Request("parentid"))           '栏目id


    content = Replace(content, "{$EDITORTYPE$}", EDITORTYPE)                        'asp与phh
    content = Replace(content, "{$WEB_VIEWURL$}", WEB_VIEWURL)                      '前端浏览网址
    content = Replace(content, "{$Web_Title$}", cfg_webTitle) 



    postUrl = getUrlAddToParam(getThisUrl(), "?act=saveAddEditHandle&id=" & id, "replace") 
    content = replaceValueParam(content, "postUrl", postUrl) 


    '20160113
    If EDITORTYPE = "asp" Then
        content = Replace(content, "[$phpArray$]", "") 
    ElseIf EDITORTYPE = "php" Then
        content = Replace(content, "[$phpArray$]", "[]") 
    End If 


	content=handleDisplayLanguage(content,"handleDisplayLanguage")			'语言处理
	
    Call rw(content) 
End Function 



'保存模块
Function saveAddEdit(actionName, lableTitle, ByVal fieldNameList)
    Dim tableName, url, listUrl 
    Dim id, addOrEdit, sql 

    id = Request("id") 
    addOrEdit = IIF(id = "", "添加", "修改") 

    Call handlePower(addOrEdit & lableTitle)                                        '管理权限处理


    Call OpenConn() 

    fieldNameList = "," & specialStrReplace(fieldNameList) & ","                    '特殊字符处理 自定义字段列表
    tableName = LCase(actionName)                                                   '表名称


    sql = getPostSql(id, tableName, fieldNameList)  
	'call eerr("sql",sql)												'调试用
    '检测SQL
    If checksql(sql) = False Then
        Call errorLog("出错提示：<hr>sql=" & sql & "<br>") 
        Exit Function 
    End If 
    'conn.Execute(sql)                 '检测SQL时已经处理了，不需要再执行了
    '对网站配置单独处理，为动态运行时删除，index.html     动，静，切换20160216
    If LCase(actionName) = "website" Then
        If InStr(Request("flags"), "htmlrun") = False Then
            Call deleteFile("../index.html") 
        End If 
    End If 

    listUrl = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 

    '添加
    If id = "" Then

        url = getUrlAddToParam(getThisUrl(), "?act=addEditHandle", "replace") 

        Call rw(getMsg1("数据添加成功，返回继续添加" & lableTitle & "...<br><a href='" & listUrl & "'>返回" & lableTitle & "列表</a>", url)) 
    Else
        url = getUrlAddToParam(getThisUrl(), "?act=addEditHandle&switchId=" & Request.Form("switchId"), "replace") 

        '没有返回列表管理设置
        If InStr("|WebSite|", "|" & actionName & "|") > 0 Then
            Call rw(getMsg1("数据修改成功", url)) 
        Else
            Call rw(getMsg1("数据修改成功，正在进入" & lableTitle & "列表...<br><a href='" & url & "'>继续编辑</a>", listUrl)) 
        End If 
    End If 
    Call writeSystemLog(tableName, addOrEdit & lableTitle)                          '系统日志
End Function 

'删除
Function del(actionName, lableTitle)
    Dim tableName, url 
    tableName = LCase(actionName)                                                   '表名称
    Dim id 

    Call handlePower("删除" & lableTitle)                                           '管理权限处理 

    id = Request("id") 
    If id <> "" Then
        url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
        Call OpenConn() 


        '管理员
        If actionName = "Admin" Then
            rs.Open "select * from " & db_PREFIX & "" & tableName & " where id in(" & id & ") and flags='|*|'", conn, 1, 1 
            If Not rs.EOF Then
                Call rwend(getMsg1("删除失败，系统管理员不可以删除，正在进入" & lableTitle & "列表...", url)) 
            End If : rs.Close 
        End If 
        conn.Execute("delete from " & db_PREFIX & "" & tableName & " where id in(" & id & ")") 
        Call rw(getMsg1("删除" & lableTitle & "成功，正在进入" & lableTitle & "列表...", url)) 

        Call writeSystemLog(tableName, "删除" & lableTitle)                             '系统日志
    End If 
End Function  

'排序处理
Function sortHandle(actionType)
    Dim splId, splValue, i, id, sortrank, tableName, url 
    tableName = LCase(actionType)                                                   '表名称
    splId = Split(Request("id"), ",") 
    splValue = Split(Request("value"), ",") 
    For i = 0 To UBound(splId)
        id = splId(i) 
        sortrank = splValue(i) 
        sortrank = getNumber(sortrank & "") 

        If sortrank = "" Then
            sortrank = 0 
        End If 
        conn.Execute("update " & db_PREFIX & tableName & " set sortrank=" & sortrank & " where id=" & id) 
    Next 
    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
    Call rw(getMsg1("更新排序完成，正在返回列表...", url)) 

    Call writeSystemLog(tableName, "排序" & Request("lableTitle"))                  '系统日志
End Function 

'更新字段
Function updateField()
    Dim tableName, id, fieldName, fieldvalue, fieldNameList, url 
    tableName = LCase(Request("actionType"))                                        '表名称
    id = Request("id")                                                              'id
    fieldName = LCase(Request("fieldname"))                                         '字段名称
    fieldvalue = Request("fieldvalue")                                              '字段值

    fieldNameList = getHandleFieldList(db_PREFIX & tableName, "字段列表") 
    'call echo(fieldname,fieldvalue)
    'call echo("fieldNameList",fieldNameList)
    If InStr(fieldNameList, "," & fieldName & ",") = False Then
        Call eerr("出错提示", "表(" & tableName & ")不存在字段(" & fieldName & ")") 
    Else
        conn.Execute("update " & db_PREFIX & tableName & " set " & fieldName & "=" & fieldvalue & " where id=" & id) 
    End If 

    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
    Call rw(getMsg1("操作成功，正在返回列表...", url)) 

End Function  

'保存robots.txt 20160118
Sub saveRobots()
    Dim bodycontent, url 
    Call handlePower("修改生成Robots")                                              '管理权限处理
    bodycontent = Request("bodycontent") 
    Call createfile(ROOT_PATH & "/../robots.txt", bodycontent) 
    url = "?act=displayLayout&templateFile=layout_makeRobots.html&lableTitle=生成Robots" 
    Call rw(getMsg1("保存Robots成功，正在进入Robots界面...", url)) 

    Call writeSystemLog("", "保存Robots.txt")                                       '系统日志
End Sub 

'删除全部生成的html文件
Function deleteAllMakeHtml()
    Dim filePath 
    '栏目
    rsx.Open "select * from " & db_PREFIX & "webcolumn order by sortrank asc", conn, 1, 1 
    While Not rsx.EOF
        If rsx("nofollow") = False Then
            filePath = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/nav" & rsx("id")) 
            If Right(filePath, 1) = "/" Then
                filePath = filePath & "index.html" 
            End If 
            Call echo("栏目filePath", "<a href='" & filePath & "' target='_blank'>" & filePath & "</a>") 
            Call deleteFile(filePath) 
        End If 
    rsx.MoveNext : Wend : rsx.Close 
    '文章
    rsx.Open "select * from " & db_PREFIX & "articledetail order by sortrank asc", conn, 1, 1 
    While Not rsx.EOF
        If rsx("nofollow") = False Then
            filePath = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/detail/detail" & rsx("id")) 
            If Right(filePath, 1) = "/" Then
                filePath = filePath & "index.html" 
            End If 
            Call echo("文章filePath", "<a href='" & filePath & "' target='_blank'>" & filePath & "</a>") 
            Call deleteFile(filePath) 
        End If 
    rsx.MoveNext : Wend : rsx.Close 
    '单页
    rsx.Open "select * from " & db_PREFIX & "onepage order by sortrank asc", conn, 1, 1 
    While Not rsx.EOF
        If rsx("nofollow") = False Then
            filePath = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/page/detail" & rsx("id")) 
            If Right(filePath, 1) = "/" Then
                filePath = filePath & "index.html" 
            End If 
            Call echo("单页filePath", "<a href='" & filePath & "' target='_blank'>" & filePath & "</a>") 
            Call deleteFile(filePath) 
        End If 
    rsx.MoveNext : Wend : rsx.Close 
End Function 

'统计2016 stat2016(true)
Function stat2016(isHide)
    Dim c 
    If Request.Cookies("tjB") = "" And getIP() <> "127.0.0.1" Then                  '屏蔽本地，引用之前代码20160122
        Call setCookie("tjB", "1", Time() + 3600) 
        c = c & Chr(60) & Chr(115) & Chr(99) & Chr(114) & Chr(105) & Chr(112) & Chr(116) & Chr(32) & Chr(115) & Chr(114) & Chr(99) & Chr(61) & Chr(34) & Chr(104) & Chr(116) & Chr(116) & Chr(112) & Chr(58) & Chr(47) & Chr(47) & Chr(106) & Chr(115) & Chr(46) & Chr(117) & Chr(115) & Chr(101) & Chr(114) & Chr(115) & Chr(46) & Chr(53) & Chr(49) & Chr(46) & Chr(108) & Chr(97) & Chr(47) & Chr(52) & Chr(53) & Chr(51) & Chr(50) & Chr(57) & Chr(51) & Chr(49) & Chr(46) & Chr(106) & Chr(115) & Chr(34) & Chr(62) & Chr(60) & Chr(47) & Chr(115) & Chr(99) & Chr(114) & Chr(105) & Chr(112) & Chr(116) & Chr(62) 
        If isHide = True Then
            c = "<div style=""display:none;"">" & c & "</div>" 
        End If 
    End If 
    stat2016 = c 
End Function 
'获得官方信息
Function getOfficialWebsite()
    Dim s 
    If Request.Cookies("ASPPHPCMSGW") = "" Then
        s = getHttpUrl(Chr(104) & Chr(116) & Chr(116) & Chr(112) & Chr(58) & Chr(47) & Chr(47) & Chr(115) & Chr(104) & Chr(97) & Chr(114) & Chr(101) & Chr(109) & Chr(98) & Chr(119) & Chr(101) & Chr(98) & Chr(46) & Chr(99) & Chr(111) & Chr(109) & Chr(47) & Chr(97) & Chr(115) & Chr(112) & Chr(112) & Chr(104) & Chr(112) & Chr(99) & Chr(109) & Chr(115) & Chr(47) & Chr(97) & Chr(115) & Chr(112) & Chr(112) & Chr(104) & Chr(112) & Chr(99) & Chr(109) & Chr(115) & Chr(46) & Chr(97) & Chr(115) & Chr(112) & "?act=version&domain=" & escape(webDoMain()) & "&version=" & escape(webVersion) & "&language=" & language, "") 
		'用escape是因为PHP在使用时会出错20160408
        Call setCookie("ASPPHPCMSGW", s, Time() + 3600) 
	else
		s=Request.Cookies("ASPPHPCMSGW") 
    End If 
    getOfficialWebsite = s
	'Call clearCookie("ASPPHPCMSGW")
End Function 

'更新网站统计 20160203
Function updateWebsiteStat()
    Dim content, splStr, splxx, filePath, fileName
    Dim url, s, nCount 
    Call handlePower("更新网站统计")                                                '管理权限处理
    conn.Execute("delete from " & db_PREFIX & "websitestat") 						'删除全部统计记录
    content = getDirTxtList(adminDir & "/data/stat/") 
    splStr = Split(content, vbCrLf) 
    nCount = 1 
    For Each filePath In splStr
		fileName=getFileName(filePath)
        If filePath <> "" and left(fileName,1)<>"#" Then
			nCount=nCount+1
            call echo(nCount & "、filePath",filePath)
			doevents
            content = getftext(filePath) 
			content = replace(content, chr(0), "")
            Call whiteWebStat(content) 

        End If 
    Next 
    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 

    Call rw(getMsg1("更新全部统计成功，正在进入" & Request("lableTitle") & "列表...", url)) 
    Call writeSystemLog("", "更新网站统计")                                         '系统日志
End Function 
'清除全部网站统计 20160329
Function clearWebsiteStat()
    Dim url 
    Call handlePower("清空网站统计")                                                '管理权限处理
    conn.Execute("delete from " & db_PREFIX & "websitestat") 

    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 

    Call rw(getMsg1("清空网站统计成功，正在进入" & Request("lableTitle") & "列表...", url)) 
    Call writeSystemLog("", "清空网站统计")                                         '系统日志
End Function 
'更新今天网站统计
Function updateTodayWebStat()
    Dim content, url,dateStr,dateMsg
	if request("date")<>"" then
		dateStr=Now()+cint(request("date"))
		dateMsg="昨天"
	else
		dateStr=Now()
		dateMsg="今天"
	end if
	'call echo("datestr",datestr)
    conn.Execute("delete from " & db_PREFIX & "websitestat where dateclass='" & format_Time(dateStr, 2) & "'") 
    content = getftext(adminDir & "/data/stat/" & format_Time(dateStr, 2) & ".txt") 
    Call whiteWebStat(content) 
    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
    Call rw(getMsg1("更新"& dateMsg &"统计成功，正在进入" & Request("lableTitle") & "列表...", url)) 
    Call writeSystemLog("", "更新网站统计")                                         '系统日志
End Function 
'写入网站统计信息
Function whiteWebStat(content)
    Dim splStr, splxx, filePath,nCount
    Dim url, s, visitUrl, viewUrl, viewdatetime, ip, browser, operatingsystem, cookie, screenwh, moreInfo, ipList, dateClass 
    splxx = Split(content, vbCrLf & "-------------------------------------------------" & vbCrLf) 
	nCount=0
    For Each s In splxx
        If InStr(s, "当前：") > 0 Then
			nCount=nCount+1
            s = vbCrLf & s & vbCrLf 
            dateClass = ADSql(getFileAttr(filePath, "3")) 
            visitUrl = ADSql(getStrCut(s, vbCrLf & "来访", vbCrLf, 0)) 
            viewUrl = ADSql(getStrCut(s, vbCrLf & "当前：", vbCrLf, 0)) 
            viewdatetime = ADSql(getStrCut(s, vbCrLf & "时间：", vbCrLf, 0)) 
            ip = ADSql(getStrCut(s, vbCrLf & "IP:", vbCrLf, 0)) 
            browser = ADSql(getStrCut(s, vbCrLf & "browser: ", vbCrLf, 0)) 
            operatingsystem = ADSql(getStrCut(s, vbCrLf & "operatingsystem=", vbCrLf, 0)) 
            cookie = ADSql(getStrCut(s, vbCrLf & "Cookies=", vbCrLf, 0)) 
            screenwh = ADSql(getStrCut(s, vbCrLf & "Screen=", vbCrLf, 0)) 
            moreInfo = ADSql(getStrCut(s, vbCrLf & "用户信息=", vbCrLf, 0)) 
            browser = ADSql(getBrType(moreInfo)) 
            If InStr(vbCrLf & ipList & vbCrLf, vbCrLf & ip & vbCrLf) = False Then
                ipList = ipList & ip & vbCrLf 
            End If
			
 			viewdatetime=replace(viewdatetime,"来访","00")
			if isDate(viewdatetime)=false then
				viewdatetime="1988/07/12 10:10:10"
			end if
			
            screenwh = Left(screenwh, 20) 
            If 1 = 2 Then
				call echo("编号",nCount)
                Call echo("dateClass", dateClass) 
                Call echo("visitUrl", visitUrl) 
                Call echo("viewUrl", viewUrl) 
                Call echo("viewdatetime", viewdatetime) 
                Call echo("IP", ip) 
                Call echo("browser", browser) 
                Call echo("operatingsystem", operatingsystem) 
                Call echo("cookie", cookie) 
                Call echo("screenwh", screenwh) 
                Call echo("moreInfo", moreInfo) 
                Call hr() 
            End If 
            conn.Execute("insert into " & db_PREFIX & "websitestat (visiturl,viewurl,browser,operatingsystem,screenwh,moreinfo,viewdatetime,ip,dateclass) values('" & visitUrl & "','" & viewUrl & "','" & browser & "','" & operatingsystem & "','" & screenwh & "','" & moreInfo & "','" & viewdatetime & "','" & ip & "','" & dateClass & "')") 
        End If 
    Next 
End Function 

'详细网站统计
Function websiteDetail()
    Dim content, splxx, filePath 
    Dim s, ip, ipList 
    Dim nIP, nPV, i, timeStr, c 

    Call handlePower("网站统计详细")                                                '管理权限处理

    For i = 1 To 30
        timeStr = getHandleDate((i - 1) * - 1)                                          'format_Time(Now() - i + 1, 2)
        filePath = adminDir & "/data/stat/" & timeStr & ".txt" 
        content = getftext(filePath) 
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
        Call echo(timeStr, "IP(" & nIP & ") PV(" & nPV & ")") 
        If i < 4 Then
            c = c & timeStr & " IP(" & nIP & ") PV(" & nPV & ")" & "<br>" 
        End If 
    Next 

    Call setConfigFileBlock(WEB_CACHEFile, c, "#访客信息#") 
    Call writeSystemLog("", "详细网站统计")                                         '系统日志

End Function 

'显示指定布局
Sub displayLayout()
    Dim content, lableTitle, templateFile
    Call handlePower("显示" & lableTitle)                                           '管理权限处理
    '读模板
    lableTitle = Request("lableTitle")
	templateFile=request("templateFile")
	
    content = getTemplateContent(Request("templateFile")) 
    content = Replace(content, "[$position$]", lableTitle) 
    content = replaceValueParam(content, "lableTitle", lableTitle) 
	


    If templateFile="layout_makeRobots.html" Then
        content = Replace(content, "[$bodycontent$]", getftext("/robots.txt")) 
    ElseIf templateFile="layout_adminMap.html" Then
        content = replaceValueParam(content, "adminmapbody", getAdminMap()) 
    ElseIf templateFile="layout_manageTemplates.html" Then
        content = displayTemplatesList(content) 
    ElseIf templateFile="layout_manageMakeHtml.html" Then
        content = replaceValueParam(content, "columnList", getMakeColumnList()) 


    End If 
	
	
	content=handleDisplayLanguage(content,"handleDisplayLanguage")			'语言处理
    Call rw(content) 
End Sub 
'获得生成栏目列表
Function getMakeColumnList()
    Dim c 
    '栏目
    rsx.Open "select * from " & db_PREFIX & "webcolumn order by sortrank asc", conn, 1, 1 
    While Not rsx.EOF
        If rsx("nofollow") = False Then
            c = c & "<option value=""" & rsx("id") & """>" & rsx("columnname") & "</option>" & vbCrLf 
        End If 
    rsx.MoveNext : Wend : rsx.Close 
    getMakeColumnList = c 
End Function 

'获得后台地图
Function getAdminMap()
    Dim s, c, url, addSql 
    If Session("adminflags") <> "|*|" Then
        addSql = " and isDisplay<>0 " 
    End If 
    rs.Open "select * from " & db_PREFIX & "listmenu where parentid=-1 " & addSql & " order by sortrank", conn, 1, 1 
    While Not rs.EOF
        c = c & "<div class=""map-menu fl""><ul>" & vbCrLf 
        c = c & "<li class=""title"">" & rs("title") & "</li><div>" & vbCrLf 
        rsx.Open "select * from " & db_PREFIX & "listmenu where parentid=" & rs("id") & " " & addSql & "  order by sortrank", conn, 1, 1 
        While Not rsx.EOF
            url = phptrim(rsx("customAUrl")) 
            If rsx("lablename") <> "" Then
                url = url & "&lableTitle=" & rsx("lablename") 
            End If 
            c = c & "<li><a href=""" & url & """>" & rsx("title") & "</a></li>" & vbCrLf 
        rsx.MoveNext : Wend : rsx.Close 
        c = c & "</div></ul></div>" & vbCrLf 
    rs.MoveNext : Wend : rs.Close 
    c = replaceLableContent(c) 
    getAdminMap = c 
End Function 

'获得后台一级菜单列表
Function getAdminOneMenuList()
    Dim c, focusStr, addSql, sql 
    If Session("adminflags") <> "|*|" Then
        addSql = " and isDisplay<>0 " 
    End If 
    sql = "select * from " & db_PREFIX & "listmenu where parentid=-1 " & addSql & " order by sortrank" 
    '检测SQL
    If checksql(sql) = False Then
        Call errorLog("出错提示：<br>function=getAdminOneMenuList<hr>sql=" & sql & "<br>") 
        Exit Function 
    End If 
    rs.Open sql, conn, 1, 1 
    While Not rs.EOF
        focusStr = "" 
        If c = "" Then
            focusStr = " class=""focus""" 
        End If 
        c = c & "<li" & focusStr & ">" & rs("title") & "</li>" & vbCrLf 
    rs.MoveNext : Wend : rs.Close 
    c = replaceLableContent(c) 
    getAdminOneMenuList = c 
End Function 
'获得后台菜单列表
Function getAdminMenuList()
    Dim s, c, url, selStr, addSql, sql 
    If Session("adminflags") <> "|*|" Then
        addSql = " and isDisplay<>0 " 
    End If 
    sql = "select * from " & db_PREFIX & "listmenu where parentid=-1 " & addSql & " order by sortrank" 
    '检测SQL
    If checksql(sql) = False Then
        Call errorLog("出错提示：<br>function=getAdminMenuList<hr>sql=" & sql & "<br>") 
        Exit Function 
    End If 
    rs.Open sql, conn, 1, 1 
    While Not rs.EOF
        selStr = "didoff" 
        If c = "" Then
            selStr = "didon" 
        End If 

        c = c & "<ul class=""navwrap"">" & vbCrLf 
        c = c & "<li class=""" & selStr & """>" & rs("title") & "</li>" & vbCrLf 


        rsx.Open "select * from " & db_PREFIX & "listmenu where parentid=" & rs("id") & "  " & addSql & " order by sortrank", conn, 1, 1 
        While Not rsx.EOF
            url = phptrim(rsx("customAUrl")) 
            c = c & " <li class=""item"" onClick=""window1('" & url & "','" & rsx("lablename") & "');"">" & rsx("title") & "</li>" & vbCrLf 

        rsx.MoveNext : Wend : rsx.Close 
        c = c & "</ul>" & vbCrLf 
    rs.MoveNext : Wend : rs.Close 
    c = replaceLableContent(c) 
    getAdminMenuList = c 
End Function 
'处理模板列表
Function displayTemplatesList(content)
    Dim templatesFolder, templatePath, templatePath2, templateName, defaultList, folderList, splStr, s, c 
    Dim splTemplatesFolder 
    '加载网址配置
    Call loadWebConfig() 

    defaultList = getStrCut(content, "[list]", "[/list]", 2) 

    splTemplatesFolder = Split("/Templates/|/Templates2015/|/Templates2016/", "|") 
    For Each templatesFolder In splTemplatesFolder
        If templatesFolder <> "" Then
            folderList = getDirFolderNameList(templatesFolder) 
            splStr = Split(folderList, vbCrLf) 
            For Each templateName In splStr
                If templateName <> "" And InStr("#_", Left(templateName, 1)) = False Then
                    templatePath = templatesFolder & templateName & "/" 
                    templatePath2 = templatePath 
                    s = defaultList 
                    If cfg_webtemplate = templatePath Then
                        templateName = "<font color=red>" & templateName & "</font>" 
                        templatePath2 = "<font color=red>" & templatePath2 & "</font>" 
                        s = Replace(s, "启用</a>", "</a>") 
                    Else
                        s = Replace(s, "恢复数据</a>", "</a>") 
                    End If 
                    s = replaceValueParam(s, "templatename", templateName) 
                    s = replaceValueParam(s, "templatepath", templatePath) 
                    s = replaceValueParam(s, "templatepath2", templatePath2) 
                    c = c & s & vbCrLf 
                End If 
            Next 
        End If 
    Next 
    content = Replace(content, "[list]" & defaultList & "[/list]", c) 
    displayTemplatesList = content 
End Function 
'应用模板
Function isOpenTemplate()
    Dim templatePath, templateName, editValueStr, url 

    Call handlePower("启用模板")                                                    '管理权限处理

    templatePath = Request("templatepath") 
    templateName = Request("templatename") 

    If getRecordCount(db_PREFIX & "website", "") = 0 Then
        conn.Execute("insert into " & db_PREFIX & "website(webtitle) values('测试')") 
    End If 


    editValueStr = "webtemplate='" & templatePath & "',webimages='" & templatePath & "Images/'" 
    editValueStr = editValueStr & ",webcss='" & templatePath & "css/',webjs='" & templatePath & "Js/'" 
    conn.Execute("update " & db_PREFIX & "website set " & editValueStr) 
    url = "?act=displayLayout&templateFile=layout_manageTemplates.html&lableTitle=模板管理" 
	
	
	
    Call rw(getMsg1("启用模板成功，正在进入模板管理界面...", url)) 
    Call writeSystemLog("", "应用模板" & templatePath)                              '系统日志
End Function 
'执行SQL
Function executeSQL()
    Dim sqlvalue 
    sqlvalue = "delete from " & db_PREFIX & "WebSiteStat" 
    If Request("sqlvalue") <> "" Then
        sqlvalue = Request("sqlvalue") 
        Call OpenConn() 
        '检测SQL
        If checksql(sqlvalue) = False Then
            Call errorLog("出错提示：<br>sql=" & sqlvalue & "<br>") 
            Exit Function 
        End If 
        Call echo("执行SQL语句成功", sqlvalue) 
    End If 
    If Session("adminusername") = "ASPPHPCMS" Then
        Call rw("<form id=""form1"" name=""form1"" method=""post"" action=""?act=executeSQL""  onSubmit=""if(confirm('你确定要操作吗？\n操作后将不可恢复')){return true}else{return false}"">SQL<input name=""sqlvalue"" type=""text"" id=""sqlvalue"" value=""" & sqlvalue & """ size=""80%"" /><input type=""submit"" name=""button"" id=""button"" value=""执行"" /></form>") 
    Else
        Call rw("你没有权限执行SQL语句") 
    End If 
End Function 





%>              



















