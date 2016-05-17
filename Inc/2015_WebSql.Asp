<% 
'快速操作网站Sql 为移植(20151029)

'产品类别追加SQL处理
'SqlGetDidSidTidSQl(Did,Sid,Tid)
'Sql =  GetWhereAnd(Sql,AddSql)
'写二种操作数据库方法，一种是操作全部数据库，另一个是带句柄操作数据库  20150723

'================ 辅助 ==========================
'类型大类小类子为自动链接
Function getDidSidTidSQl(did, sid, tid)
    Dim sql, PrefixSql                                                              '前缀SQL
    PrefixSql = " Where " 
    If did <> "" Then
        sql = sql & PrefixSql & "BigClassName='" & did & "'" 
        PrefixSql = " And " 
    End If 
    If sid <> "" Then
        sql = sql & PrefixSql & "SmallClassName='" & sid & "'" 
        PrefixSql = " And " 
    End If 
    If tid <> "" Then
        sql = sql & PrefixSql & "ThreeClassName='" & tid & "'" 
    End If 
    getDidSidTidSQl = sql 
End Function 
'检测大类存在
Function checkBigClass(did)
    checkBigClass = conn.Execute("Select Count(*) From [BigClass] Where BigClassName='" & did & "'")(0) 
End Function 
'检测小类存在
Function checkSmallClass(did, sid)
    Dim sql, addSql 
    sql = "Select Count(*) From [SmallClass]" 
    addSql = getDidSidTidSQl(did, sid, "") 
    sql = GetWhereAnd(sql, addSql) 
    checkSmallClass = conn.Execute(sql)(0) 
End Function 
'检测子类存在
Function checkThreeClass(did, sid, tid)
    Dim sql, addSql 
    sql = "Select Count(*) From [ThreeClass]" 
    addSql = getDidSidTidSQl(did, sid, tid) 
    sql = GetWhereAnd(sql, addSql) 
    checkThreeClass = conn.Execute(sql)(0) 
End Function 
'检测产品信息存在
Function checkProduct(did, sid, tid)
    Dim sql 
    sql = "Select Count(*) From [Product]" 
    If did <> "" Then sql = GetWhereAnd(sql, "BigClassName='" & did & "'") 
    If sid <> "" Then sql = GetWhereAnd(sql, "SmallClassName='" & sid & "'") 
    If tid <> "" Then sql = GetWhereAnd(sql, "ThreeClassName='" & tid & "'") 
    checkProduct = conn.Execute(sql)(0) 
End Function 
'自动添加文本信息（先判断）
Function autoAddMainInfo(title, ShowTitle, content)
    Call HandleAutoAddMainInfo(title, ShowTitle, False, content) 
End Function 
'处理自动添加文本内容
Function handleAutoAddMainInfo(title, ShowTitle, OnHtml, content)
    Dim sql 
    sql = "Select * From [MainInfo] Where Title='" & title & "'" 
    If UserId <> "" Then sql = GetWhereAnd(sql, " And UserId=" & UserId)            '指定用户
    If CheckSql(sql) = False Then Call eerr("Sql", sql) 
    rs.Open sql, conn, 1, 3 
    If rs.EOF Then
        rs.AddNew 
        rs("Title") = title 
        rs("ShowTitle") = ShowTitle 
        rs("Content") = content 
        rs("OnHtml") = OnHtml 
        rs.Update 
    End If : rs.Close 
End Function 
'自动添加大类（先判断）
Function autoAddBigClass(did)
    Dim folderName, nSort 
    If conn.Execute("Select Count(*) From [BigClass] Where BigClassName='" & did & "'")(0) = 0 Then
        folderName = PinYin(did, 3) 
        nSort = conn.Execute("Select Count(*) From [SmallClass]")(0) + 1 
        conn.Execute("Insert Into[BigClass](BigClassName,FolderName,Sort) Values('" & did & "', '" & folderName & "'," & nSort & ")") 
        autoAddBigClass = "【" & did & "】创建大类成功" 
    Else
        autoAddBigClass = "【" & did & "】大类已存在" 
    End If 
End Function 
'自动添加小类（先判断）
Function autoAddSmallClass(did, sid)
    Dim folderName, nSort 
    If conn.Execute("Select Count(*) From [SmallClass] Where BigClassName='" & did & "' And  SmallClassName='" & sid & "'")(0) = 0 Then
        folderName = PinYin(sid, 3) 
        nSort = conn.Execute("Select Count(*) From [SmallClass] Where BigClassName='" & did & "'")(0) + 1 
        conn.Execute("Insert Into[SmallClass](BigClassName,SmallClassName,FolderName,Sort) Values('" & did & "','" & sid & "', '" & folderName & "'," & nSort & ")") 
        autoAddSmallClass = "【" & did & "】创建小类成功" 
    Else
        autoAddSmallClass = "【" & did & "】小类已存在" 
    End If 
End Function 
'自动添加三级子类（先判断）
Function autoAddThreeClass(did, sid, tid)
    Dim folderName, nSort 
    If conn.Execute("Select Count(*) From [ThreeClass] Where BigClassName='" & did & "' And  SmallClassName='" & sid & "' And ThreeClassName='" & tid & "'")(0) = 0 Then
        folderName = PinYin(sid, 3) 
        nSort = conn.Execute("Select Count(*) From [SmallClass] Where BigClassName='" & did & "' And  SmallClassName='" & sid & "'")(0) + 1 
        conn.Execute("Insert Into[ThreeClass](BigClassName,SmallClassName,ThreeClassName,FolderName,Sort) Values('" & did & "','" & sid & "','" & tid & "', '" & folderName & "'," & nSort & ")") 
        autoAddThreeClass = "【" & sid & "】创建子类成功" 
    Else
        autoAddThreeClass = "【" & sid & "】子类已存在" 
    End If 
End Function 
'自动添加文章（先判断）
Function autoAddArticle(did, sid, tid, title, content)
    autoAddArticle = AutoAddArticleAction("", did, sid, tid, title, "", content) 
End Function 
'自动添加文件（先判断）
Function autoAddArticleAction(action, did, sid, tid, title, SortID, content)
    Dim addSql, sql, NewTitle,Splxx
    Dim BigImage, SmallImage, NewBigImage, NewSmallImage 
    Dim WebKeywords, WebDescription, NewWebKeywords, NewWebDescription, OriginalTitleYes, newAritcleID 

    BigImage = RParam(action, "BigImage")                                           '大图地址
    SmallImage = RParam(action, "SmallImage")                                       '小图地址
    BigImage = HandleHttpUrl(Replace(BigImage, "[$样式名称$]", WebSkins)) 
    SmallImage = HandleHttpUrl(Replace(SmallImage, "[$样式名称$]", WebSkins)) 

    WebKeywords = RParam(action, "Keywords")                                        '网站关键词
    WebDescription = RParam(action, "Description")                                  '网站描述
    WebKeywords = HandleHttpUrl(Replace(WebKeywords, "[$样式名称$]", WebSkins)) 
    WebDescription = HandleHttpUrl(Replace(WebDescription, "[$样式名称$]", WebSkins)) 

    OriginalTitleYes = LCase(RParam(action, "OriginalTitleYes"))                    '是否为原标题
	
	dim memberprice			'会员价
	dim price				'价格
    memberprice =RParam(action, "memberprice")			'会员价
    price =RParam(action, "price")						'价格
	
	

    title = Trim(title) 
    addSql = addSql & getDidSidTidSQl(did, sid, tid) 
    If title <> "" Then addSql = addSql & " And Title='" & title & "'" 
    sql = "Select * From [Product]" & GetWhereAnd(sql, addSql) 
    RsTemp.Open sql, conn, 1, 3 
    If RsTemp.EOF Then
        If OriginalTitleYes = "true" Then
            NewTitle = title 
        Else
            NewTitle = did & "-" & sid & "-" & tid & "-" & title 
        End If 
        RsTemp.AddNew 
        RsTemp("BigClassName") = did 
        RsTemp("SmallClassName") = sid 
        RsTemp("ThreeClassName") = tid 
        RsTemp("WebTitle") = title 
        RsTemp("Title") = NewTitle 


        'FileName = HtmlFileName(NewTitle & "-" & GetRnd(11))            '文件名加随机数

        NewBigImage = Replace(BigImage, "[id]", SortID) 
        NewSmallImage = Replace(SmallImage, "[id]", SortID) 
        If NewBigImage <> "" Then RsTemp("BigFiles") = NewBigImage 
        If NewSmallImage <> "" Then RsTemp("SmallFiles") = NewSmallImage 

        NewWebKeywords = Replace(WebKeywords, "[id]", NewTitle) 
        NewWebDescription = Replace(WebDescription, "[id]", NewTitle) 

        RsTemp("WebKeywords") = NewWebKeywords 
        RsTemp("WebDescription") = NewWebDescription 

        RsTemp("Content") = content
		
		call echo(action,memberprice)
		'价格
		if price<>"" then
			if instr(price,"-")>0 then
				splxx=split(price,"-")
        		RsTemp("price") = PHPRand(cint(splxx(0)),cint(splxx(1)))
			end if 
		end if
		'会员价
		if memberprice<>"" then
			if instr(memberprice,"-")>0 then
				splxx=split(memberprice,"-")
        		RsTemp("memberprice") = PHPRand(cint(splxx(0)),cint(splxx(1)))
			end if 
		end if

        RsTemp.Update 
        newAritcleID = RsTemp(0) 
        RsTemp.Close 

        RsTemp.Open "Select * From [Product] where id=" & newAritcleID, conn, 1, 3 
        If Not RsTemp.EOF Then
            RsTemp("FileName") = "article" & newAritcleID 
            RsTemp.Update 
        End If 

        autoAddArticleAction = True 
    Else
        autoAddArticleAction = False 
    End If : RsTemp.Close 
End Function 
'自动添加文章并自动创建大类
Function autoAddBigClassAndArticle(did, topNumb)
    Dim i 
    Call AutoAddBigClass(did) 
    For i = 1 To topNumb
        Call AutoAddArticle(did, "", "", did & "标题-" & i, did & "内容-" & i) 
    Next 
End Function 







'获得网站模板路径 20150114
Function getWebTemplate()
    Dim rs 
    Set rs = CreateObject("Adodb.RecordSet")
        rs.Open "Select * From [WebSite]", conn, 1, 1 
        If Not rs.EOF Then
            getWebTemplate = rs("WebTemplate") 
        End If : rs.Close 
End Function
'获得网站样式路径 20150114
Function getWebSkins()
    Dim rs 
    Set rs = CreateObject("Adodb.RecordSet")
        rs.Open "Select * From [WebSite]", conn, 1, 1 
        If Not rs.EOF Then
            getWebSkins = rs("WebSkins") 
        End If : rs.Close 
End Function
'获得网站素材 20150713
Function getWebImages()
    Dim rs 
    Set rs = CreateObject("Adodb.RecordSet")
        rs.Open "Select * From [WebSite]", conn, 1, 1 
        If Not rs.EOF Then
            getWebImages = rs("WebImages") 
        End If : rs.Close 
End Function

'获得产品信息细节页是否加Html文件前缀 230141215
Function detailPageAddHtmlPrefix(ByVal fileName)
    fileName = HandleHttpUrl(fileName) 
    If InStr(fileName, "/") = False Then
        fileName = "/Html/" & fileName 
    Else
        fileName = ReplaceN(GetWebFolderName() & fileName, "//", "/", 3) 
        'Call Echo("FileName",FileName)

    End If 
    If Left(fileName, 1) <> "/" Then fileName = "/" & fileName 
    detailPageAddHtmlPrefix = fileName                                              '前面还是加一个/吧
End Function 
'获得处理后文件夹名称 汉字转拼音
Function getHandleFolderName(folderName)
    folderName = PinYin2(folderName) 
    folderName = Replace(Replace(folderName, "/", "\"), " ", "") 
    folderName = Replace(Replace(folderName, "\\", "\"), "\\", "\") 
    getHandleFolderName = folderName 
End Function 



'获得Main信息网址
Function getMainInfoUrl(title)
    Dim fileName, url 
    rs.Open "Select * From [MainInfo] Where Title='" & title & "'", conn, 1, 1 
    If Not rs.EOF Then
        '自定义模板
        If Trim(rs("TemplatePath")) <> "" Then
            TemplatePath = WebTemplate & rs("TemplatePath") 
        Else
            TemplatePath = WebTemplate & "Main_Model.Html" 
        End If 


        '更新于20150209
        url = Trim(rs("FileName")) 

        If CheckMakeHtmlFile(url) = True Then
            fileName = GetHandleFileName(url) 
            url = detailPageAddHtmlPrefix(rs("FolderName")) & "/" & url 
            url = GetHandleUrl(url)                                                         '处理URL
        End If 

        url = WebDebug(url, "act=MainInfo&Id=" & rs("Id"))                          '调试
    End If : rs.Close 
    getMainInfoUrl = url 
End Function 
'获得Main信息网址  (赋值上面)
Function getMainUrl(title)
    getMainUrl = getMainInfoUrl(title) 
End Function 
'获得导航链接地址 获得导航文件地址 获得导航大类 获得导航小类  导航大类链接 导航小类链接 导航链接
Function getNavUrl(did, sid)
    Dim url, folderName, TempRs 

    Set TempRs = CreateObject("Adodb.RecordSet")
        TempRs.Open "Select * From [NavBigClass] Where BigClassName='" & did & "'", conn, 1, 1 
        If Not TempRs.EOF Then
            folderName = Trim(TempRs("FolderName")) 
            url = GetWebFolderName & "/" & TempRs("FolderName") & "/" & TempRs("FileName") 
            url = GetHandleUrl(url)                                                         '处理URL
            '追加判断20141226
            If CheckMakeHtmlFile(TempRs("FileName")) = False Then
                url = TempRs("FileName") 
            End If 

            url = WebDebug(url, "act=Nav&NavDid=" & did)                                '调试
        End If : TempRs.Close 
        If sid <> "" Then
            TempRs.Open "Select * From [NavSmallClass] Where BigClassName='" & did & "' And SmallClassName='" & sid & "'", conn, 1, 3 
            If Not TempRs.EOF Then
                If TempRs("FolderName") = "" Then
                    url = "/" & folderName & "/" 
                Else
                    url = "/" & TempRs("FolderName") & "/" 
                End If 
                If TempRs("FileName") = "" Then
                    TempRs("FileName") = HtmlFileName(TempRs("SmallClassName")) 
                    TempRs.Update 
                End If 
                url = GetWebFolderName & url & TempRs("FileName") 
                url = GetHandleUrl(url)                                                         '处理URL
                'Call Echo("Url", Url)
                '追加判断20141226
                If CheckMakeHtmlFile(TempRs("FileName")) = False Then
                    url = TempRs("FileName") 
                End If 

                url = WebDebug(url, "act=Nav&NavDid=" & did & "&NavSid=" & sid)         '调试
            'Call Echo("Url", Url)
            End If : TempRs.Close 
        End If 
        getNavUrl = url 
    'Call Echo("GetNavUrl", GetNavUrl)
End Function
'获得导航链接地址打开类型 获得导航文件地址 获得导航大类 获得导航小类
Function getNavTarget(did, sid)
    If sid <> "" Then
        TempRs.Open "Select * From [NavSmallClass] Where BigClassName='" & did & "' And SmallClassName='" & sid & "'", conn, 1, 1 
        If Not TempRs.EOF Then
            getNavTarget = TempRs("Target") 
        End If : TempRs.Close 
        Exit Function 
    End If 
    TempRs.Open "Select * From [NavBigClass] Where BigClassName='" & did & "'", conn, 1, 1 
    If Not TempRs.EOF Then
        getNavTarget = TempRs("Target") 
    End If : TempRs.Close 
End Function 
'获得BigClass表中类型大类文件夹
Function getBigClassFolderName(did)
    Dim TempRs 
    Set TempRs = CreateObject("Adodb.RecordSet")
        TempRs.Open "Select * From [BigClass] Where BigClassName='" & did & "'", conn, 1, 1 
        If Not TempRs.EOF Then

            getBigClassFolderName = "/" & TempRs("FolderName") & "/" & TempRs("FileName") 
            getBigClassFolderName = GetHandleUrl(getBigClassFolderName)                     '处理URL

            '追加判断20141226
            If CheckMakeHtmlFile(TempRs("FileName")) = False Then
                getBigClassFolderName = TempRs("FileName") 
            End If 

            Call WebDebug(getBigClassFolderName, "act=CreateClass&ProDid=" & TempRs("BigClassName") & "&page=1") '调试
        End If : TempRs.Close 
End Function
'调用上一函数  (辅助)
Function getBigClassFileName(did)
    getBigClassFileName = getBigClassFolderName(did) 
End Function 
'调用上一函数  (辅助)
Function getBigClassUrl(did)
    getBigClassUrl = getBigClassFolderName(did) 
End Function 
'获得类别小类文件夹
Function getSmallClassFolderName(did, sid)
    Dim url, RootFolder, DirName, TempRs 
    Set TempRs = CreateObject("Adodb.RecordSet")

        TempRs.Open "Select * From [BigClass] Where BigClassName='" & did & "'", conn, 1, 1 
        If Not TempRs.EOF Then

            RootFolder = "/" & TempRs("FolderName") & "/" 

            'If CheckMakeHtmlFile(TempRs("FileName"))=True Then
            'GetSmallClassFolderName = "/" & TempRs("FolderName") & "/"
            'GetSmallClassFolderName = GetHandleUrl(GetSmallClassFolderName)    '处理URL
            'Else
            'GetSmallClassFolderName = TempRs("FileName")
            'End If
            getSmallClassFolderName = RootFolder 
            Call WebDebug(getSmallClassFolderName, "act=CreateClass&ProDid=" & TempRs("BigClassName")) '调试
        End If : TempRs.Close 
        If sid <> "" Then
            TempRs.Open "Select * From [SmallClass] Where BigClassName='" & did & "' And SmallClassName='" & sid & "'", conn, 1, 1 
            If Not TempRs.EOF Then


                DirName = Trim(Replace(TempRs("FolderName"), "\", "/")) 
                If Left(DirName, 1) <> "/" Then
                    RootFolder = RootFolder & "/" & DirName & "/" 
                Else
                    RootFolder = "/" & TempRs("FolderName") & "/" 
                End If 

                'If CheckMakeHtmlFile(TempRs("FileName"))=True Then
                'GetSmallClassFolderName = GetSmallClassFolderName & "/" & TempRs("FolderName") & "/" & TempRs("FileName")
                'GetSmallClassFolderName = GetHandleUrl(GetSmallClassFolderName)    '处理URL
                'Else
                'GetSmallClassFolderName = TempRs("FileName")
                'End If
                If CheckMakeHtmlFile(TempRs("FileName")) = True Then
                    RootFolder = RootFolder & TempRs("FileName") 
                    RootFolder = "/" & GetWebFolderName() & GetHandleUrl(RootFolder) '处理URL         '改进
                Else
                    RootFolder = "/" & GetWebFolderName() & TempRs("FileName")       '改进
                    '判断当文件名称为网址是 则返回网址  此函数是获得小类目录名 改进 20150107
                    If GetWebSite(TempRs("FileName")) <> "" Then
                        RootFolder = TempRs("FileName") 
                    End If 
                End If 

                getSmallClassFolderName = RootFolder 

                Call WebDebug(getSmallClassFolderName, "act=CreateClass&ProDid=" & did & "&ProSid=" & TempRs("SmallClassName")) '调试
            End If : TempRs.Close 
        End If 
End Function
'调用上一函数   (辅助)
Function getSmallClassFileName(did, sid)
    getSmallClassFileName = getSmallClassFolderName(did, sid) 
End Function 
'调用上一函数   (辅助)
Function getSmallClassUrl(did, sid)
    getSmallClassUrl = getSmallClassFolderName(did, sid) 
End Function 
'获得类型子类文件夹
Function getThreeClassFolderName(did, sid, tid)
    Dim RootFolder, DirName 
    'Call Echo("显示",Did & ">>" & Sid & ">>" & Tid)
    TempRs.Open "Select * From [BigClass] Where BigClassName='" & did & "'", conn, 1, 1 
    If Not TempRs.EOF Then

        RootFolder = "/" & TempRs("FolderName") & "/" 

        'If CheckMakeHtmlFile(TempRs("FileName"))=True Then
        'GetThreeClassFolderName = "/" & TempRs("FolderName") & "/"
        'GetThreeClassFolderName = GetHandleUrl(GetThreeClassFolderName)    '处理URL
        'Else
        'GetThreeClassFolderName = TempRs("FileName")
        'End If
        getThreeClassFolderName = RootFolder 
        Call WebDebug(getThreeClassFolderName, "act=CreateClass&ProDid=" & TempRs("BigClassName")) '调试
    End If : TempRs.Close 
    If sid <> "" Then
        TempRs.Open "Select * From [SmallClass] Where BigClassName='" & did & "' And SmallClassName='" & sid & "'", conn, 1, 1 
        If Not TempRs.EOF Then

            'If CheckMakeHtmlFile(TempRs("FileName"))=True Then
            'GetThreeClassFolderName = GetThreeClassFolderName & "/" & TempRs("FolderName") & "/"
            'GetThreeClassFolderName = GetHandleUrl(GetThreeClassFolderName)    '处理URL
            'Else
            'GetThreeClassFolderName = TempRs("FileName")
            'End If

            DirName = Trim(Replace(TempRs("FolderName"), "\", "/")) 
            If Left(DirName, 1) <> "/" Then
                RootFolder = RootFolder & "/" & DirName & "/" 
            Else
                RootFolder = "/" & TempRs("FolderName") & "/" 
            End If 

            getThreeClassFolderName = RootFolder 
            Call WebDebug(getThreeClassFolderName, "act=CreateClass&ProDid=" & did & "&ProSid=" & TempRs("SmallClassName")) '调试
        End If : TempRs.Close 
    End If 
    If tid <> "" Then
        TempRs.Open "Select * From [ThreeClass] Where BigClassName='" & did & "' And SmallClassName='" & sid & "' And ThreeClassName='" & tid & "'", conn, 1, 1 
        If Not TempRs.EOF Then

            DirName = Trim(Replace(TempRs("FolderName"), "\", "/")) 
            If Left(DirName, 1) <> "/" Then
                RootFolder = RootFolder & "/" & DirName & "/" 
            Else
                RootFolder = "/" & TempRs("FolderName") & "/" 

                '判断当文件名称为网址是 则返回网址  此函数是获得子类目录名 改进 20150107
                If GetWebSite(TempRs("FileName")) <> "" Then
                    RootFolder = TempRs("FileName") 
                End If 

            End If 
            'Call Echo(TempRs("FileName"), CheckMakeHtmlFile(TempRs("FileName")))
            'If CheckMakeHtmlFile(TempRs("FileName"))=True Then
            'GetThreeClassFolderName = GetThreeClassFolderName & "/" & TempRs("FolderName") & "/" & TempRs("FileName")
            'GetThreeClassFolderName = GetHandleUrl(GetThreeClassFolderName)    '处理URL
            'Else
            'GetThreeClassFolderName = TempRs("FileName")
            'End If

            If CheckMakeHtmlFile(TempRs("FileName")) = True Then
                RootFolder = RootFolder & TempRs("FileName") 
                RootFolder = GetHandleUrl(RootFolder)                                           '处理URL
            Else
                RootFolder = TempRs("FileName") 
            End If 

            getThreeClassFolderName = RootFolder 
            Call WebDebug(getThreeClassFolderName, "act=CreateClass&ProDid=" & did & "&ProSid=" & sid & "&ProTid=" & TempRs("ThreeClassName")) '调试
        End If : TempRs.Close 
    End If 
End Function 
'调用上一函数   (辅助)
Function getThreeClassFileName(did, sid, tid)
    getThreeClassFileName = getThreeClassFolderName(did, sid, tid) 
End Function 
'调用上一函数   (辅助)
Function getThreeClassUrl(did, sid, tid)
    getThreeClassUrl = getThreeClassFolderName(did, sid, tid) 
End Function 
'获得产品类型的URL地址，包括大类与小类与子类
Function getClassUrl(did, sid, tid)
    Dim url 
    If tid <> "" Then
        url = getThreeClassFolderName(did, sid, tid) 
    ElseIf sid <> "" Then
        url = getSmallClassFolderName(did, sid) 
    Else
        url = getBigClassFolderName(did) 
    End If 
    url = Trim(Replace(url, "\", "/")) 
    If Left(url, 1) <> "/" And GetWebSite(url) = "" Then url = "/" & url 
    getClassUrl = url 
End Function 
'获得产品类型的URL地址，包括大类与小类与子类 导航用
Function getClassUrl_Nav(did, sid, tid)
    Dim url 
    url = getClassUrl(did, sid, tid) 
    url = Mid(url, 2) 
    If Right(url, 1) = "/" Then url = Left(url, Len(url) - 1) 
    getClassUrl_Nav = url 
End Function 
'获得产品链接地址
Function getProductUrl(addSql)
    TempRs.Open "Select * From [Product] " & addSql, conn, 1, 1 
    If Not TempRs.EOF Then
        getProductUrl = GetHandleUrl(TempRs("FileName")) 
        Call WebDebug(getProductUrl, "act=CreateArticle&ID=" & TempRs("Id"))        '调试
    End If : TempRs.Close 
End Function 
'自动处理产品文件名称 不重复    添加于2014,03,28
Function aotuGetProductFileName(ByVal title)
    Dim YesOK, fileName, TempTitle 
    TempTitle = title 
    title = Md5(TempTitle, 2) & Md5(TempTitle, 4) 
    YesOK = True 
    While YesOK
        TempRs.Open "Select * From [Product] Where Title='" & title & "'", conn, 1, 1 
        If TempRs.EOF Then
            YesOK = False 
            aotuGetProductFileName = title 
        Else
            TempTitle = TempTitle & "1" 
            title = Md5(TempTitle, 2) & Md5(TempTitle, 4) 
        End If : TempRs.Close 
        'Call Echo(TempTitle,Title) : Doevents
    Wend 
End Function 
'产品上一页
Function upPage(id)
    Dim add, did, sid, tid, sql, url, Target, title, AHttpUrl 
    RsTemp.Open "Select * From [Product] Where ID=" & id, conn, 1, 1 
    If Not RsTemp.Eof Then
        did = RsTemp("BigClassName") 
        sid = RsTemp("SmallClassName") 
        tid = RsTemp("ThreeClassName") 
    End If : RsTemp.Close 

    If tid <> "" Then
        add = "BigClassName='" & did & "' And SmallClassName='" & sid & "' And ThreeClassName='" & tid & "' And" 
    ElseIf sid <> "" Then
        add = "BigClassName='" & did & "' And SmallClassName='" & sid & "' And" 
    Else
        add = "BigClassName='" & did & "' And " 
    End If 
    sql = "Select * from [Product] where " & add & "  Id < " & id & " order by Id  desc" 

    RsTemp.Open sql, conn, 1, 1 
    If Not RsTemp.Eof Then

        url = RsTemp("FileName") 
        If CheckRemoteUrl(url) = False Then
            url = GetHandleUrl(RsTemp("FileName")) 
            Call WebDebug(url, "act=CreateArticle&ID=" & RsTemp("Id"))                  '调试
        End If 

        Target = "_top" 
        title = RsTemp("Title") 
        AHttpUrl = "<a " & AHref(url, title, Target) & ">上一篇：" 

        upPage = AHttpUrl & CutStr(title, 100, "...") & "</a><br>" 
    Else
        upPage = InfoColor("上一篇：没有<br>", SysStyle(0)) & "" 
    End If 
    RsTemp.Close 
End Function 
'产品下一页
Function downPage(id)
    Dim add, did, sid, tid, sql, url, Target, title, AHttpUrl 
    RsTemp.Open "Select * From [Product] Where ID=" & id, conn, 1, 1 
    If Not RsTemp.EOF Then
        did = RsTemp("BigClassName") 
        sid = RsTemp("SmallClassName") 
        tid = RsTemp("ThreeClassName") 
    End If : RsTemp.Close 

    If tid <> "" Then
        add = "BigClassName='" & did & "' And SmallClassName='" & sid & "' And ThreeClassName='" & tid & "' And" 
    ElseIf sid <> "" Then
        add = "BigClassName='" & did & "' And SmallClassName='" & sid & "' And" 

    Else
        add = "BigClassName='" & did & "' And " 
    End If 
    sql = "select * from Product where " & add & "  Id > " & id & " order by Id  asc" 
    RsTemp.Open sql, conn, 1, 1 
    If Not RsTemp.EOF Then

        url = RsTemp("FileName") 
        If CheckRemoteUrl(url) = False Then
            url = GetHandleUrl(RsTemp("FileName")) 
            Call WebDebug(url, "act=CreateArticle&ID=" & RsTemp("Id"))                  '调试
        End If 

        Target = "_top" 
        title = RsTemp("Title") 
        AHttpUrl = "<a " & AHref(url, title, Target) & ">下一篇：" 

        downPage = AHttpUrl & CutStr(title, 100, "...") & "</a><br>" 

    Else
        downPage = InfoColor("下一篇：没有<br>", SysStyle(0)) 
    End If 
    RsTemp.Close 
End Function 

'处理获得完善Html文件名称
Function getHandleFileName(ByVal fileName)
    '处理获得文件路径
    If fileName = "" Or fileName = "/" Or fileName = "\" Then
        getHandleFileName = "/Index.Html" 
    Else
        getHandleFileName = fileName & ".Html" 
    End If 
End Function 



'生成HTML文件名称处理
Function htmlFileName(ByVal fileName)
    htmlFileName = handleSaveFileName(fileName, True) 
End Function 
'处理成能保存的文件名称(20150902)
Function handleSaveFileName(ByVal fileName, isToPinYin)
    Dim ZD, i, s 
    fileName = Trim(fileName) 
    If CheckMakeHtmlFile(fileName) = True Then
        ZD = ":*?""<>|.," 
        For i = 1 To Len(ZD)
            s = Mid(ZD, i, 1) 
            fileName = Replace(fileName, s, "") 
        Next 
        If isToPinYin = True Then
            fileName = Trim(PinYin2(fileName)) 
        End If 
    End If 
    'If FileName = "" Then FileName = GetDefaultFileName()            '文件名称为空则获得默认文件名称 不能要
    handleSaveFileName = Left(fileName, 100)                                        '截取100个字符
End Function 

'检查是否可以创建文件
Function checkMakeHtmlFile(ByVal fileName)
    Dim url 
    fileName = Trim(LCase(Replace(fileName, "\", "/"))) : url = fileName 
    If InStr(fileName, "/") > 0 Then
        fileName = Trim(Mid(fileName, InStrRev(fileName, "/") + 1)) 
    End If 
    checkMakeHtmlFile = True 
    If Left(url, 7) = "http://" Or Left(url, 4) = "www." Or Left(url, 4) = "url:" Or Left(url, 5) = "/url:" Or Left(fileName, 1) = "#" Or Left(fileName, 11) = "javascript:" Then
        checkMakeHtmlFile = False 
    End If 
End Function 
'显示在线修改面板 2014 12 12
'用法：
'Url = ""& adminDir &"MainInfo.Asp?act=ShowEdit&Id=" & Id & "&n=" & GetRnd(11)
'Content = DisplayOnlineEditDialog(Url, Content)
'添加在线修改
'MainContent = DisplayOnlineEditDialog(""& adminDir &"NavManage.Asp?act=EditNavBig&Id=" & TempRs("Id") & "&n=" & GetRnd(11), MainContent)
Function displayOnlineEditDialog(url, content)
    displayOnlineEditDialog = HandleDisplayOnlineEditDialog(url, content, "", "") 
End Function 
'在线修改快速添加二   S = DisplayOnlineED2(Url, S, "<li|<a ")
Function displayOnlineED2(url, content, ReplaceStr)
    displayOnlineED2 = HandleDisplayOnlineEditDialog(url, content, "", ReplaceStr) 
End Function 
%> 
