<% 
'自动添加信息

Select Case Request("a")
    Case "ayawawa", "wawa" : Call Ayawawa() : Response.End()
    Case "row", "space" : Call HandRowSpace() : Response.End()
    Case "ShowAoutAddList" : Call ShowAoutAddList() : Response.End()                '显示列表
    Case "AutoAddWebInfo" : Call AutoAddWebInfo(Request("DataFolderPath")) : Response.End()
    Case "ClearDatabases" : Call ClearDatabases() : Response.End()                  '清空数据库
    Case "AddDateViewWebEffect" : Call AddDateViewWebEffect() : Response.End()      '写入数据查看网站效果
End Select

'显示自动添加信息选择列表20140315
Sub showAoutAddList()
    Dim splStr, i, folderPath, s, content, url 
    folderPath = "/DataDir/自动添加网站信息/" 
    Call HandlePath(folderPath)                                                     '获得完整路径
    Call echo("显示自动添加信息选择列表", folderPath) 
    content = GetDirFolderList(folderPath) 
    Call echo("清空数据库，方便查看程序是否添加成功", "<a href=?a=ClearDatabases>点击操作</a>") 
    Call echo("写入数据查看网站效果", "<a href=?a=AddDateViewWebEffect>点击操作</a>") 
    splStr = Split(content, vbCrLf) 
    For Each folderPath In splStr
        If folderPath <> "" Then
            s = Mid(folderPath, InStrRev(folderPath, "\") + 1) 
            url = "?act=AutoAddWebInfo&DataFolderPath=" & folderPath 
            Call echo(s, "<a href='" & url & "' target='_self'>" & folderPath & "</a>") 
        End If 
    Next 
End Sub 


'VB内容开始
'写入指定数据，方便查看网站整体效果(写入数据查看网站效果)
Sub addDateViewWebEffect()
    Dim NewConn, path, DidList, IDList, splStr, did, nCount, ArticleFolder, filePath, nOK, nArticle 
    Dim sql, addSql 
    Call OpenConn() 
    path = "E:\E盘\WEB网站\网站备份\Datamaiside20141201.mdb" 
    Set NewConn = CreateObject("Adodb.Connection")
        NewConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database PassWord = '';Data Source = " & path 
        '导航大类
        rs.Open "Select * From [NavBigClass] Where (NavType='产品' Or NavType='新闻') Order By NavType", conn, 1, 1 
        While Not rs.EOF
            If InStr("|" & DidList & "|", "|" & rs("BigClassName") & "|") = False Then
                DidList = DidList & rs("BigClassName") & "|" 
            End If 
        rs.MoveNext : Wend : rs.Close 

        '文章文件夹
        If 1 = 2 Then
            ArticleFolder = "E:\E盘\WEB网站\网站备份\文章\" 
            Call DeleteFolder(ArticleFolder) 
            Call CreateFolder(ArticleFolder) 
            nOK = 0 
            rs.Open "Select * From [Product] Where Len(Content)>200", NewConn, 1, 1 
            While Not rs.EOF
                nOK = nOK + 1 
                filePath = ArticleFolder & nOK & ".txt" 
                Call echo("创建文章", filePath) 
                Call CreateFile(filePath, rs("Content")) 
            rs.MoveNext : Wend : rs.Close 
            Response.End() 
        End If 

        'AddSql "And Where BigClassName<>'常见问题'"

        'Call Eerr("导航",DidList)
        nCount = NewConn.Execute("Select Count(*) From [Product] Where (BigFiles<>'' Or SmallFiles<>'')")(0) 
        Call echo("nCount", nCount) 
        '产品大类
        rs.Open "Select * From [BigClass]", conn, 1, 1 
        While Not rs.EOF
            If InStr("|" & DidList & "|", "|" & rs("BigClassName") & "|") = False Then
                DidList = DidList & rs("BigClassName") & "|" 
            End If 
        rs.MoveNext : Wend : rs.Close 
        nArticle = 0                                                                    '文章默认初值
        '给产品大类添加信息
        splStr = Split(DidList, "|") 
        For Each did In splStr
            If did <> "" And Trim(LCase(did)) <> "banner" Then
                Call echo("大类名称", did & "(添加数量" & CopyDateToWebDate(conn, NewConn, did, 18, IDList, nArticle) & ")") 
            End If 
        Next 
End Sub


'拷贝数据库数组到指定网站数据库里
Function copyDateToWebDate(conn, NewConn, did, topNumb, IdList, nArticle)
    Dim i, j, YesOK, ImgUrl, nAdd, content, IsTrue, addSql 
    conn.Execute("Delete From [Product] Where BigClassName='" & did & "'") 
    nAdd = 0 
    For i = 1 To topNumb
        addSql = " Where BigFiles<>'' Or SmallFiles<>''" 
        If did = "常见问题" Then
            addSql = " Where BigFiles='' And SmallFiles=''" 
        'Call Echo("Did",Did)
        End If 
        rsx.Open "Select * From [Product]", NewConn, 1, 3 
        YesOK = True 
        While Not rsx.EOF And YesOK = True
            IsTrue = False 
            If did = "常见问题" And rsx("BigFiles") = "" And rsx("SmallFiles") = "" Then
                IsTrue = True 
            ElseIf rsx("BigFiles") <> "" Or rsx("SmallFiles") <> "" Then
                IsTrue = True 
            End If 
            If IsTrue = True Then
                If InStr("|" & IdList & "|", "|" & rsx(0) & "|") = False Then
                    IdList = IdList & rsx(0) & "|" 
                    'Call Echo(Did & ",Id=" & Rsx(0), "IDList=" & IDList)
                    YesOK = False 
                    rss.Open "Select * From [Product]", conn, 1, 3 
                    rss.AddNew 
                    rss("BigClassName") = did 
                    rss("Title") = rsx("Title") 

                    ImgUrl = Replace(Trim(rsx("SmallFiles")), "\", "/") 
                    If Left(ImgUrl, 1) = "/" Then ImgUrl = "http://maiside.net/" & ImgUrl 
                    rss("SmallFiles") = rsx("SmallFiles") 

                    ImgUrl = Replace(Trim(rsx("BigFiles")), "\", "/") 
                    If Left(ImgUrl, 1) = "/" Then ImgUrl = "http://maiside.net/" & ImgUrl 
                    rss("BigFiles") = ImgUrl 

                    nArticle = nArticle + 1 
                    content = rsx("Content") 
                    If Len(content) < 200 Then
                        content = GetFText("E:\E盘\WEB网站\网站备份\文章\" & nArticle & ".txt") 
                    End If 

                    rss("WebTitle") = rsx("WebTitle")                                               '网站自定标题
                    rss("WebKeywords") = rsx("WebKeywords")                                         '网站关键词
                    rss("WebDescription") = rsx("WebDescription")                                   '网站描述

                    rss("FileName") = PinYin(Trim(rsx("Title")), 3) 


                    If Len(content) < 255 Then
                        content = content & GetArticleFillContent("", "", 9) 
                    End If 

                    rss("Content") = content 
                    rss("Author") = "站长" 
                    rss("ArticleSource") = "本站" 
                    nAdd = nAdd + 1 
                    rss.Update : rss.Close 
                End If 
            End If 
        rsx.MoveNext : Wend : rsx.Close 
    Next 
    copyDateToWebDate = nAdd 
End Function 



'清空数据库，方便查看程序是否添加成功
Sub clearDatabases()
    Dim i, j 
    Call OpenConn() 
    '网站
    rs.Open "Select * From [WebSite]", conn, 1, 3 
    If Not rs.EOF Then
        rs("WebTitle") = "网站标题" 
        rs("WebKeywords") = "网站关键词" 
        rs("WebSiteBottom") = "网站底部" 
        rs.Update 
    End If : rs.Close 
    '处理文本
    rs.Open "Select * From [MainInfo]", conn, 1, 3 
    For i = 1 To rs.RecordCount
        rs("ShowTitle") = "文本标题(" & i & ")" 
        rs("Content") = "文本内容(" & i & ")" 
        rs.Update 
    rs.MoveNext : Next : rs.Close 
    '产品
    rs.Open "Select * From [Product]", conn, 1, 3 
    For i = 1 To rs.RecordCount
        rs("Title") = "信息标题(" & i & ")" 
        rs("Content") = "信息内容(" & i & ")" 
        If rs("BigFiles") <> "" Then
            rs("BigFiles") = "/UploadFiles/NoImg.jpg" 
        End If 
        If rs("SmallFiles") <> "" Then
            rs("SmallFiles") = "/UploadFiles/NoImg.jpg" 
        End If 
        rs("WebTitle") = "网站自定标题(" & i & ")" 
        rs("WebKeywords") = "网站关键词(" & i & ")" 
        rs("WebDescription") = "网站描述(" & i & ")" 
        rs.Update 
    rs.MoveNext : Next : rs.Close 
    '大类导航
    rs.Open "Select * From [NavBigClass]", conn, 1, 3 
    For i = 1 To rs.RecordCount
        '小类导航
        rsx.Open "Select * From [SmallClass] Where BigClassName='" & rs("BigClassName") & "'", conn, 1, 3 
        For j = 1 To rsx.RecordCount
            rsx("BigClassName") = "大导航(" & i & ")" 
            rsx("SmallClassName") = "小导航(" & j & ")" 
            rsx.Update 
        rsx.MoveNext : Next : rsx.Close 
        rs("BigClassName") = "大导航(" & i & ")" 
        rs.Update 
    rs.MoveNext : Next : rs.Close 
    '友情链接
    rs.Open "Select * From [FriendLink]", conn, 1, 3 
    For i = 1 To rs.RecordCount
        rs("Title") = "友链接(" & i & ")" 
        rs.Update 
    rs.MoveNext : Next : rs.Close 
    '留言
    rs.Open "Select * From [GuestBook]", conn, 1, 3 
    For i = 1 To rs.RecordCount
        rs("Content") = "留言内容(" & i & ")" 
        rs("Reply") = "回复内容(" & i & ")" 
        rs.Update 
    rs.MoveNext : Next : rs.Close 

    '搜索统计
    rs.Open "Select * From [SearchCount]", conn, 1, 3 
    For i = 1 To rs.RecordCount
        rs("Title") = "搜索词(" & i & ")" 
        rs.Update 
    rs.MoveNext : Next : rs.Close 
    Call echo("操作完成", "空数据库方便调试程序成功") 
End Sub 



'自动添加网站信息 数据自动处理
Sub autoAddWebInfo(DataFolderPath)
    Dim SplList, splStr, splxx, i, j, p, s, sS, c, folderName, title, did, sid, tid, content, ContactUs, BigFolder, SmallFolder, path, fileName, url, id, ExpireTime 
    Dim nCount, DataFolder, NavPath, NavName, NavType, NavSort, NavLocation, strHtml, DivCssStr, SidSpl, SmallNavName 
    Dim ProductFolder, sort, BigFiles, ClassPath, SplDid, SplSid, SplTid, addSql, sql, ImgPath, Alt, Through, userName, Pwd, SidSort, TidSort, ClassName 
    Dim IDSort, TableName, OnHtml, IsWebTitle, ToPath, filePath, StartTime, QuanXian, BannerName 
    Dim DidName, SidName, TidName 
    Dim TemplatePath                                                                'VB里用到 20150625
    Dim msg                                                                         '回显信息

    StartTime = Now() 
    '数据文件夹路径
    DataFolder = HandlePath(DataFolderPath) & "\" 
    Call rw("<script type='text/javascript' src='/Skins/Default/Js/MyFunction.Js'></script>") 
    Call echo("数据文件夹", DataFolder) 
    DivCssStr = " style='margin-top:10px;width:250px;height:16px;overflow:hidden;background:#ECE9D8;border:1px double #716F64;'" 
    '先删除全部数据 Delete * From在Access里可以 Delete From属于通用
    conn.Execute("Delete From [NavBigClass]")                                  '删除导航大类
    conn.Execute("Delete From [NavSmallClass]")                                '删除导航小类
    conn.Execute("Delete From [BigClass]")                                     '删除产品大类
    conn.Execute("Delete From [SmallClass]")                                   '删除产品小类
    conn.Execute("Delete From [ThreeClass]")                                   '删除产品子类
    conn.Execute("Delete From [Product]")                                      '删除产品信息
    conn.Execute("Delete From [MainInfo]")                                     '删除单文本列表
    conn.Execute("Delete From [SearchCount]")                                  '删除搜索信息表
    conn.Execute("Delete From [Member]")                                       '删除会员列表
    conn.Execute("Delete From [GuestBook]")                                    '删除留言列表
    conn.Execute("Delete From [FriendLink]")                                   '删除友情链接列表
    conn.Execute("Delete From [Admin]")                                        '删除Admin管理表 修改于20140928
    'CAll Eerr( GetRecordCount("BigClass", ""),"删除大类")

    Call DeleteFolder("../UploadFiles/image") 
    Call CreateFolder("../UploadFiles/image") 

    '网站配置
    path = DataFolder & "网站配置.ini" 
    If checkFile(path) Then
        content = GetFText(path) 
        rs.Open "Select * From [WebSite]", conn, 1, 3 
        If rs.EOF Then rs.AddNew 

        If InStr(content, "WebSiteUrl=") > 0 Then
            rs("WebSiteUrl") = StrCut(content, "WebSiteUrl=", "【End】", 2) 
        End If 

        If InStr(content, "WebTitle=") > 0 Then
            rs("WebTitle") = StrCut(content, "WebTitle=", "【End】", 2) 
        End If 

        If InStr(content, "WebKeywords=") > 0 Then
            rs("WebKeywords") = StrCut(content, "WebKeywords=", "【End】", 2) 
        End If 

        If InStr(content, "WebDescription=") > 0 Then
            rs("WebDescription") = StrCut(content, "WebDescription=", "【End】", 2) 
        End If 

        If InStr(content, "WebSiteBottom=") > 0 Then
            rs("WebSiteBottom") = StrCut(content, "WebSiteBottom=", "【End】", 2) 
        End If 
        '模板路径
        If InStr(content, "WebTemplate=") > 0 Then
            rs("WebTemplate") = StrCut(content, "WebTemplate=", "【End】", 2) 
        End If 
        'Skins路径
        If InStr(content, "WebSkins=") > 0 Then
            rs("WebSkins") = StrCut(content, "WebSkins=", "【End】", 2) 
        End If 
        'Images路径
        If InStr(content, "WebImages=") > 0 Then
            rs("WebImages") = StrCut(content, "WebImages=", "【End】", 2) 
        End If 
        'Css路径
        If InStr(content, "WebCss=") > 0 Then
            rs("WebCss") = StrCut(content, "WebCss=", "【End】", 2) 
        End If 
        'Js路径
        If InStr(content, "WebJs=") > 0 Then
            rs("WebJs") = StrCut(content, "WebJs=", "【End】", 2) 
        End If 


        rs.Update : rs.Close 

        strHtml = "<div id='WebSite' " & DivCssStr & "><span onclick=""OnOffDiv(this,'WebSite')"">添加网站配置</span><br></div>" & vbCrLf 
        Call rw(strHtml) 
    End If 

    '后台账号配置
    path = DataFolder & "后台账号.ini" 
    If checkFile(path) Then
        content = GetFText(path) 
        splStr = Split(content, vbCrLf) 
        For i = 0 To UBound(splStr)
            s = splStr(i) 
            If s <> "" And Left(Trim(s), 1) <> "#" Then
                splxx = Split(s & "||||||", "|") 
                userName = Trim(splxx(0)) 
                Pwd = splxx(1) 
                If Pwd = "" Then Pwd = userName 
                QuanXian = splxx(2) 
                'Call Echo(UserName,Pwd)
                rs.Open "Select * From [Admin] Where UserName='" & userName & "'", conn, 1, 3 
                If rs.EOF Then rs.AddNew 
                rs("UserName") = userName 
                rs("Pwd") = MD5(Pwd, 2) & MD5(Pwd, 4) 
                rs("QuanXian") = QuanXian 
                rs.Update 
                strHtml = strHtml & "添加后台账号>>" & userName & "<br>" 
                rs.Close 
            End If 
        Next 
        strHtml = "<div id='AdminUser' " & DivCssStr & "><span onclick=""OnOffDiv(this,'WebSite')"">添加后台账号</span><br></div>" & vbCrLf 
        Call rw(strHtml) 
    Else
        rs.Open "Select * From [Admin] Where UserName='admin'", conn, 1, 3 
        userName = "admin" : Pwd = "admin" 
        If rs.EOF Then rs.AddNew 
        rs("UserName") = userName 
        rs("Pwd") = MD5(Pwd, 2) & MD5(Pwd, 4) 
        rs.Update : rs.Close 
    End If 
    '##############################################

    '网站导航
    NavPath = DataFolder & "导航/Nav.ini" 
    If checkFile(NavPath) = True Then
        content = GetFText(NavPath) 
        splStr = Split(content, vbCrLf) 
        strHtml = "<div id='Nav' " & DivCssStr & "><span onclick=""OnOffDiv(this,'Nav')"">添加网站导航</span><br>" & vbCrLf 
        For i = 0 To UBound(splStr)
            s = splStr(i) 
            '结束标签
            If LCase(Trim(s)) = "#end#" Or LCase(Trim(s)) = "#end" Then
                Exit For 
            ElseIf s <> "" And Left(Trim(s), 1) <> "#" Then
                splxx = Split(s & "||||||||||", "|") 

                NavName = Trim(splxx(0))                                                        '去除两边空格
                NavType = Trim(splxx(1)) 
                NavSort = Trim(splxx(2)) 
                fileName = Trim(splxx(3)) 
                TemplatePath = Trim(splxx(4)) 
                NavLocation = Trim(splxx(5)) 
                If NavType = "" Then NavType = "文本" 
                If NavSort = "" Then NavSort = i 
                folderName = splxx(6) 
                OnHtml = splxx(7) 
                OnHtml = IIF(OnHtml = "" Or OnHtml = "1", True, False) 

                rs.Open "Select * From [NavBigClass] Where BigClassName='" & NavName & "'", conn, 1, 3 
                If rs.EOF And NavName <> "" Then
                    rs.AddNew 
                    rs("BigClassName") = NavName 
                    rs("NavTop") = True 
                    rs("NavType") = NavType 
                    rs("FileName") = fileName 
                    rs("TemplatePath") = TemplatePath 
                    'Rs("FileName") = FileName
                    rs("OnHtml") = OnHtml 
                    If folderName = "NOFolderName" Or folderName = "NO文件夹" Then
                        folderName = "" 
                    ElseIf splxx(0) = "网站地图" Then                                               '单独处理网站地图
                        rs("FolderName") = "Sitemap" 
                    ElseIf splxx(1) <> "首页" Then
                        folderName = IIF(folderName = "", PinYin(splxx(0), 3), folderName) 
                        rs("FolderName") = folderName 
                    Else
                        rs("FolderName") = "/" 
                    End If 

                    'Call Echo("NavLocation",NavLocation)    '调试
                    '导航位置  默认顶部为显示 要想让它不显示则使用NO顶部
                    If InStr(NavLocation, "NO顶部") > 0 Then rs("NavTop") = False 
                    If InStr(NavLocation, "底部") > 0 Then rs("NavButtom") = True 
                    If InStr(NavLocation, "左边") > 0 Then rs("NavLeft") = True 
                    If InStr(NavLocation, "中间") > 0 Then rs("NavContent") = True 
                    If InStr(NavLocation, "右边") > 0 Then rs("NavRight") = True 
                    If InStr(NavLocation, "其它") > 0 Then rs("NavOthre") = True 

                    '给产品新闻类创建大类
                    If InStr("|产品|新闻|", "|" & NavType & "|") > 0 Then
                        Call AutoAddBigClass(splxx(0))                                                  '自动添加类型大类
                    End If 
                    content = NavName & " 详细信息内容写在这里面<hr>更多....." 
                    path = DataFolder & "导航/" & NavName & ".txt" 
                    If checkFile(path) = True Then
                        content = GetFText(path) 
                    End If 
                    rs("Content") = content 
                    rs("Sort") = NavSort 
                    strHtml = strHtml & "名称=" & NavName & "<br>类型=" & NavType & "<br>排序=" & NavSort & _ 
                     "<br>文件夹=" & folderName & "<br>静态文件名=" & fileName & "<br>模板=" & TemplatePath & "<br>" & vbCrLf 
                    rs.Update 

                    '添加网站导航小类
                    path = DataFolder & "导航/" & NavName & ".ini" 
                    If checkFile(path) = True Then
                        content = GetFText(path) 
                        SidSpl = Split(content, vbCrLf) 
                        For j = 0 To UBound(SidSpl)
                            s = SidSpl(j) 
                            If s <> "" And Left(Trim(s), 1) <> "#" Then
                                splxx = Split(s & "||||||", "|") 
                                SmallNavName = splxx(0) 
                                fileName = splxx(2) 
                                TemplatePath = splxx(3) 
                                If NavSort = "" Then NavSort = j 
                                rsx.Open "Select * From [NavSmallClass] Where BigClassName='" & NavName & "' And SmallClassName='" & SmallNavName & "'", conn, 1, 3 
                                If rsx.EOF Then
                                    rsx.AddNew 
                                    rsx("BigClassName") = NavName 
                                    rsx("SmallClassName") = SmallNavName 
                                    rsx("FolderName") = PinYin(SmallNavName, 3) 
                                    rsx("TemplatePath") = TemplatePath 

                                    content = NavName & " 详细信息内容写在这里面<hr>更多....." 
                                    path = DataFolder & "导航/" & NavName & "》" & SmallNavName & ".txt" 
                                    If checkFile(path) = True Then
                                        content = GetFText(path) 
                                    End If 
                                    rsx("Content") = content 
                                    strHtml = strHtml & " &nbsp; &nbsp; 添加导航小类 " & NavName & "》" & SmallNavName & "<br>" 
                                    rsx.Update 
                                End If : rsx.Close 
                            End If 
                        Next 
                    End If 
                    strHtml = strHtml & "<hr>" & vbCrLf 
                End If : rs.Close 
            End If 
        Next 
    End If 
    Call rw(strHtml & "</div>") 
    '##############################################


    '产品类别
    ClassPath = DataFolder & "类别/Class.ini" 
    If checkFile(NavPath) = True Then
        content = GetFText(ClassPath) 
        SplDid = Split(content, vbCrLf) 
        strHtml = "<div id='Class' " & DivCssStr & "><span onclick=""OnOffDiv(this,'Class')"">添加产品类别</span><br>" & vbCrLf 
        For i = 0 To UBound(SplDid)
            s = SplDid(i) 
            If s <> "" And Left(Trim(s), 1) <> "#" Then
                splxx = Split(s & "||||||", "|") 
                did = Trim(splxx(0)) 
                sort = splxx(1) 
                fileName = splxx(2) 
                TemplatePath = splxx(3) 
                If sort = "" Then sort = i 
                If Left(splxx(0), 2) = vbTab & vbTab Then
                    TidName = did 
                    sql = "Select * From [ThreeClass] Where BigClassName='" & DidName & "' And SmallClassName='" & SidName & "' And ThreeClassName='" & TidName & "'" 
                    If fileName = "" Then fileName = TidName 
                    msg = copystr("&nbsp;", 8) & TidName 
                ElseIf Left(splxx(0), 1) = vbTab Then
                    SidName = did : TidName = "" 
                    sql = "Select * From [SmallClass] Where BigClassName='" & DidName & "' And SmallClassName='" & SidName & "'" 
                    If fileName = "" Then fileName = SidName 
                    msg = copystr("&nbsp;", 4) & SidName 
                Else
                    DidName = did : SidName = "" : TidName = "" 
                    sql = "Select * From [BigClass] Where BigClassName='" & DidName & "'" 
                    If fileName = "" Then fileName = DidName 
                    msg = DidName 
                End If 
                rs.Open sql, conn, 1, 3 
                If rs.EOF Then
                    rs.AddNew 
                    If DidName <> "" Then rs("BigClassName") = DidName 
                    If SidName <> "" Then rs("SmallClassName") = SidName 
                    If TidName <> "" Then rs("ThreeClassName") = TidName 
                    rs("Sort") = sort 
                    rs("FolderName") = PinYin(fileName, 3) 
                    rs.Update 
                    strHtml = strHtml & msg & "<br>" 
                End If : rs.Close 


            End If 
        Next 
        Call rw(strHtml & "</div>") 
    End If 
    '##############################################

    'Product添加信息
    ProductFolder = DataFolder & "Product\" 
    If CheckFolder(ProductFolder) = True Then

        strHtml = "<div id='Product' " & DivCssStr & "><span onclick=""OnOffDiv(this,'Product')"">添加Product信息</span><br>" & vbCrLf 
        SplList = Split(GetDirIniList(ProductFolder), vbCrLf) 
        For Each path In SplList
            If path <> "" Then                                                              '#号为不操作
                s = Mid(path, InStrRev(path, "\") + 1) : s = Left(s, Len(s) - 4) & "》》" 
                splxx = Split(s, "》") 
                did = splxx(0) : sid = splxx(1) : tid = splxx(2) 

                content = GetFText(path) & vbCrLf 
                If InStr(content, "#自动添加大类") > 0 And did <> "" Then
                    Call AutoAddBigClass(did) 
                End If 
                If InStr(content, "#自动添加小类") > 0 And sid <> "" Then
                    Call AutoAddSmallClass(did, sid) 
                End If 
                If InStr(content, "#自动添加子类") > 0 And tid <> "" Then
                    Call AutoAddThreeClass(did, sid, tid) 
                End If 

                If InStr(content, "========================") > 0 Then
                    splStr = Split(content, "========================" & vbCrLf) 
                Else
                    splStr = Split(content, vbCrLf) 
                End If 
                For i = 0 To UBound(splStr)
                    s = splStr(i) 
                    If s <> "" And Left(Trim(s), 1) <> "#" And s <> vbCrLf Then
                        splxx = Split(s & "||||||", "|") 
                        title = splxx(0) 
                        sort = splxx(1) 
                        BigFiles = splxx(2) 
                        TemplatePath = splxx(3) 
                        fileName = AotuGetProductFileName(title) 
                        rs.Open "Select * From [Product] Where BigClassName='" & did & "' And SmallClassName='" & sid & "' And ThreeClassName='" & tid & "' And Title='" & title & "'", conn, 1, 3 
                        If rs.EOF And Trim(title) <> "" Then
                            rs.AddNew 
                            rs("BigClassName") = did 
                            rs("SmallClassName") = sid 
                            rs("ThreeClassName") = tid 
                            rs("Title") = title 
                            rs("BigFiles") = BigFiles 
                            rs("TemplatePath") = TemplatePath 
                            rs("FileName") = fileName 
                            content = did & ">>" & sid & ">>" & tid & ">><p>" & title & " 详细内容信息。" 
                            rs("Content") = content 
                            rs.Update 
                            strHtml = strHtml & "大类=" & did & "<br>小类=" & sid & "<br>子类=" & tid & _ 
                             "<br>标题=" & title & "<br>静态文件名=" & fileName & "<br>模板=" & TemplatePath & "<hr>" & vbCrLf 
                        End If : rs.Close 

                    ElseIf InStr(s, "#导入文件夹#") > 0 Then
                        Call LeadingImage(did, sid, tid, s) 

                    ElseIf InStr(s, "#导入新闻文件夹#") > 0 Then
                        Call LeadingNews(did, sid, tid, s) 
                    End If 
                Next 
            End If 
        Next 
        Call rw(strHtml & "</div>") 
    End If 
    '##############################################

    'MainInfo添加信息
    ProductFolder = DataFolder & "MainInfo\" 
    If CheckFolder(ProductFolder) = True Then

        strHtml = "<div id='MainInfo' " & DivCssStr & "><span onclick=""OnOffDiv(this,'MainInfo')"">添加MainInfo信息</span><br>" & vbCrLf 
        SplList = Split(GetDirIniList(ProductFolder), vbCrLf) 
        For Each path In SplList
            If path <> "" Then
                content = vbCrLf & GetFText(path) & vbCrLf 
                title = Mid(path, InStrRev(path, "\") + 1) : title = Left(title, Len(title) - 4) 

                If InStr(content, "WebTitle=") > 0 Then
                    rs("WebTitle") = StrCut(content, "WebTitle=", "【End】", 2) 
                End If 

                rs.Open "Select * From [MainInfo]", conn, 1, 3 
                rs.AddNew 
                If InStr(content, vbCrLf & "Title=") > 0 Then
                    title = StrCut(content, vbCrLf & "Title=", "【End】", 2) 
                End If 
                rs("Title") = title 
                If InStr(content, "ShowTitle=") > 0 Then
                    rs("ShowTitle") = StrCut(content, "ShowTitle=", "【End】", 2) 
                End If 
                If InStr(content, "Sort=") > 0 Then
                    rs("Sort") = StrCut(content, "Sort=", "【End】", 2) 
                End If 
                If InStr(content, "Content=") > 0 Then
                    rs("Content") = StrCut(content, "Content=", "【End】", 2) 
                End If 
                rs.Update : rs.Close 
                strHtml = strHtml & "添加文本【" & title & "】<hr>" 

            End If 
        Next 
        Call rw(strHtml & "</div>") 
    End If 
    '##############################################

    '友情链接
    path = DataFolder & "友情链接.ini" 
    If checkFile(path) = True Then
        content = GetFText(path) 
        splStr = Split(content, vbCrLf) 
        strHtml = "<div id='FriendLink' " & DivCssStr & "><span onclick=""OnOffDiv(this,'FriendLink')"">添加友情链接</span><br>" & vbCrLf 
        For i = 0 To UBound(splStr)
            s = splStr(i) : sS = LCase(Trim(s)) 
            '为结束段
            If sS = "#end" Or sS = "end" Then
                Exit For 
            ElseIf s <> "" And Left(Trim(s), 1) <> "#" Then
                splxx = Split(s & "||||||", "|") 
                title = splxx(0) 
                url = splxx(1) 
                sort = IIF(splxx(2) = "", i, splxx(2)) 
                ImgPath = splxx(3) 
                Alt = splxx(4) 
                Through = IIF(splxx(5) = "", True, splxx(5)) 


                rs.Open "Select * From [FriendLink] Where Title='" & title & "'", conn, 1, 3 
                If rs.EOF Then
                    rs.AddNew 
                    rs("Title") = title 
                    rs("Url") = url 
                    rs("Sort") = sort 
                    rs("ImgPath") = ImgPath 
                    rs("Alt") = Alt 
                    rs("Through") = Through 
                    rs("Target") = "_blank" 

                    strHtml = strHtml & "标题=" & title & "<br>网址=" & url & "<br>排序=" & sort & "<br>" & vbCrLf 
                    rs.Update 

                    strHtml = strHtml & "<hr>" & vbCrLf 
                End If : rs.Close 
            End If 
        Next 
        Call rw(strHtml & "</div>") 
    End If 
    '##############################################

    '搜索关键词
    path = DataFolder & "SearchCount.ini" 
    If checkFile(path) = True Then
        content = GetFText(path) 
        splStr = Split(content, vbCrLf) 
        strHtml = "<div id='SearchCount' " & DivCssStr & "><span onclick=""OnOffDiv(this,'SearchCount')"">添加搜索关键词</span><br>" & vbCrLf 
        For i = 0 To UBound(splStr)
            s = Trim(splStr(i)) 
            If s <> "" And Left(Trim(s), 1) <> "#" Then
                splxx = Split(s & "||||||", "|") 
                title = splxx(0) 
                sort = splxx(1) 
                Through = IIF(splxx(2) = "" Or LCase(splxx(2)) = "true", 1, 0) 
                rs.Open "Select * From [SearchCount] Where Title='" & title & "'", conn, 1, 3 
                If rs.EOF Then
                    rs.AddNew 
                    rs("Title") = title 
                    rs("Sort") = sort 
                    rs("Through") = Through 
                    rs("Target") = "_blank" 

                    strHtml = strHtml & "标题=" & title & "<br>网址=" & url & "<br>排序=" & sort & "<br>" & vbCrLf 
                    rs.Update 

                    strHtml = strHtml & "<hr>" & vbCrLf 
                End If : rs.Close 
            End If 
        Next 
        Call rw(strHtml & "</div>") 
    End If 
    '##############################################

    '添加新会员
    path = DataFolder & "Member.ini" 
    If checkFile(path) = True Then
        content = GetFText(path) 
        splStr = Split(content, vbCrLf) 
        strHtml = "<div id='Member' " & DivCssStr & "><span onclick=""OnOffDiv(this,'Member')"">添加会员</span><br>" & vbCrLf 
        For i = 0 To UBound(splStr)
            s = Trim(splStr(i)) 
            If s <> "" And Left(Trim(s), 1) <> "#" Then
                splxx = Split(s & "||||||", "|") 
                userName = splxx(0) 
                Pwd = splxx(1) 
                If Pwd = "" Then Pwd = userName                                                 '密码为空则与账号一样

                rs.Open "Select * From [Member] Where UserName='" & userName & "'", conn, 1, 3 
                If rs.EOF Then
                    rs.AddNew 
                    rs("UserName") = userName 
                    rs("Pwd") = MD5(Pwd, 2) & MD5(Pwd, 4) 

                    strHtml = strHtml & "会员：" & userName & "<br>" & vbCrLf 
                    rs.Update 

                    strHtml = strHtml & "<hr>" & vbCrLf 
                End If : rs.Close 
            End If 
        Next 
        Call rw(strHtml & "</div>") 
    End If 
    '##############################################

    '给导航和信息类别加个Banner图片
    IDSort = 0 
    splStr = Split("NavBigClass,NavSmallClass,Bigclass,SmallClass,ThreeClass", ",") 
    For Each TableName In splStr
        rs.Open "Select * From [" & TableName & "] Where ImgPath=''", conn, 1, 3 
        While Not rs.EOF

            If TableName = "NavBigClass" Or TableName = "Bigclass" Then
                BannerName = Trim(rs("BigClassName")) 
            ElseIf TableName = "NavSmallClass" Or TableName = "SmallClass" Then
                BannerName = Trim(rs("SmallClassName")) 
            ElseIf TableName = "ThreeClass" Then
                BannerName = Trim(rs("ThreeClassName")) 
            End If 
            '存在自定义的 则用自定义的
            path = "/Data/WebImg/Banner/" & BannerName & ".jpg" 

            If checkFile(path) = True Then
                rs("ImgPath") = path 
            '麦思德
            ElseIf DataFolderPath = "E:\E盘\WEB网站\至前网站\DataDir\自动添加网站信息\麦思德" Then
                rs("ImgPath") = "/Data/WebImg/Banner/麦思德理念.jpg" 
            '悠口餐饮
            Else
                IDSort = IDSort + 1 
                If IDSort > 21 Then IDSort = Int(IDSort / 21) + 1                               '规定数值间
                rs("ImgPath") = "/Skins/YouKou/Index/Banner/Banner" & IDSort & ".jpg" 
            End If 
        rs.MoveNext : Wend : rs.Close 
    Next 
    '特殊处理
    If DataFolderPath = "E:\E盘\WEB网站\至前网站\DataDir\自动添加网站信息\麦思德" Then
        Call LeadingProductClass("副导航", "汉堡加盟菜单", "", "\DataDir\自动添加网站信息\麦思德\Product\数据\产品\") 
        Call LeadingProductClass("副导航", "汉堡培训分类", "", "\DataDir\自动添加网站信息\麦思德\Product\数据\产品\") 
        '给指定类加个默认产品
        'Call ClassToAddProduct(DataFolderPath,"副导航","","")


        Call echo("当前操作网站", "麦思德") 

    Else
        If DataFolderPath = "E:\E盘\WEB网站\至前网站\DataDir\自动添加网站信息\微战略" Then
            'Call AutoAddBigClassAndArticle("网站建设注意事项",22)
            'Call AutoAddBigClassAndArticle("网站建设行业资讯",22)
            'Call AutoAddBigClassAndArticle("公司新闻",22)
            Call echo("当前操作网站", "微战略") 
        End If 

        Call CreateFolder("\UploadFiles\image\Product")                                 '创建一个文件Product夹
        '给子类加个产品相应信息，链接进入(辅助)
        content = GetDirFileList("/UploadFiles/培训图片/","") & CopyStr(vbCrLf, 99) 
        splStr = Split(content, vbCrLf) 
        rs.Open "Select * From [ThreeClass] Where BigClassName='产品展示'", conn, 1, 1 
        For i = 1 To rs.RecordCount
            addSql = " And ThreeClassName='" & rs("ThreeClassName") & "' And Title='" & rs("ThreeClassName") & "'" 
            rsx.Open "Select * From [Product] Where BigClassName='" & rs("BigClassName") & "' And SmallClassName='" & rs("SmallClassName") & "' " & addSql, conn, 1, 3 
            If rsx.EOF Then
                rsx.AddNew 
                s = splStr(i - 1) 
                s = Replace(s, Server.MapPath("/"), "") 
                rsx("BigClassName") = rs("BigClassName") 
                rsx("SmallClassName") = rs("SmallClassName") 
                rsx("ThreeClassName") = rs("ThreeClassName") 
                rsx("Title") = rs("ThreeClassName") 
                rsx("BigFiles") = s 
                rsx("FileName") = getRnd(9) 
                content = rs("BigClassName") & ">>" & rs("SmallClassName") & ">>" & rs("ThreeClassName") & ">><p>" & rs("ThreeClassName") & " 详细内容信息。" 
                rsx("Content") = content 
                rsx.Update 
            End If : rsx.Close 
        rs.MoveNext : Next : rs.Close 

        '给类别加个产品
        content = GetDirJpgList(DataFolder & "\Product\数据\产品\") 
        splStr = Split(content, vbCrLf) 
        For Each s In splStr
            If s <> "" Then
                title = GetFilePathName(s) 
                rsx.Open "Select * From [Product] Where Title='" & title & "'", conn, 1, 3 
                If rsx.EOF Then rsx.AddNew 
                path = Left(s, Len(s) - 4) & ".jpg" 
                ToPath = "/UploadFiles/image/Product/" & title & ".jpg" 
                If checkFile(ToPath) = False Then Call CopyFile(path, ToPath) 
                rsx("BigFiles") = ToPath 
                content = "" 
                'Content = "<div style='text-align:center; border:1px double #CCCCCC; padding:4px;margin-bottom:10px;'><img src='"& ToPath &"' onload='javascript:if(this.width>720) this.width=720;'></div>"
                filePath = DataFolder & "Product\数据\产品\" & s & ".txt" 
                If checkFile(filePath) = False Then filePath = DataFolder & "Product\数据\产品\默认培训.txt" 
                content = content & Replace(Replace(GetFText(filePath), vbCrLf, "<br>"), "    ", "&nbsp;&nbsp;&nbsp;&nbsp;") 
                rsx("Content") = content 
                rsx.Update 
                'Call Echo("FilePath",FilePath)
                rsx.Close 
            'Call Echo(S,Title)
            End If 
        Next 
    End If 
    '批量导入产品信息(20151008)
    Call XY_AutoBatchAddArticle("FolderPath='" & DataFolder & "\自动批量添加文章信息' ") 

    Call echo("自动生成数据文件夹路径", DataFolder) 
    Call rw(GetTimer()) 
    'Call AddLaoLuo()    '符加测试用
    '##############################################

    If Request("WebDebug") = "True" Then
        Response.Redirect("/Inc/Create_Html.Asp?act=Nav&NavType=首页&Debug=1&MackHtml=False") 
    End If 
End Sub 




'导入图片文件夹 20140302
Sub leadingImage(did, sid, tid, str)
    Dim path, content, splStr, i, s, c, nCount, ImgPath, title 
    path = Trim(Replace(str, "#导入文件夹#", "")) 
    content = GetDirFileList(path,"") 
    splStr = Split(content, vbCrLf) : nCount = 0 
    For Each s In splStr
        If s <> "" Then
            nCount = nCount + 1 
            s = Replace(s, Server.MapPath("/"), "") 
            ImgPath = "/UploadFiles/image/" & Mid(s, InStrRev(s, "\") + 1) 
            If checkFile(ImgPath) = False Then Call CopyFile(s, ImgPath) 
            rs.Open "Select * From [Product]", conn, 1, 3 
            rs.AddNew 
            title = GetFilePathName(s) 
            rs("BigClassName") = did 
            rs("SmallClassName") = sid 
            rs("ThreeClassName") = tid 
            rs("Title") = title 
            rs("BigFiles") = ImgPath 
            rs("FileName") = AotuGetProductFileName(title) 
            rs.Update : rs.Close 
        End If 
    Next 

End Sub 



'导入新闻文件夹 20140313
Sub leadingNews(did, sid, tid, str)
    Dim path, content, splStr, i, s, c, nCount, title 
    path = Trim(Replace(str, "#导入新闻文件夹#", "")) 
    content = GetDirTxtList(path) 
    splStr = Split(content, vbCrLf) : nCount = 0 
    For Each s In splStr
        If s <> "" Then
            nCount = nCount + 1 
            s = Replace(s, Server.MapPath("/"), "") 
            rs.Open "Select * From [Product]", conn, 1, 3 
            rs.AddNew 
            title = GetFilePathName(s) 
            rs("BigClassName") = did 
            rs("SmallClassName") = sid 
            rs("ThreeClassName") = tid 

            rs("Title") = title 
            rs("FileName") = AotuGetProductFileName(title) 

            content = Replace(GetFText(s), vbCrLf, "<br>") 
            rs("Content") = content 
            rs.Update : rs.Close 
        End If 
    Next 

End Sub 



'导入产品类别文件夹 20140321
Sub leadingProductClass(did, sid, tid, folderPath)
    Dim path, content, splStr, i, s, c, nCount, title, FolderList 
    FolderList = GetDirFolderList(Trim(folderPath)) 
    'Call Echo("FolderList",FolderList)
    If tid = "" Then
        splStr = Split(FolderList, vbCrLf) 
        For Each path In splStr
            If path <> "" Then
                tid = GetFilePathName(path) 
                'Call Echo(Path,Tid)
                Call LeadingImage(did, sid, tid, path) 
            End If 
        Next 
    End If 
End Sub 



'给指定类加个默认产品
Sub classToAddProduct(DataFolderPath, did, sid, tid)
    'on error resume next
    Dim content, addSql 
    'Call Eerr("",MdbPath)
    '大类
    If did <> "" Then
        addSql = " Where BigClassName='" & did & "' " 
    End If 
    If TableCount("SmallClass", addSql) = 0 Then
        rs.Open "Select * From [BigClass] " & addSql, conn, 1, 1 
        While Not rs.EOF
            rsx.Open "Select * From [Product] Where BigClassName='" & rs("BigClassName") & "'", conn, 1, 3 
            If rsx.EOF Then
                rsx.AddNew 
                rsx("BigClassName") = rs("BigClassName") 
                rsx("Title") = rs("BigClassName") 
                content = GetFText(DataFolderPath & "/Product/数据/默认文章.txt") 
                content = content & Replace(Replace(content, vbCrLf, "<br>"), "    ", "&nbsp;&nbsp;&nbsp;&nbsp;") 
                rsx("Content") = content 
                rsx("FileName") = AotuGetProductFileName(rs("BigClassName")) 
                rsx.Update 
            End If : rsx.Close 
        rs.MoveNext : Wend : rs.Close 
    End If 
    '小类
    If sid <> "" Then
        addSql = " And SmallClassName='" & sid & "' " 
    End If 

    If TableCount("ThreeClass", addSql) = 0 Then
        'Call Echo("AddSql",AddSql):Doevents
        rs.Open "Select * From [SmallClass] ", addSql, conn, 1, 1 
        While Not rs.EOF
            rsx.Open "Select * From [Product] Where BigClassName='" & rs("BigClassName") & "' And SmallClassName='" & rs("SmallClassName") & "'", conn, 1, 3 
            If rsx.EOF Then
                rsx.AddNew 
                rsx("BigClassName") = rs("BigClassName") 
                rsx("SmallClassName") = rs("SmallClassName") 
                rsx("Title") = rs("SmallClassName") 
                content = GetFText(DataFolderPath & "/Product/数据/默认文章.txt") 
                content = content & Replace(Replace(content, vbCrLf, "<br>"), "    ", "&nbsp;&nbsp;&nbsp;&nbsp;") 
                rsx("Content") = content 
                rsx("FileName") = AotuGetProductFileName(rs("SmallClassName")) 
                rsx.Update 
            End If : rsx.Close 
        rs.MoveNext : Wend : rs.Close 
    End If 
    '子类
    If tid <> "" Then
        addSql = " And ThreeClassName='" & tid & "' " 
    End If 
    rs.Open "Select * From [ThreeClass] " & addSql, conn, 1, 1 
    While Not rs.EOF
        rsx.Open "Select * From [Product] Where BigClassName='" & rs("BigClassName") & "' And SmallClassName='" & rs("SmallClassName") & "' And ThreeClassName='" & rs("ThreeClassName") & "'", conn, 1, 3 
        If rsx.EOF Then
            rsx.AddNew 
            rsx("BigClassName") = rs("BigClassName") 
            rsx("SmallClassName") = rs("SmallClassName") 
            rsx("ThreeClassName") = rs("ThreeClassName") 
            rsx("Title") = rs("ThreeClassName") 
            content = GetFText(DataFolderPath & "/Product/数据/默认文章.txt") 
            content = content & Replace(Replace(content, vbCrLf, "<br>"), "    ", "&nbsp;&nbsp;&nbsp;&nbsp;") 
            rsx("Content") = content 
            rsx("FileName") = AotuGetProductFileName(rs("ThreeClassName")) 
            rsx.Update 
        End If : rsx.Close 
    rs.MoveNext : Wend : rs.Close 
End Sub 

'VB内容结束



'Call AyawawaSina()
'Call Mydd3():Response.End()
'处理文章内容 世纪佳缘
Sub ayawawa()
    Response.AddHeader "Content-Type", "text/html; charset=gb2312" 
    Dim splStr, i, s, c, content, title, path, Char_Set, fileName, UpDateTime, SaveFolder, ToPath, nOK 
    SaveFolder = "E:\E盘\需要文件\ayawawa\" 
    '保存文件夹是否存在，不存在创建
    If CheckFolder(SaveFolder) = False Then
        Call CreateFolder(SaveFolder) 
    End If 
    content = GetDirFileList("E:\E盘\WEB网站\至前网站\函数\Tools\DataDir\http：揦揦diary。jiayuan。com揦\","") 
    nOK = nOK 
    splStr = Split(content, vbCrLf) 
    For Each path In splStr
        If path <> "" Then
            Char_Set = CheckCode(path) 
            c = ReadFile(path, 0) 
            '只采内容页
            If InStr(c, "_杨冰阳_世纪佳缘交友网</title>") > 0 Then
                title = StrCut(c, "<title>", "_杨冰阳_", 2) 

                UpDateTime = StrCut(c, "<!-- JiaThis Button END -->" & vbCrLf & "<span>", "</span></h5>", 2) 
                UpDateTime = Format_Time(UpDateTime, 2) 

                fileName = UpDateTime & " " & SetFileName(title) & ".txt" 
                content = StrCut(c, "<div class=""txt"">", "<span class=""jiathis_txt"">分享到</span>", 2) 
                content = Replace(DelHTML(content), "&nbsp;", " ") 
                content = DelHtmlNote(content)                                                  '删除HTML注释
                content = unHTMLCode(content) 
                ToPath = SaveFolder & fileName 
                Call CreateFile(ToPath, content) 
                nOK = nOK + 1 
                Call echo("保存成功【" & nOK & "】", ToPath) 
            End If 
            doEvents 
        End If 

    Next 
    Response.End() 
End Sub 



'处理掉分行中的左右空格 ayawawa
Sub handRowSpace()
    Response.AddHeader "Content-Type", "text/html; charset=gb2312" 
    Dim splStr, i, s, c, content, title, path, Char_Set, fileName, UpDateTime, SaveFolder, ToPath, nOK, splxx 
    SaveFolder = Server.MapPath("ayawawa\") & "\" 
    '保存文件夹是否存在，不存在创建
    If CheckFolder(SaveFolder) = False Then
        Call CreateFolder(SaveFolder) 
    End If 
    content = GetDirFileList("E:\E盘\WEB网站\至前网站\ayawawa","") 

    splStr = Split(content, vbCrLf) 
    For Each path In splStr
        If path <> "" Then
            Char_Set = CheckCode(path) 
            c = ReadFile(path, 1) 
            'C = RegExp_Replace(C, "(\n\s*)|(\s*$)", vbCrlf)        '去除每行两边空格
            c = unHTMLCode(c) 
            splxx = Split(c, vbCrLf) : c = "" 
            For Each s In splxx
                c = c & Trim(s) & vbCrLf 
            Next 

            fileName = Mid(path, InStrRev(path, "\")) 
            Call echo("创建成功", fileName) 
            ToPath = SaveFolder & fileName 

            Call CreateFile(path, c) 
            doEvents 
        End If 

    Next 
    Response.End() 
End Sub 



'处理文章内容 新浪
Sub ayawawaSina()
    Response.AddHeader "Content-Type", "text/html; charset=gb2312" 
    Dim splStr, i, s, c, content, title, path, Char_Set, fileName, UpDateTime, SaveFolder, ToPath, nOK 
    SaveFolder = "E:\E盘\需要文件\Ayawawa_Sina\" 
    '保存文件夹是否存在，不存在创建
    If CheckFolder(SaveFolder) = False Then
        Call CreateFolder(SaveFolder) 
    End If 
    content = GetDirFileList("E:\E盘\WEB网站\至前网站\函数\Tools\DataDir\http：揦揦blog。sina。com。cn揦\","") 
    nOK = nOK 
    splStr = Split(content, vbCrLf) 
    For Each path In splStr
        If path <> "" Then
            Char_Set = CheckCode(path) 
            c = ReadFile(path, 0) 
            title = StrCut(c, "<title>", "</title>", 2) 

            If InStr(title, "_ayawawa_") > 0 Then
                title = Mid(title, 1, InStr(title, "_ayawawa_") - 1) 
            End If 
            If InStr(title, "娃娃微问答") > 0 Or InStr(title, "娃娃微信答") > 0 Then
                title = Replace(title, "&nbsp;", " ") 
                title = Mid(title, 11) 
            'Call eerr("Title",Title)
            End If 
            If InStr(title, "答读者来信") > 0 Then
                title = Mid(title, 7) 
            End If 
            If InStr(title, "【来信】") > 0 Then
                title = Mid(title, 5) 
            End If 

            If InStr(c, "<span class=""time SG_txtc"">(") = False Then
                Call echo(path, "没有时间可截取") 
            Else
                UpDateTime = StrCut(c, "<span class=""time SG_txtc"">(", ")</span>", 2) 
                UpDateTime = Format_Time(UpDateTime, 2) 

                fileName = UpDateTime & " " & SetFileName(title) & ".txt" 
                content = StrCut(c, "<!-- 正文开始 -->", "<!-- 正文结束 -->", 2) 
                content = Replace(DelHTML(content), "&nbsp;", " ") 
                content = DelHtmlNote(content)                                                  '删除HTML注释
                content = unHTMLCode(content) 
                '去掉更多文章内容
                If InStr(content, "我的更多文章：") > 0 Then
                    content = Mid(content, 1, InStrRev(content, "我的更多文章：") - 8) 
                End If 

                ToPath = SaveFolder & fileName 
                Call CreateFile(ToPath, content) 
                nOK = nOK + 1 
                Call echo("保存成功【" & nOK & "】", ToPath) 
            End If 
            doEvents 
        End If 

    Next 
    Response.End() 
End Sub 


'老罗语录 加入数据库
Sub addLaoLuo()
    Dim path, splStr, splxx, i, s, c, content, did, sid, tid, title, IDSort 
    Dim LaoLuoYuLu, LaoLuoHuiDa 
    path = "E:\E盘\需要文件\学习\好文章\老罗语录.txt" 
    content = GetFText(path) 
    splStr = Split(content, "====== 回答记者问 ======") 
    LaoLuoYuLu = splStr(0) 
    LaoLuoHuiDa = splStr(1) 


    splStr = Split(LaoLuoYuLu, vbCrLf) 
    did = "罗永浩" : sid = "语录" 
    IDSort = conn.Execute("Select Count(*) From [Product] Where BigclassName='" & did & "' And SmallClassName='" & sid & "' ")(0) 
    For Each content In splStr
        content = Trim(content) 
        If content <> "" Then
            rs.Open "Select * From [Product] Where BigclassName='" & did & "' And SmallClassName='" & sid & "' And Content='" & content & "'", conn, 1, 3 
            If rs.EOF Then
                rs.AddNew 
                rs("BigclassName") = did 
                rs("SmallClassName") = sid 
                IDSort = IDSort + 1 
                rs("Title") = IDSort 
                rs("Sort") = IDSort 
                rs("Content") = content 
                Call echo(title, "添加成功") 
                rs.Update 
            End If : rs.Close 
        End If 
    Next 

    splStr = Split(LaoLuoHuiDa, vbCrLf & vbCrLf) 
    did = "罗永浩" : sid = "语录" 
    IDSort = conn.Execute("Select Count(*) From [Product] Where BigclassName='" & did & "' And SmallClassName='" & sid & "' ")(0) 
    For Each content In splStr
        content = Trim(content) 
        If Replace(content, vbCrLf, "") <> "" Then
            rs.Open "Select * From [Product] Where BigclassName='" & did & "' And SmallClassName='" & sid & "' And Content='" & content & "'", conn, 1, 3 
            If rs.EOF Then
                rs.AddNew 
                rs("BigclassName") = did 
                rs("SmallClassName") = sid 
                IDSort = IDSort + 1 
                rs("Title") = IDSort 
                rs("Sort") = IDSort 
                rs("Content") = content 
                Call echo(title, "添加成功") 
                rs.Update 
            End If : rs.Close 
        End If 
    Next 
End Sub 



%>   

