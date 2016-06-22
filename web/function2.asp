<% 


'调用function2文件函数
Function callFunction2()
    Select Case Request("stype")
        Case "runScanWebUrl" : runScanWebUrl()                                          '运行扫描网址
        Case "scanCheckDomain" : scanCheckDomain()                                      '检测域名有效
        Case "bantchImportDomain" : bantchImportDomain()                                '批量导入域名
		case "scanDomainHomePage" : scanDomainHomePage()								'扫描域名首页
		case "scanDomainHomePageSize" : scanDomainHomePageSize()								'扫描域名首页大小与标题
		case "isthroughTrue" : isthroughTrue()											'让审核全部为真
		case "printOKWebSite" : printOKWebSite()										'打印有效网址
		case "printAspServerWebSite" : printAspServerWebSite()										'打印asp类型网站
		
		case "clearAllData" : clearAllData()										'清除全部数据

		case "function2test" : function2test()											'测试
        Case Else : Call eerr("function2页里没有动作", Request("stype"))
    End Select
End Function

'测试
function function2test()
    Call openconn() 									
    rs.Open "select * from " & db_PREFIX & "webdomain where isdomain=true", conn, 1, 1
	call echo("共",rs.recordcount)
    While Not rs.EOF
		call echo(rs("isdomain"),rs("website"))
	rs.movenext:wend:rs.close
end function
'清除全部数据
function clearAllData()
    Call openconn() 
    conn.Execute("delete from " & db_PREFIX & "webdomain") 
    Call echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK</a>") 
end function
'打印有效网址
function printOKWebSite()	
    Call openconn()
    rs.Open "select * from " & db_PREFIX & "webdomain where isdomain=true", conn, 1, 1 
	call echo("共",rs.recordcount)
    Call echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK</a>") 
    While Not rs.EOF
		'call echo(rs("isdomain"),rs("website"))
		call rw(rs("website") & "<br>")
	rs.movenext:wend:rs.close
end function
'打印asp类型网站
function printAspServerWebSite()	
    Call openconn()
    rs.Open "select * from " & db_PREFIX & "webdomain where isasp=true and (isaspx=false and isphp=false)", conn, 1, 1 
	call echo("共",rs.recordcount)
    Call echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK</a>") 
    While Not rs.EOF
		'call echo(rs("isdomain"),rs("website"))
		call rw(rs("website") & "<br>")
	rs.movenext:wend:rs.close
end function

'让审核全部为真
function isthroughTrue()
    Call openconn() 
    conn.Execute("update " & db_PREFIX & "webdomain set isthrough=true") 
    Call echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK</a>") 
end function

'扫描首页大小
Function scanDomainHomePageSize()
    Dim url, nSetTime, isdomain, htmlDir, txtFilePath,homePageList,nThis,nCount
	dim splstr,s,c,website,nState,websize,content,startTime,webtitle,webkeywords,webdescription
	 
	nThis=request("nThis")
	if nThis="" then
		nThis=0
	else
		nThis=cint(nThis)
	end if
	
    nSetTime = 3 
    Call openconn() 
    rs.Open "select * from " & db_PREFIX & "webdomain where website<>'' and websize=0 and isdomain=true", conn, 1, 1 
	nCount=request("nCount")
	if nCount="" then
		nCount=rs.recordcount
	end if
    While Not rs.EOF
		nThis=nThis+1 
        Call echo(nThis & "/" & nCount, rs("website")) 
        doevents 
        htmlDir = "/../网站UrlScan/域名首页大小/"
        Call createDirFolder(htmlDir) 
        txtFilePath = htmlDir & "/" & setFileName(rs("website")) & ".txt"
        If checkFile(txtFilePath) = True Then 
            Call echo("类型", "本地") 
			nSetTime=1
        Else
			website=getwebsite(rs("website"))
			if website="" then
				call eerr("域名为空",httpurl)
			end if
		 	content=getHttpPage(website,rs("charset"))
	        '【PHP】$content=toGB2312Char($content);                                            //给PHP用，转成gb2312字符
			if content="" then
				content=" "
			end if
			
            Call createFile(txtFilePath, content)
            Call echo("类型", "网络") 
        End If 
		content=getftext(txtFilePath)
		webtitle=getHtmlValue(content,"webtitle")
		webkeywords=getHtmlValue(content,"webkeywords")
		webdescription=getHtmlValue(content,"webdescription")
		
		
		websize=getfsize(txtFilePath)
		call echo("webtitle",webtitle)
        '这样写是给转PHP时方便
        conn.Execute("update " & db_PREFIX & "webdomain  set webtitle='"& ADSql(webtitle) &"',webkeywords='"& webkeywords &"',webdescription='"& webdescription &"',websize="& websize &",isthrough=false,updatetime='" & Now() & "'  where id=" & rs("id") & "")

		startTime=request("startTime")
		if startTime="" then
			startTime=now()
		end if
		
		call rw(VBRunTimer(startTime) & "<hr>")
        url = getUrlAddToParam(getThisUrl(), "?nThis="& nThis &"&nCount="& nCount &"&startTime="& startTime &"&N=" & getRnd(11), "replace") 
 
        Call rw(jsTiming(url, nSetTime)) 
        Response.End() 
    rs.MoveNext : Wend : rs.Close 
    Call echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK，共("& nThis &")条</a>") 
End Function 

'扫描域名首页
Function scanDomainHomePage()
    Dim url, nSetTime, isdomain, htmlDir, txtFilePath,homePageList,nThis,nCount
	dim splstr,s,c,website,nState
	dim isAsp,isAspx,isPhp,isJsp,c2
	isAsp=0:isAspx=0:isPhp=0:isJsp=0
	nThis=request("nThis")
	if nThis="" then
		nThis=0
	else
		nThis=cint(nThis)
	end if
	
    nSetTime = 3 
    Call openconn() 
    rs.Open "select * from " & db_PREFIX & "webdomain where website<>'' and homepagelist='' and isdomain=true", conn, 1, 1 
	nCount=request("nCount")
	if nCount="" then
		nCount=rs.recordcount
	end if
    While Not rs.EOF
		nThis=nThis+1 
        Call echo(nThis & "/" & nCount, rs("website")) 
        doevents 
        htmlDir = "/../网站UrlScan/域名首页/"
        Call createDirFolder(htmlDir) 
        txtFilePath = htmlDir & "/" & setFileName(rs("website")) & ".txt"
        If checkFile(txtFilePath) = True Then
            c = phptrim(getFText(txtFilePath))
			isAsp=getstrcut(c,"isAsp=",vbcrlf,1)
			isAspx=getstrcut(c,"isAspx=",vbcrlf,1)
			isPhp=getstrcut(c,"isPhp=",vbcrlf,1)
			isJsp=getstrcut(c,"isJsp=",vbcrlf,1)
            Call echo("类型", "本地") 
			nSetTime=1
        Else
			website=getwebsite(rs("website"))
			if website="" then
				call eerr("域名为空",httpurl)
			end if
			splstr=Array("index.asp","index.aspx","index.php","index.jsp","index.htm","index.html","default.asp","default.aspx","default.jsp","default.htm","default.html")
			c2=""
			homePageList=""
			for each s in splstr
				url=website & s
				nState=getHttpUrlState(url)
				call echo(url,nState & "   ("& getHttpUrlStateAbout(nState) &")")
				doevents
				if (s="index.asp" or s="default.asp") and (nState="200" or nState="302") then
					isAsp=1
				elseif (s="index.aspx" or s="default.aspx") and (nState="200" or nState="302") then
					isAspx=1
				elseif (s="index.php" or s="default.php") and (nState="200" or nState="302") then
					isPhp=1
				elseif (s="index.jsp" or s="default.jsp") and (nState="200" or nState="302") then
					isJsp=1
				end if
				if nState="200" or nState="302" then
					homePageList=homePageList & s & "|"
				end if
				c2=c2 & s & "=" & nState & vbcrlf
			next
			c= "isAsp=" & isAsp & vbcrlf
			c= c & "isAspx=" & isAspx & vbcrlf
			c= c & "isPhp=" & isPhp & vbcrlf
			c= c & "isJsp=" & isJsp & vbcrlf & c2
			
			if homePageList="" then
				homePageList="无"
			end if
			
            Call createFile(txtFilePath, c) 
            Call echo("类型", "网络") 
        End If 
        '这样写是给转PHP时方便
        conn.Execute("update " & db_PREFIX & "webdomain  set isasp="& isAsp &",isaspx="& isAspx &",isphp="& isPhp &",isjsp="& isJsp &",isthrough=false,homepagelist='"& homePageList &"',updatetime='" & Now() & "'  where id=" & rs("id") & "") 

		startTime=request("startTime")
		if startTime="" then
			startTime=now()
		end if
		
		call rw(VBRunTimer(startTime) & "<hr>")
        url = getUrlAddToParam(getThisUrl(), "?nThis="& nThis &"&nCount="& nCount &"&startTime="& startTime &"&N=" & getRnd(11), "replace") 
 
        Call rw(jsTiming(url, nSetTime)) 
        Response.End() 
    rs.MoveNext : Wend : rs.Close 
    Call echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK，共("& nThis &")条</a>") 
End Function 

'批量导入域名
Function bantchImportDomain()
    Dim content, splStr, url, webSite, nOK 
    content = LCase(Request.Form("bodycontent")) 
    splStr = Split(content, vbCrLf) 
    nOK = 0 
    Call openconn() 
    For Each url In splStr
        webSite = getwebsite(url) 
        If webSite <> "" Then
            rs.Open "select * from " & db_PREFIX & "webdomain where website='" & webSite & "'", conn, 1, 1 
            If rs.EOF Then
                conn.Execute("insert into " & db_PREFIX & "webdomain(website,isthrough,isdomain) values('" & webSite & "',true,false)") 
                Call echo("添加成功", webSite) 
                nOK = nOK + 1 
            Else
                Call echo("website", webSite) 
            End If : rs.Close 
        End If 
    Next 
    Call echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK 共(" & nOK & ")条</a>") 
End Function 

'检测域名有效
Function scanCheckDomain()
    Dim url, nSetTime, isdomain, htmlDir, txtFilePath, nThis,nCount,startTime
    nSetTime = 3
	nThis=request("nThis")
	if nThis="" then
		nThis=0
	else
		nThis=cint(nThis)
	end if
    Call openconn()
    rs.Open "select * from " & db_PREFIX & "webdomain where isthrough=true", conn, 1, 1 
	nCount=request("nCount")
	if nCount="" then
		nCount=rs.recordcount
	end if
    While Not rs.EOF
		nThis=nThis+1
        Call echo(nThis & "/" & nCount, rs("website")) 
        doevents 
        htmlDir = "/../网站UrlScan/域名/"
        Call createDirFolder(htmlDir) 
        txtFilePath = htmlDir & "/" & setFileName(rs("website")) & ".txt"
        If checkFile(txtFilePath) = True Then
            isdomain = phptrim(getFText(txtFilePath)) 
            Call echo("类型", "本地") 
			nSetTime=1
        Else
            isdomain = IIF(checkDomainName(rs("website")), 1, 0) 
            Call createFile(txtFilePath, isdomain & " ") 			'防止PHP版写入不进去 0 这个内容
            Call echo("类型", "网络" & txtFilePath & "("& checkFile(txtFilePath) &")=" & isdomain) 
        End If 
        '这样写是给转PHP时方便
        conn.Execute("update " & db_PREFIX & "webdomain  set isthrough=false,isdomain=" & isdomain & ",updatetime='" & Now() & "'  where id=" & rs("id") & "") 

		startTime=request("startTime")
		if startTime="" then
			startTime=now()
		end if
		
		call rw(VBRunTimer(startTime) & "<hr>")
        url = getUrlAddToParam(getThisUrl(), "?nThis="& nThis &"&nCount="& nCount &"&startTime="& startTime &"&N=" & getRnd(11), "replace") 
		
        Call rw(jsTiming(url, nSetTime)) 
        Response.End() 
    rs.MoveNext : Wend : rs.Close 
    Call echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK，共("& nThis &")条</a>") 
End Function 

'扫描网址
Function runScanWebUrl()
    Dim nSetTime, setCharSet, httpUrl, url, selectWeb ,nThis,nCount,startTime
    setCharSet = "gb2312"                                                           'gb2312
    'http://www.dfz9.com/
    'http://www.maiside.net/
    'http://sharembweb.com/
    'http://www.ufoer.com/
    httpUrl = "http://sharembweb.com/" 
    'selectWeb="ufoer"
    If selectWeb = "ufoer" Then
        httpUrl = "http://www.ufoer.com/" 
        setCharSet = "utf-8" 
    End If
	
	nThis=request("nThis")
	if nThis="" then
		nThis=0
	else
		nThis=cint(nThis)
	end if
	
    Call openconn() 
    rs.Open "select * from " & db_PREFIX & "weburlscan", conn, 1, 1 
	nCount=request("nCount")
	if nCount="" then
		nCount=rs.recordcount
	end if
    If rs.EOF Then
        conn.Execute("insert into " & db_PREFIX & "weburlscan(httpurl,title,isthrough,charset) values('" & httpUrl & "','home',true,'" & setCharSet & "')") 
    End If : rs.Close 
    '循环
    rsx.Open "select * from " & db_PREFIX & "weburlscan where isThrough=true", conn, 1, 1 
    If Not rsx.EOF Then
		nThis=nThis+1
        Call echo(nThis,  rsx("httpurl"))  
        doevents 
        nSetTime = scanUrl(rsx("httpurl"), rsx("title"), rsx("charset")) 
        '这样写是给转PHP时方便
        conn.Execute("update " & db_PREFIX & "weburlscan  set isthrough=false  where id=" & rsx("id") & "") 

		startTime=request("startTime")
		if startTime="" then
			startTime=now()
		end if
		
		call VBRunTimer(startTime)
        url = getUrlAddToParam(getThisUrl(), "?nThis="& nThis &"&nCount="& nCount &"&startTime="& startTime &"&N=" & getRnd(11), "replace") 

        Call rw(jsTiming(url, nSetTime)) 
        Response.End()
    End If : rsx.Close  
    Call echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebUrlScan&addsql=order by id desc&lableTitle=网址扫描'>OK，共("& nThis &")条</a>") 
    '输入报告
    rs.Open "select * from " & db_PREFIX & "weburlscan where webstate=404", conn, 1, 1 
    While Not rs.EOF
        Call echo("<a href='" & rs("httpurl") & "' target='_blank'>" & rs("httpurl") & "</a>", "<a href='" & rs("tohttpurl") & "' target='_blank'>" & rs("tohttpurl") & "</a>") 
    rs.MoveNext : Wend : rs.Close 
End Function 
'扫描网址
Function scanUrl(httpUrl, toTitle, codeset)
    Dim splStr, i, s, content, PubAHrefList, PubATitleList, splUrl, spltitle, title, url, htmlDir, htmlFilePath, nOK, dataArray, webState, u, iniDir, iniFilePath ,websize
    Dim nSetTime, startTime, openSpeed, isLocal, isThrough
    htmlDir = "/../网站UrlScan/" & setFileName(getwebsite(httpUrl)) 
    Call createDirFolder(htmlDir) 
    htmlFilePath = htmlDir & "/" & setFileName(httpUrl) & ".html" 
    iniDir = htmlDir & "/conifg" 
    Call createfolder(iniDir) 
    iniFilePath = iniDir & "/" & setFileName(httpUrl) & ".txt" 

    'httpurl="http://maiside.net/"

    webState = 0 
    nSetTime = 1 
    openSpeed = 0 
    If checkFile(htmlFilePath) = False Then
        startTime = Now() 
        Call echo("codeset", codeset) 
        dataArray = handleXmlGet(httpUrl, codeset) 
        content = dataArray(0) 
        '【PHP】$content=toGB2312Char($content);                                            //给PHP用，转成gb2312字符

        webState = dataArray(1) 
        openSpeed = DateDiff("s", startTime, Now()) 
        'content=gethttpurl(httpurl,codeset)
        'call createfile(htmlFilePath,content)
        Call writeToFile(htmlFilePath, content, codeset) 
        Call createFile(iniFilePath, webState & vbCrLf & openSpeed) 
        nSetTime = 3 
        isLocal = 0 
    Else
        'content=getftext(htmlFilePath)
        content = readFile(htmlFilePath, codeset) 
		 '【PHP】$content=toGB2312Char($content);                                            //给PHP用，转成gb2312字符
        splStr = Split(getftext(iniFilePath), vbCrLf) 
        webState = CInt(splStr(0)) 
        openSpeed = CInt(splStr(0)) 
        isLocal = 1 
    End If 
	websize=getFSize(htmlFilePath)
	if websize="" then
		websize=0
	end if
    Call echo("isLocal", isLocal) 
    rs.Open "select * from " & db_PREFIX & "weburlscan where httpurl='" & httpUrl & "'", conn, 1, 1 
    If rs.EOF Then
        conn.Execute("insert into " & db_PREFIX & "weburlscan(httpurl,title,charset) values('" & httpUrl & "','" & toTitle & "','" & codeset & "')") 
    End If : rs.Close 
    conn.Execute("update " & db_PREFIX & "weburlscan  set webstate=" & webState & ",websize=" & websize & ",openspeed=" & openSpeed & ",charset='" & codeset & "'  where httpurl='" & httpUrl & "'") 

	'strLen(content)  不用这个，不精准

    s = getContentAHref("", content, PubAHrefList, PubATitleList) 
    s = handleScanUrlList(httpUrl, s) 

    'call echo("httpurl",httpurl)
    'call echo("s",s)
    'call echo("PubATitleList",PubATitleList)
    nOK = 0 
    splUrl = Split(PubAHrefList, vbCrLf) 
    spltitle = Split(PubATitleList, vbCrLf) 
    For i = 1 To UBound(splUrl)
        title = spltitle(i) 
        url = splUrl(i) 
        '去掉#号后台的字符20160506
        If InStr(url, "#") > 0 Then
            url = Mid(url, 1, InStr(url, "#") - 1) 
        End If 
        If url = "" Then
            If title <> "" Then
                Call echo("网址为空", title) 
            End If 
        Else
            url = handleScanUrlList(httpUrl, url) 
            url = handleWithWebSiteList(httpUrl, url) 
            If url <> "" Then
                rs.Open "select * from " & db_PREFIX & "weburlscan where httpurl='" & url & "'", conn, 1, 1 
                If rs.EOF Then
                    u = LCase(url)
                    If InStr(u, "tools/downfile.asp") > 0 Or InStr(u, "/url.asp?") > 0 Or InStr(u, "/aspweb.asp?") > 0 Or InStr(u, "/phpweb.php?") > 0 Or u = "http://www.maiside.net/qq/" Or InStr(u, "mailto:") > 0 Or InStr(u, "tel:") > 0 Or InStr(u, ".html?replytocom") > 0 Then'.html?replytocom  王通网站
                        isThrough = 0 
                    Else
                        isThrough = 1 '不用true 因为写入数据会有问题
                    End If 
                    conn.Execute("insert into " & db_PREFIX & "weburlscan(tohttpurl,totitle,httpurl,title,isthrough,charset) values('" & httpUrl & "','" & toTitle & "','" & url & "','" & Left(title, 255) & "'," & isThrough & ",'" & codeset & "')") 
                    nOK = nOK + 1 
                    Call echo(i, url) 
                Else
                    Call echo(title, url) 
                End If : rs.Close 
            End If 
        End If 
    Next
	
    scanUrl = nSetTime 
End Function 


%>                  

