<%
'网络操作 网址      简单简单，越简单越好

Class NetWorkClass
	'网络网址处理类
	Dim MDBPath, DatabaseType
	Dim Conn
	Dim Rs
	Dim Rss
	Dim Rst
	Dim Rsx
	Dim Rsd
	Dim TempRs
	Dim RsTemp
	Dim SplSpecialUrl                   '分割无需处理网址
	Dim CacheDir                        '缓存目录路径
	
	dim WebSite							'域名
	dim SourceUrlList					'源网址列表
	dim CleanUrlList					'干净网址列表	
	dim IdenticalUrlList				'相同网址列表
	dim DifferentUrlList				'不同网址列表
	dim DifferentParameterUrlList		'不同参数网址列表
	dim SqlInUrlList					'可注入网址列表
	dim WebState						'网站状态
	dim WebFileSize						'请求文件大小
	dim pubHttpUrl						'当前网址
	dim pubWebTitle						'网站标题
	dim pubWebSite						'网站名称
	dim isReadCacheFile					'是否读缓冲文件
	dim Depth							'深度
	

	'获得网址内容中网址列表  CustomWebSite 为域名判断，为空则自动提取域名
	Function GetHttpUrlContentUrlList(ByVal HttpUrl,setCode,CustomWebSite,toUrl,toTitle)
		Dim Content,HttpArrar,saveFolder,saveFilePath
		HttpUrl = HandleUrlComplete(HttpUrl)                            '处理网址完整性

		HttpUrl = HandleSpecialUrl(HttpUrl)                             '处理无需处理网址 如新浪百度等
		HttpUrl = HandleInvalidUrl(HttpUrl)                             '处理无效网址 如javascript:;
		HttpUrl = RemoveNonWebpage(HttpUrl)                             '屏蔽非网页网址如.jpg.gif不需要
 	
		if HttpUrl<>"" then
			if CustomWebSite<>"" then
				WebSite = CustomWebSite		
			else
				WebSite=getWebSite(httpUrl)
			end if
			'Content = getHttpUrl(HttpUrl,"")                              	   '下载网址中内容本地存在则直接读出
			
			if isReadCacheFile=true then			
				saveFolder = CacheDir & setFileName(getWebSite(HttpUrl)) & "/"
				call createFolder(saveFolder)
				saveFilePath = saveFolder & setFileName(HttpUrl)
				
				if checkfile(saveFilePath)=true then
					HttpArrar=split(getFText(saveFilePath & ".txt"),vbCrlf)
					HttpArrar(0) = getFText(saveFilePath)
					toUrl = HttpArrar(2)											    '状态
					toTitle = HttpArrar(3)											    '状态
				end if
			end if
			if IsEmpty(HttpArrar) then
				HttpArrar = handleXmlGet(HttpUrl, setCode)
					call echo(HttpUrl,saveFilePath):doevents							'显示下载网址信息
					call createFile(saveFilePath, HttpArrar(0))
					call createFile(saveFilePath & ".txt", vbcrlf & HttpArrar(1) & vbcrlf & toUrl & vbcrlf & toTitle )
			end if
			Content = HttpArrar(0)												'内容
			WebState = HttpArrar(1)											    '状态
			WebFileSize = stringLength(Content)								    '文件大小
			pubHttpUrl=HttpUrl													'网址网址
			pubWebTitle = RegExpGetStr("<TITLE>([^<>]*)</TITLE>", Content, 1)	'网站标题
			pubWebSite = getWebSite(HttpUrl)									'网站域名
			
			'call rwend(Content)
			if 1=2 then
			Content = GetContentAHref(HttpUrl, Content)                         '获得内容里链接网址 第一种（自己写）
			else
			Content = GetAUrlTitleList(Content,"")                         		 '获得内容里链接网址 第二种（正则）
			Content = batchFullHttpUrl(WebSite,Content)							 '批量处理网址完成 第二种需要
			'content = handleFullHttpUrl(WebSite,Content)						 '处理网址完整地址
			end if
			
			
			
			SourceUrlList = Content		'源网址列表
			Content = HandleSpecialUrl(Content)                                 '处理无需处理网址 如新浪百度等
			Content = HandleInvalidUrl(Content)                                 '处理无效网址 如javascript:;
			Content = RemoveNonWebpage(Content)                                 '屏蔽非网页网址如.jpg.gif不需要
			Content = RemoveWidthUrl(Content)                                 	'去除同相网址，超强版
			
			
			CleanUrlList = content		'干净网址列表
			
			IdenticalUrlList = GetIdenticalWebSiteUrlList(WebSite,content)		'相同域名网址列表
			DifferentUrlList = GetDifferentWebSiteUrlList(WebSite,content)		'不同域名网址列表
			
			'call echo(webSite,DifferentUrlList)
			
		end if
 	end function
	
    '获得 源网址列表
    Public Property Get getSourceUrlList()
        getSourceUrlList = SourceUrlList
    End Property
    '获得 干净网址列表
    Public Property Get getCleanUrlList()
        getCleanUrlList = CleanUrlList
    End Property
    '获得 相同网址列表
    Public Property Get getIdenticalUrlList()
        getIdenticalUrlList = IdenticalUrlList
    End Property
    '获得 不同网址列表
    Public Property Get getDifferentUrlList()
        getDifferentUrlList = DifferentUrlList
    End Property
    '获得 不同参数网址列表
    Public Property Get getDifferentParameterUrlList()
        getDifferentParameterUrlList = handleDifferentParameterUrlList(IdenticalUrlList)
    End Property
    '获得 注入网址列表
    Public Property Get getSqlInUrlList()
        getSqlInUrlList = handleSqlInUrlList(IdenticalUrlList)
    End Property
    '获得 发送状态
    Public Property Get getWebState()
        getWebState =WebState
    End Property
    '获得 网址文件大小
    Public Property Get getWebFileSize()
        getWebFileSize =WebFileSize
    End Property
	'设置 是否读缓冲文件
	Public Property Let setisReadCacheFile(Str)
		isReadCacheFile = Str
	End Property
	
	'处理数组里网址完整
	function handleFullHttpUrl(WebSite,urlList)						 '处理图片完整地址
		dim splUrl,url,title,s,c,splxx
		WebSite=getWebSite(WebSite)
		SplUrl = Split(urlList, vbCrLf)
		urlList=""
		For Each s In SplUrl
			if instr(s,"$Array$")>0 then
				splxx=split(s,"$Array$")
				url=splxx(0)
				title=splxx(1)
			else
				url = s
				title=""
			end if
			url = FullHttpUrl(WebSite,url)
			if instr(vbcrlf & urlList & vbCrlf, vbcrlf & url & vbCrlf)=false then
				urlList=urlList & url & vbcrlf
				if c<>"" then c=c & vbCrlf
				c=c & url & "$Array$" & title
			end if
		next
		handleFullHttpUrl = c
	end function
		
	'获得存在字符的网址列表
	public function getSearchUrl(urlList,searchStr)
		dim splUrl,url,c,isAdd
		SplUrl = Split(urlList, vbCrLf)
		For Each url In SplUrl
			if instr(vbcrlf & c & vbcrlf, vbcrlf & url & vbcrlf)=false and instr(url,searchStr)>0  then
				if c<>"" then c=c & vbCrlf
				c=c & url
			end if
		next
		getSearchUrl=c
	end function		                            	
	'去除同相网址，超强版
	function RemoveWidthUrl(urlList)
		dim splUrl,url,c,isAdd
		SplUrl = Split(urlList, vbCrLf)
		For Each url In SplUrl
			isAdd=false
			if instr(vbcrlf & c & vbcrlf, vbcrlf & url & vbcrlf)=false  then
				isAdd=true
				if right(url,1)<>"/" then
					if instr(vbcrlf & c & vbcrlf, vbcrlf & url & "/" & vbcrlf)>0  then
						isAdd=false
					end if
				end if
				if isAdd=true then
					'call echo(url,"url")
					if c<>"" then c=c & vbCrlf
					c=c & url
				end if
			end if
		next
		RemoveWidthUrl=c
	end function	
	'处理 注入网址列表
	function handleSqlInUrlList(urlList)
		dim splUrl,url,tempUrl,c
		SplUrl = Split(urlList, vbCrLf)
		urlList=""
		For Each url In SplUrl
			tempUrl=remoteHttpUrlParameter(url)
			'call echo(url,tempUrl)
			if instr(tempUrl,"?")>0 and instr(vbcrlf & urlList & vbcrlf, vbcrlf & tempUrl & vbcrlf)=false then				
				urlList=urlList & tempUrl & vbCrlf
				if c<>"" then c=c & vbCrlf
				c=c & url
			end if
		next
		handleSqlInUrlList=c
	end function	
	'处理 不同参数网址列表
	function handleDifferentParameterUrlList(urlList)
		dim splUrl,url,tempUrl,c
		SplUrl = Split(urlList, vbCrLf)
		urlList=""
		For Each url In SplUrl
			tempUrl=remoteHttpUrlParameter(url)
			'call echo(url,tempUrl)
			if len(tempUrl)>3 and instr(vbcrlf & urlList & vbcrlf, vbcrlf & tempUrl & vbcrlf)=false then				
				urlList=urlList & tempUrl & vbCrlf
				if c<>"" then c=c & vbCrlf
				c=c & url
			end if
		next
		handleDifferentParameterUrlList=c
	end function	
	'批量处理无需操作网址列表 如新浪百度等
	Function HandleSpecialUrl(urlList)
		Dim Path, Url, SplUrl, HttpUrl, UrlYes
		SplUrl = Split(urlList, vbCrLf)
		urlList=""			'清空
		For Each HttpUrl In SplUrl
			UrlYes = True
			For Each Url In SplSpecialUrl
				If Url <> "" And Left(Url, 1) <> "#" Then
					If InStr(HttpUrl, Url) > 0 Then UrlYes = False: Exit For
				End If
			Next
			If UrlYes = True Then
				if UrlList<>"" then
					UrlList=UrlList & vbCrlf
				end if
				UrlList = UrlList & HandleHttpUrl(HttpUrl)
			End If
		Next
		HandleSpecialUrl = UrlList
	End Function
	'处理无效网址  如javascript:;      '已经去掉重复网址
	Function HandleInvalidUrl(urlList)
		Dim SplUrl, HttpUrl, SuffixUrl, HandleYes
		SplUrl = Split(urlList, vbCrLf)
		urlList=""
		For Each HttpUrl In SplUrl
			HttpUrl = HandleHttpUrl(HttpUrl)
			SuffixUrl = Mid(HttpUrl, InStrRev(HttpUrl, "/") + 1)
			HandleYes = True
			
			'处理最后为/#
			If HandleYes = True Then
				HandleYes = len(HttpUrl)>3
			End If
			'处理最后为/#
			If HandleYes = True Then
				HandleYes = Left(SuffixUrl, 1) <> "#"
			End If
			'处理是否在列表里
			If HandleYes = True Then
				HandleYes = InStr(vbCrLf & UrlList & vbCrLf, vbCrLf & HttpUrl & vbCrLf) = False
			End If
			'判断最后为/javascript:
			If HandleYes = True Then
				HandleYes = Left(LCase(SuffixUrl), 11) <> "javascript:"
			End If
			'追加，搜索这个javascript: 屏蔽它  网址中有'屏蔽
			If HandleYes = True Then
				If InStr(LCase(HttpUrl), "javascript:") > 0 Or InStr(LCase(HttpUrl), "'") > 0 Then
					HandleYes = False
				End If
			End If
			
			'网址列表累加
			If HandleYes = True Then
				if UrlList<>"" then
					UrlList=UrlList & vbCrlf
				end if
				UrlList = UrlList & HttpUrl
			End If
		Next
		HandleInvalidUrl = UrlList
	End Function
	'屏蔽非网页网址如.jpg.gif不需要
	Function RemoveNonWebpage(Content)
		Dim SplStr, HttpUrl, LowerCaseUrl, UrlList, SuffixStr, SuffixType, UrlYes
		SplStr = Split(Content, vbCrLf)
		For Each HttpUrl In SplStr
			HttpUrl = Trim(HttpUrl)
			UrlYes = True
			If HttpUrl <> "" Then
				HttpUrl = HandleHttpUrl(HttpUrl)
				LowerCaseUrl = LCase(HttpUrl)
				SuffixStr = Mid(LowerCaseUrl, InStrRev(LowerCaseUrl, "/") + 1)
				'有?号则 取?号前面值
				If InStr(SuffixStr, "?") Then
					SuffixStr = Mid(SuffixStr, 1, InStr(SuffixStr, "?") - 1)
				End If
				If InStr(SuffixStr, ".") > 0 Then
					SuffixType = Mid(SuffixStr, InStrRev(SuffixStr, ".") + 1)
					If InStr("|jpg|gif|png|bmp|zip|rar|js|xml|doc|pdf|ppt|xlsx|xls|exe|txt|", "|" & SuffixType & "|") > 0 Then
						UrlYes = False
					End If
				End If
			End If
			If UrlYes = True Then
				If InStr(vbCrLf & UrlList & vbCrLf, vbCrLf & HttpUrl & vbCrLf) = False Then
					if UrlList<>"" then
						UrlList=UrlList & vbCrlf
					end if
					UrlList = UrlList & HttpUrl
				End If
			End If
		Next
		RemoveNonWebpage = UrlList
	End Function
	'获得 同域名列表
	Function GetIdenticalWebSiteUrlList(ByVal WebSite, ByVal urlList)
		GetIdenticalWebSiteUrlList = GetIdenticalOrDifferentWebSite(WebSite, urlList, "identica")
	End Function
	'获得 不同域名列表
	Function GetDifferentWebSiteUrlList(ByVal WebSite, ByVal urlList)
		GetDifferentWebSiteUrlList = GetIdenticalOrDifferentWebSite(WebSite, urlList, "different")
	End Function
	'处理是否获得同域名网址列表与非同域名网址列表 SType=identica   different
	Function GetIdenticalOrDifferentWebSite(ByVal WebSite, ByVal urlList, SType)
		Dim SplUrl, HttpUrl, SuffixUrl, ThisWebSite, YesNO
		SType = LCase(SType)                                    '类型转成小写 同时自动为字符类型
'		WebSite = LCase(GetWebSite(WebSite))                    '获得域名
		SplUrl = Split(urlList, vbCrLf)
		urlList=""
		For Each HttpUrl In SplUrl
			HttpUrl = Trim(HttpUrl)
			ThisWebSite = LCase(GetWebSite(HttpUrl))
			YesNO = False
			if SType = "identica" Or SType = "1" then
				If ThisWebSite = WebSite Or instr(ThisWebSite,WebSite)>0 then
					yesNo=true
				end if
			else
				If ThisWebSite <> WebSite  and instr(ThisWebSite,WebSite)=false then
					'call echo(ThisWebSite,WebSite)
					yesNo=true
				end if
			end if
			'处理累加网址列表
			If YesNO = True Then
				If InStr(vbCrLf & UrlList & vbCrLf, vbCrLf & HttpUrl & vbCrLf) = False Then
					if UrlList<>"" then
						UrlList=UrlList & vbCrlf
					end if
					UrlList = UrlList & HttpUrl
				End If
			End If
		Next
		GetIdenticalOrDifferentWebSite = UrlList
	End Function	
	
	
	'----------------------------------------------------------------------------------------
	
	'初始化
	Private Sub Class_Initialize()
		SplSpecialUrl = Split(GetFText("\VB工程\Config\不处理域名列表.ini"), vbCrLf)           '分割无需处理网址
		'Call Echo("SplSpecialUrl",Ubound(SplSpecialUrl))
		CacheDir = "E:\E盘\WEB网站\网站UrlScan\"                          '缓存目录路径
		DatabaseType = "Access"                                           '默认设置数据库
		isReadCacheFile=true										      '为读缓冲文件
		if checkFolder(CacheDir)=false then
			call eerr("缓存目录路径不存在", cacheDir)
		end if

		Set Rs = CreateObject("Adodb.RecordSet")
		Set Rsx = CreateObject("Adodb.RecordSet")
		Set Rss = CreateObject("Adodb.RecordSet")
		Set Rst = CreateObject("Adodb.Recordset")
		Set Rsd = CreateObject("Adodb.Recordset")
		Set TempRs = CreateObject("Adodb.RecordSet")
		Set TempRs2 = CreateObject("Adodb.RecordSet")
		Set RsTemp = CreateObject("Adodb.RecordSet")	
		
	End Sub
	'终止
	Private Sub Class_Terminate()
		
	End Sub
	'设置缓存目录路径
	Public Property Let SetCacheDir(FolderPath)
		CacheDir = FolderPath
	End Property
	'设置 数据库类型
	Public Property Let setDatabaseType(Str)
		DatabaseType = Str
	End Property
    '获得函数中无意义变量数
    Public Property Get GetTxtFilePath(Url)
        GetTxtFilePath = CacheDir & SetFileName(Url)                           '文件存放在本地路径
    End Property 	
	
	
	'打开数据库
	Sub OpenConn()
		Dim SqlDatabaseName, SqlPassword, SqlUsername, SqlLocalName, ConnStr
		If DatabaseType = "Access" Then
			MDBPath = "/../网站备份\数据库/WebUrlScan.mdb"
			Call HandlePath(MDBPath)    '获得完整路径
		End If
		'连接MMD数据库
		If MDBPath <> "" Then
			Call HandlePath(MDBPath)    '获得完整路径
			Set Conn = CreateObject("Adodb.Connection")
			Conn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database PassWord = '';Data Source = " & MDBPath
		Else
			If DatabaseType = "SqlServerWebData" Then
				SqlDatabaseName = "WebData" 'SQL数据库名
				SqlUsername = "sa" 'SQL数据库用户名
				SqlPassword = "aaa" 'SQL数据库用户密码
				SqlLocalName = "127.0.0.1,1433" 'SQL主机IP地址（服务器名）在2008版本使用
				
			ElseIf DatabaseType = "SqlServerLocalData" Then
				SqlDatabaseName = "LocalData" 'SQL数据库名
				SqlUsername = "sa" 'SQL数据库用户名
				SqlPassword = "aaa" 'SQL数据库用户密码
				SqlLocalName = "127.0.0.1,1433" 'SQL主机IP地址（服务器名）在2008版本使用
			
			ElseIf DatabaseType = "RemoteSqlServer" Then
				'远程SqlServer数据库
				SqlDatabaseName = "qds0140159_db"
				SqlUsername = "qds0140159": SqlPassword = "L4dN4eRd"
				SqlLocalName = "qds-014.hichina.com"
			End If
			ConnStr = " Password = " & SqlPassword & "; user id =" & SqlUsername & "; Initial Catalog =" & SqlDatabaseName & "; data source =" & SqlLocalName & ";Provider = sqloledb;"
			Set Conn = CreateObject("Adodb.Connection")
			Conn.Open ConnStr
		End If
		
		Set Rs = CreateObject("Adodb.Recordset")
		Set Rss = CreateObject("Adodb.Recordset")
		Set Rst = CreateObject("Adodb.Recordset")
		Set Rsx = CreateObject("Adodb.Recordset")
		Set Rsd = CreateObject("Adodb.Recordset")
		Set TempRs = CreateObject("Adodb.RecordSet")
		Set RsTemp = CreateObject("Adodb.RecordSet")
	End Sub
	'清空数据库
	Sub ClearDatabases()
		Conn.Execute ("Delete From [WebUrlScan]")                       '清空网址列表
	End Sub
  	'---------------------------------------------------------------------------------------- 操作数据库
	'追加网站
	function addUrl(toUrl,ToTitle)
		call OpenConn()
		addUrl=false
		rs.open"Select * From [WebUrlScan] Where HttpUrl='"& pubHttpUrl &"'",conn,1,3
		if rs.eof then
			rs.addnew
			
			rs("WebSite")=pubWebSite
			rs("HttpUrl")=pubHttpUrl
			rs("Title")=pubWebTitle
			rs("Title")=pubWebTitle
			rs("WebState")=WebState
			rs("WebFileSize")=WebFileSize
			
			rs("toUrl")=toUrl
			rs("ToTitle")=ToTitle
			
			rs.update
			addUrl=true
		end if :rs.close
	end function
	'判断网站
	function checkUrl(url)
		call OpenConn()
		rs.open"Select * From [WebUrlScan] Where HttpUrl='"& url &"'",conn,1,1
		checkUrl=false
		if not rs.eof then
			checkUrl=true
		end if
		rs.close
	end function
	'批量操作网址列表
	function batchAddUrl(toUrl,urlList)
		dim splstr,url,c,isAddOK,splxx,toTitle
		splstr=split(urlList,vbCrlf)
		'call rw(urlList)
		for each url in splstr			
			url=phpTrim(url)	
			if instr(url,"$Array$")	>0 then
				splxx=split(url,"$Array$")
				url = splxx(0)
				toTitle=splxx(1)
			end if
			if len(url)>3 then
				'网址为false则操作
				if checkUrl(url)=false then
					call GetHttpUrlContentUrlList(url,"","",toUrl,toTitle)	'127.0.0.1本地测试时，加utf-8编码则出错
					isAddOK = addUrl(toUrl,toTitle)
					c=c & url & "("& isAddOK &")" & vbCrlf
				else				
					c=c & url & "(exist)" & vbCrlf
				end if
			end if
		next
		batchAddUrl = c
	end function
End Class

%>

