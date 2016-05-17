<%
'������� ��ַ      �򵥼򵥣�Խ��Խ��

Class NetWorkClass
	'������ַ������
	Dim MDBPath, DatabaseType
	Dim Conn
	Dim Rs
	Dim Rss
	Dim Rst
	Dim Rsx
	Dim Rsd
	Dim TempRs
	Dim RsTemp
	Dim SplSpecialUrl                   '�ָ����账����ַ
	Dim CacheDir                        '����Ŀ¼·��
	
	dim WebSite							'����
	dim SourceUrlList					'Դ��ַ�б�
	dim CleanUrlList					'�ɾ���ַ�б�	
	dim IdenticalUrlList				'��ͬ��ַ�б�
	dim DifferentUrlList				'��ͬ��ַ�б�
	dim DifferentParameterUrlList		'��ͬ������ַ�б�
	dim SqlInUrlList					'��ע����ַ�б�
	dim WebState						'��վ״̬
	dim WebFileSize						'�����ļ���С
	dim pubHttpUrl						'��ǰ��ַ
	dim pubWebTitle						'��վ����
	dim pubWebSite						'��վ����
	dim isReadCacheFile					'�Ƿ�������ļ�
	dim Depth							'���
	

	'�����ַ��������ַ�б�  CustomWebSite Ϊ�����жϣ�Ϊ�����Զ���ȡ����
	Function GetHttpUrlContentUrlList(ByVal HttpUrl,setCode,CustomWebSite,toUrl,toTitle)
		Dim Content,HttpArrar,saveFolder,saveFilePath
		HttpUrl = HandleUrlComplete(HttpUrl)                            '������ַ������

		HttpUrl = HandleSpecialUrl(HttpUrl)                             '�������账����ַ �����˰ٶȵ�
		HttpUrl = HandleInvalidUrl(HttpUrl)                             '������Ч��ַ ��javascript:;
		HttpUrl = RemoveNonWebpage(HttpUrl)                             '���η���ҳ��ַ��.jpg.gif����Ҫ
 	
		if HttpUrl<>"" then
			if CustomWebSite<>"" then
				WebSite = CustomWebSite		
			else
				WebSite=getWebSite(httpUrl)
			end if
			'Content = getHttpUrl(HttpUrl,"")                              	   '������ַ�����ݱ��ش�����ֱ�Ӷ���
			
			if isReadCacheFile=true then			
				saveFolder = CacheDir & setFileName(getWebSite(HttpUrl)) & "/"
				call createFolder(saveFolder)
				saveFilePath = saveFolder & setFileName(HttpUrl)
				
				if checkfile(saveFilePath)=true then
					HttpArrar=split(getFText(saveFilePath & ".txt"),vbCrlf)
					HttpArrar(0) = getFText(saveFilePath)
					toUrl = HttpArrar(2)											    '״̬
					toTitle = HttpArrar(3)											    '״̬
				end if
			end if
			if IsEmpty(HttpArrar) then
				HttpArrar = handleXmlGet(HttpUrl, setCode)
					call echo(HttpUrl,saveFilePath):doevents							'��ʾ������ַ��Ϣ
					call createFile(saveFilePath, HttpArrar(0))
					call createFile(saveFilePath & ".txt", vbcrlf & HttpArrar(1) & vbcrlf & toUrl & vbcrlf & toTitle )
			end if
			Content = HttpArrar(0)												'����
			WebState = HttpArrar(1)											    '״̬
			WebFileSize = stringLength(Content)								    '�ļ���С
			pubHttpUrl=HttpUrl													'��ַ��ַ
			pubWebTitle = RegExpGetStr("<TITLE>([^<>]*)</TITLE>", Content, 1)	'��վ����
			pubWebSite = getWebSite(HttpUrl)									'��վ����
			
			'call rwend(Content)
			if 1=2 then
			Content = GetContentAHref(HttpUrl, Content)                         '���������������ַ ��һ�֣��Լ�д��
			else
			Content = GetAUrlTitleList(Content,"")                         		 '���������������ַ �ڶ��֣�����
			Content = batchFullHttpUrl(WebSite,Content)							 '����������ַ��� �ڶ�����Ҫ
			'content = handleFullHttpUrl(WebSite,Content)						 '������ַ������ַ
			end if
			
			
			
			SourceUrlList = Content		'Դ��ַ�б�
			Content = HandleSpecialUrl(Content)                                 '�������账����ַ �����˰ٶȵ�
			Content = HandleInvalidUrl(Content)                                 '������Ч��ַ ��javascript:;
			Content = RemoveNonWebpage(Content)                                 '���η���ҳ��ַ��.jpg.gif����Ҫ
			Content = RemoveWidthUrl(Content)                                 	'ȥ��ͬ����ַ����ǿ��
			
			
			CleanUrlList = content		'�ɾ���ַ�б�
			
			IdenticalUrlList = GetIdenticalWebSiteUrlList(WebSite,content)		'��ͬ������ַ�б�
			DifferentUrlList = GetDifferentWebSiteUrlList(WebSite,content)		'��ͬ������ַ�б�
			
			'call echo(webSite,DifferentUrlList)
			
		end if
 	end function
	
    '��� Դ��ַ�б�
    Public Property Get getSourceUrlList()
        getSourceUrlList = SourceUrlList
    End Property
    '��� �ɾ���ַ�б�
    Public Property Get getCleanUrlList()
        getCleanUrlList = CleanUrlList
    End Property
    '��� ��ͬ��ַ�б�
    Public Property Get getIdenticalUrlList()
        getIdenticalUrlList = IdenticalUrlList
    End Property
    '��� ��ͬ��ַ�б�
    Public Property Get getDifferentUrlList()
        getDifferentUrlList = DifferentUrlList
    End Property
    '��� ��ͬ������ַ�б�
    Public Property Get getDifferentParameterUrlList()
        getDifferentParameterUrlList = handleDifferentParameterUrlList(IdenticalUrlList)
    End Property
    '��� ע����ַ�б�
    Public Property Get getSqlInUrlList()
        getSqlInUrlList = handleSqlInUrlList(IdenticalUrlList)
    End Property
    '��� ����״̬
    Public Property Get getWebState()
        getWebState =WebState
    End Property
    '��� ��ַ�ļ���С
    Public Property Get getWebFileSize()
        getWebFileSize =WebFileSize
    End Property
	'���� �Ƿ�������ļ�
	Public Property Let setisReadCacheFile(Str)
		isReadCacheFile = Str
	End Property
	
	'������������ַ����
	function handleFullHttpUrl(WebSite,urlList)						 '����ͼƬ������ַ
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
		
	'��ô����ַ�����ַ�б�
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
	'ȥ��ͬ����ַ����ǿ��
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
	'���� ע����ַ�б�
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
	'���� ��ͬ������ַ�б�
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
	'�����������������ַ�б� �����˰ٶȵ�
	Function HandleSpecialUrl(urlList)
		Dim Path, Url, SplUrl, HttpUrl, UrlYes
		SplUrl = Split(urlList, vbCrLf)
		urlList=""			'���
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
	'������Ч��ַ  ��javascript:;      '�Ѿ�ȥ���ظ���ַ
	Function HandleInvalidUrl(urlList)
		Dim SplUrl, HttpUrl, SuffixUrl, HandleYes
		SplUrl = Split(urlList, vbCrLf)
		urlList=""
		For Each HttpUrl In SplUrl
			HttpUrl = HandleHttpUrl(HttpUrl)
			SuffixUrl = Mid(HttpUrl, InStrRev(HttpUrl, "/") + 1)
			HandleYes = True
			
			'�������Ϊ/#
			If HandleYes = True Then
				HandleYes = len(HttpUrl)>3
			End If
			'�������Ϊ/#
			If HandleYes = True Then
				HandleYes = Left(SuffixUrl, 1) <> "#"
			End If
			'�����Ƿ����б���
			If HandleYes = True Then
				HandleYes = InStr(vbCrLf & UrlList & vbCrLf, vbCrLf & HttpUrl & vbCrLf) = False
			End If
			'�ж����Ϊ/javascript:
			If HandleYes = True Then
				HandleYes = Left(LCase(SuffixUrl), 11) <> "javascript:"
			End If
			'׷�ӣ��������javascript: ������  ��ַ����'����
			If HandleYes = True Then
				If InStr(LCase(HttpUrl), "javascript:") > 0 Or InStr(LCase(HttpUrl), "'") > 0 Then
					HandleYes = False
				End If
			End If
			
			'��ַ�б��ۼ�
			If HandleYes = True Then
				if UrlList<>"" then
					UrlList=UrlList & vbCrlf
				end if
				UrlList = UrlList & HttpUrl
			End If
		Next
		HandleInvalidUrl = UrlList
	End Function
	'���η���ҳ��ַ��.jpg.gif����Ҫ
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
				'��?���� ȡ?��ǰ��ֵ
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
	'��� ͬ�����б�
	Function GetIdenticalWebSiteUrlList(ByVal WebSite, ByVal urlList)
		GetIdenticalWebSiteUrlList = GetIdenticalOrDifferentWebSite(WebSite, urlList, "identica")
	End Function
	'��� ��ͬ�����б�
	Function GetDifferentWebSiteUrlList(ByVal WebSite, ByVal urlList)
		GetDifferentWebSiteUrlList = GetIdenticalOrDifferentWebSite(WebSite, urlList, "different")
	End Function
	'�����Ƿ���ͬ������ַ�б����ͬ������ַ�б� SType=identica   different
	Function GetIdenticalOrDifferentWebSite(ByVal WebSite, ByVal urlList, SType)
		Dim SplUrl, HttpUrl, SuffixUrl, ThisWebSite, YesNO
		SType = LCase(SType)                                    '����ת��Сд ͬʱ�Զ�Ϊ�ַ�����
'		WebSite = LCase(GetWebSite(WebSite))                    '�������
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
			'�����ۼ���ַ�б�
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
	
	'��ʼ��
	Private Sub Class_Initialize()
		SplSpecialUrl = Split(GetFText("\VB����\Config\�����������б�.ini"), vbCrLf)           '�ָ����账����ַ
		'Call Echo("SplSpecialUrl",Ubound(SplSpecialUrl))
		CacheDir = "E:\E��\WEB��վ\��վUrlScan\"                          '����Ŀ¼·��
		DatabaseType = "Access"                                           'Ĭ���������ݿ�
		isReadCacheFile=true										      'Ϊ�������ļ�
		if checkFolder(CacheDir)=false then
			call eerr("����Ŀ¼·��������", cacheDir)
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
	'��ֹ
	Private Sub Class_Terminate()
		
	End Sub
	'���û���Ŀ¼·��
	Public Property Let SetCacheDir(FolderPath)
		CacheDir = FolderPath
	End Property
	'���� ���ݿ�����
	Public Property Let setDatabaseType(Str)
		DatabaseType = Str
	End Property
    '��ú����������������
    Public Property Get GetTxtFilePath(Url)
        GetTxtFilePath = CacheDir & SetFileName(Url)                           '�ļ�����ڱ���·��
    End Property 	
	
	
	'�����ݿ�
	Sub OpenConn()
		Dim SqlDatabaseName, SqlPassword, SqlUsername, SqlLocalName, ConnStr
		If DatabaseType = "Access" Then
			MDBPath = "/../��վ����\���ݿ�/WebUrlScan.mdb"
			Call HandlePath(MDBPath)    '�������·��
		End If
		'����MMD���ݿ�
		If MDBPath <> "" Then
			Call HandlePath(MDBPath)    '�������·��
			Set Conn = CreateObject("Adodb.Connection")
			Conn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database PassWord = '';Data Source = " & MDBPath
		Else
			If DatabaseType = "SqlServerWebData" Then
				SqlDatabaseName = "WebData" 'SQL���ݿ���
				SqlUsername = "sa" 'SQL���ݿ��û���
				SqlPassword = "aaa" 'SQL���ݿ��û�����
				SqlLocalName = "127.0.0.1,1433" 'SQL����IP��ַ��������������2008�汾ʹ��
				
			ElseIf DatabaseType = "SqlServerLocalData" Then
				SqlDatabaseName = "LocalData" 'SQL���ݿ���
				SqlUsername = "sa" 'SQL���ݿ��û���
				SqlPassword = "aaa" 'SQL���ݿ��û�����
				SqlLocalName = "127.0.0.1,1433" 'SQL����IP��ַ��������������2008�汾ʹ��
			
			ElseIf DatabaseType = "RemoteSqlServer" Then
				'Զ��SqlServer���ݿ�
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
	'������ݿ�
	Sub ClearDatabases()
		Conn.Execute ("Delete From [WebUrlScan]")                       '�����ַ�б�
	End Sub
  	'---------------------------------------------------------------------------------------- �������ݿ�
	'׷����վ
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
	'�ж���վ
	function checkUrl(url)
		call OpenConn()
		rs.open"Select * From [WebUrlScan] Where HttpUrl='"& url &"'",conn,1,1
		checkUrl=false
		if not rs.eof then
			checkUrl=true
		end if
		rs.close
	end function
	'����������ַ�б�
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
				'��ַΪfalse�����
				if checkUrl(url)=false then
					call GetHttpUrlContentUrlList(url,"","",toUrl,toTitle)	'127.0.0.1���ز���ʱ����utf-8���������
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

