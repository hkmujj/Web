<!--#Include File = "Inc/_Config.Asp"-->            
<!--#Include File = "web/function.Asp"-->      
<% 

Select Case Request("act")
    Case "scanurl" : Call scanHttpUrl(Request("url"))
    Case "displayUrlList" : Call displayUrlList()
    Case "displayLisks" : Call displayLisks()		'��ʾ����
	
	case "displayDataInfo" : call displayDataInfo()		'��ʾ���ݿ�
	
    Case "testUrl" : Call testUrl()
    Case Else : Call displayDefaultLayout()
End Select

'��ʾĬ�����
function displayDefaultLayout()
	call echo("��ʾURL�б�","<a href='?act=displayUrlList&nThread=1'>����1</a>|<a href='?act=displayUrlList&nThread=2'>����2</a>|<a href='?act=displayUrlList&nThread=3'>����3</a>|<a href='?act=displayUrlList&nThread=12'>����12</a>")
	call echo("�����б�","<a href='?act=displayLisks'>�������</a>")
	call echo("�����б�","<a href='?act=displayDataInfo'>displayDataInfo</a>")
	call echo("������ַ","<a href='?act=testUrl'>�������</a>")
end function

'��÷�������ַ
Function getServerUrl(ByVal nIndex)
    Dim serverName 
    serverName = Array("bb", "cc", "dd", "ee", "ff", "gg","z1","z2","z3","z4","z5","z6","") 
    nIndex = nIndex Mod UBound(serverName)
    getServerUrl = "http://" & serverName(nIndex) & "/atemp." & EDITORTYPE & "?act=scanurl"
End Function 

'��ʾ���ݿ�
function displayDataInfo()		
    Dim i, url, s ,urlList,c
    Call openconn()  
    rs.Open "select top 130999 * from xy_webdomain  ", conn, 1, 1 
    For i = 1 To rs.RecordCount
		call rw("<li>")
        call echo(i,"<a href="""& rs("website") & """ target='_blank'>"& phptrim(rs("webtitle")) &"</a>")
		call rw("</li>")
    rs.MoveNext : Next : rs.Close 
end function
'��ʾ����
function displayLisks()
    Dim i, url, s ,urlList,c
    Call openconn() 
	rs.Open "select * from xy_webdomain where links<>''", conn, 1, 1 
	call echo("���м�¼",rs.RecordCount)
	rs.close
	call deletefile("1.txt")
    rs.Open "select top 130999 * from xy_webdomain where links<>'' order by nlinks desc", conn, 1, 1 
    For i = 1 To rs.RecordCount
        '��PHP��$rs=mysql_fetch_array($rsObj);
		call echo(rs("nlinks"), "<a href='" & rs("website") & "' target='_blank'>" & rs("website") & "</a>")
'		call rwend(rs("links"))
		call createAddFile2("1.txt",handleScanUrl(rs("links")))
    rs.MoveNext : Next : rs.Close 
end function
'����ɨ����ַ�б�
function handleScanUrl(urlList)
	dim splstr,url,c
	splstr=split(urlList,vbcrlf)
	for each url in splstr
		'call echo(url,checkScanUrl(url))
		if instr(vbcrlf & c & vbcrlf, vbcrlf & url & vbcrlf)=false and checkScanUrl(url)=true then
			c=c & url & vbcrlf
		end if	
	next
	handleScanUrl=c
end function
'����Ƿ�Ϊ��ɨ����ַ
function checkScanUrl(url)
	dim website,c,splstr,s
	website=lcase(getwebsite(url))
	website=replace(website,"/",".")		'���⴦��
	c="alipay|sina|qq|google|baidu|so|bing|sogou|yahoo|haosou|youdao|163|360|39|ifeng|hao123|51|letv|sohu|pptv|taobao|tv|renren|gov|github|admin5|chinaz"
	c=c & "|cnzz|mycodes|hicode|asp300|662p|codesky|php|asp|jsp|thinkphp|dedecms|jb51|codefans|aspjzy|ip138|"
	
	splstr=split(c,"|")
	for each s in splstr
		if s <>"" then
			if instr(website,"." & s & ".com.")>0 or instr(website,"." & s & ".cn.")>0 or instr(website,"." & s & ".net.")>0  or instr(website,"." & s & ".la.")>0 then
				checkScanUrl=false
				exit function
			end if
		end if	
	next
	checkScanUrl=true
end function

'��ʾ��ַ�б�
Function displayUrlList()
    Dim i, url, s ,urlList,c
    Call openconn() 
	rs.Open "select * from xy_webdomain where isthrough<>0", conn, 1, 1 
	call echo("���м�¼",rs.RecordCount)
	rs.close
    rs.Open "select top 130999 * from xy_webdomain where isthrough<>0", conn, 1, 1 
    For i = 1 To rs.RecordCount
        '��PHP��$rs=mysql_fetch_array($rsObj);
        url = getServerUrl(i) & "&url=" & rs("website") 
        
		if c<>"" then c=c & ","
		c=c & "'"& url &"'"
		's = "<iframe src='" & url & "' height='100' width='100%' frameborder='1' scrolling='yes'></iframe>" & vbcrlf
        'Call rw(s) 
        'Call echo(url, rs("website"))
    rs.MoveNext : Next : rs.Close 
	dim nThread
	nThread=handleNumber(request("nThread"))
	if nThread="" then
		nThread=12
	end if
	call rw(batchJsHandle(c,nThread))
End Function 

'������ַ
Function testUrl()
    Dim httpurl 
    httpurl = "http://127.0.0.1/1.html" 
    httpurl = "http://sharembweb.com" 
    Call scanHttpUrl(httpurl) 
End Function 
'ɨ����ַ
Function scanHttpUrl(httpurl)
    Dim rootDir, webSite, folderDir, configFile, configContent, webstate,tempWebstate, msgStr, content, splStr, s, url 
    Dim htmlFile, websize, webtitle, webkeywords, webdescription, isdomain, startTime, openspeed,homepagelist,lists,PubAHrefList, PubATitleList
	dim isasp,isaspx,isphp,isjsp,ishtm,ishtml,links,nlinks
	isasp=0:isaspx=0:isphp=0:isjsp=0:ishtm=0:ishtml=0
    rootDir = "/../��վUrlScan/httpurl2016/" 
    Call createDirFolder(rootDir) 

    webSite = getWebSite(httpurl) 
    Call echo("��ַ", httpurl) 
    s = "<a href='"& webSite &"' target='_blank'>" & webSite & "</a>" 
    Call echo("����", s) 
    folderDir = rootDir & setFileName(webSite) & "/" & setFileName(httpurl) & "/" 
    Call createDirFolder(folderDir) 
    configFile = folderDir & "config.txt" 

    configContent = vbCrLf & getftext(configFile) & vbCrLf 

    isdomain = getStrCut(configContent, vbCrLf & "isdomain=", vbCrLf, 2) 
    If isdomain = "" Then
        isdomain = IIF(checkDomainName(webSite), "1", "0") 
        Call createAddFile2(configFile, "isdomain=" & isdomain & vbCrLf) 
    End If 
    Call echo("������Ч��", isdomain) 
    If isdomain = "0" Then
		call openconn()
		conn.execute("update xy_webdomain set isthrough=0 where website='"& httpurl &"'")
        Call eerr("������Ч", webSite) 
    End If 

    webstate = getStrCut(configContent, vbCrLf & "webstate=", vbCrLf, 2) 
    openspeed = getStrCut(configContent, vbCrLf & "openspeed=", vbCrLf, 2) 
    msgStr = "���ػ���" 
    If webstate = "" Then
        startTime = Now() 
        webstate = getHttpUrlState(httpurl) 
        openspeed = DateDiff("s", startTime, Now()) 
        Call createAddFile2(configFile, "webstate=" & webstate & vbCrLf) 
        Call createAddFile2(configFile, "openspeed=" & openspeed & vbCrLf) 
        msgStr = "��ȡ����" 
    End If 
    Call echo("��״̬��" & msgStr, webstate & "(" & getHttpUrlStateAbout(webstate) & ")") 
    Call echo("openspeed", openspeed) 

    htmlFile = folderDir & setFileName(httpurl) & ".html" 
    If checkFile(htmlFile) = False Then
        content = getHttpUrl(httpurl, "") 
        '��PHP��$content=toGB2312Char($content);                                            //��PHP�ã�ת��gb2312�ַ�
		if content="" then content="��Ϊ�ա�"
        Call createFile(htmlFile, content) 
    End If 
    content = getFText(htmlFile) 
    websize = getFSize(htmlFile) 

    webtitle = getHtmlValue(content, "webtitle") 
    webkeywords = getHtmlValue(content, "webkeywords") 
    webdescription = getHtmlValue(content, "webdescription") 
    Call echo("webtitle", webtitle) 
    Call echo("webkeywords", webkeywords) 
    Call echo("webdescription", webdescription) 
    Call echo("��ҳ��С", printSpaceSize(websize)) 
	
	nlinks=0
    lists = getContentAHref("", content, PubAHrefList, PubATitleList) 
    lists = handleDifferenceWebSiteList(httpurl, lists)		'��ò�ͬ��ַ�б�	
	splstr=split(lists,vbcrlf)
	for each url in splstr		
		url=getwebsite(url)
		if url<>"" and instr(vbcrlf & links & vbcrlf,vbcrlf & lists & vbcrlf)=false then
			links=links & url & vbcrlf
			nlinks=nlinks+1
		end if
	next
	links=ADSql(links)
	call echo(nlinks,links)

    '��ҳ״̬
    splStr = Array("index.asp", "index.aspx", "index.php", "index.jsp", "index.htm", "index.html", "default.asp", "default.aspx", "default.jsp", "default.htm", "default.html") 
    For Each s In splStr
        url = webSite & s 
        tempWebstate = getStrCut(configContent, s & "=", vbCrLf, 2) 
        If tempWebstate = "" Then
            tempWebstate = getHttpUrlState(url)
            Call createAddFile2(configFile, s & "=" & tempWebstate & vbCrLf) 
        End If
		if (s="index.asp" or s="default.asp") and (tempWebstate="200" or tempWebstate="301") then
			isasp=1
		end if
		if (s="index.aspx" or s="default.aspx") and (tempWebstate="200" or tempWebstate="301") then
			isaspx=1
		end if
		if (s="index.php" or s="default.php") and (tempWebstate="200" or tempWebstate="301") then
			isphp=1
		end if
		if (s="index.jsp" or s="default.jsp") and (tempWebstate="200" or tempWebstate="301") then
			isjsp=1
		end if
		if (s="index.htm" or s="default.htm") and (tempWebstate="200" or tempWebstate="301") then
			ishtm=1
		end if
		if (s="index.html" or s="default.html") and (tempWebstate="200" or tempWebstate="301") then
			ishtml=1
		end if
		if tempWebstate="200" then
			homepagelist=homepagelist & s & "|"
		end if
        Call echo(s, tempWebstate) 
    Next
	call openconn()
	s=",homepagelist='"& homepagelist &"',isasp="& isasp &",isaspx="& isaspx &",isphp="& isphp &",isjsp="& isjsp &",ishtm="& ishtm &",ishtml="& ishtml &",links='"& links &"',nlinks="& nlinks &",webtitle='"& webtitle &"',webkeywords='"& webkeywords &"',webdescription='"& webdescription &"'"	
	conn.execute("update xy_webdomain set isthrough=0,webstate="& webstate &",openspeed="& openspeed &",websize="& websize &",isdomain="& isdomain & s &" where website='"& httpurl &"'")
End Function 

function batchJsHandle(urlList,nThread)
%>
<script type="text/javascript" src="/Jquery/Jquery.Min.js"></script>
</head>
<body>
<div id="msg"></div>
<%
dim i,s,c
for i =1 to nThread
	call rw("<iframe src='' id='iframe"& i &"' height='100' width='100%' frameborder='1' scrolling='yes'></iframe>" & vbcrlf)
	if i>1 then
		s="index++" & vbcrlf & "callIframe(""iframe"& i &""",urlArray[index],callback)"
	end if
	c=c & s & vbcrlf
next
%>
<script>
var urlArray=Array(<%=urlList%>);
var index=0
var nRun=0
function callIframe(id,url,callback) { 
    $('iframe#'+id).attr('src', url);  
    $('iframe#'+id).load(function(){
        callback(id,this); 
    }); 
} 
function callback(id,e){
	if(index<urlArray.length-1){
		//alert(index + '<' + (urlArray.length-1))
		index++
		
		callIframe(id,urlArray[index],callback)	
	}
	nRun++
	$("#msg").html( index + '/' + (urlArray.length-1) + ",���д�����"+ nRun +"��" )
	
	if(index >=(urlArray.length-1) && urlArray.length>0){

 		$("#msg").html( $("#msg").html() + "���ö�ʱ����8 ���ˢ�µ�ǰҳ��")
		var picTimer = setInterval(function() {
				location.reload()   
				clearInterval(picTimer);  
		 
		},8000); //��4000�����Զ����ŵļ������λ������

	}
	
}

callIframe("iframe1",urlArray[index],callback)
<%=c%>
</script>
<%end function%>