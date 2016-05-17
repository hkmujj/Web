<!--#Include File = "Inc/_Config.Asp"-->          
<!--#Include File = "web/function.Asp"-->    
<%
dim pub_httpurl
pub_httpurl="http://maiside.net/"
'保存数据处理页
Select Case Request("act")
	case "displayGetDefaultIndex" : displayGetDefaultIndex(request("httpurl"))	'显示获得默认首页界面
	case "getDefaultIndex" : getDefaultIndex()	'获得默认首页
	
	case "displayGetWebSiteDir" : displayGetWebSiteDir(request("httpurl"))	'显示获得网站目录
	case "getWebSiteDir" : getWebSiteDir()	'获得默认首页
	
	case "scanWebSiteSafe" : scanWebSiteSafe("http://127.0.0.1")		'扫描见网站安全
	
	case else : displayDefaultLayout()		'显示默认布局
End Select

'扫描见网站安全
function scanWebSiteSafe(httpurl)
	dim website,splstr,url,s,c,nState,isAsp,isAspx,isPhp,isJsp
	call echo("扫描网址", httpurl)
	website=getwebsite(httpurl)
	if website="" then
		call eerr("域名为空",httpurl)
	end if
	call echo("域名",website)
	isAsp=0:isAspx=0:isPhp=0:isJsp=0
	'检测网站使用那种语言
	splstr=Array("index.asp","index.aspx","index.php","index.jsp","default.asp","default.aspx","default.jsp")
	for each s in splstr
		url=website & s
		nState=getHttpUrlState(url)
		nState=cstr(nState)
		call echo(url,nState & "   ("& getHttpUrlStateAbout(nState) &")" & typename(nState))
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
	next
	call echo("isAsp",isAsp)
	call echo("isAspx",isAspx)
	call echo("isPhp",isPhp)
	call echo("isJsp",isJsp)
end function


'显示默认布局
function displayDefaultLayout()
	call echo("显示获得默认首页界面","<a href='?act=displayGetDefaultIndex'>进入</a>")
	call echo("显示获得网站目录","<a href='?act=displayGetWebSiteDir'>进入</a>")
	
	call echo("检测网站安全","<a href='?act=scanWebSiteSafe'>进入</a>")
	
end function
'获得目录列表
function getWebSiteDir(httpurl)
	dim url,splstr,s,website,nState
	website=getwebsite(httpurl)
	if website="" then
		call eerr("域名为空",httpurl)
	end if
	splstr=Array("inc/","admin/","manage/","images/","css/","js/")
	for each s in splstr
		url=website & s
		nState=getHttpUrlState(url)
		call echo(url,nState & "   ("& getHttpUrlStateAbout(nState) &")")
		doevents
	next
end function

'获得默认首页界面
function getDefaultIndex(httpurl)
	dim url,splstr,s,website,nState
	website=getwebsite(httpurl)
	if website="" then
		call eerr("域名为空",httpurl)
	end if
	splstr=Array("index.asp","index.aspx","index.php","index.jsp","index.htm","index.html","default.asp","default.aspx","default.jsp","default.htm","default.html")
	for each s in splstr
		url=website & s
		nState=getHttpUrlState(url)
		call echo(url,nState & "   ("& getHttpUrlStateAbout(nState) &")")
		doevents
	next
end function
'显示获得默认首页界面
sub displayGetDefaultIndex()
%>
<form id="form1" name="form1" method="post" action="?act=getDefaultIndex">
  <p>获得域名默认首页</p>
  <p>域名
    <input type="text" name="httpurl" id="httpurl" value="<%=pub_httpurl%>" />
    <input type="submit" name="button" id="button" value="提交" />
  </p>
</form>
<%
end sub
'显示获得网站目录
sub displayGetWebSiteDir()
%>
<form id="form1" name="form1" method="post" action="?act=getWebSiteDir">
  <p>获得网站目录</p>
  <p>域名
    <input type="text" name="httpurl" id="httpurl" value="<%=pub_httpurl%>" />
    <input type="submit" name="button" id="button" value="提交" />
  </p>
</form>
<%
end sub
%>