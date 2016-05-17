<!--#Include virtual = "/Inc/_Config.Asp"--> 
<%
'保存数据处理页
Select Case Request("act")
	case "displayGetDefaultIndex" : displayGetDefaultIndex()	'显示获得默认首页界面
	case "getDefaultIndex" : getDefaultIndex()	'获得默认首页界面
	case else : displayDefaultLayout()		'显示默认布局
End Select

'显示默认布局
function displayDefaultLayout()
	call echo("显示获得默认首页界面","<a href='?act=displayGetDefaultIndex'>进入</a>")
	
end function

'获得默认首页界面
function getDefaultIndex()
	dim url,splstr,s,httpurl,website
	httpurl=request("httpurl")
	website=getwebsite(httpurl)
	if website="" then
		call eerr("域名为空",httpurl)
	end if
	splstr=Array("index.asp","index.aspx","index.php","index.jsp","index.htm","index.html","default.asp","default.aspx","default.jsp","default.htm","default.html")
	for each s in splstr
		url=website & s
		call echo(url,XMLGetStatus(url))
		doevents
	next
end function
'显示获得默认首页界面
sub displayGetDefaultIndex()
%>
<form id="form1" name="form1" method="post" action="?act=getDefaultIndex">
  <p>获得域名默认首页</p>
  <p>域名
    <input type="text" name="httpurl" id="httpurl" value="http://www.dfz9.com/ " />
    <input type="submit" name="button" id="button" value="提交" />
  </p>
</form>
<%
end sub
%>