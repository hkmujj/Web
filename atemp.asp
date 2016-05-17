<!--#Include File = "Inc/_Config.Asp"-->          
<!--#Include File = "web/function.Asp"-->    
<%
dim pub_httpurl
pub_httpurl="http://maiside.net/"
'�������ݴ���ҳ
Select Case Request("act")
	case "displayGetDefaultIndex" : displayGetDefaultIndex(request("httpurl"))	'��ʾ���Ĭ����ҳ����
	case "getDefaultIndex" : getDefaultIndex()	'���Ĭ����ҳ
	
	case "displayGetWebSiteDir" : displayGetWebSiteDir(request("httpurl"))	'��ʾ�����վĿ¼
	case "getWebSiteDir" : getWebSiteDir()	'���Ĭ����ҳ
	
	case "scanWebSiteSafe" : scanWebSiteSafe("http://127.0.0.1")		'ɨ�����վ��ȫ
	
	case else : displayDefaultLayout()		'��ʾĬ�ϲ���
End Select

'ɨ�����վ��ȫ
function scanWebSiteSafe(httpurl)
	dim website,splstr,url,s,c,nState,isAsp,isAspx,isPhp,isJsp
	call echo("ɨ����ַ", httpurl)
	website=getwebsite(httpurl)
	if website="" then
		call eerr("����Ϊ��",httpurl)
	end if
	call echo("����",website)
	isAsp=0:isAspx=0:isPhp=0:isJsp=0
	'�����վʹ����������
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


'��ʾĬ�ϲ���
function displayDefaultLayout()
	call echo("��ʾ���Ĭ����ҳ����","<a href='?act=displayGetDefaultIndex'>����</a>")
	call echo("��ʾ�����վĿ¼","<a href='?act=displayGetWebSiteDir'>����</a>")
	
	call echo("�����վ��ȫ","<a href='?act=scanWebSiteSafe'>����</a>")
	
end function
'���Ŀ¼�б�
function getWebSiteDir(httpurl)
	dim url,splstr,s,website,nState
	website=getwebsite(httpurl)
	if website="" then
		call eerr("����Ϊ��",httpurl)
	end if
	splstr=Array("inc/","admin/","manage/","images/","css/","js/")
	for each s in splstr
		url=website & s
		nState=getHttpUrlState(url)
		call echo(url,nState & "   ("& getHttpUrlStateAbout(nState) &")")
		doevents
	next
end function

'���Ĭ����ҳ����
function getDefaultIndex(httpurl)
	dim url,splstr,s,website,nState
	website=getwebsite(httpurl)
	if website="" then
		call eerr("����Ϊ��",httpurl)
	end if
	splstr=Array("index.asp","index.aspx","index.php","index.jsp","index.htm","index.html","default.asp","default.aspx","default.jsp","default.htm","default.html")
	for each s in splstr
		url=website & s
		nState=getHttpUrlState(url)
		call echo(url,nState & "   ("& getHttpUrlStateAbout(nState) &")")
		doevents
	next
end function
'��ʾ���Ĭ����ҳ����
sub displayGetDefaultIndex()
%>
<form id="form1" name="form1" method="post" action="?act=getDefaultIndex">
  <p>�������Ĭ����ҳ</p>
  <p>����
    <input type="text" name="httpurl" id="httpurl" value="<%=pub_httpurl%>" />
    <input type="submit" name="button" id="button" value="�ύ" />
  </p>
</form>
<%
end sub
'��ʾ�����վĿ¼
sub displayGetWebSiteDir()
%>
<form id="form1" name="form1" method="post" action="?act=getWebSiteDir">
  <p>�����վĿ¼</p>
  <p>����
    <input type="text" name="httpurl" id="httpurl" value="<%=pub_httpurl%>" />
    <input type="submit" name="button" id="button" value="�ύ" />
  </p>
</form>
<%
end sub
%>