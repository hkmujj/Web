<!--#Include virtual = "/Inc/_Config.Asp"--> 
<%
'�������ݴ���ҳ
Select Case Request("act")
	case "displayGetDefaultIndex" : displayGetDefaultIndex()	'��ʾ���Ĭ����ҳ����
	case "getDefaultIndex" : getDefaultIndex()	'���Ĭ����ҳ����
	case else : displayDefaultLayout()		'��ʾĬ�ϲ���
End Select

'��ʾĬ�ϲ���
function displayDefaultLayout()
	call echo("��ʾ���Ĭ����ҳ����","<a href='?act=displayGetDefaultIndex'>����</a>")
	
end function

'���Ĭ����ҳ����
function getDefaultIndex()
	dim url,splstr,s,httpurl,website
	httpurl=request("httpurl")
	website=getwebsite(httpurl)
	if website="" then
		call eerr("����Ϊ��",httpurl)
	end if
	splstr=Array("index.asp","index.aspx","index.php","index.jsp","index.htm","index.html","default.asp","default.aspx","default.jsp","default.htm","default.html")
	for each s in splstr
		url=website & s
		call echo(url,XMLGetStatus(url))
		doevents
	next
end function
'��ʾ���Ĭ����ҳ����
sub displayGetDefaultIndex()
%>
<form id="form1" name="form1" method="post" action="?act=getDefaultIndex">
  <p>�������Ĭ����ҳ</p>
  <p>����
    <input type="text" name="httpurl" id="httpurl" value="http://www.dfz9.com/ " />
    <input type="submit" name="button" id="button" value="�ύ" />
  </p>
</form>
<%
end sub
%>