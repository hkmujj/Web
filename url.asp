<%
'elseif request("title")="1111111111" then
'	response.Redirect("11111")

dim url
'�÷�  /url.asp?title=qq313801120
if request("title")="�ֻ�Bվ" then
	response.Redirect("http://sharembweb.com/Tools/mobile/bwebold")
	
elseif request("title")="�ֻ�Cվ" then
	response.Redirect("http://sharembweb.com/Tools/mobile/cweb") 

elseif request("title")="������(2015)" then
	response.Redirect("http://www.dfz9.com/")

elseif request("title")="�Ͼ�΢ս��(2014)" then
	response.Redirect("http://www.wzl99.com/")
	
elseif request("title")="�Ͼ���Ȱ�¡(2014)" then
	response.Redirect("http://www.863health.com/")
	
elseif request("title")="�Ͼ�Ԫ��(2014)" then
	response.Redirect("http://ylkj11.com/")
	
elseif request("title")="����(2014)" then
	response.Redirect("http://www.jfh6666.com/")
	
	
	
elseif request("title")="�Ͼ���˼��(2013)" then
	response.Redirect("http://maiside.net/")
	
elseif request("title")="��ҵģ����վһ(2010)" then
	response.Redirect("http://www.wxjiebao.com/")
	
elseif request("title")="��ҵģ����վ��(2010)" then
	response.Redirect("http://www.laxiang8.com/")
	
elseif request("title")="��ҵģ����վ��(2010)" then
	response.Redirect("http://www.021chaijiu.com/")
  
elseif request("title")="qqȺ35915100" then
	response.Redirect("http://shang.qq.com/wpa/qunwpa?idkey=253822bd485c454811141c731156d2ecd4dba04ecf647ce81dc97d16a563137b")	

elseif request("gotourl")<>"" then
	url = "http://" & request("gotourl")
	response.Redirect(url)

'20����ASP��֤�� v1.0(ASPԴ��)
elseif request("down")="admin5_20vericode" then				'/url.asp?down=fwvv_20vericode
	response.Redirect("http://down.admin5.com/asp/106437.html")	
elseif request("down")="chinaz_20vericode" then
	response.Redirect("http://down.chinaz.com/soft/35264.htm")	
elseif request("down")="csdn_20vericode" then
	response.Redirect("http://download.csdn.net/detail/mydd3/6712723")	
elseif request("down")="jb51_20vericode" then
	response.Redirect("http://www.jb51.net/codes/118319.html")	
elseif request("down")="codesky_20vericode" then
	response.Redirect("http://www.codesky.net/codedown/html/29007.htm")	
elseif request("down")="onlinedown_20vericode" then
	response.Redirect("http://www.onlinedown.net/softdown/537626_2.htm")	
elseif request("down")="mycodes_20vericode" then
	response.Redirect("http://www.mycodes.net/40/5912.htm")	
elseif request("down")="fwvv_20vericode" then
	response.Redirect("http://www.fwvv.net/Software/View-Software-33811.shtml")
	
'ASPתPHP���� v1.0(ASPԴ��) 			'/url.asp?down=wei2008_asptophpv1
elseif request("down")="codedown_asptophpv1" then
	response.Redirect("http://www.codesky.net/codedown/html/30224.htm")
elseif request("down")="csdn_asptophpv1" then
	response.Redirect("http://download.csdn.net/detail/mydd3/9399270")
elseif request("down")="jb51_asptophpv1" then
	response.Redirect("http://www.jb51.net/codes/419460.html")
elseif request("down")="codesc_asptophpv1" then
	response.Redirect("http://www.codesc.net/source/6026.shtml")
elseif request("down")="asp300_asptophpv1" then
	response.Redirect("http://www.asp300.com/SoftView/10/SoftView_59634.html")
elseif request("down")="gpxz_asptophpv1" then
	response.Redirect("http://www.gpxz.com/yuanma/asp/qita/936898.html")
elseif request("down")="662p_asptophpv1" then
	response.Redirect("http://code.662p.com/view/12607.html")
elseif request("down")="wei2008_asptophpv1" then
	response.Redirect("http://www.wei2008.com/downinfo/90364.html")
elseif request("down")="ylwap_asptophpv1" then				'20160304
	response.Redirect("http://ylwap.cn/thread-65.htm")
	
'ASPPHPCMS V1.0
'elseif request("down")="jzku_ASPPHPCMSv1" then		 
'	response.Redirect("http://www.jzku.com/asp/2016/0226/719.html")
 
	
'��������
elseif request("baidu")<>"" then			'/url.asp?yahoo=ASP��֤�����
	response.Redirect("https://www.baidu.com/s?ie=gb2312&word=" & request("baidu"))	
elseif request("haosou")<>"" then 
	response.Redirect("http://www.haosou.com/s?ie=gb2312&q=" & request("haosou"))	
elseif request("sogou")<>"" then
	response.Redirect("https://www.sogou.com/sogou?query=" & request("sogou"))	
elseif request("yahoo")<>"" then
	response.Redirect("https://search.yahoo.com/search;_ylt=A86.JmbkJatWH5YARmebvZx4?p="& request("yahoo") &"&toggle=1&cop=mss&ei=gb2312&fr=yfp-t-901&fp=1")	
elseif request("act")="downaspvbs" then
	response.Redirect("http://sharembweb.com/Tools/downfile.asp?act=download&downfile=�z�ĸs�R���t�s�Q�r���t")
elseif request("act")="fangzhan" then
	url = request("selectServer") & "?act=downweb&httpurl=" & request("httpurl") & "&verificationTime=" &  XorEnc(Now(), 31380) & "&Char_Set=" & request("Char_Set")
	response.Redirect(url)
end if

'http://sharembweb.com/Tools/downfile.asp?act=download&downfile=%u7ABB%u7AC0%u7AFB%u7AFB%u7AF8%u7AE7%u7ABB%u7AEE%u7AFD%u7AE4%u7ABB%u7AE7%u7AFC%u7AF5%u7AE6%u7AF1%u7AF9%u7AF6%u7AE3%u7AF1%u7AF6%u7AD5%u7AC4%u7AC4%u7ABA%u7AF5%u7AE4%u7AFF


'Xor����
Function xorEnc(code, n)
    Dim c, s1, s2, s3, i 
    c = code 
    s1 = Len(c) : s3 = ""
    For i = 0 To s1 - 1
        s2 = AscW(Right(c, s1 - i)) Xor n 
        s3 = s3 & ChrW(Int(s2)) 
    Next 
    'Chr(34) ���ǵ���(") ��ֹ���� ��Ϊ"��ASP�����
    s3 = Replace(s3, ChrW(34), "��") 
    xorEnc = s3 
End Function

%>