<% 
'添加于 20160203

'获取浏览器类型(可以判断:47种浏览器;GoogLe,Grub,MSN,Yahoo!蜘蛛;十种常见IE插件)  Response.Write getBrType("")
Function getBrType(theInfo)
    Dim strType, tmp1, s 
    s = "Other Unknown" 
    If theInfo = "" Then
        theInfo = UCase(Request.ServerVariables("HTTP_USER_AGENT")) 
    End If 
    If InStr(theInfo, UCase("mozilla")) > 0 Then s = "Mozilla" 
    If InStr(theInfo, UCase("icab")) > 0 Then s = "iCab" 
    If InStr(theInfo, UCase("lynx")) > 0 Then s = "Lynx" 
    If InStr(theInfo, UCase("links")) > 0 Then s = "Links" 
    If InStr(theInfo, UCase("elinks")) > 0 Then s = "ELinks" 
    If InStr(theInfo, UCase("jbrowser")) > 0 Then s = "JBrowser" 
    If InStr(theInfo, UCase("konqueror")) > 0 Then s = "konqueror" 
    If InStr(theInfo, UCase("wget")) > 0 Then s = "wget" 
    If InStr(theInfo, UCase("ask jeeves")) > 0 Or InStr(theInfo, UCase("teoma")) > 0 Then s = "Ask Jeeves/Teoma" 
    If InStr(theInfo, UCase("wget")) > 0 Then s = "wget" 
    If InStr(theInfo, UCase("opera")) > 0 Then s = "opera" 
    If InStr(theInfo, UCase("NOKIAN")) > 0 Then s = "NOKIAN(诺基亚手机)" 
    If InStr(theInfo, UCase("SPV")) > 0 Then s = "SPV(多普达手机)" 
    If InStr(theInfo, UCase("Jakarta Commons")) > 0 Then s = "Jakarta Commons-HttpClient" 
    If InStr(theInfo, UCase("Gecko")) > 0 Then
        strType = "[Gecko] " 
        s = "Mozilla Series" 
        If InStr(theInfo, UCase("aol")) > 0 Then s = "AOL" 
        If InStr(theInfo, UCase("netscape")) > 0 Then s = "Netscape" 
        If InStr(theInfo, UCase("firefox")) > 0 Then s = "FireFox" 
        If InStr(theInfo, UCase("chimera")) > 0 Then s = "Chimera" 
        If InStr(theInfo, UCase("camino")) > 0 Then s = "Camino" 
        If InStr(theInfo, UCase("galeon")) > 0 Then s = "Galeon" 
        If InStr(theInfo, UCase("k-meleon")) > 0 Then s = "K-Meleon" 
        s = strType & s 
    End If 
    If InStr(theInfo, UCase("bot")) > 0 Or InStr(theInfo, UCase("crawl")) > 0 Then
        strType = "[Bot/Crawler]" 
        If InStr(theInfo, UCase("grub")) > 0 Then s = "Grub" 
        If InStr(theInfo, UCase("googlebot")) > 0 Then s = "GoogleBot" 
        If InStr(theInfo, UCase("msnbot")) > 0 Then s = "MSN Bot" 
        If InStr(theInfo, UCase("slurp")) > 0 Then s = "Yahoo! Slurp" 
        s = strType & s 
    End If 
    If InStr(theInfo, UCase("applewebkit")) > 0 Then
        strType = "[AppleWebKit]" 
        s = "" 
        If InStr(theInfo, UCase("omniweb")) > 0 Then s = "OmniWeb" 
        If InStr(theInfo, UCase("safari")) > 0 Then s = "Safari" 
        s = strType & s 
    End If 
    If InStr(theInfo, UCase("msie")) > 0 Then
        strType = "[MSIE" 
        tmp1 = Mid(theInfo,(InStr(theInfo, UCase("MSIE")) + 4), 6) 
        tmp1 = Left(tmp1, InStr(tmp1, ";") - 1) 
        strType = strType & tmp1 & "]" 
        s = "Internet Explorer" 
        s = strType & s 
    End If 
    If InStr(theInfo, UCase("msn")) > 0 Then s = "MSN" 
    If InStr(theInfo, UCase("aol")) > 0 Then s = "AOL" 
    If InStr(theInfo, UCase("webtv")) > 0 Then s = "WebTV" 
    If InStr(theInfo, UCase("myie2")) > 0 Then s = "MyIE2" 
    If InStr(theInfo, UCase("maxthon")) > 0 Then s = "Maxthon(傲游浏览器)" 
    If InStr(theInfo, UCase("gosurf")) > 0 Then s = "GoSurf(冲浪高手浏览器)" 
    If InStr(theInfo, UCase("netcaptor")) > 0 Then s = "NetCaptor" 
    If InStr(theInfo, UCase("sleipnir")) > 0 Then s = "Sleipnir" 
    If InStr(theInfo, UCase("avant browser")) > 0 Then s = "AvantBrowser" 
    If InStr(theInfo, UCase("greenbrowser")) > 0 Then s = "GreenBrowser" 
    If InStr(theInfo, UCase("slimbrowser")) > 0 Then s = "SlimBrowser" 
    If InStr(theInfo, UCase("360SE")) > 0 Then s = s & "-360SE(360安全浏览器)" 
    If InStr(theInfo, UCase("QQDownload")) > 0 Then s = s & "-QQDownload(QQ下载器)" 
    If InStr(theInfo, UCase("TheWorld")) > 0 Then s = s & "-TheWorld(世界之窗浏览器)" 
    If InStr(theInfo, UCase("icafe8")) > 0 Then s = s & "-icafe8(网维大师网吧管理插件)" 
    If InStr(theInfo, UCase("TencentTraveler")) > 0 Then s = s & "-TencentTraveler(腾讯TT浏览器)" 
    If InStr(theInfo, UCase("baiduie8")) > 0 Then s = s & "-baiduie8(百度IE8.0)" 
    If InStr(theInfo, UCase("iCafeMedia")) > 0 Then s = s & "-iCafeMedia(网吧网媒趋势插件)" 
    If InStr(theInfo, UCase("DigExt")) > 0 Then s = s & "-DigExt(IE5允许脱机阅读模式特殊标记)" 
    If InStr(theInfo, UCase("baiduds")) > 0 Then s = s & "-baiduds(百度硬盘搜索)" 
    If InStr(theInfo, UCase("CNCDialer")) > 0 Then s = s & "-CNCDialer(数控拨号)" 
    If InStr(theInfo, UCase("NOKIAN85")) > 0 Then s = s & "-NOKIAN85(诺基亚手机)" 
    If InStr(theInfo, UCase("SPV_C600")) > 0 Then s = s & "-SPV_C600(多普达C600)" 
    If InStr(theInfo, UCase("Smartphone")) > 0 Then s = s & "-Smartphone(Windows Mobile for Smartphone Edition 操作系统的智能手机)" 
    getBrType = s 
End Function 
%> 
