<% 
'����� 20160203

'��ȡ���������(�����ж�:47�������;GoogLe,Grub,MSN,Yahoo!֩��;ʮ�ֳ���IE���)  Response.Write getBrType("")
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
    If InStr(theInfo, UCase("NOKIAN")) > 0 Then s = "NOKIAN(ŵ�����ֻ�)" 
    If InStr(theInfo, UCase("SPV")) > 0 Then s = "SPV(���մ��ֻ�)" 
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
    If InStr(theInfo, UCase("maxthon")) > 0 Then s = "Maxthon(���������)" 
    If InStr(theInfo, UCase("gosurf")) > 0 Then s = "GoSurf(���˸��������)" 
    If InStr(theInfo, UCase("netcaptor")) > 0 Then s = "NetCaptor" 
    If InStr(theInfo, UCase("sleipnir")) > 0 Then s = "Sleipnir" 
    If InStr(theInfo, UCase("avant browser")) > 0 Then s = "AvantBrowser" 
    If InStr(theInfo, UCase("greenbrowser")) > 0 Then s = "GreenBrowser" 
    If InStr(theInfo, UCase("slimbrowser")) > 0 Then s = "SlimBrowser" 
    If InStr(theInfo, UCase("360SE")) > 0 Then s = s & "-360SE(360��ȫ�����)" 
    If InStr(theInfo, UCase("QQDownload")) > 0 Then s = s & "-QQDownload(QQ������)" 
    If InStr(theInfo, UCase("TheWorld")) > 0 Then s = s & "-TheWorld(����֮�������)" 
    If InStr(theInfo, UCase("icafe8")) > 0 Then s = s & "-icafe8(��ά��ʦ���ɹ�����)" 
    If InStr(theInfo, UCase("TencentTraveler")) > 0 Then s = s & "-TencentTraveler(��ѶTT�����)" 
    If InStr(theInfo, UCase("baiduie8")) > 0 Then s = s & "-baiduie8(�ٶ�IE8.0)" 
    If InStr(theInfo, UCase("iCafeMedia")) > 0 Then s = s & "-iCafeMedia(������ý���Ʋ��)" 
    If InStr(theInfo, UCase("DigExt")) > 0 Then s = s & "-DigExt(IE5�����ѻ��Ķ�ģʽ������)" 
    If InStr(theInfo, UCase("baiduds")) > 0 Then s = s & "-baiduds(�ٶ�Ӳ������)" 
    If InStr(theInfo, UCase("CNCDialer")) > 0 Then s = s & "-CNCDialer(���ز���)" 
    If InStr(theInfo, UCase("NOKIAN85")) > 0 Then s = s & "-NOKIAN85(ŵ�����ֻ�)" 
    If InStr(theInfo, UCase("SPV_C600")) > 0 Then s = s & "-SPV_C600(���մ�C600)" 
    If InStr(theInfo, UCase("Smartphone")) > 0 Then s = s & "-Smartphone(Windows Mobile for Smartphone Edition ����ϵͳ�������ֻ�)" 
    getBrType = s 
End Function 
%> 
