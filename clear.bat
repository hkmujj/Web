#del JumboECMS.*.dll/s
#del *.pdb/s
#del *.csproj.*.txt/s
#del thumb*.db/s

#ɾ�����
#del Project1.exe
#�ƶ����
#ren Project11.exe Project1.exe
#�������
#Project1.exe



��װ���
#Regsvr32 system/andsow.dll
#ж�����
#Regsvr32/u system/andsow.dll

#ֹͣIIS
#iisreset/stop
#����IIS
#iisreset/start

#�رս���
#taskkill /f /im explorer.exe
#�����Ҫ�رպ����¿����������
#taskkill /f /im explorer.exe && c:\windows\explorer.exe
#ɾ������ ���±�
#taskkill /f /im notepad.exe
#�򿪽��� ���±�
#start notepad

#�رս��� ��������
taskkill /f /im DTLService.exe
#�رս��� ΢����豸�������ֹ������ �����ǽ��װ�ȫ
taskkill /f /im DhPluginMgr.exe
#�رս��� ΢����豸�������ֹ������ �����ǽ��װ�ȫ
taskkill /f /im DhMachineSvc.exe
#�رս��� ΢��MicrosoftWindows�������͹���վ��װ��س���,����֧������ؼ���
taskkill /f /im unsecapp.exe
#�رս��� ����Ͱͷ��������
taskkill /f /im TaobaoProtect.exe
#�رս��� ������Ƶ
taskkill /f /im PPAP.exe
#�رս���  UC���������
taskkill /f /im UCService.exe
#�رս���  UC�����
taskkill /f /im UCBrowser.exe
#�رս���  ctfmon.exe��Microsoft Office��Ʒ��װ��һ���֣����й����뷨��һ����ִ�г���
taskkill /f /im ctfmon.exe
#�رս���  ֧����
taskkill /f /im Alipaybsm.exe
#�رս���  ��Ƶ�豸ͼ�θ���
taskkill /f /im audiodg.exe
#�رս���  
taskkill /f /im myev.exe

#�ر�����������
taskkill /f /im hdtool.exe

#�رս���  ����
taskkill /f /im PresentationFontCache.exe
taskkill /f /im igfxCUIService.exe
taskkill /f /im mDNSResponder.exe
taskkill /f /im TBSecSvc.exe
taskkill /f /im LiveUpdate360.exe
taskkill /f /im QQProtect.exe
taskkill /f /im RAVCpl64.exe

#WPS Clouds
taskkill /f /im wpscloudsvr.exe
taskkill /f /im PresentationFontCache.exe

#ϵͳ�Դ��Ĳ�����
taskkill wmpnetwk.exe


start D:\"Program Files (x86)"\Adobe\"Adobe Dreamweaver CS3"\Dreamweaver.exe
start C:\"Program Files (x86)"\Tencent\QQ\QQProtect\Bin\QQProtect.exe
#start C:\"Program Files"\MySQL\"MySQL Server 5.1"\bin\mysqld.exe

