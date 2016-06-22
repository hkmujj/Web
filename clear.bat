#del JumboECMS.*.dll/s
#del *.pdb/s
#del *.csproj.*.txt/s
#del thumb*.db/s

#删除软件
#del Project1.exe
#移动软件
#ren Project11.exe Project1.exe
#运行软件
#Project1.exe



安装组件
#Regsvr32 system/andsow.dll
#卸载组件
#Regsvr32/u system/andsow.dll

#停止IIS
#iisreset/stop
#启动IIS
#iisreset/start

#关闭进程
#taskkill /f /im explorer.exe
#如果想要关闭后重新开启用下面的
#taskkill /f /im explorer.exe && c:\windows\explorer.exe
#删除进程 记事本
#taskkill /f /im notepad.exe
#打开进程 记事本
#start notepad

#关闭进程 驱动人生
taskkill /f /im DTLService.exe
#关闭进程 微软的设备健康助手管理服务 好像是交易安全
taskkill /f /im DhPluginMgr.exe
#关闭进程 微软的设备健康助手管理服务 好像是交易安全
taskkill /f /im DhMachineSvc.exe
#关闭进程 微软MicrosoftWindows服务器和工作站套装相关程序,用于支持其相关兼容
taskkill /f /im unsecapp.exe
#关闭进程 阿里巴巴反钓鱼服务
taskkill /f /im TaobaoProtect.exe
#关闭进程 服务视频
taskkill /f /im PPAP.exe
#关闭进程  UC浏览器服务
taskkill /f /im UCService.exe
#关闭进程  UC浏览器
taskkill /f /im UCBrowser.exe
#关闭进程  ctfmon.exe是Microsoft Office产品套装的一部分，是有关输入法的一个可执行程序
taskkill /f /im ctfmon.exe
#关闭进程  支付宝
taskkill /f /im Alipaybsm.exe
#关闭进程  音频设备图形隔离
taskkill /f /im audiodg.exe
#关闭进程  
taskkill /f /im myev.exe

#关闭驱动里的软件
taskkill /f /im hdtool.exe

#关闭进程  其它
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

#系统自带的播放器
taskkill wmpnetwk.exe


start D:\"Program Files (x86)"\Adobe\"Adobe Dreamweaver CS3"\Dreamweaver.exe
start C:\"Program Files (x86)"\Tencent\QQ\QQProtect\Bin\QQProtect.exe
#start C:\"Program Files"\MySQL\"MySQL Server 5.1"\bin\mysqld.exe

