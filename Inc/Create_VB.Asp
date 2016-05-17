<!--#Include File = "Config.Asp"--> 
<% 
'选择处理 
Select Case Request("act")
	Case "CreateVBDialog","vb1" : CreateVBDialog()	'创建VB需要的面板
End Select
Response.End()
Sub CreateVBDialog()
	Dim SplStr,I,S,C,Content,RootFolder,zFormFolder,DataFolder,Path,ToPath,FormName,BasList,FrmList,ToolsMenu,FrmListStr,OpenFrom,nTools,ToolsClick,MenuTitle
	Dim YesSortPlug,TempStr,TableConfig,FormStrList,FormSetList,TableName,SoftVersion,ThisMDBPath,ThisAccessPass,CtlFileList,StartStr,EndStr
	Dim ConfigFolder
	Dim HelpFolder			'帮助文件夹
	Dim ToolsName			'工具名称 给那个动态菜单加菜单20150218 
	Dim nTools2				'工具二排序
	Dim nTools3				'网站制作排序
	dim tempContent
	SoftVersion = "Version1.1"
	RootFolder = "/VB工程/"
	'复制文件夹 免得忘记保存而后悔20150313
	Dim ToFolder			'复制到文件夹
	If CheckFolder(RootFolder) Then
		For I=0 To 90
			S = "/VB工程_"& I &"/"
			If CheckFolder(S)=False And ToFolder="" Then
				ToFolder=S
			ElseIf CheckFolder(S)=True And ToFolder<>"" Then
				ToFolder=""
			End If
		Next
	End If
	'Call EErr(RootFolder, ToFolder)
	Call MoveFolder(RootFolder, ToFolder)
	Call DeleteFolder(RootFolder)
	Call CreateFolder(RootFolder)
	zFormFolder = RootFolder & "zForm/"
	Call CreateFolder(zFormFolder)
	HelpFolder = RootFolder & "Help/"
	Call CreateFolder(HelpFolder)	
	Call CopyFile("/DataDir/VB模块/Help/SystemLogin.ini", HelpFolder & "SystemLogin.ini")
	Call CopyFile("/DataDir/VB模块/Help/Js处理.html", HelpFolder & "Js处理.html")
	
	Call CopyFile("/DataDir/VB模块/右键下拉菜单代码.txt", RootFolder & "右键下拉菜单代码.txt")
	Call CopyFile("/DataDir/VB模块/快速登录.vbs", RootFolder & "快速登录.vbs")
	
	
	
	'ConfigFolder = RootFolder & "Config/"
	'Call CreateFolder(ConfigFolder)
	
	c="|Print.Asp|Access.Asp|Cai.Asp|EncDec.Asp|Function.Asp|GetWord.Asp|RegExp.Asp|StringNumber.Asp|XML.Asp|Time.Asp|URL.Asp|FSO.Asp|Check.Asp|Html.Asp|AdodbStream.Asp|Dictionary.Asp|PinYin.Asp|2014_Array.Asp|2014_Css.Asp|2014_GBUTF.Asp|2015_Formatting.Asp|2015_Param.Asp|2015_ASP.Asp|2015_Color.Asp|FunHTML.Asp|2015_PHP.Asp|2015_WebSql.Asp|2015_hexSwitch.Asp|ASPPHPAccess.Asp|2014_Author.Asp|2016_WebControl.asp|2014_News.Asp|2014_MainInfo.Asp|2014_SiteMap.Asp|2014_Template.Asp|"
	c=c & "2014_Search.Asp|2014_Class.Asp|Incpage.Asp|2014_Banner.Asp|2014_Links.Asp|2014_Doc.Asp|URL.Asp|2015_APGeneral.Asp|2015_ToMyPHP.Asp|IE.asp|"
	
	dim searchStr,replaceStr
	'添加模块列表
	SplStr=Split(c,"|")'|ZClassAspCode.Asp
	For Each Path In SplStr
		If Path <> "" And 1=1 Then
			FormName = Mid(Path,1,InStrRev(Path,".")-1)
			'判断文件以数据打头，则加前缀 否则VB调用会出错
			If InStr("0123456789",Left(FormName,1))>0 Then
				FormName="y_" & FormName							'区部它是前面有数字
			End If
			Content = ReadFile(Path,"")		
			
			searchStr = getStrCut(Content,"'footer start","'footer end",2)
			if searchStr<>"" then
				replaceStr=vbcrlf & getftext("\DataDir\VB模块\2015_Param.Asp.txt")
				Content=replace(Content,searchStr,replaceStr)
				call echo(path,"替换footer成功")
			end if
			
			for i =1 to 10
			searchStr = getStrCut(Content,"'vbdel start","'vbdel end",1)
				if searchStr<>"" then
					Content=replace(Content,searchStr,"")
				end if
			next
			
			'组件
			If Path = "ZClassAspCode.Asp" Then
				ToPath = zFormFolder & FormName & ".cls"
				Content = StrCut(Content,"'VB开始","'VB结束",2) 
				C=""
				C=C & "VERSION 1.0 CLASS" & vbCrlf
				C=C & "BEGIN" & vbCrlf
				C=C & "  MultiUse = -1  'True" & vbCrlf
				C=C & "  Persistable = 0  'NotPersistable" & vbCrlf
				C=C & "  DataBindingBehavior = 0  'vbNone" & vbCrlf
				C=C & "  DataSourceBehavior  = 0  'vbNone" & vbCrlf
				C=C & "  MTSTransactionMode  = 0  'NotAnMTSObject" & vbCrlf
				C=C & "END" & vbCrlf
				C=C & "Attribute VB_Name = ""Z_"& FormName &"""" & vbCrlf			'免得名称与里面函数重名
				C=C & "Attribute VB_GlobalNameSpace = False" & vbCrlf
				C=C & "Attribute VB_Creatable = False" & vbCrlf
				C=C & "Attribute VB_PredeclaredId = False" & vbCrlf
				C=C & "Attribute VB_Exposed = False" & vbCrlf
				Content=C & Content
				
				BasList=BasList & "class="& FormName &"; zForm\"& FormName &".cls" & vbCrlf
			'模块
			Else
				ToPath = zFormFolder & FormName & ".bas"
				Content = StrCut(Content,"<" & "%","%" & ">",2)
				Content = "Attribute VB_Name = ""Z_"& FormName &"""" & vbCrlf & Content
				BasList=BasList & "Module=Z_"& FormName &"; zForm\"& FormName &".bas" & vbCrlf
			End If
			tempContent=content
			
			content=ASPToVB(content) 
			
			if Content="" then
				call eerr(FormName,tempContent)
			end if
			
			startStr="Function GetThisUrl()"
			endStr="End Function"
			s=GetStrCut(content,startStr,endStr,1)
			if s<>"" then
				content=replace(content,s,  startStr & vbcrlf & endStr)
			end if
			
			
			Content = Replace(Content,vbTab,"    ")			'替换Tab
			Call Echo(FormName,ToPath)
			Call CreateFile(ToPath,Content)
		End If
	Next
	Path = "/DataDir/VB模块/Manage.vbp"
	Content = GetFText(Path)
	'处理掉不需要的注释代码
	Splstr=Split(Content,vbCrlf) : Content = ""
	For Each S In SplStr
		If Left(Trim(S),1)<>"'" Then
			Content=Content & S & vbCrlf
		End If
	Next
	Content = Replace(Content,"{Bas列表}",BasList)
	'添加窗体列表
	FrmListStr = "MDIForm1|Notepad|FrmHide|Login|IE|AutoAddTable|WebLog|WebIP|Eval|SortPlug|Test|Share|ShareEdition2|QQShareConfig|QQShareUrl|FrmEdit|QQShareSendUrl|YingZheng|"
	FrmListStr="MDIForm1|Notepad|FrmHide|Login|IE|Share|ShareEdition2|QQShareConfig|QQShareUrl|QQShareSendUrl|FrmEdit"
	'全二
	FrmListStr="MDIForm1|FrmHide|FrmEdit|Login|TableDialog|Notepad|IE|-|ShareEdition2|QQShareConfig|QQShareUrl|QQShareSendUrl|-|FrmCai|CaiManage|UpDateWebConfig|-|XiCiWebDialog|XiCiWebMDB|-|CaiBaiduUrl|BaiduCai|-|#SortPlug|WebCrack|Test|DownList|IEDebug|-|AutoAddTable|Ying|-|SortPlug|PseudoOriginal|-|Scan|WebShell|Eval|WebServer|-|QQ|Qzone|-|"
	'QQ操作 'FrmListStr="MDIForm1|FrmHide|FrmEdit|Login|TableDialog|Notepad|IE|InputLocalFile|InputSqlTable|-|HTML5CSS3|FrmCrateNewWeb|TemplateCodeCreate|-|Test|IEDebug|-|WebLog|WebIP|WebSite|-|Eval|-|FrmCai|CaiManage|UpDateWebConfig|-|SinglePage|IEAutomatic|-|SearchEngine|SearchEngineInfoList|WebAdminScan|StatisticalData|-|WebManage|WebIeBox|-|MYProgram|-|SmallTool|ColorPicker|TxtEdit|-|ShieldWebSiteList|WebSiteList|SqlServerWebUrl|WebSiteInfo|WebShell|-|FrmSortPlug|jiayuan|"
	FrmListStr=      "MDIForm1|FrmHide|FrmEdit|Login|TableDialog|Notepad|IE|InputLocalFile|InputSqlTable|-|FrmCrateNewWeb|WebUrlScan|FrmCreateFunction|WebIeBox|-|Test|IEDebug|IEAutomatic|IEAutoHandle|-|Eval|-|FrmCai|CaiManage|-|WebAdminScan|ExcellentUrl|SmallTool|ColorPicker|TxtEdit|-|FrmSortPlug|AskDa|ZNRobot|JobWeb|JobCompany|FormJobCai|ComputerDir|analogBatchUP|" 
	FrmListStr=FrmListStr & "|FrmCrateNewHtmlPage|SinglePage|FrameLayout|FrmWebTemplateView|FrmCreateWebLable|UserMakeNav|--2|FileScan2015|HTML5CSS3|XyWebUp|MYProgram|--2|FrmAutoAddWebDefaultInfo|--2|FrmControlDialog|FrmSqlIn|FrmProgrameTransformation|IMitationWebInGoogleSEO|WebShell|FrmEvScan|FrmTestDialog1|FrmTestDialog2|FrmDown|FrmCrateNewHtmlPageOLD|HTML5CSS3|"

	nTools = 0	'工具菜单默认为零
	SplStr=Split(FrmListStr,"|")
	For Each S In SplStr
		'不为空  #不用
		If S <> "" And Left(S,1)<>"#" Then
			'设置操作菜单
			If InStr("|-|IE|AutoAddTable|WebLog|WebIP|WebSite|Eval|SortPlug|Test|Share|QQShareConfig|QQShareUrl|QQShareSendUrl|ShareEdition2|FrmCai|CaiManage|UpDateWebConfig|XiCiWebDialog|XiCiWebMDB|IEDebug|IEAutomatic|IEAutoHandle|BaiduCai|CaiBaiduUrl|Ying|PseudoOriginal|Scan|WebShell|WebServer|WebCrack|DownList|QQ|QQName|QQVisitors|Qzone|BaiduZhidao|BaiduZhidaoWords|StatisticalData|BaiduZhidaoData|BaiduZhidaoWordsData|BaiduZhidaoMain|SearchEngine|SearchEngineInfoList|WebAdminScan|WebManage|WebIeBox|SinglePage|FrmSMS|SMSInfo|LocaDir|MYProgram|SmallTool|HTML5CSS3|FrmCrateNewWeb|FrmCrateNewHtmlPageOLD|TemplateCodeCreate|WebSiteList|SqlServerWebUrl|ShieldWebSiteList|AsynchDown|WebSiteInfo|FrmSortPlug|ExcellentUrl|WebUrlScan|FrmHangUp|XyWebUp|FrameLayout|FrmCreateFunction|QQJournal|AskDa|ZNRobot|JobWeb|FormJobCai|UserMakeNav|JobCompany|IMitationWebInGoogleSEO|FrmCrateNewHtmlPage|FileScan2015|FrmCreateWebLable||FrmWebTemplateView|ComputerDir|FrmAutoAddWebDefaultInfo|FrmControlDialog|FrmSqlIn|FrmProgrameTransformation|analogBatchUP|jiayuan|FrmEvScan|FrmTestDialog1|FrmTestDialog2|FrmDown|","|"& S &"|") > 0 Or Left(S,2)="--" Then
				MenuTitle="" : TableConfig="" : TableName="" : ThisAccessPass="" : ThisMDBPath=""
				
	

				If S = "SortPlug" Then
					MenuTitle = "模拟点击"
					YesSortPlug = "OK"				'为作弊软件特殊使用					
					Call CopyFile("/DataDir/VB模块/Action.ini", RootFolder & "Action.ini")			'拷贝配置文件
					
				
				ElseIf S = "UserMakeNav" Then
					MenuTitle = "网站导航生成"
				ElseIf S = "FileScan2015" Then
					MenuTitle = "文件搜索2015"
				ElseIf S = "FrmControlDialog" Then
					MenuTitle = "控制面板"
				ElseIf S = "FrmSqlIn" Then
					MenuTitle = "Sql注入"
				ElseIf S = "FrmProgrameTransformation" Then
					MenuTitle = "程序转换"
				ElseIf S = "analogBatchUP" Then
					MenuTitle = "模似批量上传" 
				ElseIf S = "FrmEvScan" Then
					MenuTitle = "EvScan"
				ElseIf S = "FrmTestDialog1" Then
					MenuTitle = "测试面板1" 
				ElseIf S = "FrmTestDialog2" Then
					MenuTitle = "测试面板2"
				ElseIf S = "FrmDown" Then
					MenuTitle = "下载文件"
					
					
					
					
				ElseIf S = "IMitationWebInGoogleSEO" Then
					MenuTitle = "仿站在Google优化"
					 
					Call CopyFolder("/DataDir/VB模块/IMitationWebInGoogleSEO",RootFolder & "/IMitationWebInGoogleSEO")			'拷贝文件夹
					
					
					
				ElseIf S = "XyWebUp" Then
					MenuTitle = "批量上传网站"
				ElseIf S = "FrmCreateFunction" Then
					MenuTitle = "生成函数面板"
				ElseIf S = "FrmSortPlug" Then
					MenuTitle = "模拟操作软件"				
				ElseIf S = "FrameLayout" Then
					MenuTitle = "网页框架布局"
				ElseIf S = "FrmHangUp" Then
					MenuTitle = "禁止窗口必需验证"					
				ElseIf S = "MYProgram" Then
					MenuTitle = "我的程序"
					Call CopyFile("/DataDir/VB模块/AspFunction.ini", HelpFolder & "/AspFunction.ini")			'拷贝配置文件
				ElseIf S = "SmallTool" Then
					MenuTitle = "小工具"
				ElseIf S = "HTML5CSS3" Then
					MenuTitle = "Html5+Css3快速生成工具"
					CtlFileList = CtlFileList & "UserControl=zForm\ControlDialog.ctl" & vbCrlf		
					Call CopyFile("\DataDir\VB模块\ControlDialog.ctl", zFormFolder & "ControlDialog.ctl")			'复制文件
					Call CopyFile("\DataDir\VB模块\ControlDialog.ctx", zFormFolder & "ControlDialog.ctx")			'复制文件				
				ElseIf S = "FrmCrateNewWeb" Then
					MenuTitle = "创建新网页"				
				ElseIf S = "FrmCrateNewHtmlPageOLD" Then
					MenuTitle = "创建新网页(旧版)"		
							
				ElseIf S = "FrmCreateWebLable" Then
					MenuTitle = "生成网站嵌套标签"	 
				ElseIf S = "FrmCrateNewHtmlPage" Then
					MenuTitle = "创建HTML页面"
				ElseIf S = "TemplateCodeCreate" Then
					MenuTitle = "模板代码创建生成器"
				ElseIf S = "FrmWebTemplateView" Then
					MenuTitle = "网站模板预览" 
				ElseIf S = "FrmAutoAddWebDefaultInfo" Then
					MenuTitle = "自动添加网站默认信息"  
					
				ElseIf S = "IE" Then
					MenuTitle = "浏览器"
				ElseIf S = "Test" Then
					MenuTitle = "测试窗口"
				ElseIf S = "AutoAddTable" Then
					MenuTitle = "自动填表"
				ElseIf S = "YingZheng" Then
					MenuTitle = "破解验证码"
				ElseIf S = "Ying" Then
					MenuTitle = "简易验证码识别系统"					
					Call CopyFolder("/DataDir/VB模块/图片破解", RootFolder & "图片破解")		'加载配置文件夹
				ElseIf S = "XiCiWebDialog" Then
					MenuTitle = "西祠论坛群发"	
				ElseIf S = "FrmSMS" Then
					MenuTitle = "短信"					
				ElseIf S = "SMSInfo" Then
					MenuTitle = "短信管理"	
					TableName = "SMSInfo"
					ThisAccessPass="****"			
					ThisMDBPath = " ""I:\个人制作\个人\短信\SMS.mdb"""
					TableConfig = TableConfig & "发信人|DXName|1500|True[分]"
					TableConfig = TableConfig & "收信人|ToName|1500|True[分]"
					TableConfig = TableConfig & "发信时期|AddDate|1000|True[分]" 	
					TableConfig = TableConfig & "留言内容|Content|4000|True[分]" 		
				ElseIf S = "LocaDir" Then
					MenuTitle = "本地目录"	
					TableName = "LocaDir" 	
					ThisMDBPath = "App.Path & ""/../../网站备份\数据库\LocaDir.Mdb""" 
					ThisMDBPath="""SqlServer数据库"""
					TableConfig = TableConfig & "文件夹|FolderPath|5000|True[分]"	
					TableConfig = TableConfig & "文件名称|FileName|3500|True[分]"	
					TableConfig = TableConfig & "文件大小|FileSize|1200|True[分]"	
					TableConfig = TableConfig & "文件类型|FileType|1200|True[分]" 
				ElseIf S = "ComputerDir" Then
					MenuTitle = "电脑Scan"	
					TableName = "ComputerDir" 	
					ThisMDBPath = "App.Path & ""/../../网站备份\数据库\ComputerDir.mdb"""  
					TableConfig = TableConfig & "大类名称|BigClassName|1600|True[分]"			
					TableConfig = TableConfig & "文件路径|FilePath|3500|True[分]"	
					TableConfig = TableConfig & "文件名称|FileName|1200|True[分]"	
					TableConfig = TableConfig & "文件大小|FileSize|1200|True[分]" 
					TableConfig = TableConfig & "文件类型|FileType|1200|True[分]" 
					
					
					
				ElseIf S = "XiCiWebMDB" Then
					MenuTitle = "西祠论坛数据管理"	
					TableName = "XiCiWeb"
					TableConfig = TableConfig & "网址|HttpUrl|6000|True[分]"
					TableConfig = TableConfig & "标题|Title|3000|True[分]"
					TableConfig = TableConfig & "是否使用|Through|1000|True[分]" 			 
				ElseIf S = "WebLog" Then
					MenuTitle = "网站日志"	
					TableName = "WebLog"					
					ThisMDBPath = "GetIniS(""init"", ""DefaultQQAccessPath"", """") & ""\WebLog.Mdb"""
					TableConfig = TableConfig & "来访网址|GoUrl|3500|True[分]"
					TableConfig = TableConfig & "当前浏览网址|HttpUrl|3500|True[分]"
					TableConfig = TableConfig & "访问时间|GoDateTime|2000|True[分]" 
					TableConfig = TableConfig & "搜索引擎类型|SearchEngineType|1500|True[分]"
					TableConfig = TableConfig & "搜索关键词|SearchWord|2000|True[分]"
					TableConfig = TableConfig & "IP|IP|1600|True[分]" 
					TableConfig = TableConfig & "IE|IE|2200|True[分]" 	
					TableConfig = TableConfig & "日期|AddDate|1200|True[分]"	
					TableConfig = TableConfig & "域名|WebSite|2000|True[分]" 			
				ElseIf S = "WebIP" Then
					MenuTitle = "网站流量"
					TableName = "WebIP"
					ThisMDBPath = "GetIniS(""init"", ""DefaultQQAccessPath"", """") & ""\WebLog.Mdb"""
					TableConfig = TableConfig & "添加日期|AddDate|3000|True[分]"
					TableConfig = TableConfig & "浏览次数|PVCount|1000|True[分]"
					TableConfig = TableConfig & "域名数|WebSiteCount|1000|True[分]" 
					TableConfig = TableConfig & "IP|IPCount|1000|True[分]" 
					TableConfig = TableConfig & "IE|IECount|1000|True[分]" 					
					TableConfig = TableConfig & "百度数|Baidu|1000|True[分]" 
					TableConfig = TableConfig & "Google数|Google|1000|True[分]" 
					TableConfig = TableConfig & "Yahoo数|Yahoo|1000|True[分]" 
					TableConfig = TableConfig & "搜搜数|Soso|1000|True[分]" 
					TableConfig = TableConfig & "360数|Cn360|1000|True[分]" 	
					TableConfig = TableConfig & "QQ空间数|Qzone|1000|True[分]" 	
					TableConfig = TableConfig & "微信数|WeiXin|1000|True[分]" 	
					

				ElseIf S = "jiayuan" Then
					MenuTitle = "世纪佳缘"
					TableName = "jiayuan"
					ThisMDBPath = """../../网站备份/数据库/jiayuan.mdb"""
					TableConfig = TableConfig & "会员ID|uid|1000|True[分]"
					TableConfig = TableConfig & "昵称|nickname|1500|True[分]"
					TableConfig = TableConfig & "性别|sex|800|True[分]" 
					TableConfig = TableConfig & "年龄|Age|1000|True[分]" 
					TableConfig = TableConfig & "身高|height|800|True[分]" 	
					TableConfig = TableConfig & "体重|Weight|1000|True[分]" 	 
					TableConfig = TableConfig & "婚姻|marriage|1500|True[分]" 		
					TableConfig = TableConfig & "教育|education|1500|True[分]" 		
					TableConfig = TableConfig & "工资|minMonthlySalary|1000|True[分]" 
					TableConfig = TableConfig & "兄弟姐妹|BrothersAndSisters|1500|True[分]" 
					TableConfig = TableConfig & "个人介绍|shortnote|2000|True[分]" 
					
					
					
				ElseIf S = "AskDa" Then
					MenuTitle = "遇到问题和解决"
					TableName = "AskDa"
					ThisMDBPath = """../../网站备份/数据库/AskDa.mdb"""
					TableConfig = TableConfig & "问|Ask|5000|True[分]" 
					TableConfig = TableConfig & "答|Da|5000|True[分]" 
					TableConfig = TableConfig & "添加日期|AddDate|3000|True[分]" 
				ElseIf S = "ZNRobot" Then
					MenuTitle = "智慧机器人"
					TableName = "ZNRobot"
					ThisMDBPath = """../../网站备份/数据库/ZNRobot.mdb"""
					TableConfig = TableConfig & "大类名称|BigClassName|1600|True[分]"			
					TableConfig = TableConfig & "标题|Title|4000|True[分]"	
					TableConfig = TableConfig & "相关标签|RelatedTags|4000|True[分]"
				ElseIf S = "JobWeb" Then
					MenuTitle = "招聘岗位管理"
					TableName = "JobWeb"
					ThisMDBPath = """../../网站备份/数据库/JobAccess.mdb"""
					TableConfig = TableConfig & "大类名称|BigClassName|1600|True[分]" 				
					TableConfig = TableConfig & "搜索关键词|SearchWord|1300|True[分]" 				
					TableConfig = TableConfig & "标题|Title|3500|True[分]"	
					TableConfig = TableConfig & "城市|City|800|True[分]"
					TableConfig = TableConfig & "区域|Area|1000|True[分]"
					TableConfig = TableConfig & "公司规模|CompanyScale|1200|True[分]"
					TableConfig = TableConfig & "经验|Experience|1200|True[分]"
					TableConfig = TableConfig & "学历|Education|1200|True[分]"
					TableConfig = TableConfig & "职位月薪|MonthlySalary|1200|True[分]"
					TableConfig = TableConfig & "招聘时间|JobTime|1200|True[分]" 
					TableConfig = TableConfig & "公司名称|CompanyName|3800|True[分]" 
				ElseIf S = "JobCompany" Then
					MenuTitle = "招聘公司管理"
					TableName = "JobCompany"
					ThisMDBPath = """../../网站备份/数据库/JobAccess.mdb"""			
					TableConfig = TableConfig & "公司名称|CompanyName|3000|True[分]" 
					TableConfig = TableConfig & "公司网址|CompanyUrl|3000|True[分]"
					TableConfig = TableConfig & "城市|City|800|True[分]"
					TableConfig = TableConfig & "区域|Area|1000|True[分]"
					TableConfig = TableConfig & "公司规模|CompanyScale|1200|True[分]"
					TableConfig = TableConfig & "公司性质|CompanyNature|1200|True[分]"
					TableConfig = TableConfig & "公司行业|CompanyIndustry|4000|True[分]" 
					TableConfig = TableConfig & "最新发布|NewJobTime|1600|True[分]"					
					TableConfig = TableConfig & "发布天数|ReleaseJobDay|1200|True[分]"
					
				ElseIf S = "FormJobCai" Then
					MenuTitle = "招聘网站采集" 
				ElseIf S = "WebSite" Then
					MenuTitle = "域名统计"
					TableName = "WebSite"
					ThisMDBPath = "GetIniS(""init"", ""DefaultQQAccessPath"", """") & ""\WebLog.Mdb"""
					TableConfig = TableConfig & "域名|WebSite|5000|True[分]" 				
					TableConfig = TableConfig & "数量|WebSiteCount|1000|True[分]"
					
				ElseIf S = "WebUrlScan" Then
					MenuTitle = "网站链接Scan"
					TableName = "WebUrlScan"
					ThisMDBPath = "App.Path & ""\..\..\网站备份\数据库\WebUrlScan.Mdb"""
					TableConfig = TableConfig & "大类名称|BigClassName|1600|True[分]" 				
					TableConfig = TableConfig & "标题|Title|4000|True[分]"
					TableConfig = TableConfig & "网页状态|WebState|1200|True[分]"	
					TableConfig = TableConfig & "打开时间|OpenSpeed|1200|True[分]"			
					TableConfig = TableConfig & "文件大小|WebFileSize|1200|True[分]"
					TableConfig = TableConfig & "链接类型|LinkType|1200|True[分]"		'分内链与外链
				ElseIf S = "ExcellentUrl" Then
					MenuTitle = "优秀网站"
					TableName = "ExcellentUrl"
					ThisMDBPath = "App.Path & ""\..\..\网站备份\数据库\ExcellentUrl.Mdb"""
					'TableConfig = TableConfig & "大类|BigClassName|1600|True[分]" 			
					'TableConfig = TableConfig & "网址|HttpUrl|4000|True[分]"					
					TableConfig = TableConfig & "域名|WebSite|3000|True[分]"
					TableConfig = TableConfig & "标题|Title|8000|True[分]"
					TableConfig = TableConfig & "IP|IP|2200|True[分]"
					'TableConfig = TableConfig & "打开时间|OpenSpeed|1200|True[分]"			
					'TableConfig = TableConfig & "文件大小|WebFileSize|1200|True[分]" 
					
				ElseIf S = "SearchEngine" Then
					MenuTitle = "搜索引擎"
					TableName = "SearchEngine"
					'ThisMDBPath = "App.Path & ""/../../网站备份\数据库\SearchEngine.Mdb"""	'本地Access数据库
					ThisMDBPath = """SqlServer"""	'本地SqlServer数据库
					'TableConfig = TableConfig & "大类|BigClassName|1500|True[分]" 				
					TableConfig = TableConfig & "网址|HttpUrl|4000|True[分]"
					TableConfig = TableConfig & "标题|Title|6000|True[分]"
					TableConfig = TableConfig & "域名|WebSite|3000|True[分]"
					TableConfig = TableConfig & "收录时间|IncludedTime|1500|True[分]"
					
					TableConfig = TableConfig & "Table数|TableNumb|1300|True[分]"
					TableConfig = TableConfig & "Div数|DivNumb|1000|True[分]"
					TableConfig = TableConfig & "Span数|SpanNumb|1000|True[分]"
					TableConfig = TableConfig & "Ul数|UlNumb|1000|True[分]"
					TableConfig = TableConfig & "Dl数|DlNumb|1000|True[分]"
				ElseIf S = "SqlServerWebUrl" Then
					MenuTitle = "网站网址"
					TableName = "WebUrl"
					ThisMDBPath = """SqlServer"""	'本地SqlServer数据库
					TableConfig = TableConfig & "网址|HttpUrl|4000|True[分]"
					TableConfig = TableConfig & "域名|WebSite|3000|True[分]"
					TableConfig = TableConfig & "域名ID|WebSiteID|1000|True[分]"
					TableConfig = TableConfig & "状态|WebStatus|1200|True[分]"
					TableConfig = TableConfig & "打开时间|OpenTime|1800|True[分]"
					TableConfig = TableConfig & "处理|Handle|1000|True[分]"
				ElseIf S = "WebSiteList" Then
					MenuTitle = "域名列表"
					TableName = "WebSiteList"
					ThisMDBPath = """SqlServer"""	'本地SqlServer数据库
					
					TableConfig = TableConfig & "大类名称|BigClassName|1500|True[分]"
					TableConfig = TableConfig & "域名|WebSite|3000|True[分]"
					TableConfig = TableConfig & "状态|WebStatus|1000|True[分]"
					TableConfig = TableConfig & "IP|IP|2200|True[分]"
					TableConfig = TableConfig & "IP地址|IPAddress|2000|True[分]"
					TableConfig = TableConfig & "打开时间|OpenTime|1400|True[分]"
					TableConfig = TableConfig & "处理|Handle|1000|True[分]"
				ElseIf S = "WebSiteInfo" Then
					MenuTitle = "域名列表"
					TableName = "WebSiteInfo"
					ThisMDBPath = """SqlServer"""	'本地SqlServer数据库
					TableConfig = TableConfig & "域名|WebSite|3000|True[分]"
					TableConfig = TableConfig & "总页数|WebCountPage|1000|True[分]"
					TableConfig = TableConfig & "首页内链数|HomeStationLink|1500|True[分]"
					TableConfig = TableConfig & "首页外链数|HomeInboundLink|1500|True[分]"
					TableConfig = TableConfig & "全部内链数|StationLink|1500|True[分]"
					TableConfig = TableConfig & "全部外链数|InboundLink|1500|True[分]"					
					
				ElseIf S = "ShieldWebSiteList" Then
					MenuTitle = "屏蔽域名列表"
					TableName = "ShieldWebSiteList"
					ThisMDBPath = """SqlServer"""	'本地SqlServer数据库
					TableConfig = TableConfig & "域名|WebSite|3000|True[分]"
					TableConfig = TableConfig & "IP|IP|2000|True[分]"
					TableConfig = TableConfig & "状态|WebStatus|1200|True[分]"
					TableConfig = TableConfig & "打开时间|OpenTime|1800|True[分]"
					TableConfig = TableConfig & "处理|Handle|1000|True[分]"						
				ElseIf S = "SearchEngineInfoList" Then
					MenuTitle = "搜索引擎列表"
					TableName = "SearchEngineInfoList"
					ThisMDBPath = "App.Path & ""/../../网站备份\数据库\SearchEngineInfoList.Mdb"""
					TableConfig = TableConfig & "大类|BigClassName|1500|True[分]" 				
					'TableConfig = TableConfig & "网址|HttpUrl|4000|True[分]"
					TableConfig = TableConfig & "标题|Title|6000|True[分]"
					TableConfig = TableConfig & "域名|WebSite|3000|True[分]"
					TableConfig = TableConfig & "收录时间|IncludedTime|1300|True[分]"
					TableConfig = TableConfig & "页数|PageNumber|800|True[分]"
					TableConfig = TableConfig & "排名|Ranking|800|True[分]"					
				ElseIf S = "WebAdminScan" Then
					MenuTitle = "网站后台Scan"
					TableName = "WebAdminScan"
					ThisMDBPath = "App.Path & ""/../../网站备份\数据库\WebAdminScan.Mdb"""
					'TableConfig = TableConfig & "大类|BigClassName|1500|True[分]" 				
					TableConfig = TableConfig & "网址|HttpUrl|7000|True[分]"					
					TableConfig = TableConfig & "响应状态|WebState|1200|True[分]"						
					TableConfig = TableConfig & "打开时间|OpenSpeed|1200|True[分]"			
					TableConfig = TableConfig & "文件大小|WebFileSize|1200|True[分]"	 
				ElseIf S = "Eval" Then
					MenuTitle = "Eval"
				ElseIf S = "WebManage" Then
					MenuTitle = "模板网页制作" 
				ElseIf S = "WebIeBox" Then
					MenuTitle = "模板样式管理"  
'				ElseIf S = "Share" Then
'					MenuTitle = "自动分享"
				ElseIf S = "ShareEdition2" Then
					MenuTitle = "QQ点赞批量分享 改进版"  
				ElseIf S = "QQShareConfig" Then
					MenuTitle = "QQ分享配置"
					TableName = "QQShareConfig"				 
					TableConfig = TableConfig & "网站标题|WebTitle|3000|True[分]"
					TableConfig = TableConfig & "网站描述|WebDescription|3000|True[分]"
					TableConfig = TableConfig & "网站图片地址|ImgPath|3000|True[分]"
					TableConfig = TableConfig & "评语|PingYu|2000|True[分]"
					TableConfig = TableConfig & "是否使用|Through|2000|True[分]"					
				ElseIf S = "QQShareUrl" Then
					MenuTitle = "QQ分享网址"
					TableName = "QQShareUrl"				 
					TableConfig = TableConfig & "网址|HttpUrl|6000|True[分]" 
					TableConfig = TableConfig & "审核|Through|800|True[分]"
					
				ElseIf S = "QQShareSendUrl" Then
					MenuTitle = "已分享网址"
					TableName = "QQShareSendUrl"				 
					TableConfig = TableConfig & "已分享网址|HttpUrl|6000|True[分]" 
					TableConfig = TableConfig & "通过|Through|800|True[分]"
					
				ElseIf S = "FrmCai" Then
					MenuTitle = "采集"	
				ElseIf S = "CaiManage" Then
					MenuTitle = "采集管理"
					TableName = "Cai"
					ThisMDBPath = "App.Path & ""/../../网站备份\数据库\Cai.Mdb"""
					TableConfig = TableConfig & "大类名称|BigClassName|2000|True[分]"
					TableConfig = TableConfig & "小类名称|SmallClassName|2000|True[分]"
					TableConfig = TableConfig & "子类名称|ThreeClassName|2000|True[分]"
					TableConfig = TableConfig & "标题|Title|4000|True[分]"
					TableConfig = TableConfig & "网址|HttpUrl|4000|True[分]" 
					TableConfig = TableConfig & "内容|Content|4000|True[分]" 
				ElseIf S = "UpDateWebConfig" Then
					MenuTitle = "批量上传到网站配置"
					TableName = "UpDateWebConfig"
					TableConfig = TableConfig & "批量上传到网站|WebSite|4000|True[分]"
					TableConfig = TableConfig & "网站标题|WebTitle|2000|True[分]"
					TableConfig = TableConfig & "网站关键词|WebKeywords|2000|True[分]"
					TableConfig = TableConfig & "网站描述|WebDescription|2000|True[分]"
					TableConfig = TableConfig & "文章追加后缀|SuffixCustom|2000|True[分]"
					TableConfig = TableConfig & "上传|Through|800|True[分]" 
				ElseIf S = "BaiduCai" Then
					MenuTitle = "百度网址管理"
					TableName = "BaiduCai"
					TableConfig = TableConfig & "大类名称|BigClassName|2000|True[分]"
					TableConfig = TableConfig & "小类名称|SmallClassName|2000|True[分]"
					TableConfig = TableConfig & "子类名称|ThreeClassName|2000|True[分]"
					TableConfig = TableConfig & "标题|Title|2000|True[分]"
					TableConfig = TableConfig & "百度网址|BaiduUrl|2000|True[分]"
					TableConfig = TableConfig & "网址|HttpUrl|2000|True[分]" 
					TableConfig = TableConfig & "域名|WebSite|2000|True[分]" 
				ElseIf S = "BaiduZhidaoMain" Then
					MenuTitle = "百度知道Main" 					
				ElseIf S = "BaiduZhidaoData" Then
					MenuTitle = "百度知道数据"
					TableName = "BaiduZhidaoData"					
					ThisMDBPath = "DefaultQQAccessPath & ""\BaiduZhidaoData.Mdb"""
					TableConfig = TableConfig & "百度ID|ZhidaoID|2200|True[分]"
					TableConfig = TableConfig & "标题|Title|6000|True[分]"
					TableConfig = TableConfig & "回答数|ReplyNumb|2000|True[分]"
					TableConfig = TableConfig & "支持顶数|SupportNumb|2000|True[分]" 
				ElseIf S = "BaiduZhidaoWordsData" Then
					MenuTitle = "百度知道关键词数据"
					TableName = "BaiduZhidaoWordsData"					
					ThisMDBPath = "DefaultQQAccessPath & ""\BaiduZhidaoWordsData.Mdb"""
					TableConfig = TableConfig & "关键词|Word|2000|True[分]"
					
				ElseIf S = "WebShell" Then
					MenuTitle = "WebShell管理"
					TableName = "WebShell"
					TableConfig = TableConfig & "大类名称|BigClassName|1800|True[分]"
					TableConfig = TableConfig & "小马|Ev1|6500|True[分]"
					TableConfig = TableConfig & "域名|WebSite|3200|True[分]"				
					TableConfig = TableConfig & "通过|Through|800|True[分]"					
				ElseIf S = "WebServer" Then
					MenuTitle = "WebServer网站服务器"
					TableName = "WebServer" 
					TableConfig = TableConfig & "IP|IP|2000|True[分]" 					
					TableConfig = TableConfig & "组件|Assembly|2000|True[分]" 		
					TableConfig = TableConfig & "安全级别|AnQuan|1200|True[分]"		
				ElseIf S = "QQ" Then
					MenuTitle = "QQ管理"
					TableName = "QQ" 
					ThisMDBPath = "DefaultQQAccessPath & ""\QQ.Mdb"""
					TableConfig = TableConfig & "QQ|QQ|1500|True[分]" 					
					TableConfig = TableConfig & "QQ名称|QQName|2000|True[分]" 		
					TableConfig = TableConfig & "性别|QQSex|1000|True[分]"
					TableConfig = TableConfig & "年龄|Age|1000|True[分]"
					'TableConfig = TableConfig & "QQ年龄|QQAge|1000|True[分]"
					TableConfig = TableConfig & "空间极限|SpacePurview|1000|True[分]"
					TableConfig = TableConfig & "相册数|PhotoNumb|1000|True[分]"
					TableConfig = TableConfig & "说说数|ShuaShuaNumb|1000|True[分]"
					TableConfig = TableConfig & "日志数|BlogNumb|1000|True[分]"
					'TableConfig = TableConfig & "首页显访问|HomeVisitors|1000|True[分]"
					'TableConfig = TableConfig & "访问权限|SpacePurview|1200|True[分]"
					TableConfig = TableConfig & "赞|QzoneZai|800|True[分]"
					TableConfig = TableConfig & "等级|QzoneLevel|800|True[分]"
					TableConfig = TableConfig & "积分|QzoneIntegral|800|True[分]"
					
				ElseIf S = "QQName" Then
					MenuTitle = "QQ名称管理"
					TableName = "QQName" 
					ThisMDBPath = "DefaultQQAccessPath & ""\QQName.Mdb"""
					TableConfig = TableConfig & "QQ|QQ|1500|True[分]" 					
					TableConfig = TableConfig & "QQ名称|QQName|2000|True[分]"
				ElseIf S = "QQVisitors" Then
					MenuTitle = "QQ访客管理"
					TableName = "QQVisitors" 
					ThisMDBPath = "DefaultQQAccessPath & ""\QQVisitors.Mdb"""
					TableConfig = TableConfig & "QQ|QQ|1500|True[分]" 					
					TableConfig = TableConfig & "访客QQ|FQQ|15000|True[分]"
					
					
				ElseIf S = "QQJournal" Then
					MenuTitle = "QQ日志"
					TableName = "QQJournal" 
					ThisMDBPath = """SqlServer"""	'本地SqlServer数据库
					
					TableConfig = TableConfig & "QQ|QQ|1500|True[分]" 	
							
					TableConfig = TableConfig & "类型|cate|1500|True[分]" 		
					TableConfig = TableConfig & "日志ID|blogId|1500|True[分]" 	
					TableConfig = TableConfig & "标题|title|3000|True[分]" 		
					TableConfig = TableConfig & "日志ID|blogId|1500|True[分]" 	
					TableConfig = TableConfig & "发布时间|pubTime|1500|True[分]" 		 
					
					
					
				ElseIf S = "BaiduZhidao" Then
					MenuTitle = "百度知道管理"
					TableName = "BaiduZhidao" 
					ThisMDBPath = "DefaultQQAccessPath & ""\BaiduZhidao.Mdb"""
					TableConfig = TableConfig & "大类名称|BigClassName|2000|True[分]"
					TableConfig = TableConfig & "提问标题|Title|4000|True[分]"
					TableConfig = TableConfig & "提问网址|HttpUrl|2000|True[分]"
					TableConfig = TableConfig & "发布者|Author|1800|True[分]"
					TableConfig = TableConfig & "回答数|ReplyNumb|1000|True[分]"
					TableConfig = TableConfig & "回答者|ReplyName|1800|True[分]"
					TableConfig = TableConfig & "采纳率|AdoptionRate|1000|True[分]"
					TableConfig = TableConfig & "支持数|SupportNumb|1000|True[分]" 
					TableConfig = TableConfig & "浏览次数|BrowseNumb|1200|True[分]"						
					TableConfig = TableConfig & "通过|Through|800|True[分]"		
				ElseIf S = "BaiduZhidaoWords" Then
					MenuTitle = "百度知道相关词管理"
					TableName = "BaiduZhidaoWords" 
					ThisMDBPath = "DefaultQQAccessPath & ""\BaiduZhidaoWords.Mdb"""
					TableConfig = TableConfig & "搜索词|SearchWord|3000|True[分]"
					TableConfig = TableConfig & "相关词|RelatedWord|3000|True[分]"
					TableConfig = TableConfig & "排序|nSort|800|True[分]"
					
				ElseIf S = "StatisticalData" Then
					MenuTitle = "数据统计管理"
					TableName = "StatisticalData" 
					ThisMDBPath = "DefaultQQAccessPath & ""\StatisticalData.Mdb"""
					TableConfig = TableConfig & "大类名称|BigClassName|2000|True[分]"
					TableConfig = TableConfig & "标题|Title|4000|True[分]"
					TableConfig = TableConfig & "数值|CountNumb|2000|True[分]"
					
				ElseIf S = "CaiBaiduUrl" Then
					MenuTitle = "采集百度网址"
				ElseIf S = "IEDebug" Then
					MenuTitle = "IE调试"
				ElseIf S = "WebCrack" Then
					MenuTitle = "网站破解"
				ElseIf S = "DownList" Then
					MenuTitle = "批量下载"					
							
				ElseIf S = "PseudoOriginal" Then
					MenuTitle = "仿原创工具"
					Call CreateFolder(RootFolder & "/Data")
					Call CopyFile("\DataDir\VB模块\服务器\Dict字典未用\默认字典.txt", RootFolder & "/DataDir/默认字典.txt")			'默认字典
					Call CopyFile("\DataDir\VB模块\服务器\Dict字典未用\字符替换字典.txt", RootFolder & "/DataDir/字符替换字典.txt")			'字符替换字典
					Call CopyFile("\DataDir\VB模块\服务器\Dict字典未用\简体繁体转换.txt", RootFolder & "/DataDir/简体繁体转换.txt")			'简体繁体转换 
				ElseIf S = "Scan" Then
					MenuTitle = "Scan扫描"				
				ElseIf S = "Qzone" Then
					MenuTitle = "QQ空间"
				ElseIf S = "SinglePage" Then
					MenuTitle = "网站单页制作"
				ElseIf S = "IEAutomatic" Then
					MenuTitle = "IE自动操作"
					Call CopyFile("\DataDir\VB模块\任务动作代码.ini", RootFolder & "任务动作代码.ini")			'复制文件
					Call CopyFile("\DataDir\VB模块\问问关键词列表.ini", RootFolder & "问问关键词列表.ini")		'复制文件
					Call CopyFile("\DataDir\VB模块\自动运行动作.ini", RootFolder & "自动运行动作.ini")			'复制文件
					Call CopyFile("\DataDir\VB模块\网站后台登录字典.ini", RootFolder & "网站后台登录字典.ini")	'复制文件
					
				ElseIf S = "IEAutoHandle" Then
					MenuTitle = "IE自动处理2015"
				ElseIf S = "AsynchDown" Then
					MenuTitle = "测试自动下载"
					
				ElseIf S = "-" Or Left(S,2)="--" Then
					MenuTitle = "-"								
				End If
				
				'排除分割线 顶部显示菜单名称
				If S <>"-" And Left(S,2) <> "--" Then 
					ToolsMenu=ToolsMenu & "    ElseIf S = """& MenuTitle &""" Then" & vbCrlf 
					ToolsMenu=ToolsMenu & "        " & S &".Show: "& S &".SetFocus" & vbCrlf
					'为加载窗体则加个配置
					If TableConfig <> "" Then
						
						ToolsMenu=ToolsMenu & "        If " & S & ".Caption <> ""$表面板$"" Then Exit Sub" & vbCrlf '判断不重复加载
						ToolsMenu=ToolsMenu & "        " & S & ".ThisFormName=""" & S & """" & vbCrlf				'窗体名称
						ToolsMenu=ToolsMenu & "        " & S & ".Caption=""" & MenuTitle & """" & vbCrlf				'显示窗体名称
						ToolsMenu=ToolsMenu & "        " & S & ".FrameManage.Caption=""" & MenuTitle & """" & vbCrlf				'显示窗体框架名称						
						ToolsMenu=ToolsMenu & "        " & S & ".TableName=""" & TableName & """" & vbCrlf			'表名称
						
						
						If ThisAccessPass<>"" Then ToolsMenu=ToolsMenu & "        " & S & ".ThisAccessPass=""" & ThisAccessPass & """" & vbCrlf			'数据库密码
						'If ThisMDBPath <> "" Then ToolsMenu=ToolsMenu & "        " & S & ".ThisMDBPath=HandlePath(" & ThisMDBPath & ")" & vbCrlf					'数据库路径
						If ThisMDBPath <> "" Then ToolsMenu=ToolsMenu & "        " & S & ".ThisMDBPath="& ThisMDBPath & vbCrlf					'数据库路径
						 
						
						ToolsMenu=ToolsMenu & "        " & S & ".TableConfig=""" & TableConfig & """" & vbCrlf
						ToolsMenu=ToolsMenu & "        " & S & ".TableConfig=""" & TableConfig & """" & vbCrlf
						
						ToolsMenu=ToolsMenu & "        Call " & S & ".FormLoad()" & vbCrlf
						
					End If
				End If
				'加载数据窗体
				If TableConfig<> "" Then
					FormStrList = FormStrList & vbCrlf & "Public "& S &" As Form"		'表格窗体定义
					FormSetList = FormSetList & vbCrlf & "Set " & S & "=New TableDialog"	'加载窗体
				End If 
				 '这里设置下拉按钮放在哪里一组20150610
				 If InStr("|AskDa|ZNRobot|JobWeb|JobCompany|FormJobCai|FrmTestDialog1|FrmTestDialog2|FrmDown|","|"& S &"|")>0 Then
					'工具菜单
					If nTools2 = 0 Then
						ToolsClick=ToolsClick & "    if checkAdminQuanXian("""& MenuTitle &""") then Frm.Tools2Menu(0).Caption = """& MenuTitle &"""" & vbCrlf
					Else
						ToolsClick=ToolsClick & "    if checkAdminQuanXian("""& MenuTitle &""") then Load Frm.Tools2Menu("& nTools2 &"): Frm.Tools2Menu("& nTools2 &").Caption = """& MenuTitle &"""" & vbCrlf
					End if 
					nTools2 = nTools2 + 1 
				 '网站制作分组20150610
				 ElseIf InStr("|UserMakeNav|SinglePage|FrmCrateNewHtmlPage|FileScan2015|FrameLayout|FrmCreateWebLable|HTML5CSS3|XyWebUp|MYProgram|FrameLayout|FrmCreateWebLable|FrmWebTemplateView|--2|FrmAutoAddWebDefaultInfo|FrmControlDialog|FrmSqlIn|FrmProgrameTransformation|WebShell|FrmEvScan|","|"& S &"|")>0 Then
					'工具菜单
					If nTools3 = 0 Then
						ToolsClick=ToolsClick & "    if checkAdminQuanXian("""& MenuTitle &""") then Frm.Tools3Menu(0).Caption = """& MenuTitle &"""" & vbCrlf
					Else
						ToolsClick=ToolsClick & "    if checkAdminQuanXian("""& MenuTitle &""") then Load Frm.Tools3Menu("& nTools3 &"): Frm.Tools3Menu("& nTools3 &").Caption = """& MenuTitle &"""" & vbCrlf
					End if 
					nTools3 = nTools3 + 1 
				Else
					'工具菜单
					If nTools = 0 Then
						ToolsClick=ToolsClick & "    if checkAdminQuanXian("""& MenuTitle &""") then Frm.ToolsMenu(0).Caption = """& MenuTitle &"""" & vbCrlf
					Else
						ToolsClick=ToolsClick & "    if checkAdminQuanXian("""& MenuTitle &""") then Load Frm.ToolsMenu("& nTools &"): Frm.ToolsMenu("& nTools &").Caption = """& MenuTitle &"""" & vbCrlf
					End if 
					nTools = nTools + 1 
				End If
			End If
			
			'20150328			
			If InStr("|WebManage|HTML5CSS3|FrmCrateNewWeb|FrmCrateNewHtmlPageOLD|","|"& S &"|")>0 Then
				'拷贝Template文件夹
				Call CopyFolder("\DataDir\VB模块\Template", RootFolder & "/Template")			'复制文件夹 
			End If
			 
			'默认打开窗体  排除分割线
			If S <> "MDIForm1" And S <> "-"  And Left(S,2) <> "--" Then
				OpenFrom=OpenFrom & "    ElseIf Tag = LCase("""& S &""") Then" & vbCrlf & "        " & S &".Show: "& S &".SetFocus" & vbCrlf
			End If
			'排除分割线，并且数据配置表为空
			If S <>"-" And Left(S,2) <> "--" And TableConfig = "" Then
				S = S & ".frm"
				If CheckFile("/DataDir/VB模块/"& S)=False Then
					Call Eerr("没有窗体","/DataDir/VB模块/"& S)
				End If
				FrmList = FrmList & "Form=zForm\"& S & vbCrlf
				Call CopyFile("/DataDir/VB模块/"& S, zFormFolder & S)
				'拷贝数数
				S=Replace(S,".frm",".frx")
				Call CopyFile("/DataDir/VB模块/"& S, zFormFolder & S)
			End If
		End If
	Next

	'异步下载模块
	CtlFileList = CtlFileList & "UserControl=zForm\Downloader.ctl" & vbCrlf		
	Call CopyFile("\DataDir\VB模块\Downloader.ctl", zFormFolder & "Downloader.ctl")			'复制文件
	Call CopyFile("\DataDir\VB模块\Downloader.ctx", zFormFolder & "Downloader.ctx")		    '复制文件
	'自定义参数模块
	CtlFileList = CtlFileList & "UserControl=zForm\ControlParam.ctl" & vbCrlf		
	Call CopyFile("\DataDir\VB模块\ControlParam.ctl", zFormFolder & "ControlParam.ctl")			'复制文件
	Call CopyFile("\DataDir\VB模块\ControlParam.ctx", zFormFolder & "ControlParam.ctx")		    '复制文件
	'自定义网址模块
	CtlFileList = CtlFileList & "UserControl=zForm\ControlUrl.ctl" & vbCrlf		
	Call CopyFile("\DataDir\VB模块\ControlUrl.ctl", zFormFolder & "ControlUrl.ctl")			'复制文件
	Call CopyFile("\DataDir\VB模块\ControlUrl.ctx", zFormFolder & "ControlUrl.ctx")		    '复制文件
  
	
	
	Content = Replace(Content,"Manage.exe","网站制作软件.exe")
	'Content = Replace(Content,"QQ313801120","www.maiside.net")		'替换成相应联系方式20150316	
'	Call Eerr("FrmList",FrmList)
	Content = Replace(Content,"{Frm列表}",FrmList & CtlFileList)
	Call CreateFile(RootFolder & "Manage.vbp", Content)
	Call Echo("FormStrList",FormStrList)
	'MDIForm1.frm处理
	Path = "/DataDir/VB模块/MDIForm1.frm"
	Content = GetFText(Path) 
	Content = Replace(Content,"{工具}",ToolsMenu)
	Content = Replace(Content,"{工具导航}",ToolsClick)
	Content = Replace(Content,"{默认打开窗体}",OpenFrom) 
	Content = Replace(Content,"{加载窗体}",FormSetList)
	Content = Replace(Content,"Version1.0","网站制作软件")   			'SoftVersion版本号
	'伪原创按钮处理
	If InStr(FrmListStr,"PseudoOriginal")=0 Then
		'Call CreateFile("测试文件一会删除.txt",Content)			'这是因为MDIForm1.frm文件被替换的原因
		'Call Rw(Content):Response.End()
		TempStr = StrCut(Content, vbCrlf & "'伪原创","End Sub",1)
		Content = Replace(Content,TempStr, vbCrlf)
	End If
	
	
	If YesSortPlug = "OK" Then		
		Content = Replace(Replace(Content,"[SortPlug开始]",""),"[SortPlug结束]","")		
	Else
		TempStr = StrCut(Content,"[SortPlug开始]","[SortPlug结束]",1)
		Content = Replace(Content,TempStr,"")
	End If 
	Call CreateFile(zFormFolder & "MDIForm1.frm", Content)
	
	'复制基础模块
	SplStr=Split("VB_Img.bas|CloseEXE.bas|Ying1.bas|Ying2.bas|VB_Config.bas|API.bas|VB_Conn.bas|WinMain.bas|Md5.bas|VB_Common.bas|Tp.bas|FindEXE.bas|Input.bas|Rwini.bas|clsWaitableTimer.cls|GPS.cls|Request.cls|AspCode.cls|NetWorkClass.cls|modExcel.bas|VB_ImportAndExportWeb.bas|WebSiteGetIP.bas|cx_extractor.bas|PingIP.bas|RT.bas|DomainNameToIP.bas|VB_Computer.bas|VB_CaiClass.bas|VB_VBControl.bas|ClassRichTextBox.cls|VBJSScript.bas|AspToPhpClass.cls|AspActionClass.cls|ControlTemplateStyle.ctl|VB_ModuleFunction.bas|VB_CallCommand.bas","|")
	For Each S In SplStr
		If S <> "" Then
			Call CopyFile("/DataDir/VB模块/"& S, zFormFolder & S)
			'BasList=BasList & "Module="& FormName &"; zForm\"& FormName &".bas" & vbCrlf
		End If
	Next
	Call CopyFolder("/DataDir/VB模块/Images", RootFolder & "Images") 
	Call CopyFolder("/DataDir/VB模块/Edit", RootFolder & "Edit") 	
	Call CopyFolder("/DataDir/VB模块/Config", RootFolder & "Config")
	'Call CopyFolder("/DataDir/VB模块/dll", RootFolder & "dll")
	 
	
	Call CopyFile("/DataDir/VB模块/SysInfo.ini", RootFolder & "SysInfo.ini") 
	
	
	'定义全局窗体
	Path = zFormFolder & "/VB_Config.bas"
	Content = GetFText(Path)
	S = StrCut(Content,"'{$窗体开始$}","'{$窗体结束$}",1)
	Content = Replace(Content, S , "'{$窗体开始$}" & FormStrList & vbCrlf & "'{$窗体结束$}")
	'替换版本号
	S = StrCut(Content,"Public Const SoftVersion = """,""" '软件版本",1)
	Content = Replace(Content, S , "Public Const SoftVersion = """ & SoftVersion & """ '软件版本")
	Call CreateFile(Path,Content)
	If 1=2 Then
		DataFolder = RootFolder & "DataDir/"
		Call CreateFolder(DataFolder)
		Call CreateFolder(DataFolder & "Log")	'日志文件夹
	
		Call CopyFile("/DataDir/VB模块/最近打开文档.ini", DataFolder & "最近打开文档.ini")
		'Call CopyFolder("/DataDir/VB模块/Eval", RootFolder & "Eval") 
		Call CopyFile("/DataDir/VB模块/sql.txt", DataFolder & "sql.txt")
		Call CopyFile("/DataDir/VB模块/Desk.jpg", DataFolder & "Desk.jpg")		
	End If
End Sub
%>