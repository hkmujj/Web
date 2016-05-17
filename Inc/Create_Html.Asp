<!--#Include File = "Config.Asp"--> 
<%
'VB不引用 Start
Dim UserID                                                             '显示指定网站信息
Dim WebBottom, DisplayYes, pubCode,Code, Did, Sid, Tid, ShowTitle			   '网站信息
Dim TempDid, TempSid, TempTid                                          '给文章用
Dim DefaultWebTitle, DefaultWebKeywords, DefaultWebDescription 		   '默认网站标题与关键词与描述
Dim WebNavType                                                         '网站导航类型
DisplayYes = True                                                      '是否显示HTML内容
Server.ScriptTimeout = 1800                                            '脚本超时
Dim WebActionType                                                      '网站动作类型，给产品展示列表用
Dim WebNavDid, WebNavSid                                               '给产品展示列表用
Dim DefaultWebBanner                                                   '默认网站Banner
Dim WebFolderName                                                      '网站目录名称
Dim LocalTestFolder                                                    '本地测试文件夹 给不同网站生成到不同的文件夹里
Dim NavDid, NavSid, ProDid, ProSid, ProTid                             '导航大类，导航小类，产品大类，产品小类，产品子类
Dim ArticleDid,ArticleSid,ArticleTid,ArticleTitle					   '文章大类名，文章小类名称,文章子类名称，文章标题
Dim ArticleID,ArticleDateTime,ArticleContent,ArticleAuthor			   '文章ID,文章时间,文章内容,文章作者
Dim ArticleBigFiles,ArticleSmallFiles								   '文章大图,文章小图
Dim WebTemplate, WebSkins											   '网站模板文件夹与网站样式文件夹 
Dim WebImages,WebCss,WebJs											   '网站图片路径，网站CSS路径，网站Js路径
Dim TemplateHTMLCode												   '原始HTMLCode
Dim JsUrl,JsContent													   'JS网址与JS内容
Dim Tplpath,TplContent,TplListContent								   '模板块配置路径与内容,列表内容
Dim WebModuleSkins													   '网站模块皮肤20150115
Dim WebFindTpl														   '网站模板被替换内容列表20150115
Dim WebReplaceTpl													   '网站模板替换内容列表20150115
Dim WebCodeFindTpl													   '网站网页被替换内容列表20150115
Dim WebCodeReplaceTpl												   '网站网页替换内容列表20150115
dim isHtmlFormatting													'是否格式化HTML
dim DB_PREFIX															'表前缀20160201
 
Call OpenConn() 
Dim ObjYK : Set ObjYK = New ClassYouKou                                '悠口餐饮技类
Dim ObjMSD : Set ObjMSD = New Maiside								   '悠口麦思德
dim ModuleReplaceArray(99,99)									   '模块替换数组20150712

'&FindTpl=style='312'&ReplaceTpl=style='165'    皮肤切换20150110

'http://127.0.0.1/Inc/Create_Html.Asp?act=Nav&NavDidType=首页&MackHtml=False&Debug=True      &Debug=True去掉 网页则显示静态链接

'处理导航配置
'选择处理
Select Case Request("act")
    Case "auto", "autoweb","AutoAddWebInfo" : AutoAddWebInfo(Request("DataFolderPath"))	'AutoHandleWeb()	'自动生成网站信息	
    Case "Nav"	'导航于20141011 加强了可以指定导航类型 显示网页
		NavDid = Request("NavDid")
		If NavDid="" And Request("NavDidType")<>"" Then
			NavDid = GetNavTypeInName(Request("NavDidType"))
		End If
		'不需要处理这个 是我html调用弄错了 call handleNavConfig(Request("NavDid"),Request("NavSid"))			'处理导航配置
		Call MakeWebHtml(NavDid,Request("NavSid"),"","","","Nav","","","")		'生成指定大类导航
	Case "ShowNavList" : ShowNavList()			'显示生成导航页面 
	Case "CreateNavList" : CreateNavList()		'生成导航列表
	Case "ShowClassList" : ShowClassList()		'生成类型列表
    Case "CreateInfo" : 
		Call MakeWebHtml(Rq("NavDid"),Rq("NavSid"),Rq("ProDid"),Rq("ProSid"),Rq("ProTid"),Rq("ActionType"),Rq("ID"),Rq("ClassName"),Rq("WD"))
	Case "CreateClass" : Call MakeWebHtml("","",Request("ProDid"),Request("ProSid"),Request("ProTid"),"Class","","","")		'生成指定大类导航
	Case "CreateClassList" : CreateClassList()		'生成信息类别列表
	Case "ShowArticleList" : ShowArticleList()		'显示生成文章列表
	Case "CreateArticle" : Call MakeWebHtml("","","","","","Article", Request("Id"),"","")			'生成一篇文章
	Case "CreateArticleList" : CreateArticleList()			'生成全部文章
 	Case "Search" : Call MakeWebHtml("","",Request("Did"),Request("Sid"),Request("Tid"),"Search","",Request("ClassName"),Request("wd"))	'产品搜索
	
    Case "MainInfo" : Call MakeWebHtml("","","","","","MainInfo",Request("Id"),"","")		'生成指定大类导航
	Case "MakeSiteMap" : Call MakeSiteMap():Response.End()		'生成网站地图	
	
	Case "ShowMemberList" : ShowMemberList()		'生成会员列表
	Case "ShowDelHtmllist" : ShowDelHtmllist()		'显示删除HTML面板（创建于20140307）
	Case "DelFile": DelFile(Request("FileName"))	'删除指定文件
	 
	Case "ShowWorkList","工作","work","笔记","工作列表" : ShowWorkList()	'显示工作记录列表(20140310)
	Case "ShowWork" : ShowWork(Request("FilePath")) '显示工作记录内容(20140310)	
	
	Case "news" : Rw(GetFText("/Templates/MaiSiDe/QQShare.Html")) : Response.End()			'直接跳转到首页，为软件作QQ分享使用 添加于20140411
	Case "AddProduct" : AddProduct()		'VB软件上传产品信息 20140415
	Case "DeleteProduct" : DeleteProduct("Where Id=" & Request("Id"))			'删除VB软件上传的产品ID 20140429
	
	Case "AotuAddKeyWord" : AotuAddKeyWord()	'自动添关键词
	Case "WebNavigationUsrList" : WebNavigationUsrList()		'显示网站导航网站列表
	Case "WebStat" : WebStat(""& adminDir &"Data/Stat/")		'网站统计
	Case "FormSubmit" : FormSubmit()					'表单提交
	Case "RunFile" : RunFile()							'运行指定文件
	Case "showweb" : Response.Redirect("/Inc/Create_Html.Asp?act=Nav&NavDidType=首页&MackHtml=False&Debug=True" & IIF(Request("gl")<>"","&gl=" & Request("gl"),""))	'2014 12 01
	Case "aboutauthor","about","author" : Call AboutAuthor()				'显示作者信息
	case "default" : Call Default() 
    'Case Else Default()			'方便其它页面加载
End Select
'生成HTML程序
Sub MakeWebHtml(NavDid,NavSid,ProDid,ProSid,ProTid, ActionType,ID,ClassName,WD)
	Dim TempCode,TempCreateHtmlPath,I,CreateHtmlYes,NavType,S,C,NewFile
	Dim TempMainContent				'暂存主内容，主要是因为在生成产品新闻列表页时 第一页会显示到最后一页
	Dim CustomTemplateFile			'自定义模板文件路径20150109
	dim content						'内容添加于20151215
	dim isCheckHtmlLabel : isCheckHtmlLabel=true			'是否检测Html标签成功性
	'给全局变量
	PubNavDid = NavDid
	PubNavSid = NavSid
	PubProDid = ProDid
	PubProSid = ProSid
	PubProTid = ProTid
	PubActionType = ActionType				'MainInfo|Article
	PubID = ID
	PubClassName = ClassName
	PubWD = WD
	Call HandleWebConfig()		'处理网站配置
	JsUrl = "/Inc/Create_Html.Asp?act=CreateInfo&NavDid=" & NavDid & "&NavSid=" & NavSid
	JsUrl = JsUrl & "&ProDid=" & ProDid & "&ProSid=" & ProSid & "&ProTid=" & ProTid
	JsUrl = JsUrl & "&ActionType=" & ActionType & "&ID=" & ID & "&ClassName=" & ClassName & "&WD=" & WD
	JsUrl = JsUrl & "&MackHtml=" & Rq("MackHtml") & "&Debug=" & Rq("Debug")
	JsContent = "<script type='text/javascript'>AjaxNoReturn('"& JsUrl &"');</script>" 
	CreateHtmlYes=True		'默认创建HTML为真
	Call GetWebConfig()	'获得网站配置
	'WebFolderName = GetEndUrlHandleName(WebSkins)'获得网站目录文件夹名称 \Templates\WeiZhanLue\  得到WeiZhanLue   不需要这样处理了，因为网站后里增加了这个选项，可自己改
	LocalTestFolder = GetWebFolderName()						'获得本地文件夹名称
	 
	WebActionType = ActionType		'收集网站动作类型
	'获得导航模板        
	If ActionType="Nav" Then
		If NavDid = "" Then NavType = Request("NavType") 
    	TemplatePath = GetNavTemplateText(NavDid, NavSid, NavType)    '获得模板路径
		WebNavDid=NavDid:WebNavSid=NavSid		
	'获得类型模板
	ElseIf ActionType="Class" Then	
    	TemplatePath = GetClassTemplateText(ProDid, ProSid, ProTid)    '获得模板路径
		'Call EErr("CreateHtmlPath",CreateHtmlPath)
	'文章
	ElseIf ActionType = "Article" Then
		Call CreateArticle(ID)  
	'文本
	ElseIf ActionType = "MainInfo" Then
		Call CreateMainInfo(ID) 
	End If
	'默认模板
	If TemplatePath = "" Then 
		TemplatePath = WebTemplate & "Main_Model.Html"
	Else
		TemplatePath = Replace(TemplatePath,"/","\")
		'不是绝对路径则加上默认路径地址
		If Left(TemplatePath,1)<>"\" Then TemplatePath = WebTemplate & TemplatePath
	End If
	'暂时不要 操作于20140724
	'If InStr(TemplatePath,WebTemplate)=0 Then TemplatePath = WebTemplate & TemplatePath		'给完整模板路径	
	If Rq("HttpUrl")<>"" Then	
		Code = GetHttpPage(Rq("HttpUrl"),"")
	
	ElseIf Request("Template")<>"" Then
		'待开发
		Dim StartStr,EndStr
		If Request("Template")="表单验证" Then
			Code = GetFText(WebTemplate & "Index_Model.Html")
			If InStr(Code,"</head>")>0 Then			
				Code = Mid(Code,1,InStr(Code,"</head>")+7)
			End If
			S = "{$FormSubmit did='"& Request("Did") &"' tablename='InputValue' formname='formdialog' classname='formdialog' isverificationcode='true'$}"
			code=code & "<script src=""/Jquery/Function.js""></script>" & vbcrlf
			Code = Code & "<body>" & S
			If Request("formedit")="true" Then
				Code = Code & "<div>复制下面代码放到网页里 就可以显示上面的表单了。</div><div>"& Replace(S,"{","&#x7B;") &"</div>"
			End If
		Else		
			CustomTemplateFile = TrimVbCrlf(HandleFileUrl(Request("Template")))
			If Left(CustomTemplateFile,1)<>"\" Then
				CustomTemplateFile = WebTemplate & CustomTemplateFile
			End If
			'Call Eerr(WebTemplate,CustomTemplateFile)
			'Code = GetFText(WebTemplate & Request("Template"))
			'Call Echo("CustomTemplateFile",CustomTemplateFile) 
			
			If Request("Code")<>"" Then
				Call CreateFile(CustomTemplateFile,Request("Code"))
			End If
			Code = GetFText(CustomTemplateFile)
			'call eerr(CustomTemplateFile,code)
		End If
		
		'Call Eerr(CheckFile(WebTemplate & Request("Template")), WebTemplate & Request("Template")) 
		Tplpath = Replace(Request("Template"),"/","\")
		If InStr(Tplpath,"\")>0 Then Tplpath = Mid(Tplpath,InStrRev(Tplpath,"\")+1)
		If InStr(Tplpath,".")>0 Then Tplpath = Mid(Tplpath, 1, InStrRev(Tplpath,".")-1)
		Tplpath = WebTemplate & "Config\" & Tplpath & ".tpl"
		TplContent = GetFText(TplPath)
		StartStr="#默认网页模板列表# start" : EndStr="#默认网页模板列表# end"
		If InStr(TplContent,StartStr)>0 And InStr(TplContent,EndStr)>0 Then
			TplListContent = StrCut(TplContent,StartStr,EndStr,2)  
			'Call Eerr("TplListContent",TplListContent)
		End If
		'Call Die(Code)
	Else
	    Code = GetFText(TemplatePath)    '读出模板内容
	End If
	Call HandleCodeReplace(Code)				'替换Code里面内容，功能：为于网站皮夫切换20150110
	pubCode = Code							    '记录原始Code内容 20150815
	
	
	
	'不创建本地HTML文件
	If Request("MackHtml") <> "False" Then
	    Call Echo("模板路径TemplatePath", TemplatePath)
	End If	 
	'当前位置
	If TempDid <> "" Then ProDid = TempDid : ProSid = TempSid :  ProTid = TempTid			'暂时留着不知道什么用
    'Code = GetPosition(Code, "{$Position$}")
 	TemplateHTMLCode = Code					'给分页使用
 	 
	If WebSkins = "\Skins\YouKou\" Then
		'============ 悠口餐饮 =================
		'导航 网站顶部导航 	
		'Code = Replace(Code, "{$WebTopNav$}", ObjYK.WebNavigation("Where NavTop=True  Order By Sort Asc"))    
		'导航 网站底部导航
		Code = Replace(Code, "{$WebBottomNav$}", ObjYK.WebNavigation("Where NavButtom=True  Order By Sort Asc"))
		'图片 Banner
		'Code = Replace(Code, "{$Banner$}", JsBanner2011("Banner", 6, 980, 230, True, "Order By Sort"))
		Code = Replace(Code, "{$Banner$}", ObjYK.BannerTab("Banner",6,"","",""))
		
		'热搜关键词
		Code = Replace(Code, "{$热搜关键词$}", HotSearch(""))
		'类别 热搜关键词
		Code = Replace(Code, "{$产品类别展示$}", ReadClassList("产品展示", ProSid, ProTid, "","菜单移动显示"))
		'文本
		Code = Replace(Code, "{$为什么在悠口培训$}", ReadMainInfo("为什么在悠口培训"))
		'文本 滚动学员列表
		
		Code = Replace(Code, "{$最新学员$}", ObjYK.ShowHandleUserList(100))
		Code = Replace(Code, "{$最新学员$}", ReadScrollMemberList(100,""))
		'图片 案例展示
		Code = Replace(Code, "{$热门技术推荐$}", CaseList("产品展示",12," Where Recommend=True And BigFiles<>'' ")) 
		'链接 友情链接
		Code = Replace(Code, "{$友情链接$}", ObjYK.FriendLink()) 
		
		'新闻 新闻列表
		Code = Replace(Code, "{$悠口培训新闻$}", ReadNewList("悠口培训新闻",6, 40 ,"Order By Id Desc")) 
		Code = Replace(Code, "{$餐饮行业新闻$}", ReadNewList("餐饮行业新闻",6, 40 ,"Order By Id Desc")) 
		Code = Replace(Code, "{$餐饮店面经营管理$}", ReadNewList("餐饮店面经营管理",6, 40 ,"Order By Id Desc")) 
		Code = Replace(Code, "{$饮食文化$}", ReadNewList("饮食文化",6, 40 ,"Order By Id Desc")) 
		Code = Replace(Code, "{$地方特色美食$}", ReadNewList("地方特色美食",6, 40 ,"Order By Id Desc")) 
		Code = Replace(Code, "{$厨艺培训交流$}", ReadNewList("厨艺培训交流",6, 40 ,"Order By Id Desc")) 
		'链接地址
		Code = Replace(Code, "{$网站地图$}", GetNavUrl("网站地图","")) 
		Code = Replace(Code, "{$网站地图Target$}", GetNavTarget("网站地图",""))	
		
		Code = Replace(Code, "{$培训场景More$}", GetBigClassUrl("培训场景")) 
		Code = Replace(Code, "{$热门技术推荐More$}", GetBigClassUrl("热门技术推荐")) 
		Code = Replace(Code, "{$悠口培训基地一览More$}", GetBigClassUrl("悠口培训基地一览")) 
		Code = Replace(Code, "{$学员风采More$}", GetBigClassUrl("学员风采")) 
		Code = Replace(Code, "{$诱人美味More$}", GetBigClassUrl("诱人美味"))	
		Code = Replace(Code, "{$悠口培训新闻More$}", GetBigClassUrl("悠口培训新闻")) 
		Code = Replace(Code, "{$餐饮行业新闻More$}", GetBigClassUrl("餐饮行业新闻")) 
		Code = Replace(Code, "{$餐饮店面经营管理More$}", GetBigClassUrl("餐饮店面经营管理")) 
		Code = Replace(Code, "{$饮食文化More$}", GetBigClassUrl("饮食文化")) 
		Code = Replace(Code, "{$地方特色美食More$}", GetBigClassUrl("地方特色美食")) 
		Code = Replace(Code, "{$厨艺培训交流More$}", GetBigClassUrl("厨艺培训交流")) 
		'Main详细页处理
		DefaultWebBanner = "/Skins/YouKou/Index/Banner.jpg" 
		Code = Replace(Code, "{$MainBanner$}",  ObjYK.MainBanner(NavDid,NavSid,ProDid,ProSid,ProTid))
	
		Code = Replace(Code, "{$培训场景$}", ScrollPhoto("培训场景",12,160,130,"培训场景","") ) 
		Code = Replace(Code, "{$培训场景JS$}", JsPhotoScroll("培训场景",740,158) ) 
		
		Code = Replace(Code, "{$诱人美味$}", ScrollPhoto("诱人美味",12,160,130,"诱人美味","") ) 
		Code = Replace(Code, "{$诱人美味JS$}", JsPhotoScroll("诱人美味",960,158) ) 
		
		Code = Replace(Code, "{$学员风采$}", ScrollPhoto("学员风采",12,160,130,"学员风采","") ) 
		Code = Replace(Code, "{$学员风采JS$}", JsPhotoScroll("学员风采",960,158) ) 
		
		Code = Replace(Code, "{$悠口培训基地一览$}", ScrollPhoto("悠口培训基地一览",12,160,130,"悠口培训基地一览","") ) 
		Code = Replace(Code, "{$悠口培训基地一览JS$}", JsPhotoScroll("悠口培训基地一览",960,158) ) 
		

		'QQ样式二
		S = StrCut(GetFText("/Templates/PublicModule/QQOnline2/2.html"),"[开始]","[结束]",2) 
		If S = "" Then Call Eerr("QQ样式二","为空")
		Code = Replace(Code, "<body>", "<body>" & S) 
		
		'右上角翻页关闭代码
		'S = StrCut(GetFText("/Templates/PublicModule/右上角翻页关闭代码/index.html"),"[开始]","[结束]",2) 
		'If S = "" Then Call Eerr("QQ右上角翻页关闭代码","为空")
		'Code = Replace(Code, "<body>", "<body>" & S) 
		
		
		'意见反馈和返回顶部(JS有错误)
		'S = StrCut(GetFText("/Templates/PublicModule/意见反馈和返回顶部/index.html"),"[开始]","[结束]",2) 
		'If S = "" Then Call Eerr("意见反馈和返回顶部","为空")
		'Code = Replace(Code, "<body>", "<body>" & S) 
		
		 
		
		'抽屉式图片展示效果
		'S = StrCut(GetFText("/Templates/PublicModule/抽屉式图片展示效果/index.html"),"[开始]","[结束]",2) 
		'If S = "" Then Call Eerr("抽屉式图片展示效果","为空")
		'Code = Replace(Code, "<!--BannerTest-->", S) 
		
		
		S = StrCut(GetFText("../Templates/PublicModule/右下角浮动层.html"),"[开始]","[结束]",2)
		Code = Replace(Code, "<body>", "<body>" & S)
		Code = Replace(Code, "{$最新动态标题$}", "悠口最新动态") 
		Code = Replace(Code, "{$最新动态内容$}", LiTypeNewList("悠口最新动态",6, 40 ,"")) 
	
 		TemplateHTMLCode = Code					'给分页使用 特殊网站处理需要加这个
	
	
	ElseIf WebSkins = "\Skins\MaiSiDe\" Then
		'============ 麦思德 =================
		'导航 网站顶部导航 	
		Code = Replace(Code, "{$WebTopNav$}", ObjYK.NewWebNavigation(NavDid,"Where NavTop=True  Order By Sort Asc"))    
		'导航 网站底部导航
		Code = Replace(Code, "{$WebBottomNav$}", ObjYK.WebNavigation("Where NavButtom=True  Order By Sort Asc"))
		'图片 Banner
		'Code = Replace(Code, "{$Banner$}", JsBanner2011("Banner", 6, 980, 230, True, "Order By Sort"))
		Code = Replace(Code, "{$Banner$}", ObjYK.BannerTab("Banner",6,"","",""))
		
		'热搜关键词
		Code = Replace(Code, "{$热搜关键词$}", HotSearch(""))
		'类别 热搜关键词
		Code = Replace(Code, "{$产品类别展示$}", ObjMSD.LeftClassList("副导航", ProSid, ProTid, ""))
		'文本
		Code = Replace(Code, "{$为什么在麦思德培训$}", ReadMainInfo("为什么在麦思德培训"))
		'文本 滚动学员列表
		
		Code = Replace(Code, "{$最新学员$}", ObjYK.ShowHandleUserList(100))
		Code = Replace(Code, "{$最新学员$}", ReadScrollMemberList(100,""))
		'图片 案例展示
		Code = Replace(Code, "{$热门技术推荐$}", CaseList("产品展示",12," Where Recommend=True And BigFiles<>'' ")) 
		'链接 友情链接
		Code = Replace(Code, "{$友情链接$}", ObjYK.FriendLink()) 
		
		'新闻 新闻列表
		Code = Replace(Code, "{$麦思德汉堡新闻$}", ReadNewList("麦思德汉堡新闻",6, 40 ," Order By AddDateTime Desc")) 
		Code = Replace(Code, "{$餐饮行业新闻$}", ReadNewList("餐饮行业新闻",6, 40 ," Order By AddDateTime Desc")) 
		Code = Replace(Code, "{$餐饮店面经营管理$}", ReadNewList("餐饮店面经营管理",6, 40 ," Order By AddDateTime Desc")) 
		Code = Replace(Code, "{$饮食文化$}", ReadNewList("饮食文化",6, 40 ," Order By Sort Asc")) 
		Code = Replace(Code, "{$地方特色美食$}", ReadNewList("地方特色美食",6, 40 ," Order By AddDateTime Desc")) 
		Code = Replace(Code, "{$常见问题解答$}", ReadNewList("常见问题解答",6, 40 ," Order By AddDateTime Desc")) 
		'链接地址
		Code = Replace(Code, "{$网站地图$}", GetNavUrl("网站地图","")) 
		Code = Replace(Code, "{$网站地图Target$}", GetNavTarget("网站地图",""))	
		
		Code = Replace(Code, "{$麦思德汉堡新闻链接$}", GetClassUrl_Nav("麦思德汉堡新闻","",""))
		Code = Replace(Code, "{$成功案例链接$}", GetClassUrl_Nav("成功案例","",""))	
		Code = Replace(Code, "{$汉堡加盟菜单副导航链接$}", GetClassUrl_Nav("副导航","汉堡加盟菜单",""))	
 
		Code = Replace(Code, "{$麦思德汉堡新闻More$}", GetBigClassUrl("麦思德汉堡新闻")) 
		Code = Replace(Code, "{$培训场景More$}", GetBigClassUrl("培训场景")) 
		Code = Replace(Code, "{$热门技术推荐More$}", GetBigClassUrl("热门技术推荐")) 
		Code = Replace(Code, "{$麦思德培训基地一览More$}", GetBigClassUrl("麦思德培训基地一览")) 
		Code = Replace(Code, "{$学员风采More$}", GetBigClassUrl("学员风采")) 
		Code = Replace(Code, "{$诱人美味More$}", GetBigClassUrl("诱人美味"))	
		Code = Replace(Code, "{$麦思德培训新闻More$}", GetBigClassUrl("麦思德培训新闻")) 
		Code = Replace(Code, "{$餐饮行业新闻More$}", GetBigClassUrl("餐饮行业新闻")) 
		Code = Replace(Code, "{$餐饮店面经营管理More$}", GetBigClassUrl("餐饮店面经营管理")) 
		Code = Replace(Code, "{$饮食文化More$}", GetBigClassUrl("饮食文化")) 
		Code = Replace(Code, "{$地方特色美食More$}", GetBigClassUrl("地方特色美食")) 
		Code = Replace(Code, "{$常见问题解答More$}", GetBigClassUrl("常见问题解答")) 
		Code = Replace(Code, "{$麦思德培训基地一览More$}", GetBigClassUrl("麦思德培训基地一览")) 
		Code = Replace(Code, "{$成功案例More$}", GetBigClassUrl("成功案例")) 
		'Main详细页处理
		DefaultWebBanner = "/DataDir/WebImg/Banner/麦思德理念.jpg" 
		Code = Replace(Code, "{$MainBanner$}",  ObjYK.MainBanner(NavDid,NavSid,ProDid,ProSid,ProTid))
	
		Code = Replace(Code, "{$培训场景$}", ScrollPhoto("培训场景",12,160,130,"培训场景"," Order By AddDateTime Desc") ) 
		Code = Replace(Code, "{$培训场景JS$}", JsPhotoScroll("培训场景",740,158) ) 
	
		Code = Replace(Code, "{$成功案例$}", ScrollPhoto("成功案例",12,160,130,"成功案例"," Order By AddDateTime Desc") ) 
		Code = Replace(Code, "{$成功案例JS$}", JsPhotoScroll("成功案例",740,158) ) 
		
		Code = Replace(Code, "{$诱人美味$}", ScrollPhoto("诱人美味",12,160,130,"诱人美味"," Order By AddDateTime Desc") ) 
		Code = Replace(Code, "{$诱人美味JS$}", JsPhotoScroll("诱人美味",960,158) ) 
		
		Code = Replace(Code, "{$学员风采$}", ScrollPhoto("学员风采",12,160,130,"学员风采"," Order By AddDateTime Desc") ) 
		Code = Replace(Code, "{$学员风采JS$}", JsPhotoScroll("学员风采",960,158) ) 
		
		Code = Replace(Code, "{$麦思德培训基地一览$}", ScrollPhoto("麦思德培训基地一览",12,160,130,"麦思德培训基地一览"," Order By AddDateTime Desc") ) 
		Code = Replace(Code, "{$麦思德培训基地一览JS$}", JsPhotoScroll("麦思德培训基地一览",960,158) ) 
		
		
		Code = ManyTimesHandleTemplateAction(Code)			'麦思德先处理20150113  在下面这个先处理，要不然会出错
		If 1=1 Then
		S = StrCut(GetFText("../Templates/PublicModule/右下角浮动层.html"),"[开始]","[结束]",2)
		Code = Replace(Code, "<body>", "<body>" & S)
		Code = Replace(Code, "{$最新动态标题$}", "麦思德最新动态") 
		Code = Replace(Code, "{$最新动态内容$}", LiTypeNewList("麦思德最新动态",6, 40 ,"")) 
		
		'QQ样式一
		'S = StrCut(GetFText("../Templates/PublicModule/QQOnline/QQOnLine.html"),"[开始]","[结束]",2) 
		'Code = Replace(Code, "<body>", "<body>" & S) 
		
		'QQ样式二
		'S = StrCut(GetFText("/Templates/PublicModule/QQOnline2/2.html"),"[开始]","[结束]",2) 
		'If S = "" Then Call Eerr("QQ样式二","为空")
		'Code = Replace(Code, "<body>", "<body>" & S) 
		
		
		'旧版视频
		'S = StrCut(GetFText("../Templates/PublicModule/视频播放.html"),"[开始]","[结束]",2) 
		'Code = Replace(Code, "<!--视频-->", S) 
		
		S = StrCut(GetFText("../Templates/PublicModule/餐饮菜单图片.html"),"[开始]","[结束]",2) 
		Code = Replace(Code, "<!--餐饮图片-->", S) 
		
		S = StrCut(GetFText("../Templates/PublicModule/展示Banner.html"),"[开始]","[结束]",2) 
		Code = Replace(Code, "<!--展示Banner-->", S) 
		
		S = StrCut(GetFText("../Templates/PublicModule/展示Banner.html"),"[JS开始]","[JS结束]",2) 
		Code = Replace(Code, "<!--JS展示Banner-->", S) 
		
		
		'S = StrCut(GetFText("../Templates/PublicModule/QQ图片版.html"),"[开始]","[结束]",2)
		'Code = Replace(Code, "<body>", "<body>" & S)
		
		S = StrCut(GetFText("../Templates/PublicModule/麦思德视频.html"),"[开始]","[结束]",2)
		Code = Replace(Code, "<body>", "<body>" & S)
		
		End If  
		
		
		
 		TemplateHTMLCode = Code					'给分页使用 特殊网站处理需要加这个
 	End If 
	
 	'产品新闻分类
	If ActionType="Class" Or InStr("|产品|新闻|下载|","|" & WebNavType & "|")>0  Then 
		MainContent = ProductList(ProDid, ProSid, ProTid,ActionType,"","",1)
		TempCode = Code
		'多次处理模板动作
		Code = ManyTimesHandleTemplateAction(TempCode)					'处理类动作，因为它特殊 因为它下面将创建文件
		'处理第一页
		MainContent = ProductList(ProDid, ProSid, ProTid,ActionType,"","", 1) : TempMainContent = MainContent
		Code = Replace(TemplateHTMLCode, "{$MainContent$}", MainContent)    '主体内容			只要替换MainContent就OK
		
		'类当创建HTML文件名为空则提示，并终止 20140417
		If CreateHtmlPath = "" Then
			Call Eerr("Class","创建类出问题，可能没有这个类")
		End If
		If CheckMakeHtmlFile(CreateHtmlPath) = False Then 
			Call WriteToFile(CreateHtmlPath, ManyTimesHandleTemplateAction(TempCode),1)	'创建产品第一页 处理动作
			call handleMakeHtmlFileEditTime(CreateHtmlPath)		'处理生成html文件修改时间
		end if
		'Call Echo("CreatePage",CreatePage)
		If CreatePage > 1 Then
			For I=2 To CreatePage
				If (Request("Page")=CStr(I) Or Request("Page")="") And InStr(TemplateHTMLCode,"{$MainContent$}")>0 Then
					TempCreateHtmlPath = CreateHtmlPath					' Replace(CreateHtmlPath,".Html", I & ".Html")
					
					MainContent = ProductList(ProDid, ProSid, ProTid,ActionType,"","", I)
					Code = Replace(TemplateHTMLCode, "{$MainContent$}", MainContent)    '主体内容			只要替换MainContent就OK
					Code = ManyTimesHandleTemplateAction(Code)		 '处理动作，用这种
					Call CreateDirFolder(TempCreateHtmlPath)		'创建一组文件夹20141215
					
					NewFile=TempCreateHtmlPath
					NewFile=Mid(TempCreateHtmlPath,1,InStr(TempCreateHtmlPath,".")-1)
					NewFile=NewFile & I & Mid(TempCreateHtmlPath,InStr(TempCreateHtmlPath,"."))
					
					If Request("MackHtml") <> "False" Then
						Call Echo("当前第【"& I &"】页",NewFile)
					End If
					if makeHtmlFileToLCase = true then
						NewFile=lcase(NewFile)		'文件小写
					end if
					
					content=HandleWebCodeReplaceTpl(ManyTimesHandleTemplateAction(Code))
					'add20151215
					if isCheckHtmlLabel=true then
						if checkHtmlFormatting(content)=false then
							call echoRed("检测html标签出错", CreateHtmlPath)
						end if
					end if
				
					
					Call WriteToFile(NewFile, content,1) 	'创建产品其它页面 处理动作
		 			call handleMakeHtmlFileEditTime(NewFile)		'处理生成html文件修改时间
				End If
				DoEvents
			Next 
			If Request("Page")="" Then 
				Code = Replace(TemplateHTMLCode, "{$MainContent$}", TempMainContent)    '主体内容			只要替换MainContent就OK
				Code = ManyTimesHandleTemplateAction(Code)		 '处理动作，用这种
			End If
		End If
		CreateHtmlYes=True 			'为真，因为他要生成首页
		
		'不创建本地HTML文件
		If Request("MackHtml") <> "False" Then
			Call Echo(ProDid & ">>" & ProSid & ">>" & ProTid ,"共生成【"& CreatePage &"】页")
		End If 
	ElseIf ActionType = "Search" Then
		Page = IIF(Request("Page")<>"",Request("Page"),1)
		MainContent =  ProductList(ProDid, ProSid, ProTid,ActionType, ClassName, wd, Page)	
		CreateHtmlYes=False
	End If
	'对远程模板处理
	If Rq("HttpIP")<>"" Then	 
		Code = Replace(Code, GetServicerIP, Rq("HttpIP"))
	End If
	
	'多次处理模板动作
	Code = ManyTimesHandleTemplateAction(Code)

	'不创建本地HTML文件 
	If Request("MackHtml") <> "False" Then
		'判断创建html文件是否为文件夹
		if checkfolder(CreateHtmlPath)=true then
			CreateHtmlYes=false
		end if
		'对生成文件权限处理
		If GetIP="127.0.0.1" Or Session("Administrator") <> "" Or MD5(Request("P"), 1) = "d7213640" Then
			Call Echo("生成Html路径CreateHtmlPath", "<a href='"& HandleHttpUrl(WebDoMain() & CreateHtmlPath) &"'"& ATarget("_blank") &">"& handlePath(CreateHtmlPath) &"</a>") 			
			DoEvents
			Call Echo("CreateHtmlYes",CreateHtmlYes)
			If CheckFile(TemplatePath)=True Then
				'生成Html文件
				
				'Call Eerr("Code2",Code)
				If CreateHtmlYes=True Then
					If CheckMakeHtmlFile(CreateHtmlPath)=True Then 
						Call CreateDirFolder(CreateHtmlPath)		'创建一组文件夹20141215
						if makeHtmlFileToLCase = true then
							CreateHtmlPath=lcase(CreateHtmlPath)		'文件小写
						end if
						content=HandleWebCodeReplaceTpl(ManyTimesHandleTemplateAction(Code))
						'add20151215
						if isCheckHtmlLabel=true then
							if checkHtmlFormatting(content)=false then
								call echoRed("检测html标签出错", CreateHtmlPath)
							end if
						end if
						'call rwend(CreateHtmlPath & vbcrlf & "<hr>" & content)
						Call WriteToFile(CreateHtmlPath, content,1)			'处理动作
						'Call Eerr("CreateHtmlPath",CreateHtmlPath)						
						call handleMakeHtmlFileEditTime(CreateHtmlPath)		'处理生成html文件修改时间
					End If
				End If
			Else 
				Call SetColorEcho("Green","生成失败", "模板文件【"& TemplatePath  &"】不存在")
			End If
		Else
			Call Echo("生成文件失败","未登录")
		End If
		If Request("gotoHtml") = "true" Then 
			call eerr("跳转网址" & HandleHttpUrl(WebDoMain() & CreateHtmlPath), "<script>javascript:location='"& HandleHttpUrl(WebDoMain() & CreateHtmlPath) &"'</script>")
		end if 
	End If 
	
	If Request("DisplayYes")="False"  Then DisplayYes = False
    If DisplayYes = True Then 
		'自动跳转到当前静态页面
		If Rq("ToHtml")="true" Then
			'Response.Redirect(CreateHtmlPath)
			Call Die("<script>location.href='"& CreateHtmlPath &"'</script>")
			Call eerr("CreateHtmlPath",CreateHtmlPath)		
		'给控制面板加Css样式
		ElseIf Rq("gl") = "edit" Then			
			If InStr(Code,"</head>")>0 Then
				Code = Replace(Code,"</head>", "<script src=""/Jquery/Callcontext_menu.js""></script></head>")
			Else
				Code = Code & ControlDialogCss()
			End If
			If InStr(Code,"<body>")>0 Then
				Code = Replace(Code,"<body>", "<body onLoad=""ContextMenu.intializeContextMenu()"">")
			Else
				Code = Code & ControlDialogCss()
			End If
		End If
		Call Rw(HandleWebCodeReplaceTpl(Code))		
	End If
End Sub

'修改生成文件 的最后修改时间  配置后端(20150909)			'修改时间完全成功，不用再修改了(20151218)
Function handleMakeHtmlFileEditTime(CreateHtmlPath)
    If checkfolder("\VB工程\") = False Then Exit Function                           '没有这个文件夹，则退出
    If 1 = 2 Then Exit Function                                                     '开启与关闭
	
    Dim testPath, newTestPath, templateFileEditTime 	
    testPath = WebTemplate & "/AVBConfig/" & HandleFilePathArray(TemplatePath)(3) & ".tpl" 
	
	
	
    templateFileEditTime = GetFileEditDate(testPath) 
    'call echo(testPath,templateFileEditTime)
    newTestPath = GetStrCut(getftext(testPath), "<VB IncludeTemplateFile path='", "'", 2)
    If newTestPath <> "" Then
		'call echo(CreateHtmlPath,newTestPath)
        newTestPath = GetStrCut(getftext("\VB工程\SysInfo.ini"), "DefaultTemplateFolder=", vbCrLf, 2) & newTestPath
		
		dim fileTime				'文件时间
		fileTime = GetFileEditDate(newTestPath)
		
		templateFileEditTime=fileTime
		
		if templateFileEditTime<>"" then			
			if cdate(fileTime)<cdate(templateFileEditTime) then
				templateFileEditTime=fileTime
			end if
		else
			templateFileEditTime=fileTime
		end if
        
    End If
	'call echo(testPath,templateFileEditTime)
    If templateFileEditTime <> "" Then
        Call ModifyFileModifydate(CreateHtmlPath, templateFileEditTime) 
    End If 
End Function


'默认显示导航列表
Sub Default()	
	If GetIP<>"127.0.0.1" And Session("Administrator") = "" Then
		Call Echo("提示","管理员未登录")
		If Request("VBPass") <> "" Then
			Session("Administrator") = "vbLogin"
		End If
		Exit Sub
	End If
	Dim C,Click
    C=C & "<script language=""javascript"">" & vbCrlf
    C=C & "// 确定操作" & vbCrlf
    C=C & "function Confirm(){" & vbCrlf
    C=C & "	if(confirm(""你确定要操作吗？\n操作后将不可恢复""))" & vbCrlf
    C=C & "	return true;" & vbCrlf
    C=C & "	else" & vbCrlf
    C=C & "	return false;" & vbCrlf
    C=C & "}" & vbCrlf
    C=C & "</script>" & vbCrlf
	Call Rw(C)
	Call Rw("<div style=""height:400px;"">")
	Click=" onClick='return Confirm();'"
	C = "全部生成</a> &nbsp; | <a "& AHref("?act=CreateNavList&File=False","","_blank") & Click &">生成全部未生成</a>"
	C = C & "  | <a href=http://www.maiside.net/Inc/Create_Html.Asp target='_blank'>新页打开</a>"
	Call Echo("导航生成","<a "& AHref("?act=ShowNavList","","_blank") &">点击进入</a> &nbsp; | <a "& AHref("?act=CreateNavList","","_blank") & Click & ">" & C)	
	Call Rw("<hr>")
	
	C = "全部生成</a> &nbsp; | <a "& AHref("?act=CreateClassList&File=False","","_blank") & Click &">生成全部未生成</a>"
	Call Echo("类别生成","<a "& AHref("?act=ShowClassList","","_blank") &">点击进入</a> &nbsp; | <a "& AHref("?act=CreateClassList","","_blank") & Click &">" & C)	
	Call Rw("<hr>")
	
	C = "全部生成</a> &nbsp; | <a "& AHref("?act=CreateArticleList&File=False","","_blank") & Click &">生成全部未生成</a>"	
	Call Echo("文章生成","<a "& AHref("?act=ShowArticleList","","_blank") &">点击进入</a> &nbsp; | <a "& AHref("?act=CreateArticleList","","_blank") & Click &">" & C)
	Call Rw("<hr>")
	
	Call Echo("会员生成","<a "& AHref("?act=ShowMemberList","","_blank") &">点击进入</a> &nbsp; | <a "& AHref("?act=CreateMemberList","","_blank") & Click &">全部生成</a>")
	Call Rw("<hr>")
	
	Call Echo("生成RSS/Feed.Xml","<a "& AHref("Rss.Asp","","_blank") & Click &">点击生成</a>")
	Call Rw("<hr>")
	Call Echo("删除文件","<a "& AHref("?act=ShowDelHtmllist","","_blank") &">点击进入</a> &nbsp; | <a "& AHref("?act=ShowDelHtmllist&All=True","","_blank") & Click &">删除全部HTML</a>")
	Call Hr()
	Call Echo("生成网站地图","<a "& AHref("?act=MakeSiteMap","","_blank") &">点击进入</a>")
	Call Rw("</div>")
End Sub
'显示 准备生成导航列表
Sub ShowNavList()
	Dim Url,S
	Call Rw("<title>显示生成导航列表</title>")
    Rs.Open "Select * From [NavBigClass] Order By Sort Asc", Conn, 1, 1 
    While Not Rs.Eof		
		If Rs("OnHtml") = True Then
					
			Url = Rs("FileName")
			If CheckMakeHtmlFile(Url)=True Then
				S =  "<a href=""?act=Nav&NavDid="& Rs("BigClassName") &"&DisplayYes=False"" "& ATarget("_blank") &">点击生成</a>&nbsp;|&nbsp;"
				S=S &  "<a href=""?Debug=1&act=Nav&NavDid="& Rs("BigClassName") &"&MackHtml=False"" "& ATarget("_blank") &">动态预览</a>"
				Call Echo( Rs("BigClassName") & "("& Rs("NavType") &")",S)
			Else
				Call Rw(Rs("BigClassName") & " &nbsp; &nbsp; 为网址("& Url &")，不生成<br>")
			End If
			
		Else
			Call Rw(Rs("BigClassName") & " &nbsp; &nbsp; 生成HTML未开启<br>")
		End If
    	Rsx.Open "Select * From [NavSmallClass] Where BigClassName='"& Rs("BigClassName") &"' Order By Sort Asc", Conn, 1, 1 
		While Not Rsx.Eof
			
			If Rsx("OnHtml") = True Then
			
				Url = Rsx("FileName")
				If CheckMakeHtmlFile(Url)=True Then
					S = "<a href=""?act=Nav&NavDid="& Rsx("BigClassName") &"&NavSid="& Rsx("SmallClassName") &""" "& ATarget("_blank") &">点击生成</a>&nbsp;|&nbsp;"
					S=S &  "<a href=""?Debug=1&act=Nav&NavDid="& Rs("BigClassName") &"&NavSid="& Rsx("SmallClassName") &"&MackHtml=False"" "& ATarget("_blank") &">动态预览</a>"
					Call Echo( " &nbsp; &nbsp;[小类] " & Rsx("SmallClassName"), S)
				Else
					Call Rw(" &nbsp; &nbsp;[小类] " & Rsx("SmallClassName") & "为网址，不生成<br>")
				End If
				
				
			Else
				Call Rw(" &nbsp; &nbsp;[小类] " & Rsx("SmallClassName") & "生成HTML未开启<br>")
			End If
		Rsx.MoveNext : Wend : Rsx.Close 
    Rs.MoveNext : Wend : Rs.Close 
	Call Echo("全部生成", "<a href=""?act=CreateNavList"" "& ATarget("_blank") &">点击生成全部导航</a>")
End Sub
'VB不引用 End

'替换Code里面内容，功能：为于网站皮夫切换20150110
Function HandleCodeReplace(ByVal Code)
	Dim FindTpl,ReplaceTpl,SplFindTpl,SplReplaceTpl,I,ReplaceStr
	
	'Dim WebModuleSkins WebFindTpl WebReplaceTpl
	
	FindTpl = Request("FindTpl")
	ReplaceTpl = Request("ReplaceTpl")
	
	'被替换样式为空则用默认样式
	If FindTpl = "" Then
		FindTpl = "style='312'"
	End If
	'追加后台自选模块模板
	If WebModuleSkins <> "" Then
		ReplaceTpl = ReplaceTpl & "style='"& WebModuleSkins &"'"		
	End If
	If ReplaceTpl = "" Then HandleCodeReplace=Code : Exit Function
	
	'Call Echo("FindTpl",FindTpl)
	'Call Echo("ReplaceTpl",ReplaceTpl)
	SplFindTpl=Split(FindTpl,"|")
	SplReplaceTpl=Split(ReplaceTpl,"|")
	For I=0 To Ubound(SplFindTpl)
		ReplaceStr=""
		If Ubound(SplReplaceTpl)>=I Then
			ReplaceStr=SplReplaceTpl(I)
		End If
		Code=Replace(Code,SplFindTpl(I),ReplaceStr)
	Next
	
	'网站模板替换指定内容
	SplFindTpl=Split(WebFindTpl,"[$Array$]")
	SplReplaceTpl=Split(WebReplaceTpl,"[$Array$]")
	For I=0 To Ubound(SplFindTpl)
		ReplaceStr=""
		If Ubound(SplReplaceTpl)>=I Then
			ReplaceStr=SplReplaceTpl(I)
		End If
		Code=Replace(Code,SplFindTpl(I),ReplaceStr)
	Next
	HandleCodeReplace = Code
	 
End Function 
'处理网页代码内容(批量) 20150115
Function HandleWebCodeReplaceTpl(Content)
	Dim SplFindTpl,SplReplaceTpl,I,ReplaceStr
	'网站模板替换指定内容
	SplFindTpl=Split(WebCodeFindTpl,"[$Array$]")
	SplReplaceTpl=Split(WebCodeReplaceTpl,"[$Array$]")
	For I=0 To Ubound(SplFindTpl)
		ReplaceStr=""
		If Ubound(SplReplaceTpl)>=I Then
			ReplaceStr=SplReplaceTpl(I)
		End If
		Content=Replace(Content,SplFindTpl(I),ReplaceStr)
	Next
	HandleWebCodeReplaceTpl = Content
End Function
'运行指定文件  测试用
Sub RunFile()
	Dim FilePath,Content
	FilePath=Request("File")
	Content=GetFText(FilePath)
	Content = ManyTimesHandleTemplateAction(Content)					'处理类动作，因为它特殊 因为它下面将创建文件
	Rw(Content)
	Response.End()
End Sub

'生成导航列表
Function  CreateNavList()
	Dim RsDid, RsSid,StartTime,nThis,nCount,IsFile,Url
	nCount = 0					'生成文件总数
	StartTime = Now () 
	DisplayYes = False			'回显为假
	Set RsDid = CreateObject("Adodb.RecordSet")
	Set RsSid = CreateObject("Adodb.RecordSet")
	RsDid.Open "Select * From [NavBigClass] Where OnHtml=True Order By Sort", Conn, 1, 1 
	nCount = nCount + RsDid.RecordCount
	While Not RsDid.Eof
		IsFile = False
		If LCase(Request("File"))="false" Then
			Call GetNavTemplateText(RsDid("BigClassName"), "", "")    '获得模板路径
			IsFile = CheckFile(CreateHtmlPath)
		End If
		nThis = nThis + 1
		Call Echo("当前位置",nThis & "/" & nCount)
		
		If IsFile = False Then 
			Url = RsDid("FileName")
			If CheckMakeHtmlFile(Url)=True Then			
				Call Echo("导航大类", RsDid("BigClassName"))
				Call MakeWebHtml(RsDid("BigClassName"),"","","","","Nav","","","")
			Else
				Call Rw("导航大类：" & RsDid("BigClassName") & " &nbsp; &nbsp; 当前为网址，不生成<br>")
			End If
			
			
		Else
			Call Echo("导航大类("& RsDid("BigClassName") &")", "<a href='"& CreateHtmlPath &"'"& ATarget("_blank") &">"& CreateHtmlPath &"</a>" & "已经存在")
		End If
		
		RsSid.Open "Select * From [NavSmallClass] Where BigClassName='"& RsDid("BigClassName") &"' And OnHtml=True Order By Sort Asc", Conn, 1, 1 
		nCount = nCount + RsSid.RecordCount
		While Not RsSid.Eof				
			IsFile = False
			If LCase(Request("File"))="false" Then
				Call GetNavTemplateText(RsSid("BigClassName"), RsSid("SmallClassName"), "")    '获得模板路径
				IsFile = CheckFile(CreateHtmlPath)
			End If
			nThis = nThis + 1
			If IsFile = False Then 
				Url = RsSid("FileName")
				If CheckMakeHtmlFile(Url)=True Then			
					Call Echo(" &nbsp; &nbsp;导航小类", RsSid("SmallClassName"))
					Call MakeWebHtml(RsSid("BigClassName"),RsSid("SmallClassName"),"","","","Nav","","","")
				Else
					Call Rw("导航小类：" & RsSid("SmallClassName") & " &nbsp; &nbsp; 当前为网址，不生成<br>")
				End If
				
			Else
				Call Echo(" &nbsp; &nbsp;导航小类", "<a href='"& CreateHtmlPath &"'"& ATarget("_blank") &">"& CreateHtmlPath &"</a>" & "已经存在")
			End If
		RsSid.MoveNext : Wend : RsSid.Close 
		DoEvents
		Call HR()
	RsDid.MoveNext:Wend:RsDid.Close
	Set RsDid = Nothing
	Set RsSid = Nothing
	Call Echo("操作完成", "导航，共生成文件【"& nCount &"】个，用时" & PrintTimeValue(DateDiff("s", StartTime, Now())) )
End Function
'显示生成类型列表(大小子三类)
Sub ShowClassList()
	Dim Url,S	
	Call Rw("<title>显示生成产品类别列表</title>")
    Rs.Open "Select * From [BigClass] Order By Sort Asc", Conn, 1, 1 
    While Not Rs.Eof
	
		Url = Rs("FileName")
		If CheckMakeHtmlFile(Url)=True Then
			S = "<a href=""?act=CreateClass&ProDid="& Rs("BigClassName") &""" "& ATarget("_blank") &">点击生成</a>&nbsp;|&nbsp;"
			S=S & "<a href=""?Debug=1&act=CreateClass&ProDid="& Rs("BigClassName") &"&MackHtml=False"" "& ATarget("_blank") &">动态预览</a>"
			Call Echo( Rs("BigClassName") & "(大类)", S)
		Else
			Call Rw("类别大类：" & Rs("BigClassName") & " &nbsp; &nbsp; 当前为网址，不生成<br>")
		End If
		
    	Rsx.Open "Select * From [SmallClass] Where BigClassName='"& Rs("BigClassName") &"' Order By Sort Asc", Conn, 1, 1 
		While Not Rsx.Eof
		
			Url = Rsx("FileName")
			If CheckMakeHtmlFile(Url)=True Then
				S = "<a href=""?act=CreateClass&ProDid="& Rsx("BigClassName") &"&ProSid="& Rsx("SmallClassName") &""" "& ATarget("_blank") &">点击生成</a>&nbsp;|&nbsp;"
				S=S & "<a href=""?Debug=1&act=CreateClass&ProDid="& Rs("BigClassName") &"&ProSid="& Rsx("SmallClassName") &"&MackHtml=False"" "& ATarget("_blank") &">动态预览</a>"
				
				Call Echo( " &nbsp; &nbsp;"&  Rsx("SmallClassName") &"[小类] ", S)
			Else
				Call Rw("类别小类：" & Rsx("SmallClassName") & " &nbsp; &nbsp; 当前为网址，不生成<br>")
			End If
			
			
			
			
			Rss.Open "Select * From [ThreeClass] Where BigClassName='"& Rsx("BigClassName") &"' And SmallClassName='"& Rsx("SmallClassName") &"' Order By Sort Asc", Conn, 1, 1 
			While Not Rss.Eof
			
			
			Url = Rss("FileName")
			If CheckMakeHtmlFile(Url)=True Then
				S = "<a href=""?act=CreateClass&ProDid="& Rss("BigClassName") &"&ProSid="& Rss("SmallClassName") &"&ProTid="& Rss("ThreeClassName") &""" "& ATarget("_blank") &">点击生成</a>&nbsp;|&nbsp;"
				
S=S & "<a href=""?Debug=1&act=CreateClass&ProDid="& Rs("BigClassName") &"&ProSid="& Rsx("SmallClassName") &"&ProTid=" & Rss("ThreeClassName") &"&MackHtml=False"" "& ATarget("_blank") &">动态预览</a>"
				
				
				
				Call Echo( " &nbsp; &nbsp; &nbsp; &nbsp;"&  Rss("ThreeClassName") &"[子类] ", S)
			Else
				Call Rw("类别子类：" & Rss("ThreeClassName") & " &nbsp; &nbsp; 当前为网址，不生成<br>")
			End If 
			Rss.MoveNext : Wend : Rss.Close 
		Rsx.MoveNext : Wend : Rsx.Close 
    Rs.MoveNext : Wend : Rs.Close 
	Call Echo("全部生成", "<a href=""?act=CreateClassList"" "& ATarget("_blank") &">点击生成全部信息类别列表</a>")
End Sub
'生成产品类型列表
Function  CreateClassList()
	Dim RsDid, RsSid, RsTid,StartTime,nThis,nCount,AddSql,IsFile,Url
	nCount = 0					'生成文件总数
	nThis = 0
	StartTime = Now () 
	DisplayYes = False			'回显为假
	Set RsDid = CreateObject("Adodb.RecordSet")
	Set RsSid = CreateObject("Adodb.RecordSet")
	Set RsTid = CreateObject("Adodb.RecordSet")
	RsDid.Open "Select * From [BigClass] Where OnHtml=True Order By Sort", Conn, 1, 1 
	nCount = nCount + RsDid.RecordCount
	While Not RsDid.Eof
		IsFile = False
		If LCase(Request("File"))="false" Then
			Call GetClassTemplateText(RsDid("BigClassName"), "", "")    '获得模板路径
			IsFile = CheckFile(CreateHtmlPath) 
		End If	
		nThis = nThis + 1
		Call Echo("当前位置",nThis & "/" & nCount)
		If IsFile = False Then	
		
			Url = RsDid("FileName")
			If CheckMakeHtmlFile(Url)=True Then
				Call Echo("类别大类", RsDid("BigClassName"))
				Call MakeWebHtml("","",RsDid("BigClassName"),"","","Class","","","")
			Else
				Call Rw("类别大类：" & RsDid("BigClassName") & " &nbsp; &nbsp; 当前为网址，不生成<br>")
			End If
		Else
			CreateHtmlPath = "<a href='"& CreateHtmlPath &"' target='_blank'>"& CreateHtmlPath &"</a>"
			Call Echo("类别大类("& RsDid("BigClassName") &")", CreateHtmlPath & "已经存在")
		End If
		Call Hr()
		DoEvents
		
		RsSid.Open "Select * From [SmallClass] Where BigClassName='"& RsDid("BigClassName") &"' And OnHtml=True Order By Sort Asc", Conn, 1, 1 
		nCount = nCount + RsSid.RecordCount
		While Not RsSid.Eof		
			IsFile = False
			If LCase(Request("File"))="false" Then
				Call GetClassTemplateText(RsSid("BigClassName"), RsSid("SmallClassName"), "")    '获得模板路径
				IsFile = CheckFile(CreateHtmlPath) 
			End If		
			nThis = nThis + 1
			Call Echo("当前位置",nThis & "/" & nCount)
			If IsFile = False Then	 
			
			Url = RsSid("FileName")
			If CheckMakeHtmlFile(Url)=True Then
				Call Echo(" &nbsp; &nbsp;类别小类", RsSid("SmallClassName"))
				Call MakeWebHtml("","",RsSid("BigClassName"),RsSid("SmallClassName"),"","Class","","","")
			Else
				Call Rw("类别子类：" & RsSid("SmallClassName") & " &nbsp; &nbsp; 当前为网址，不生成<br>")
			End If 
				
			Else
				CreateHtmlPath = "<a href='"& CreateHtmlPath &"' target='_blank'>"& CreateHtmlPath &"</a>"
				Call Echo(" &nbsp; &nbsp;类别小类("& RsSid("SmallClassName") &")", CreateHtmlPath & "已经存在")
			End If
			Call Hr()
			DoEvents 
			AddSql = " And SmallClassName='"& RsSid("SmallClassName") &"' And OnHtml=True Order By Sort Asc"
			RsTid.Open "Select * From [ThreeClass] Where BigClassName='"& RsDid("BigClassName") &"' " & AddSql, Conn, 1, 1
			nCount = nCount + RsTid.RecordCount
			While Not RsTid.Eof		
				IsFile = False
				If LCase(Request("File"))="false" Then
					Call Echo("WebTitle",WebTitle)
					Call GetClassTemplateText(RsTid("BigClassName"), RsTid("SmallClassName"), RsTid("ThreeClassName"))    '获得模板路径
					IsFile = CheckFile(CreateHtmlPath) 
				End If
				nThis = nThis + 1
				Call Echo("当前位置",nThis & "/" & nCount)
				If IsFile = False Then
				


				Url = RsTid("FileName")
				If CheckMakeHtmlFile(Url)=True Then
					Call Echo(" &nbsp; &nbsp; &nbsp; &nbsp;类别子类", RsTid("SmallClassName"))
					Call MakeWebHtml("","",RsTid("BigClassName"),RsTid("SmallClassName"), RsTid("ThreeClassName"),"Class","","","")
				Else
					Call Rw("类别子类：" & RsTid("ThreeClassName") & " &nbsp; &nbsp; 当前为网址，不生成<br>")
				End If
				
				
				Else
					CreateHtmlPath = "<a href='"& CreateHtmlPath &"' target='_blank'>"& CreateHtmlPath &"</a>"
					Call Echo(" &nbsp; &nbsp; &nbsp; &nbsp;类别子类("& RsTid("ThreeClassName") &")", CreateHtmlPath & "已经存在")
				End If
				Call Hr()
				DoEvents
			RsTid.MoveNext : Wend : RsTid.Close
			DoEvents			
		RsSid.MoveNext : Wend : RsSid.Close 		
	RsDid.MoveNext:Wend:RsDid.Close
	Set RsDid = Nothing
	Set RsSid = Nothing
	Set RsTid = Nothing 
	Call Echo("操作完成", "产品类别，共生成文件【"& nCount &"】个，用时" & PrintTimeValue(DateDiff("s", StartTime, Now())) )
End Function
'显示要生成会员信息列表
Sub ShowMemberList()
	Dim Url
    Rs.Open "Select * From [Member] Order By Sort Asc", Conn, 1, 1 
    While Not Rs.Eof
    	Call Echo( Rs("MyName") & "(会员账号)", "<a href=""?act=CreateMember&ID="& Rs("ID") &""" "& ATarget("_blank") &">点击生成</a>")
    Rs.MoveNext : Wend : Rs.Close 
	Call Echo("全部生成", "<a href=""?act=CreateMemeberList"" "& ATarget("_blank") &">点击生成全部文章信息</a>")
End Sub
Sub ShowDelHtmllist()
	Dim LocalYes,Url,SplStr,Content,FilePath,FolderName
	LocalYes=False
	If LocalYes = False Then
		If GetIP()="127.0.0.1" Then LocalYes=True
	End If
	Url = Replace(Replace(LCase(GetHost()),"http://",""),"/","")
	If LocalYes = False Then
		If Url="127.0.0.1" Or Url = "localhost" Or Left(Url,8)="192.168." Then LocalYes=True
	End If
	If LocalYes=True Then
		Call Echo("提示", "为本地，不可操作")		
		Response.End()
		Exit Sub
	End If
	Content=GetDirHtmlList("/")
	SplStr=Split(Content,vbCrlf)
	For Each FilePath In SplStr
		Call Echo("删除文件成功", FilePath)
		Call DeleteFile(FilePath)
	Next
	Content =GetDirFolderNameList("/")
	SplStr=Split(Content,vbCrlf)
	For Each FolderName In SplStr
		If InStr(LCase("|Action|Admin|AspBox|aspnet_client|BBS|DataDir|Data|Editor|Help|Images|Inc|Jquery|M|Maiside|maisideAPP|Mobile|Module|OtherAsp|QQ|Skins|Templates|UploadFiles|User|WebAdmin|Xinxi|XSS|CGI-BIN|HttpErrors|Web_Soft|"),"|"& LCase(FolderName) &"|")=False Then
			Call Echo("删除文件夹成功", FolderName)
			Call DeleteFolder("/" & FolderName)
		End If
	Next
	Call Echo("删除全部HTML","完成")
End Sub
'显示删除HTML面板 旧版本
Sub ShowDelHtmllist2() 
	Dim RsDid, RsSid,StartTime,nThis,nCount,IsFile,FileName,C,RsTid,AddSql,RsTemp 
	nCount = 0					'生成文件总数
	StartTime = Now () 
	DisplayYes = False			'回显为假
	Set RsDid = CreateObject("Adodb.RecordSet")
	Set RsSid = CreateObject("Adodb.RecordSet")
	Set RsTid = CreateObject("Adodb.RecordSet")
	Set RsTemp = CreateObject("Adodb.RecordSet")
	'删除导航
	RsDid.Open "Select * From [NavBigClass] Where OnHtml=True Order By Sort", Conn, 1, 1
	nCount=nCount+RsDid.RecordCount
	While Not RsDid.Eof
		nThis = nThis + 1 
		FileName = GetNavUrl(RsDid("BigClassName"),"")
		If Right(FileName,1) = "/" Then FileName=FileName & "Index.html"		
		C = IIF(CheckFile(FileName)=True And Request("All") <> "True", " &nbsp; <a href='?act=DelFile&FileName="& FileName &"' target='_blank'>删除文件</a>", "")
		C= "<a href='"& FileName &"' target='_blank'>"& FileName &"</a>" & C
		Call Echo(RsDid("BigClassName"), C )
		If Request("All")= "True" Then DelFile(FileName) : Call Echo("当前位置",nThis & "/" & nCount)
		Call HR()		
		DoEvents
		
		RsSid.Open "Select * From [NavSmallClass] Where BigClassName='"& RsDid("BigClassName") &"' And OnHtml=True Order By Sort Asc", Conn, 1, 1
		nCount=nCount+RsSid.RecordCount
		While Not RsSid.Eof
		nThis = nThis + 1 
			FileName =  GetNavUrl(RsSid("BigClassName"),RsSid("SmallClassName"))
			If Right(FileName,1) = "/" Then FileName=FileName & "Index.html"
			If Request("All")= "True" Then DelFile(FileName) : Call Echo("当前位置",nThis & "/" & nCount) 
			
			C = IIF(CheckFile(FileName)=True And Request("All") <> "True", " &nbsp; <a href='?act=DelFile&FileName="& FileName &"' target='_blank'>删除文件</a>", "")
			Call Echo(RsSid("SmallClassName"),FileName & C)
			Call HR()
			DoEvents
		RsSid.MoveNext : Wend : RsSid.Close 
	RsDid.MoveNext:Wend:RsDid.Close 
	'删除类别
	RsDid.Open "Select * From [BigClass] Where OnHtml=True Order By Sort", Conn, 1, 1 
	nCount=nCount+RsDid.RecordCount
	While Not RsDid.Eof
		nThis = nThis + 1 
		FileName = GetBigClassFolderName(RsDid("BigClassName"))
		If Right(FileName,1) = "/" Then FileName=FileName & "Index.html"			
		C = IIF(CheckFile(FileName)=True And Request("All") <> "True", " &nbsp; <a href='?act=DelFile&FileName="& FileName &"' target='_blank'>删除文件</a>", "")
		C= "<a href='"& FileName &"' target='_blank'>"& FileName &"</a>" & C
		Call Echo(RsDid("BigClassName"), C )
		If Request("All")= "True" Then DelFile(FileName) : Call Echo("当前位置",nThis & "/" & nCount)
		Call HR()
		DoEvents
		
		RsSid.Open "Select * From [SmallClass] Where BigClassName='"& RsDid("BigClassName") &"' And OnHtml=True Order By Sort Asc", Conn, 1, 1 
		nCount=nCount+RsSid.RecordCount
		While Not RsSid.Eof	
		nThis = nThis + 1 
			FileName = GetSmallClassFolderName(RsSid("BigClassName"),RsSid("SmallClassName") )
			If Right(FileName,1) = "/" Then FileName=FileName & "Index.html"			
			C = IIF(CheckFile(FileName)=True And Request("All") <> "True", " &nbsp; <a href='?act=DelFile&FileName="& FileName &"' target='_blank'>删除文件</a>", "")
			C = "<a href='"& FileName &"' target='_blank'>"& FileName &"</a>" & C
			Call Echo(" &nbsp; &nbsp; " & RsSid("SmallClassName"), C )
			If Request("All")= "True" Then DelFile(FileName) : Call Echo("当前位置",nThis & "/" & nCount)
			Call HR()
			DoEvents
					
			AddSql = " And SmallClassName='"& RsSid("SmallClassName") &"' And OnHtml=True Order By Sort Asc"
			RsTid.Open "Select * From [ThreeClass] Where BigClassName='"& RsSid("BigClassName") &"' " & AddSql, Conn, 1, 1
			nCount=nCount+RsTid.RecordCount
			While Not RsTid.Eof	
				nThis = nThis + 1 
				FileName = GetThreeClassFolderName(RsTid("BigClassName"),RsTid("SmallClassName"),RsTid("ThreeClassName") )
				If Right(FileName,1) = "/" Then FileName=FileName & "Index.html"
				C = IIF(CheckFile(FileName)=True And Request("All") <> "True", " &nbsp; <a href='?act=DelFile&FileName="& FileName &"' target='_blank'>删除文件</a>", "")
				C = "<a href='"& FileName &"' target='_blank'>"& FileName &"</a>" & C
				Call Echo(" &nbsp; &nbsp; &nbsp; &nbsp; " & RsTid("ThreeClassName"), C )
				If Request("All")= "True" Then DelFile(FileName) : Call Echo("当前位置",nThis & "/" & nCount)
				Call HR()
				DoEvents
				
			RsTid.MoveNext : Wend : RsTid.Close		
		RsSid.MoveNext : Wend : RsSid.Close 
	RsDid.MoveNext:Wend:RsDid.Close
	
	'删除文章
    RsTemp.Open "Select * From [Product] Where OnHtml=True Order By Sort Asc", Conn, 1, 1 
	nCount=nCount+RsTemp.RecordCount
    While Not RsTemp.Eof
		nThis = nThis + 1
		
		FileName = GetProductUrl("Where Id=" & RsTemp("ID"))
		If Right(FileName,1) = "/" Then FileName=FileName & "Index.html"
		C = IIF(CheckFile(FileName)=True And Request("All") <> "True", " &nbsp; <a href='?act=DelFile&FileName="& FileName &"' target='_blank'>删除文件</a>", "")
		C = "<a href='"& FileName &"' target='_blank'>"& FileName &"</a>" & C
		Call Echo(RsTemp("ID"), C )
		If Request("All")= "True" Then DelFile(FileName) : Call Echo("当前位置",nThis & "/" & nCount)
		Call HR()
		DoEvents
    RsTemp.MoveNext : Wend : RsTemp.Close 
	
	If Request("All")= "True" Then
		Call Echo("操作完成", "删除，共删除文件【"& nCount &"】个，用时" & PrintTimeValue(DateDiff("s", StartTime, Now())) )
	Else
		Call Echo("全部删除","<a "& AHref("?act=ShowDelHtmllist&All=True","","_blank") &">点击删除全部HTML</a>")
	End If
End Sub
'删除指定文件
Function DelFile(FileName)
	Dim C
	C = DeleteFile(FileName)
	Call Echo("删除文件", FileName & " &nbsp;" & C)
End Function
'显示要生成文章信息列表
Sub ShowArticleList()
	Dim Url,OnHtml
    Rs.Open "Select * From [Product] Where OnHtml=True Order By Id Desc", Conn, 1, 1 		'Order By Sort Asc
    While Not Rs.Eof	
		Url = Rs("FileName")
		If CheckMakeHtmlFile(Url)=True Then
			OnHtml = IIF(Rs("OnHtml") = True,"<a href=""?act=CreateArticle&ID="& Rs("ID") &""" "& ATarget("_blank") &">点击生成</a>&nbsp;|&nbsp;","")				
			
			OnHtml = OnHtml & "<a href=""?Debug=1&act=CreateArticle&ID="& Rs("ID") &"&MackHtml=False"" "& ATarget("_blank") &">动态预览</a>"
			
			Url = GetHandleUrl( Rs("FolderName") & Rs("FileName"))
			Url = " &nbsp; <a href='"& Url &"' target='_blank'>"& Url &"</a>," & CheckFile(Url)		
			
			Call Echo( Rs("Title") & "(信息标题)", OnHtml & Url & ",日期" & Format_Time(Rs("AddDateTime"),4))
		Else
			Call Rw("文章信息：" &  Rs("Title") & " &nbsp; &nbsp; 当前为网址，不生成<br>")
		End If		
		
    Rs.MoveNext : Wend : Rs.Close 
	Call Echo("全部生成", "<a href=""?act=CreateArticleList"" "& ATarget("_blank") &">点击生成全部文章信息</a>")
End Sub
'创建一篇文章
Sub CreateArticle(Id)
	Dim FileName,ImgFile,TempImgFile,S
	Dim ConfigContent,ConfigValue,StartStr,EndStr
	Dim OriginalPhotoWidth		'原图片宽
	Dim nDetailPhotoWidth		'详细图片宽
	Dim DetailPhotoWidthStr		'详细图片宽值
	Dim TempStartStr,TempEndStr,ProductModuleTypeList,NewsModuleTypeList,DownloadModuleTypeList,DisplayPhoto,FindS
	Dim Author			'作者
		
	Rs.Open"Select * From [Product] Where Id=" & Id,Conn,1,1
	If Not Rs.Eof Then
		TempDid = Rs("BigClassName")
		TempSid = Rs("SmallClassName")
		TempTid = Rs("ThreeClassName")		
		
		'赋值给全局变量 20141121
		PubProDid=TempDid
		PubProSid=TempSid
		PubProTid=TempTid 
		
		MainTitle = Rs("Title")					'信息标题名称
				
		ImgFile = Trim(IIF(IsNul(Rs("BigFiles")),Rs("SmallFiles"),Rs("BigFiles")))
		
		
		ImgFile = ReplaceGlobleLable(ImgFile) : TempImgFile = ImgFile
		
		Dim TemplateSourceID				'模板资源ID
		TemplateSourceID = GetArticleInfoStyle(TempDid,TempSid,TempTid,Id) 
		S = ReadArticleInfoStyleSource(TemplateSourceID)
				
		If (TempImgFile <> "" And len(TempImgFile)>5)  Then
			If CheckUrl(TempImgFile) = False And checkFile(TempImgFile)=True Then
				nDetailPhotoWidth = GetArticleInfoPhotoWidth(TempDid,TempSid,TempTid,Id)					'设置图片宽
				nDetailPhotoWidth = Cint(nDetailPhotoWidth)			'转成数字类型
'				Call EErr("nDetailPhotoWidth",nDetailPhotoWidth)
				
				'Call Echo(TemplateSourceID,nDetailPhotoWidth)
				
				OriginalPhotoWidth = GetImageWidthValue(TempImgFile)					'原图片宽
				nDetailPhotoWidth = IIF(OriginalPhotoWidth>nDetailPhotoWidth,nDetailPhotoWidth,OriginalPhotoWidth)
				 
				If nDetailPhotoWidth<>"" Then DetailPhotoWidthStr = " width="""& nDetailPhotoWidth &""""
				
			Else
				nDetailPhotoWidth = ""				'对图片进行判断 700判断吧
				DetailPhotoWidthStr = ""
			End If
		End If
		
		Author = Rs("Author")
		If Author="" Then Author="本站"
		
 
		
		
		S = Replace(S,"[$updownpage$]",UpPage(Rs(0)) & DownPage(Rs(0))) 			'替换上一页下一页	
		'S = Replace(S,"[$adddatetime-1$]", Format_Time(Rs("AddDateTime"),1))		
		S = Replace(S,"[$detilphotowidth$]",nDetailPhotoWidth)						'详细图片宽
		S = Replace(S,"[$detilphotowidthstr$]",DetailPhotoWidthStr)					'详细图片宽内容
		
		S = ReplaceValueParam(S,"did",TempDid)
		S = ReplaceValueParam(S,"id",Rs("Id"))			
		S = ReplaceValueParam(S,"title",Rs("title"))
		S = ReplaceValueParam(S,"imgurl",ImgFile)
		S = ReplaceValueParam(S,"price",Rs("Price"))
		S = ReplaceValueParam(S,"description",Rs("WebDescription"))									
		S = ReplaceValueParam(S,"adddatetime",Rs("AddDateTime"))
		S = ReplaceValueParam(S,"author",Author) 
		S = ReplaceValueParam(S,"content",Rs("Content"))
		MainContent = S
		'Call Echo(DetailPhotoWidthStr,nDetailPhotoWidth)
		'Die MainContent
		If nDetailPhotoWidth = "" Then
			StartStr = "<!--#photo start#-->" : EndStr = "<!--#photo end#-->"
			'查找
			If InStr(MainContent,StartStr)>0 And InStr(MainContent,EndStr)>0 Then
				S = StrCut(MainContent,StartStr,EndStr,2)
				'Call Eerr("S",S)
			End If
		End If
		
		
		'添加在线修改
		MainContent = DisplayOnlineEditDialog(""& adminDir &"Product.Asp?act=ShowEditProduct&Id=" & Rs("Id") & "&n=" & GetRnd(11), MainContent)
		
		
		
		ArticleDid=Rs("BigClassName")				'大类名称
		ArticleSid=Rs("SmallClassName")				'小类名称
		ArticleTid=Rs("ThreeClassName")				'子类名称
		ArticleTitle=Rs("Title")					'标题
		ArticleBigFiles=Rs("BigFiles")				'大类图片
		ArticleSmallFiles=Rs("SmallFiles")			'小类图片
		
		ArticleID=Rs("Id")							'文章ID
		ArticleDateTime=Rs("AddDateTime")			'添加时间
		ArticleContent=Rs("Content")				'文章内容
		ArticleAuthor = Rs("Author")				'作者
		 
		'自定义模板
		If Trim(Rs("TemplatePath"))<>"" Then
			TemplatePath = WebTemplate & Rs("TemplatePath") 
		Else
			TemplatePath = WebTemplate & "Main_Model.Html"
		End If
		FileName = Rs("FileName")
		If CheckMakeHtmlFile(FileName)=True Then
			FileName = GetHandleFileName(FileName)
			FileName = DetailPageAddHtmlPrefix(Rs("FolderName")) & "/" & FileName
		End If
		CreateHTMLPath = FileName
		
		'显示这篇文章的网站标题关键词描述
		WebTitle = IIf(Rs("Title")<>"",Rs("Title"),WebTitle)
		WebTitle = IIf(Rs("WebTitle")<>"",Rs("WebTitle"),WebTitle)
		WebKeywords = IIf(Rs("WebKeywords")<>"",Rs("WebKeywords"),WebKeywords)
		WebDescription = IIf(Rs("WebDescription")<>"",Rs("WebDescription"),WebDescription) 
	End If:Rs.Close 
End Sub
'创建Main一页面
Sub CreateMainInfo(Id)
	Dim FileName,ImgPath 
	Rs.Open"Select * From [MainInfo] Where Id=" & Id,Conn,1,1
	If Not Rs.Eof Then 
		'这一种废弃掉 20141230
		'MainContent = "<div class='ArticleTitle'>"& Rs("ShowTitle") &"</div>" 
		'MainContent = MainContent & Rs("Content")
		'直接显示内容就好了   20141230
		MainContent = Rs("Content") 
		MainContent = DisplayOnlineEditDialog(""& adminDir &"MainInfo.Asp?act=ShowEdit&Id=" & Rs("Id") & "&n=" & GetRnd(11), MainContent)			'任意修改20150202
		MainTitle = Rs("ShowTitle")
		
		'自定义模板
		If Trim(Rs("TemplatePath"))<>"" Then
			TemplatePath = WebTemplate & Rs("TemplatePath") 
		Else
			TemplatePath = WebTemplate & "Main_Model.Html"
		End If
		
		FileName = Trim(Rs("FileName"))
		'If FileName = "" Then FileName = "MainInfo" & Rs("Id")			'文件名称为空则使用默认
		FileName = GetHandleFileName(FileName)
		'创建一组文件夹(导航文件夹路径)
		Call CreateDirFolder("/Html")
		CreateHTMLPath = DetailPageAddHtmlPrefix(Rs("FolderName")) & "/" & FileName
		CreateHTMLPath = GetHandleUrl(CreateHTMLPath)	'处理URL 
		
		'显示这篇文章的网站标题关键词描述
		WebTitle = IIf(Rs("Title")<>"",Rs("Title"),WebTitle)
		WebTitle = IIf(Rs("WebTitle")<>"",Rs("WebTitle"),WebTitle)
		WebKeywords = IIf(Rs("WebKeywords")<>"",Rs("WebKeywords"),WebKeywords)
		WebDescription = IIf(Rs("WebDescription")<>"",Rs("WebDescription"),WebDescription) 
	End If:Rs.Close 
End Sub
'生成全部文章信息
Sub CreateArticleList()
	Dim RsPro,StartTime,nCount,IsFile,Url
	nCount = 0					'生成文件总数
	StartTime = Now () 
	DisplayYes = False			'回显为假
	Set RsPro=CreateObject("Adodb.RecordSet")
    RsPro.Open "Select Top 2999 * From [Product] Where ONHtml=True", Conn, 1, 1 
    While Not RsPro.Eof
		IsFile = False
		call echo("文章ID", RsPro("ID"))
		If LCase(Request("File"))="false" Then
			Call CreateArticle(RsPro("ID")) 
			IsFile = CheckFile(CreateHtmlPath)
		End If
		nCount = nCount + 1
		Doevents
		If IsFile = False Then
			Call Echo("当前位置", nCount & "/" & RsPro.RecordCount)
			Doevents
			
			Url = RsPro("FileName")
			If CheckMakeHtmlFile(Url)=True Then
				Call MakeWebHtml("","","","","","Article", RsPro("Id"),"","")            
			Else
				Call Rw("文章信息：" & RsPro("Title") & " &nbsp; &nbsp; 当前为网址，不生成<br>")
			End If 
		Else
			CreateHtmlPath = "<a href='"& CreateHtmlPath &"' target='_blank'>"& CreateHtmlPath &"</a>"
			Call Echo("文章ID("& RsPro("ID") &")", CreateHtmlPath & "已经存在，当前位置("& nCount & "/" & RsPro.RecordCount &")")
		End If
		Call Hr()
		DoEvents
    RsPro.MoveNext : Wend : RsPro.Close 
	Set RsPro=Nothing
	Call Echo("操作完成", "文章，共生成文件【"& nCount &"】个，用时" & PrintTimeValue(DateDiff("s", StartTime, Now())) )
End Sub 
'获得网站配置信息
Sub GetWebConfig()
	Dim Url, Id
	'打开过就用缓冲内容
	If DefaultWebTitle <> "" Then		
		WebTitle = DefaultWebTitle
		WebKeywords = DefaultWebKeywords
		WebDescription = DefaultWebDescription
		Exit Sub
	End If
    '获得网站配置
    Rs.Open "Select * From [WebSite]", Conn, 1, 1 
    If Not Rs.Eof Then
        WebTemplate = HandleHttpUrl(Rs("WebTemplate")) 
        WebSkins = HandleHttpUrl(Rs("WebSkins")) 		
        WebImages = HandleHttpUrl(Rs("WebImages")) 		
        WebCss = HandleHttpUrl(Rs("WebCss")) 		
        WebJs = HandleHttpUrl(Rs("WebJs")) 
		
		WebFolderName = PHPTrim(Rs("WebFolderName"))			'网站文件夹名称20150721
		
		'是否格式化生成后的HTML
		isHtmlFormatting = LCase(Rs("isHtmlFormatting"))
		isHtmlFormatting=IIF(isHtmlFormatting="0" Or isHtmlFormatting="false",false,true)
		'是否网页闭合标签
		isWebLabelClose = LCase(Rs("isWebLabelClose"))
		isWebLabelClose=IIF(isWebLabelClose="0" Or isWebLabelClose="false",false,true) 
		
		WebUrl = Rs("WebSiteUrl")
		WebTitle = Rs("WebTitle") 
        WebKeywords = Rs("WebKeywords") 
        WebDescription = Rs("WebDescription") 
		WebBottom = Rs("WebSiteBottom")
		
		WebModuleSkins = Rs("ModuleSkins")				'网站模块皮肤20150115
		WebFindTpl = Rs("FindTpl")				        '网站模板被替换内容列表20150115
		WebReplaceTpl = Rs("ReplaceTpl")		        '网站模板替换内容列表20150115
		
		
		WebCodeFindTpl = Rs("WebCodeFindTpl")		        '网站网页替换内容列表20150115
		WebCodeReplaceTpl = Rs("WebCodeReplaceTpl")		        '网站网页替换内容列表20150115]
		'内容为Null处理 要不然在处理替换时出错 20150119
		If IsNul(WebCodeFindTpl) Then WebCodeFindTpl=""
		If IsNul(WebCodeReplaceTpl) Then WebCodeReplaceTpl=""
		
		
		Id=Rs("Id")
		Url = ""& adminDir &"WebSiteConfig.Asp?n=" & GetRnd(11)	
		WebBottom = DisplayOnlineEditDialog(Url, WebBottom)
		
		
		
    End If : Rs.Close 
End Sub
'==================================
'		代码核心区域
'==================================

'显示导航列表 【NavBigClass NavSmallClass】
Function Navigation(AddSql)
    Dim I,S, C, Url, FileName,ClassStr,BigFolder,Target,Did,ShowDid,Sid,ShowSid
    Rs.Open "Select * From [NavBigClass]" & AddSql, Conn, 1, 1 
	For I = 1 To Rs.RecordCount
        FileName = Rs("FileName") 
        If FileName <> "" Then FileName = FileName & ".html" 
		BigFolder = "/" & Rs("FolderName") & "/"
        Url = BigFolder & FileName 
        Url = Replace(Replace(Url, "//", "/"), "//", "/")
		Did = Rs("BigClassName")
		ShowDid = FontColorFontB(Did,Rs("FontB"),Rs("FontColor")) 
		Target = Rs("Target")
		ClassStr = "" 
      	If I = 1 Then ClassStr = " class='CurrentLi'" 
		S = "<li"& ClassStr &"><a "& AHref(Url,"",Target) &"{下拉菜单}>" & ShowDid & "</a></li> "
		If Rs("NavType") = "文本" Then
			Rss.Open "Select * From [NavSmallClass] Where BigClassName='"& Rs("BigClassName") &"'", Conn, 1, 1
			If Not Rs.Eof Then
				S = S & "<div id='m"& I &"' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'>"
				While Not Rss.Eof
					FileName = Rss("FileName") 
					If FileName <> "" Then FileName = FileName & ".html" 
					Url = BigFolder & FileName
					Url = BigFolder & "/" & Rss("FolderName") & "/" & FileName 
					Url = Replace(Replace(Url, "//", "/"), "//", "/")
					Sid = Rss("SmallClassName")
					ShowSid = FontColorFontB(Sid,Rss("FontB"),Rss("FontColor")) 
					Target = Rss("Target")
					S = S & "<a "& AHref(Url,"",Target) &">" & ShowSid & "</a>"
				Rss.MoveNext:Wend:Rss.Close
				S = S & "</div>"
				S = Replace(S,"{下拉菜单}", " onmouseover=mopen('m"& I &"') onmouseout='mclosetime()' ")
			End If
		ElseIf Rs("NavType") = "产品" Or Rs("NavType") = "新闻" Then
			Rss.Open "Select * From [SmallClass] Where BigClassName='"& Rs("BigClassName") &"'", Conn, 1, 1
			If Not Rs.Eof Then
				S = S & "<div id='m"& I &"' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'>"
				While Not Rss.Eof
					FileName = Rss("FileName") 
					If FileName <> "" Then FileName = FileName & ".html" 
					Url = BigFolder & FileName
					Url = BigFolder & "/" & Rss("FolderName") & "/" & FileName 
					Url = Replace(Replace(Url, "//", "/"), "//", "/")
					Sid = Rss("SmallClassName")
					ShowSid = FontColorFontB(Sid,Rss("FontB"),Rss("FontColor")) 
					Target = Rss("Target")
					S = S & "<a "& AHref(Url,"",Target) &">" & ShowSid & "</a>"
				Rss.MoveNext:Wend:Rss.Close
				S = S & "</div>"
				S = Replace(S,"{下拉菜单}", " onmouseover=mopen('m"& I &"') onmouseout='mclosetime()' ")
			End If
		 
		End If
		S = Replace(S,"{下拉菜单}","")		'去掉下拉菜单定义
        C = C & S 
    Rs.MoveNext : Next : Rs.Close 
    Navigation = C 
End Function
'网站底部导航
Function BottomNavigation(AddSql)
    Dim I,S, C, Url, FileName,ClassStr,BigFolder,Did,ShowDid,Target
    Rs.Open "Select * From [NavBigClass]" & AddSql, Conn, 1, 1 
	For I = 1 To Rs.RecordCount
        FileName = Rs("FileName") 
        If FileName <> "" Then FileName = FileName & ".html" 
		BigFolder = "/" & Rs("FolderName") & "/"
        Url = BigFolder & FileName 
        Url = Replace(Replace(Url, "//", "/"), "//", "/")
		Did = Rs("BigClassName")
		ShowDid = FontColorFontB(Did,Rs("FontB"),Rs("FontColor")) 
		Target = Rs("Target") 
		C=C & "<a "& AHref(Url,"",Target) &"{下拉菜单}>" & ShowDid & "</a>"
		If I <> Rs.RecordCount Then C = C & " &nbsp;| &nbsp; "
		
	Rs.MoveNext : Next : Rs.Close 
	BottomNavigation = C
End Function
 
'文本【MainInfo】
Function GetMainInfo(Code, Did, Label) 
	Dim Url,C
	Rs.Open"Select * From [MainInfo] Where Title='"& Did &"' And Recommend=True Order By Sort,ID",Conn,1,1
	If Not Rs.Eof Then  
		Code = Replace(Code, Replace(Label,"$}","Title$}"), Rs("Title"))
		Code = Replace(Code, Replace(Label,"$}","ShowTitle$}"), Rs("ShowTitle"))
		Code = Replace(Code, Replace(Label,"$}","Content$}"), Rs("Content"))		  
	End If : Rs.Close  
	GetMainInfo = Code
End Function
'Logo【Product】
Function Logo(Did)
	Dim Url,C
	Rs.Open"Select * From [Product] Where BigClassName='"& Did &"' And Recommend=True Order By Sort,ID",Conn,1,1
	If Not Rs.Eof Then
		Url = GetHandleUrl( Rs("FileName") )
		C = "<a href='"& Url &"' target='_blank'><img border=0 src='"& Rs("BigFiles") &"' alt='"& Rs("Title") &"'></a>"
	End If : Rs.Close 
	Logo = C
End Function
'Banner【Product】
Function Banner(Did)
	Dim Url,C
	Rs.Open"Select * From [Product] Where BigClassName='"& Did &"' And Recommend=True Order By Sort",Conn,1,1
	While Not Rs.Eof
		Url = GetHandleUrl( Rs("FileName") )
		C=C & "<li><a href='"& Url &"' target='_blank'><img border=0 src='"& Rs("BigFiles") &"' alt='"& Rs("Title") &"'></a></li>"	
	Rs.MoveNext : Wend : Rs.Close 
	Banner = C
End Function
'新闻列表【Product】
Function NewList(Code,Did,TopNumb,TitleLabel, MoreLabel, ContentLabel)
	Dim Url,C,Sql,ShowTitle
	Sql = "Select Top "& TopNumb &" * From [Product] Where BigClassName='"& Did &"' And Recommend=True Order By Sort" 
	Rs.Open Sql,Conn,1,1 
	While Not Rs.Eof
		Url = GetHandleUrl( Rs("FileName") )
		ShowTitle = CutStr(Rs("Title"),28,"... ")
		ShowTitle = FontColorFontB(ShowTitle,Rs("FontB"),Rs("FontColor"))		
		C=C & "<li><a "& AHref(Url,Rs("Title"),Rs("Target")) &">"& ShowTitle &"</a><span class=""Time"">" & Format_Time(Rs("UpDateTime"), 14) & "</span></li>"				 
	Rs.MoveNext : Wend : Rs.Close	
	Code = Replace(Code, TitleLabel, Did)
	If MoreLabel <> "" Then Code = Replace(Code, MoreLabel, "")
	If ContentLabel <> "" Then Code = Replace(Code, ContentLabel, C  )
	NewList = Code
End Function
'首页案例展示【Product】
Function HomeCaseList(Code,Did,TopNumb,CaseLabel, ModulePath)
	Dim Url,C,Sql,ModuleContent,I
	Dim InfoBigFiles,InfoTitle,InfoFileName,InfoContent,ShowTitle
	 
	ModuleContent = GetFText(ModulePath)	'模块路径
	Sql = "Select * From [Product] Where BigClassName='"& Did &"' And Recommend=True Order By Sort" 
	Rs.Open Sql,Conn,1,1 
	For I=1 To TopNumb
		If Rs.Eof Then Exit For
		'获得第一条产品类信息
		If InfoTitle = "" Then
			InfoTitle = CutStr(Rs("Title"),28,"... ")	 
			InfoBigFiles = Rs("BigFiles")
			InfoFileName = GetHandleUrl(Rs("FileName"))
		End If
		Url = GetHandleUrl( Rs("FileName") )
		ShowTitle = CutStr(Rs("Title"),28,"... ")
		ShowTitle = FontColorFontB(ShowTitle,Rs("FontB"),Rs("FontColor")) 	'处理文本颜色与加粗	
			
		C=C & "<li><a "&  AHref(Url,Rs("Title"),Rs("Target")) &">"& ShowTitle &"</a></li>"
		'<span class=""Time"">" & Format_Time(Rs("UpDateTime"), 14) & "</span>   显示时间的
	Rs.MoveNext : Next : Rs.Close 
	 
	ModuleContent = Replace(ModuleContent,"{$Case_ProductListInfoTitle$}", InfoTitle)
	ModuleContent = Replace(ModuleContent,"{$Case_ProductListInfoBigFiles$}", InfoBigFiles)
	ModuleContent = Replace(ModuleContent,"{$Case_ProductListInfoFileName$}", InfoFileName)
	
	ModuleContent = Replace(ModuleContent,"{$Case_ProductListTitle$}", Did)
	ModuleContent = Replace(ModuleContent,"{$Case_ProductListFolderName$}", GetBigClassFolderName(Did)) 
	ModuleContent = Replace(ModuleContent,"{$Case_ProductListContent$}", C)
	  
	Code = Replace(Code, CaseLabel,  ModuleContent) 
	HomeCaseList = Code 
End Function

'显示一张图片【Product】
Function ShowImg(Code, Did, Label)
	Dim FileName,Title,Target,BigFiles
	Rs.Open"Select * From [Product] Where BigClassName='"& Did &"' And Recommend=True Order By Sort",Conn,1,1
	If Not Rs.Eof Then
		FileName = GetHandleUrl( Rs("FileName") )
		Title = Rs("Title")
		Target = ATarget(Rs("Target")) 
		BigFiles = Rs("BigFiles")
		Code = Replace(Code, Replace(Label,"$}","FileName$}"), FileName)
		Code = Replace(Code, Replace(Label,"$}","Title$}"), Title)
		Code = Replace(Code, Replace(Label,"$}","Target$}"), Target)
		Code = Replace(Code, Replace(Label,"$}","BigFiles$}"), BigFiles)
	End If : Rs.Close 
	ShowImg = Code
End Function

'首页新闻列表【Product】
Function HomeNewsList(ByVal Did, TopNumb)
    Dim SplStr, I, J, S, C,ModNumb, FileName,Target,Url,Title,ShowTitle,ShowSid
    C=C & "	<ul class=""NewsListBox"">" & vbCrlf
	Rs.Open"Select * From [SmallClass] Where BigClassName='"& Did &"' Order By Sort",Conn,1,1
	For I=1 To Rs.RecordCount
		ModNumb = (I-1) Mod 3
		'Call Echo("",Rs("FolderName"))
		FileName = GetHandleUrl(GetSmallClassFolderName(Did,Rs("SmallClassName")) )
		Target = ATarget(Rs("Target"))
		Sid = Rs("SmallClassName")
		ShowSid = FontColorFontB(Sid,Rs("FontB"),Rs("FontColor")) 	'处理文本颜色与加粗	
    C=C & "    	<li class=""News-Box"& ModNumb &""">" & vbCrlf
    C=C & "        " & vbCrlf
    C=C & "            <div class=""tab"">" & vbCrlf
    C=C & "                <ul class=""tab-hd""><li><A "& AHref(FileName,"",Target) &">"& ShowSid &"</A></li>  </ul>" & vbCrlf
    C=C & "                <ul class=""tab-bd"">" & vbCrlf
    C=C & "                    <li>" & vbCrlf
    C=C & "                        <div class='DivList'>" & vbCrlf
		Rsx.Open"Select * From [Product] Where BigClassName='"& Did &"' And SmallClassName='"& Rs("SmallClassName") &"' And Recommend=True Order By Sort",Conn,1,1
		For J=1 To TopNumb
		If Rsx.Eof Then Exit For
			Url =  GetHandleUrl( Rsx("FileName") )
			Target = ATarget(Rsx("Target"))
			Title = Rsx("Title")
			ShowTitle = CutStr(Rsx("Title"),42,"... ")
			ShowTitle = FontColorFontB(ShowTitle,Rsx("FontB"),Rsx("FontColor")) 	'处理文本颜色与加粗	
		
    		C=C & "                            <div class='DivLi'>· <a "& AHref(Url,Rsx("Title"),Rsx("Target")) &">"& ShowTitle &"</a></div>" & vbCrlf
		Rsx.MoveNext:Next:Rsx.Close
	
    C=C & "                            <div class='clear'></div>" & vbCrlf
    C=C & "                        </div>                " & vbCrlf
    C=C & "                    </li>" & vbCrlf
    C=C & "                </ul>" & vbCrlf
    C=C & "            </div>" & vbCrlf
    C=C & "        " & vbCrlf
    C=C & "        </li>" & vbCrlf
	If I Mod 3=0 Then C=C & "</ul><div class=""clear8""></div><ul class=""NewsListBox"">"
	Rs.MoveNext:Next:Rs.Close
    C=C & "    </ul>" & vbCrlf
    C=C & "	<div class=""clear8""></div>" & vbCrlf
	
    HomeNewsList=C
End Function  
'首页案例展示
Function HomeScrollCaseList(Code,Label, Did,TopNumb)
	Dim Url,Title,Target,BigFiles,I,C,ShowTitle

	Rsx.Open"Select * From [Product] Where BigClassName='"& Did &"' And Recommend=True Order By Sort",Conn,1,1
	For I=1 To TopNumb
	If Rsx.Eof Then Exit For
		Url =  GetHandleUrl( Rsx("FileName") )
		Target = ATarget(Rsx("Target"))
		Title = Rsx("Title")
			ShowTitle = CutStr(Rsx("Title"),32,"... ")
			ShowTitle = FontColorFontB(ShowTitle,Rsx("FontB"),Rsx("FontColor")) 	'处理文本颜色与加粗	
			
		BigFiles = Rsx("BigFiles")
		C=C & "<div class='box'><a class='imgBorder' href='"& Url &"' "& Target &" title='"& Title &"'><IMG src='"& BigFiles &"' alt='"& Title &"'></A> <p><A href='"& Url &"' "& Target & " title='"& Title &"'>"& ShowTitle &"</a> </p></div>" & vbCrlf
	Rsx.MoveNext:Next:Rsx.Close 
	Code = Replace(Code, Replace(Label,"$}","FileName$}"), Did)
	Code = Replace(Code, Replace(Label,"$}","FolderName$}"), GetBigClassFolderName(Did))
	Code = Replace(Code, Replace(Label,"$}","Content$}"), C)
	HomeScrollCaseList = Code
End Function

 
'产品类型列表展示【SmallClass】
Function ProductClassList(Code,Did,Label)
    Dim SplStr, I, S, C,Sid,ShowSid,BigFileName,Target,Url,Title,ShowTid
    C=C & "    <ul id='suckertree1'>" & vbCrlf
	Rs.Open"Select * From [SmallClass] Where BigClassName='"& Did &"' Order By Sort",Conn,1,1
	While Not Rs.Eof
		Sid = Rs("SmallClassName") 
		ShowSid = FontColorFontB(Sid,Rs("FontB"),Rs("FontColor")) 	'处理文本颜色与加粗	
		BigFileName = GetSmallClassFileName( Rs("BigClassName"), Rs("SmallClassName") )
		Target = ATarget(Rs("Target"))
		C=C & "        <li><A "& AHref(BigFileName,"",Target) &">"& ShowSid &"</span></A>" & vbCrlf
		Rsx.Open"Select * From [ThreeClass] Where BigClassName='"& Did &"' And SmallClassName='"& Rs("SmallClassName") &"' Order By Sort,ID Asc",Conn,1,1
			If Not Rsx.Eof Then
				C=C & "            <ul>" & vbCrlf
				For I=1 To Rsx.RecordCount
					Url =  GetThreeClassFileName(Rsx("BigClassName"),Rsx("SmallClassName"),Rsx("ThreeClassName"))
					Target = ATarget(Rsx("Target"))
					Tid = Rsx("ThreeClassName") 
					ShowTid = FontColorFontB(Tid,Rsx("FontB"),Rsx("FontColor")) 	'处理文本颜色与加粗	
					C=C & "                <li><a "& AHref(Url,"",Target) &">"& ShowTid &"</a></li>" & vbCrlf
				Rsx.MoveNext:Next
				C=C & "            </ul>        " & vbCrlf
			End If:Rsx.Close
		C=C & "        </li>" & vbCrlf
	Rs.MoveNext:Wend:Rs.Close
    C=C & "    </ul>" & vbCrlf
	 
		Code = Replace(Code, Replace(Label,"$}","Content$}"), C)
		Code = Replace(Code, Replace(Label,"$}","BigFileName$}"),GetBigClassFileName(Did))

    ProductClassList=Code
End Function 

'Main内页左边导航 产品与导航小类【Class NavClass】
Function MainClassList(Code, Label, NavDid, NavSid, ProDid, ProSid, ProTid)
	Dim BigFolder,SmallFolder,Url,I,S,C,Sel,Did,Sid,Tid,Target,ClassTitle,ClassUrl
	C = ""
	If NavDid <> "" Then
		ClassTitle = NavDid
		'导航大类
		TempRs.Open"Select * From [NavBigClass] Where BigClassName='"& NavDid &"' And NavType<>'首页'",Conn,1,1
		If Not TempRs.Eof Then
			ClassUrl = "/" & TempRs("FolderName") & "/" & TempRs("FileName")
			ClassUrl = GetHandleUrl(ClassUrl)	'处理URL 
			BigFolder = "/" & TempRs("FolderName") & "/" 
		End If:TempRs.Close 
		'导航小类 
		TempRs.Open"Select * From [NavSmallClass] Where BigClassName='"& NavDid &"' Order By Sort",Conn,1,1
		While Not TempRs.Eof
			Sid = TempRs("SmallClassName")
			Url = "/" & BigFolder & TempRs("FolderName") & "/" & TempRs("FileName")
			Url = GetHandleUrl(Url)	'处理URL
			Target = TempRs("Target")
			Sel = IIF(Sid = NavSid,StyleClass("This"),"") 
			C = C & "<li><a "& AHref(Url,"",Target) & Sel & ">"& Sid &"</a></li>"  
		TempRs.MoveNext:Wend:TempRs.Close
	End If 
	If ProDid<>"" Or C = "" Then
		'导航为空 
		If C = "" And ProDid = "" Then
			ProDid = NavDid
			ProSid = NavSid
		End If 
		'类别大类
		If ProDid <> "" Then
			ClassTitle = ProDid 
			TempRs.Open"Select * From [BigClass] Where BigClassName='"& ProDid &"'",Conn,1,1
			If Not TempRs.Eof Then
				ClassUrl = "/" & TempRs("FolderName") & "/" & TempRs("FileName")
				ClassUrl = GetHandleUrl(ClassUrl)	'处理URL 
				BigFolder = "/" & TempRs("FolderName") & "/" 
			End If:TempRs.Close
		End If
		'类别小类
		If ProSid <> "" Then
			ClassTitle = ProSid 
			TempRs.Open"Select * From [SmallClass] Where BigClassName='"& ProDid &"' And SmallClassName='"& ProSid &"'",Conn,1,1
			If Not TempRs.Eof Then
				ClassUrl = "/" & BigFolder & TempRs("FolderName") & "/" & TempRs("FileName")
				ClassUrl = GetHandleUrl(ClassUrl)	'处理URL 
				SmallFolder = "/" & TempRs("FolderName") & "/" 
			End If:TempRs.Close 
		End If
		
		'类别子类
		If ProSid <> "" Then
			TempRs.Open"Select * From [ThreeClass] Where BigClassName='"& ProDid &"' And SmallClassName='"& ProSid &"' Order By Sort",Conn,1,1
			While Not TempRs.Eof 				
				Tid = TempRs("ThreeClassName")
				Url = "/" & BigFolder & SmallFolder & TempRs("FolderName") & "/" & TempRs("FileName")
				Url = GetHandleUrl(Url)	'处理URL
				Target = TempRs("Target")
				Sel = IIF(Tid = ProTid,StyleClass("This"),"") 
				C = C & "<li><a "& AHref(Url,"",Target) & Sel & ">"& Tid &"</a></li>"  				
			TempRs.MoveNext:Wend:TempRs.Close
		End If
		'当ProSid不为空，但子类没有列表 则显示相类
		If C = "" Then		
			TempRs.Open"Select * From [SmallClass] Where BigClassName='"& ProDid &"' Order By Sort",Conn,1,1
			While Not TempRs.Eof 
				Sid = TempRs("SmallClassName")
				Url = "/" & BigFolder & TempRs("FolderName") & "/" & TempRs("FileName")
				Url = GetHandleUrl(Url)	'处理URL
				Target = TempRs("Target")
				Sel = IIF(Sid = ProSid,StyleClass("This"),"") 
				C = C & "<li><a "& AHref(Url,"",Target) & Sel & ">"& Sid &"</a></li>"  
			TempRs.MoveNext:Wend:TempRs.Close
			
		End If
	End If

	Code = Replace(Code, Replace(Label,"$}","InfoContent$}"), C)
	Code = Replace(Code, Replace(Label,"$}","Title$}"), ClassTitle)
	Code = Replace(Code, Replace(Label,"$}","FileName$}"), ClassUrl)
	
	MainClassList = Code
End Function
'产品高级搜索
function ProductSearch(addSql, ClassName, wd)
    dim i, s, c, Sel, url, title 
    c = c & "    <div class=""SearchBox"">" & vbCrLf 
    c = c & "        <form action=""/Inc/Create_Html.Asp"" method=""get"" name=""f"" onsubmit=""return CheckSearch(this);"">" & vbCrLf 
    c = c & "        <div class=""SearchLeft"">" & vbCrLf 
    c = c & "            <span class=""Title"">热门搜索:</span> " & vbCrLf 
    rs.open "Select * From [SearchCount] Where Through=True Order By Sort", conn, 1, 1 
    for i = 1 to 10
        if rs.eOF then exit for 
        title = rs("Title") 
        url = "/Inc/Create_Html.Asp?act=Search&wd=" & title 
        ShowTitle = title 
        ShowTitle = FontColorFontB(ShowTitle, rs("FontB"), rs("FontColor"))             '处理文本颜色与加粗
        if title = wd then
            c = c & "            " & TextFontB(title, true) & "&nbsp;" & vbCrLf 
        else
            c = c & "            <a " & AHref(url, title, rs("Target")) & ">" & ShowTitle & "</a>&nbsp;" & vbCrLf 
        end if 
    rs.moveNext : next : rs.close 
    c = c & "       </div>" & vbCrLf 
    c = c & "        <div class=""SearchRight"">" & vbCrLf 
    c = c & "            <select name=""ClassName"" id=""ClassName"">" & vbCrLf 
    c = c & "                <option value=""{分}{分}{分}"" Selected>所有栏目</option>" & vbCrLf 
    rs.open "Select * From [BigClass] " & addSql & " Order By Sort", conn, 1, 1 
    while not rs.eOF
        s = rs("BigClassName") & "{分}{分}{分}" 
        Sel = IIf(s = ClassName, " selected", "") 
        c = c & "                <option value=""" & s & """" & Sel & ">" & rs("BigClassName") & "</option>" & vbCrLf 
        rss.open "Select * From [SmallClass] Where BigClassName='" & rs("BigClassName") & "' Order By Sort", conn, 1, 1 
        while not rss.eOF
            s = rss("BigClassName") & "{分}" & rss("SmallClassName") & "{分}{分}" 
            Sel = IIf(s = ClassName, " selected", "") 
            c = c & "                <option value=""" & s & """" & Sel & "> &nbsp; ├-" & rss("SmallClassName") & "</option>" & vbCrLf 
            rsx.open "Select * From [ThreeClass] Where BigClassName='" & rss("BigClassName") & "' And SmallClassName='" & rss("SmallClassName") & "' Order By Sort", conn, 1, 1 
            while not rsx.eOF
                s = rsx("BigClassName") & "{分}" & rsx("SmallClassName") & "{分}" & rsx("ThreeClassName") & "{分}" 
                Sel = IIf(s = ClassName, " selected", "") 
                c = c & "                <option value=""" & s & """" & Sel & "> &nbsp;  &nbsp; ├-" & rsx("ThreeClassName") & "</option>" & vbCrLf 
            rsx.moveNext : wend : rsx.close 
        rss.moveNext : wend : rss.close 
    rs.moveNext : wend : rs.close 
    wd = IIF(wd <> "", wd, "请输入关键词") 
    c = c & "            </select> <input type=""text"" name=""wd"" id=""search-text"" size=""15"" onblur=""if(this.value=='') this.value='请输入关键词';"" onfocus=""if(this.value=='请输入关键词') this.value='';"" value=""" & wd & """ />  <input type=""submit"" id=""search-submit"" value=""搜索"" />" & vbCrLf 
    c = c & "        </div>" & vbCrLf 
    c = c & "    </form> " & vbCrLf 
    c = c & "     </div>" & vbCrLf 
    c = c & "     <div class=""clear2""></div>" & vbCrLf 

    ProductSearch = c 
end function

'================ 工作记录 ==========================
'显示工作列表
Sub ShowWorkList()
	Dim FolderPath,Content,SplStr,I,S,C,FilePath
	FolderPath = "/Help/工作记录/"
	If CheckFolder(FolderPath) = False Then Call CreateFolder(FolderPath)	'文件夹不存在 则创建
	Content = GetDirTxtList(FolderPath)
	Call Echo("工作记录","显示全部工作记录")
	FilePath = FolderPath & Format_Time(Now(),2) & ".txt"
	Call Echo("FilePath",FilePath)
	If CheckFile(FilePath) = False Then
		Call WriteToFile(FilePath,"=========== "& Format_Time(Now(),4) &" 工作日志 =============" & vbCrlf,1)
	End If
	SplStr=Split(Content,vbCrlf)
	For Each S In SplStr
		If S <> "" Then
			FilePath = S
			S = Right(S,Len(S)-InStrRev(S,"\"))
			Call Echo(S,"<a href=""?act=ShowWork&FilePath="& FilePath &""">点击查看</a>")
		End If
	Next
End Sub
'查看工作内容
Sub ShowWork(FilePath)
	Dim Content,C
	Content = GetFtext(FilePath) 
	Content = Replace(Content,"<br><br>","<br>")
	Call Rw(CreateAjax())               
	FilePath = Replace(FilePath,"\","*")
	C = Replace(OnLineEditJS(),"onblur=","onblur=Ajax('ajax.asp?act=SaveFile&FilePath="& FilePath &"','AjaxShow');")	 
	Call Rw(C)
	Call Echo("内容",OnLineEdit(Content,"Content"))
	Call Rw("<div id='AjaxShow'></div>")
End Sub
'VB软件上传产品信息 20140415
Sub AddProduct()
	'on error resume next
	Dim Did,Sid,Tid,Title,BigFiles,Content,ID,WebTitle,WebKeywords,WebDescription,DelStr,ZhuiJianWord

	Did = Request("BigClassName")
	Sid = Request("SmallClassName")
	Tid = Request("ThreeClassName")
	Title = Request("Title")
	BigFiles = Request("ImgPath")
	Content = Request("Content") 
	
	
	WebTitle = Request("WebTitle") 
	WebKeywords = Request("WebKeywords") 
	WebDescription = Request("WebDescription") 
	
	'随机词三个
	ZhuiJianWord = " " & GetMaisideKeyWord(3)
	WebTitle = Replace(WebTitle,"{$随机词3$}",ZhuiJianWord)
	WebKeywords = Replace(WebKeywords,"{$随机词3$}",ZhuiJianWord)
	WebDescription = Replace(WebDescription,"{$随机词3$}",ZhuiJianWord)
	
	
	Call Echo("上传信息","")
	Call Echo("上传网址",WebDoMain())
	Call Echo("BigClassName",Request("BigClassName"))
	Call Echo("SmallClassName",Request("SmallClassName"))
	Call Echo("ThreeClassName",Request("ThreeClassName"))
	Call Echo("Title",Request("Title"))
	Call Echo("WebTitle",WebTitle)
	Call Echo("WebKeywords",WebKeywords)
	Call Echo("WebDescription",WebDescription)
	Call Echo("ImgPath",Request("ImgPath"))
	Call Echo("Content",Request("Content"))
	Call OpenConn()
	Dim AddSql
	AddSql= "Where BigclassName='"& Did &"' And SmallClassName='"& Sid &"'"
	AddSql=AddSql & "And ThreeClassName='"& Tid &"' And Title='"& Title &"'"
	
	Rs.Open"Select * From [Product] " & AddSql,Conn,1,3
	If Rs.Eof Then
		Rs.AddNew
		Rs("BigClassName") = Did
		Rs("SmallClassName") = Sid
		Rs("ThreeClassName") = Tid
		Rs("Title") = Title
		Rs("BigFiles") = BigFiles
		Rs("Content") = Content
		'HTML文件名称
		Rs("FileName") = SetFileName(Title)
		Rs("WebTitle") = WebTitle
		Rs("WebKeywords") = WebKeywords
		Rs("WebDescription") = WebDescription
		Rs.Update
		ID = Rs("ID")
		DelStr = "，<a href='"& GetThisUrl &"?act=DeleteProduct&ID="& ID &"&P="& Request("P") &"'>点击删除</a>"
		Call Echo("添加成功", Title & "，ID" & ID & DelStr)
	Else
		ID = Rs("ID")
		DelStr = "，<a href='"& GetThisUrl &"?act=DeleteProduct&ID="& ID &"&P="& Request("P") &"'>点击删除</a>"
		Call Echo("添加失败", Title & "，已经存在"& DelStr)
		Exit Sub
	End If:Rs.Close
	'不显示回显HTML
	DisplayYes = False
	'生成文章
	Call MakeWebHtml("","","","","","Article", ID,"","")
	Call Hr()
	Doevents
	CreateHtmlPath = ""
	'生成指定类列表
	Call MakeWebHtml("","",Did,Sid,Tid,"Class","","","")
	
	If Err<>0 Then call doError(Err.Description, "Msg") : Call Echo("程序出错",Err.Description)
End Sub
'删除VB上传上去的产品
Sub DeleteProduct(AddSql)
	Dim Did,Sid,Tid,FilePath
	'密码为vbpass
	If md5(Request("P"),2) <> "134fe1ed" Then
		Call Eerr("删除失败","密码错误")
	End If
	Call OpenConn()
	Rs.Open"Select * From [Product] "&AddSql,Conn,1,3
	If Not Rs.Eof Then 
		FilePath = GetHandleFileName(Rs("FileName"))
		FilePath = DetailPageAddHtmlPrefix(Rs("FolderName")) & "/" & FilePath	
		'Call Echo("FilePath",FilePath)
		'删除文件
		Call DeleteFile(FilePath)
		Call Echo("删除","File=" & FilePath)
		Call Echo("删除","ID=" & Rs("Id"))
		Rs.Delete
		Rs.Update
	End If:Rs.Close
	Call Hr()
	Doevents
	CreateHtmlPath = ""
	'生成指定类列表
	Call MakeWebHtml("","",Did,Sid,Tid,"Class","","","")	
End Sub
'辅助区Start====================
'自动添关键词
Sub AotuAddKeyWord()
	Dim KeyWord,nOK,Title
	Call OpenConn()
	nOK = 0
	Rs.Open"Select * From [Product]",Conn,1,3
	While Not Rs.Eof
		If InStr(Rs("WebTitle"),"," )=0 Then
			KeyWord = GetMaisideKeyWord(3)
			If Rs("WebTitle") = "" Then
				Title = Rs("Title") & " " & KeyWord
			Else
				Title = Rs("WebTitle") & " " & KeyWord
			End If
			Rs("WebTitle") = Title
			Rs("WebKeywords") = Title
			Rs("WebDescription") = Title
			Rs.Update
			nOK = nOK + 1
		End If 
	Rs.MoveNext:Wend:Rs.Close
	Call Echo("给文章信息自动添加关键词完成","共添加【"& nOK &"】")
End Sub
'获得关键词
Function GetMaisideKeyWord(nCount)
	Dim Path,Content,SplStr,I,S,C,nOK
	Path = "../DataDir/麦思德关键词.txt"
	Content = GetFText(Path)
	SplStr=Split(Content, ",")
	nOK=0
	Randomize
	While nOK<nCount
		S=SplStr(Cint(Rnd*Ubound(SplStr))) & ","
		If InStr(","& C &",", ","& S & ",") = 0 Then
			C=C & S
			nOK=nOK+1
		End If
	Wend
	If C <> "" Then C=Left(C,Len(C)-1)
	GetMaisideKeyWord = C
End Function
'显示网站导航网站列表 在别人网站放链接
Function WebNavigationUsrList()
	Dim I, C, Url, FileName, BigFolder, Did, ShowDid, Target,AddSql
	AddSql = "Order By Sort Asc"
	Rs.Open "Select * From [NavBigClass] " & AddSql, Conn, 1, 1 
	For I = 1 To Rs.RecordCount
		FileName = Rs("FileName") 
		If FileName <> "" And Right(FileName,1)<>"/" Then FileName = FileName & ".html"  
	
		BigFolder = "/" & Rs("FolderName") & "/" 
		Did = Rs("BigClassName") 
		ShowDid = FontColorFontB(Did, Rs("FontB"), Rs("FontColor")) 
		Target = Rs("Target") 
		
		Url = Rs("FileName")
		If CheckMakeHtmlFile(Url)=True Then		
			Url = BigFolder & FileName 
			Url = Replace(Replace(Url, "//", "/"), "//", "/")
			Url = Left(Host,Len(Host)-1) & Url
			Call WebDebug(Url,"act=Nav&NavDid=" & Did)'调试
		End If
		
		C = C & "<a " & AHref(Url, "", "") & ">" & ShowDid & "</a>("& Url &") " 
	Rs.MoveNext : Next : Rs.Close 
	C = Replace(C, "{$麦思德汉堡新闻链接$}", GetClassUrl_Nav("麦思德汉堡新闻","",""))
	C = Replace(C, "{$成功案例链接$}", GetClassUrl_Nav("成功案例","",""))	
	C = Replace(C, "{$汉堡加盟菜单副导航链接$}", GetClassUrl_Nav("副导航","汉堡加盟菜单",""))	
	C="<div id=""tesi""><strong style=""font-weight: 400"">麦思德友情链接：" & vbCrlf & C & vbCrlf
	C=C & "<script>document.getElementById(""t""+""e""+""s""+""i"").style.display='none';</script>"
	Call Rw(C)
End Function
'辅助区End====================
%> 






