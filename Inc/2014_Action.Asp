<%
'http://127.0.0.1/获得网站Inc里全部函数给VB用.Asp    运行它 更新VB


'清空模块替换数组
function clearModuleReplaceArray()
	dim i
	for i=1 to ubound(ModuleReplaceArray)-1
		ModuleReplaceArray(i,0)=""
		ModuleReplaceArray(0,i)=""
	next
end function
'打印模块替换数组
function printModuleReplaceArray()
	dim i
	for i=1 to ubound(ModuleReplaceArray)-1
		if ModuleReplaceArray(i,0)="" then
			exit function
		end if
		call echo(ModuleReplaceArray(i,0),ModuleReplaceArray(0,i))
	next
end function


'多次处理模板动作
Function ManyTimesHandleTemplateAction(ByVal Content)
	'Php global $isHtmlFormatting;
	Dim I
	Content = ReplaceGlobleLable(Content)						'替换全局全部标签
	Content = RunHtmlAspCode(Content)							'运行Html里Asp代码
	For I =1 To 6
		Content = HandleTemplateAction(Content, True)
	Next
	Content = DelTemplateMyNote(Content)						'删除模板中我的注释 和无需要代码	
	Content = Replace(Content, "href=""/", "href=""" & Host())	'让网址路径完整
	Content = TrimVbCrlf(Content)								'去除两边空格与空行
	Content = ReplaceGlobleLable(Content)						'替换全局全部标签
	
	if isHtmlFormatting=true then
		content = HandleHtmlFormatting(content,false,0,"删除空行")
	end if
	
	'闭合标签开启
	if isWebLabelClose=true then
		'第一种，淘汰
		if 1=2 then
			Content=handleImgClose(Content,true)
			Content=handleBrClose(Content,false)
			Content=handleHrClose(Content,false)
			Content=handleHrClose(Content,false)
			Content=handleHtmlLabelClose(Content,"input",false)		
			Content=replace(Content,"onClick=","onclick=")
			Content=replace(Content,"onSubmit=","onsubmit=")
		else
			'第二种 超强版 不处理JS等改进
			Content=handleCloseHtml(Content,true,"")			'图片自动加alt  "|*|",
		end if		
	end if
	
	
	
	
	ManyTimesHandleTemplateAction = Content
End Function


'替换全部标签
Function ReplaceGlobleLable(Content)
	'Php global $WebSkins,$WebTemplate,$WebImages,$WebCss,$WebJs,$PubNavDid,$PubProDid;
	'Call Eerr(PubNavDid,PubProDid)
	Content = Replace(Content,"[$样式目录$]",WebSkins)			'网站样式目录
	Content = Replace(Content,"[$WebSkins$]",WebSkins)			'网站样式目录
	
	Content = Replace(Content,"[$WebSkinsIndex$]",WebSkins & "Index\")			'网站样式目录
	Content = Replace(Content,"[$样式Index目录$]",WebSkins & "Index\")			'网站样式目录
	
	Content = Replace(Content,"[$模版文件夹$]",WebTemplate)			'网站模板目录
	Content = Replace(Content,"[$模版路径$]",WebTemplate)			'网站模板目录
	Content = Replace(Content,"[$模板目录$]",WebTemplate)			'网站模板目录
	Content = Replace(Content,"[$WebTemplate$]",WebTemplate)		'网站模板目录
	
	Content = Replace(Content,"[$图片目录$]",WebImages)			'图片路径
	Content = Replace(Content,"[$WebImages$]",WebImages)		'图片路径
	
	Content = Replace(Content,"[$Css目录$]",WebCss)			'Css路径
	Content = Replace(Content,"[$WebCss$]",WebCss)			'Css路径
	
	Content = Replace(Content,"[$Js目录$]",WebJs)			'Js路径
	Content = Replace(Content,"[$WebJs$]",WebJs)		    'Js路径
	Content = Replace(Content,"[$UploadFiles$]","/UploadFiles/")		    		'上传路径
	Content = Replace(Content,"[$UploadFilesPic$]","/UploadFiles/Pic/")		    '上传Pic路径
	
	
	
	Content = Replace(Content,"[$模块目录$]",WebTemplate & "Module\")			'网站模块目录
	Content = Replace(Content,"[$TemplateModule$]",WebTemplate & "Module\")		'网站模块目录
	
	Content = Replace(Content,"[$域名$]", Host())		'域名
	Content = Replace(Content,"[$DoMain$]", Host())		'域名	
	
	Content = Replace(Content,"[$Did$]", IIF(PubNavDid<>"",PubNavDid,PubProDid))		'自定义大类名称	
	ReplaceGlobleLable = Content
End Function

'运行Html里Asp代码 20141221修改(之前是生成文件，再远程读取，现在改成本页执行)
Function RunHtmlAspCode(ByVal  Content)
	Dim AspCode,ExecuteAspCode, I
	For I = 1 To 8
		'获得内容里一节ASP代码
		AspCode = HandleGetAspCode(Content,"获取第一段")
		'Call Echo("AspCode",AspCode)
		'对这一节ASP代码处理成干净
		ExecuteAspCode = HandleGetAspCode(AspCode,"|第一段|去ASP开始结束标签|去注释部分|每行不为空|处理成Execute代码|")
		'Call Echo("ExecuteAspCode",ExecuteAspCode)
		'去除两边空格，其它上面已经处理了，再处理一下也没什么了。
		ExecuteAspCode=TrimVbCrlf(ExecuteAspCode)
		'为一行时，则追加一个ExecuteResut返回变量，好调用
		If Left(ExecuteAspCode,1)="=" Then
			ExecuteAspCode =  "ExecuteResut" & ExecuteAspCode
		End If
		
		'Call die(Replace(ExecuteAspCode,"""",""""""))
		'处理ASP代码 注意(处理elseif不行，会出错20141229)
		Execute(ExecuteAspCode)
		'Call eerr("ExecuteResut",ExecuteResut)
		
		If AspCode<>"" Then	
			'AspCode = "<!--#" & "Include File=""Inc/Config.Asp""-->" & vbCrlf & AspCode' & "<%" & AspCode & "%" & ">"
			'Call CreateFile("/Temp.Asp", AspCode)
			'AspCode=GetHttpPage(Host & "/Temp.Asp",1)
			'替换执行后ASP代码			
			Content = Replace(Content,AspCode,ExecuteResut)
		'否则退出Asp执行
		Else
			Exit For
		End If
	Next
	RunHtmlAspCode = Content
End Function
'模板处理动作 以后可以考虑用这种方式<@SetUp sdf@>
Function HandleTemplateAction(Content,CodeYes)
	Dim StartStr,EndStr,ActionList,SplStr,Action,S,HandYes
	dim isResetRunAction					'是否为重复运行动作	
	isResetRunAction = false				'默认为假
	StartStr = "{\$" : EndStr = "\$}"
	dim tempContent
	tempContent=Content	
	tempContent=DelTemplateMyNote(tempContent)			'让数据框里内容不运行，免得出错
	ActionList = GetArray(tempContent, StartStr, EndStr, True, True)
	SplStr=Split(ActionList,"$Array$")
	For Each S In SplStr
		Action=Trim(S)
		'Call Echo("Action",Action)
		Action = HandleInModule(Action,"start") 						'处理动作
		If Action<>"" Then
			
			Action = Trim(Mid(Action,3,Len(Action)-4)) & " "'因为上面对字符处理过了，所以不能用 Trim(Mid(Action, Len(StartStr), Len(Action) - Len(StartStr) - Len(EndStr) + 1))

			HandYes = True			'处理为真
			'屏蔽指定标签
			If pingBiLable(Action) Then
				Action = ""
			'{$#$} 这种是放在图片路径里，目的是为了在VB里不处理这个路径
			ElseIf CheckFunValue(Action, "# ")=True Then
				Action = ""
				
			'读模板样式并设置标题与内容   软件里有个栏目Style进行设置
			ElseIf CheckFunValue(Action, "ReadColumeSetTitle ")=True Then
				Action = XY_ReadColumeSetTitle(Action)
			'读模板资源内容
			ElseIf CheckFunValue(Action, "ReadTemplateSource ")=True Then
				Action = XY_ReadTemplateSource(Action)
			'读模块内容
			ElseIf CheckFunValue(Action, "ReadTemplateModule ")=True Then
				Action = XY_ReadTemplateModule(Action)
			'获得内容模块 20150108
			ElseIf CheckFunValue(Action, "GetContentModule ")=True Then
				Action = XY_GetContentModule(Action)
			'替换模板内容 20150109
			ElseIf CheckFunValue(Action, "ReplaceTemplateContent ")=True Then
				Action = XY_ReplaceTemplateContent(Content,S,Action)
			'显示模板模块20150209
			ElseIf CheckFunValue(Action, "ShowTemplateModule ")=True Then
				Action = XY_ShowTemplateModule(Action)
            '显示布局
            ElseIf CheckFunValue(action, "Layout ") = True Then
                action = XY_Layout(action)
            '显示模块
            ElseIf CheckFunValue(action, "Module ") = True Then
                action = XY_Module(action)
				
				
			'加载文件
			ElseIf CheckFunValue(Action, "Include ")=True Then
				Action = XY_Include(Action)
				
			'------------------- 文本区 -----------------------	
			'文本内容
			ElseIf CheckFunValue(Action, "MainInfo ")=True Then
				Action = XY_MainInfo(Action) 				
			'文本链接A
			ElseIf CheckFunValue(Action, "MainInfoA ")=True Then
				Action = XY_MainInfoA(Action)				
			'文本链接Url
			ElseIf CheckFunValue(Action, "MainInfoUrl ")=True Then
				Action = XY_MainInfoUrl(Action)
			'文本显示标题
			ElseIf CheckFunValue(Action, "MainInfo_GetShowTitle ")=True Then
				Action = XY_MainInfo_GetShowTitle(Action)
			'显示当前模板里内容
			ElseIf CheckFunValue(Action, "ShowThisText ")=True Then
				Action = XY_ShowThisText(Action)
			
			'------------------- 导航区 -----------------------		
			'导航大类链接A
			ElseIf CheckFunValue(Action, "NavDidA ")=True Then
				Action = XY_NavDidA(Action)
			'大类导航与小类导航链接
			ElseIf CheckFunValue(Action, "GetNavUrl ")=True Then
				Action = XY_GetNavUrl(Action)
			'判断导航大类是否为当前
			ElseIf CheckFunValue(Action, "NavDidExists ")=True Then
				Action = XY_NavDidExists(Action)
				
				
			
				
			'------------------- 链接区 -----------------------	
 
				
			'普通链接A
			ElseIf CheckFunValue(Action, "HrefA ")=True Then
				Action = XY_HrefA(Action)
				
			'Banner旗帜
			ElseIf CheckFunValue(Action, "BannerSwitch ")=True Then
				Action = XY_BannerSwitch(Action)
			'Banner旗帜2
			ElseIf CheckFunValue(Action, "BannerSwitch2 ")=True Then
				Action = XY_BannerSwitch2(Action)
			'详细页Banner图
			ElseIf CheckFunValue(Action, "DetailBanner ")=True Then
				Action = XY_DetailBanner(Action)
				
				
			
			'------------------- 新闻区 -----------------------					
			'News新闻
			ElseIf CheckFunValue(Action, "News ")=True Then
				Action = XY_News(Action)
			'显示新闻列表
			ElseIf CheckFunValue(Action, "ShowNewsList ")=True Then
				Action = XY_ShowNewsList(Action)
			'显示新闻相关信息
			ElseIf CheckFunValue(Action, "RelatedInformationList ")=True Then
				Action = XY_RelatedInformationList(Action)				
			'自定义新闻产品列表
			ElseIf CheckFunValue(Action, "CustomInfoList ")=True Then
				Action = XY_CustomInfoList(Action)	 
			'获得当前新闻信息
			ElseIf CheckFunValue(Action, "getArticleInfo ")=True Then
				Action = XY_getArticleInfo(Action)	
			
			'------------------- 分类区 -----------------------	
			'自定义留言列表
			ElseIf CheckFunValue(Action, "CustomGuestBookList ")=True Then
				Action = XY_CustomGuestBookList(Action)						
								 	
			'显示类别小类列表
			ElseIf CheckFunValue(Action, "SmallClassList ")=True Then
				Action = XY_SmallClassList(Action)
			'自定义显示类别小类列表
			ElseIf CheckFunValue(Action, "CustomSmallClassList ")=True Then
				Action = XY_CustomSmallClassList(Action)
			'显示类别子类列表
			ElseIf CheckFunValue(Action, "ThreeClassList ")=True Then
				Action = XY_ThreeClassList(Action)
			'显示自定义二级目录类型列表
			ElseIf CheckFunValue(Action, "CustomTwoDirectoryClassList ")=True Then
				Action = XY_CustomTwoDirectoryClassList(Action)
			'推荐导航，给内页左边使用
			ElseIf CheckFunValue(Action, "RecommendNavigation ")=True Then
				Action = XY_RecommendNavigation(Action)
				
			'自定义导航
			ElseIf CheckFunValue(Action, "CustomNavList ")=True Then
				Action = XY_CustomNavList(Action)
				
				
				
				
			'------------------- 链接区 -----------------------		
			'友情链接
			ElseIf CheckFunValue(Action, "Links ")=True Then
				Action = XY_Links(Action)
			'链接列表处理
			ElseIf CheckFunValue(Action, "LinkList ")=True Then
				Action = XY_LinkList(Action)
			'友情链接LI方式显示			
			ElseIf CheckFunValue(Action, "FriendLinkLI ")=True Then
				Action = XY_FriendLinkLI(Action)
				
					
			'------------------- URL区 -----------------------			
			'大类链接
			ElseIf CheckFunValue(Action, "GetDidAHrefUrl ")=True Then
				Action = XY_GetDidAHrefUrl(Action)
			'小类链接
			ElseIf CheckFunValue(Action, "GetSidAHrefUrl ")=True Then
				Action = XY_GetSidAHrefUrl(Action)
			'子类链接
			ElseIf CheckFunValue(Action, "GetTidAHrefUrl ")=True Then
				Action = XY_GetTidAHrefUrl(Action)
			'产品链接
			ElseIf CheckFunValue(Action, "GetProductUrl ")=True Then
				Action = XY_GetProductUrl(Action)
				
				
				
				
			'热门搜索词列表
			ElseIf CheckFunValue(Action, "PopularSearchesList ")=True Then
				Action = XY_PopularSearchesList(Action)				
				
				
			'------------------- 自动添加信息 -----------------------			
			'批量添加产品类别
			ElseIf CheckFunValue(Action, "BatchAutoAddProductClass ")=True Then
				Action = XY_BatchAutoAddProductClass(Action)
			'批量添加导航类别
			ElseIf CheckFunValue(Action, "BatchAutoAddNavClass ")=True Then
				Action = XY_BatchAutoAddNavClass(Action)
				
				
				
			'自动添加大类
			ElseIf CheckFunValue(Action, "AutoAddDid ")=True Then
				Action = XY_AutoAddDid(Action)
			'自动添加小类
			ElseIf CheckFunValue(Action, "AutoAddSid ")=True Then
				Action = XY_AutoAddSid(Action)
			'自动添加子类
			ElseIf CheckFunValue(Action, "AutoAddTid ")=True Then
				Action = XY_AutoAddTid(Action)
			'自动添加产品信息
			ElseIf CheckFunValue(Action, "AutoAddProduct ")=True Then
				Action = XY_AutoAddProduct(Action)
			'自动添加留言信息
			ElseIf CheckFunValue(Action, "AutoAddGuestBook ")=True Then
				Action = XY_AutoAddGuestBook(Action)
			'更新大类导航信息
			ElseIf CheckFunValue(Action, "UpdateBigNav ")=True Then
				Action = XY_UpdateBigNav(Action)
			'自动批量添加文章信息
			ElseIf CheckFunValue(Action, "AutoBatchAddArticle ")=True Then
				Action = XY_AutoBatchAddArticle(Action)
			'自动添加表单配置
			ElseIf CheckFunValue(Action, "AutoAddInputConfig ")=True Then
				Action = XY_AutoAddInputConfig(Action)
				
			'------------------- 图片 -----------------------			
			'图片列表
			ElseIf CheckFunValue(Action, "ShowPhotoList ")=True Then
				Action = XY_ShowPhotoList(Action)
			'显示批量模板样式展示列表
			ElseIf CheckFunValue(Action, "ShowAppointProductList ")=True Then
				Action = XY_ShowAppointProductList(Action)
			'图片缩放
			ElseIf CheckFunValue(Action, "zoomImage ")=True Then
				Action = XY_zoomImage(Action)
			
			 
			'------------------- QQ区 -----------------------	
			'浮动QQ列表显示
			ElseIf CheckFunValue(Action, "QQLIStyleListShow ")=True Then
				Action = XY_QQLIStyleListShow(Action)
			'在线QQ弹窗
			ElseIf CheckFunValue(Action, "OnLineQQPopup ")=True Then
				Action = XY_OnLineQQPopup(Action)
			'批量点击在线QQ网址
			ElseIf CheckFunValue(Action, "GetOnLineQQUrl ")=True Then
				Action = XY_GetOnLineQQUrl(Action)
				
				
			'------------------- 表单 -----------------------
			'显示提交表单(通用)20150211
			ElseIf CheckFunValue(Action, "FormSubmit ")=True Then
				Action = XY_FormSubmit(Action)
				
			'------------------- 网站地图区 -----------------------	
			'显示网站地图
			ElseIf CheckFunValue(Action, "SiteMap ")=True Then
				Action = XY_SiteMap(Action)
				
			'------------------- 文章区 -----------------------
			'文章大类名称
			ElseIf CheckFunValue(Action, "Article_Did ")=True Then
				Action = ArticleDid
			'文章小类名称
			ElseIf CheckFunValue(Action, "Article_Sid ")=True Then
				Action = ArticleSid
			'文章子类名称
			ElseIf CheckFunValue(Action, "Article_Tid ")=True Then
				Action = ArticleTid
			'文章标题
			ElseIf CheckFunValue(Action, "Article_Title ")=True Then
				Action = ArticleTitle
			'文章大图
			ElseIf CheckFunValue(Action, "Article_BigFiles ")=True Then
				Action = ArticleBigFiles
			'文章小图
			ElseIf CheckFunValue(Action, "Article_SmallFiles ")=True Then
				Action = ArticleSmallFiles
			'文章ID
			ElseIf CheckFunValue(Action, "Article_Id ")=True Then
				Action = ArticleID
			'文章时间
			ElseIf CheckFunValue(Action, "Article_DateTime ")=True Then
				Action = ArticleDateTime
			'文章作者
			ElseIf CheckFunValue(Action, "Article_Author ")=True Then
				Action = ArticleAuthor
			'文章内容
			ElseIf CheckFunValue(Action, "Article_Content ")=True Then
				Action = ArticleContent
				
			'------------------- Table表处理 -----------------------
			'获得数据库表列表
			ElseIf CheckFunValue(Action, "GetTableList ")=True Then
				Action = XY_GetTableList(Action)
			'删除Table表
			ElseIf CheckFunValue(Action, "DeleteTable ")=True Then
				Action = XY_DeleteTable(Action)
			'执行Table表
			ElseIf CheckFunValue(Action, "executeSQL ")=True Then
				Action = XY_executeSQL(Action)
				
				
			'------------------- 循环处理 -----------------------
			'For循环处理
			ElseIf CheckFunValue(Action, "ForArray ")=True Then
				Action = XY_ForArray(Action)
			
			'------------------- 参数处理 -----------------------
			'导航参数处理
			ElseIf CheckFunValue(Action, "GetNavParam ")=True Then
				Action = XY_GetNavParam(Action)
			'产品类型参数处理
			ElseIf CheckFunValue(Action, "GetClassParam ")=True Then
				Action = XY_GetClassParam(Action)
			'文章信息参数处理
			ElseIf CheckFunValue(Action, "GetArticleParam ")=True Then
				Action = XY_GetArticleParam(Action)
			'文本内容参数处理
			ElseIf CheckFunValue(Action, "GetMainParam ")=True Then
				Action = XY_GetMainParam(Action)
				
			'------------------- 缓冲处理 -----------------------	
			'清除缓冲
			ElseIf CheckFunValue(Action, "clearCache ")=True Then
				Action = XY_clearCache(Action)
				
				
				
				
			'------------------- 特殊处理一区 -----------------------
			'JS自动更新网页
			ElseIf CheckFunValue(Action, "JsAutoUpdateWebPage ")=True Then
				Action = JsContent
			'当前位置
			ElseIf CheckFunValue(Action, "Position ")=True Then
				Action = XY_Position(Action)
			'快速导航
			ElseIf CheckFunValue(Action, "FastNavigation ")=True Then
				Action = XY_FastNavigation(Action)
				
			'------------------- 特殊处理一区 -----------------------
			'Js调试网站皮肤
			ElseIf CheckFunValue(Action, "DebugWebSkin ")=True Then
				Action = XY_DebugWebSkin(Action)
				
			'------------------- 控制区 -----------------------
			'重复运行动作
			ElseIf CheckFunValue(Action, "ResetRunAction ")=True Then
				isResetRunAction=XY_ResetRunAction(Action)
				Action = ""
				
			'Js在线显示面板
			ElseIf CheckFunValue(Action, "onLineViewDialog ")=True Then
				Action = XY_onLineViewDialog(Action)
				 			
				
			'------------------- VB处理区 -----------------------
			'拷贝模板素材
			ElseIf CheckFunValue(Action, "copyTemplateMaterial ")=True Then
				Action = VB_copyTemplateMaterial(Action)			'拷贝模板素材
				
			
			'------------------- 特殊处理常用区 -----------------------
			'获得表字段内容
			ElseIf CheckFunValue(Action, "GetTableFieldValue ")=True Then
				Action = GetTableFieldValue(Action)			'获得表字段内容 				
			'网站顶部导航
			ElseIf CheckFunValue(Action, "WebTopNav ")=True Then				
				Action = HandleNavigation(Action,"")
			'网站底部导航
			ElseIf CheckFunValue(Action, "WebBottomNav ")=True Then					
				Action = XY_WebBottomNav(Action)							
			'网站标题
			ElseIf CheckFunValue(Action, "Web_Title ")=True Then
				Action = XY_WebTitle
			'网站关键词
			ElseIf CheckFunValue(Action, "Web_KeyWords ")=True Then
				Action = XY_WebKeywords
			'网站描述
			ElseIf CheckFunValue(Action, "Web_Description ")=True Then
				Action = XY_WebDescription
			'网站底部内容
			ElseIf CheckFunValue(Action, "WebBottom ")=True Or Action = "WebBottom" Then
				Action = XY_WebBottom(Action)			
			'主体导航名称
			ElseIf CheckFunValue(Action, "MainNav ")=True Then
				Action = MainNav
			'主体标题
			ElseIf CheckFunValue(Action, "MainTitle ")=True Then
				Action = MainTitle
			'主体内容
			ElseIf CheckFunValue(Action, "MainContent ")=True Then
				Action = MainContent
			'Js版网站统计
			ElseIf CheckFunValue(Action, "JsWebStat ")=True Then
				Action = XY_JsWebStat(action)
				
				
			Else			
				HandYes = False			'处理为假
			End If
			'注意这样，有的则不显示 晕 And IsNul(Action)=False
			If IsNUl(Action) =True Then Action=""
			If HandYes = True Then
				Content = Replace(Content,S,Action)
				'处理替换程序 20150712
				'call printModuleReplaceArray()
				Content = handleModuleReplaceArray(Content)
				call clearModuleReplaceArray()
				
				'为真则把内容给Code
				If CodeYes=True Then
					Code=Content			'把内容赋值给Code，因为Include加载这块会出错
				End If
			End If
			'重复运行动作，因为代码被处理了20150712
			if isResetRunAction then
				'call rw(content)
				HandleTemplateAction = HandleTemplateAction(Content,CodeYes)
				Exit Function
			end if
		End If
	Next
	HandleTemplateAction = Content
End Function

'删除Table表
function XY_DeleteTable(Action)
	dim sql,addsql,tableName
	tableName = RParam(Action,"tablename")
	addsql = RParam(Action,"addsql")
	sql = "select * from ["& tableName &"] "
	Sql = GetWhereAnd(Sql & " ",addsql)
	if CheckSql(Sql)=false then
		XY_DeleteTable = "sql出错=" & sql
		exit function
	end if
	sql = "delete from ["& tableName &"] "
	Sql = GetWhereAnd(Sql & " ",addsql)
	conn.execute(sql)
end function
'执行SQL(20150831)
function XY_executeSQL(Action)
	dim sql
	sql = RParam(Action,"sql")
	call CheckSql(Sql)
end function
'获得文章信息
function XY_getArticleInfo(Action)
	dim Default,s,Url,ImgFile,AStr,AUrl,ImgStr
	if ArticleId="" then
		XY_getArticleInfo = ""	
		exit function
	end if
	s = GetDefaultValue(Action)
	rs.open"Select * From [Product] Where id=" & ArticleId ,conn,1,1
	if not rs.eof then
		Url = Rs("FileName")
		If CheckMakeHtmlFile(Url)=True Then
			Url = GetHandleUrl( Rs("FileName") )
			Call WebDebug(Url,"act=CreateArticle&ID=" & Rs("Id") )'调试
		End If
		
		ImgFile = IIF(Rs("SmallFiles")<>"",Rs("SmallFiles"),Rs("BigFiles"))		
		AStr = "<a src='" & Rs("BigFiles") & "'" & SetHtmlParam(Action, "imgtitle|imgalt|imgid|imgclass|imgwidth|imgheight") & ">"
		AUrl = AHref(Url, Rs("Title"), Rs("Target")) & SetHtmlParam(Action, "aclass")
		ImgStr = "<img src='" & ImgFile & "'" & SetHtmlParam(Action, "imgtitle|imgalt|imgid|imgclass|imgwidth|imgheight") & ">"
	
	
		S = ReplaceValueParam(S,"id",  rs("id"))		
		S = ReplaceValueParam(S,"title",Rs("Title"))
		S = ReplaceValueParam(S,"showtitle",ShowTitle)
		S = ReplaceValueParam(S,"adddatetime",Rs("AddDateTime"))
		S = ReplaceValueParam(S,"imgfile",ImgFile)
		S = ReplaceValueParam(S,"imgurl",ImgFile)
		S = ReplaceValueParam(S,"url",url)
		S = ReplaceValueParam(S,"aurl",AUrl)
		S = ReplaceValueParam(S,"astr",AStr)
		S = ReplaceValueParam(S,"description",Rs("WebDescription"))
		S = ReplaceValueParam(S,"content",Rs("Content"))
		S = ReplaceValueParam(S,"price",Rs("Price"))
		S = ReplaceValueParam(S,"author",Rs("Author"))
		S = ReplaceValueParam(S,"nViewCount",Rs.RecordCount)		
	end if:rs.close
	XY_getArticleInfo = s
end function
'重复运行动作
function XY_ResetRunAction(Action)
	dim operate
	operate = LCase(RParam(Action,"operate"))
	if operate="true" then
		XY_ResetRunAction=true
	else
		XY_ResetRunAction = false
	end if
end function
'For循环处理
function XY_ForArray(Action)
	dim arrayList,splitStr,Default,splstr,forI,title,s,c,nloop
	arrayList = atRParam(Action,"arraylist")				'atRParam获得结果处理动作，但不替换动作内容
	splitStr = RParam(Action,"splitstr")
	nloop = RParam(Action,"nloop")				'循环数
	
	if arrayList="" then
		arrayList=copyStr("循环" & splitStr,nloop)
	end if
	
	Default = GetDefaultValue(Action) 
	
	splstr=split(arrayList,splitStr)
	for forI = 0 to Ubound(splstr)
		title=splstr(forI)
		if title <>"" then 
			s=default
			S = ReplaceValueParam(S,"fortitle",  title)
			S = ReplaceValueParam(S,"forid",  fori)
			S = ReplaceValueParam(S,"fori",  fori)
			S = ReplaceValueParam(S,"forcount",  ubound(splstr)+1)
			c=c & s
		end if
	next 
	XY_ForArray =c
end function
'获得数据库表列表
function XY_GetTableList(Action)
	dim sql,splitstr,fieldName,fieldNameList,splField,c
	sql = RParam(Action,"sql")
	splitstr = RParam(Action,"splitstr")
	fieldNameList = RParam(Action,"fieldnamelist")
	if lcase(splitstr)="vbcrlf" then splitstr=vbcrlf
	'call echo("fieldNameList",fieldNameList)
	if CheckSql(Sql)=false then
		XY_GetTableList = "sql出错=" & sql
		exit function
	end if
	if fieldNameList="" then
		XY_GetTableList = ""
		exit function
	end if
	fieldNameList=replace(fieldNameList,"，",",")
	splField=split(fieldNameList,",")
	rs.open sql,conn,1,1
	'call echo("记录总数",rs.recordcount)
	while not rs.eof
		for each fieldName in splField
			fieldName=trim(fieldName)
			'call echo(fieldName,"f")
			if c<>"" then c=c & splitstr
			c=c & rs(fieldName)
		next
	rs.movenext:wend:rs.close
	
	XY_GetTableList = c
end function
'提交表单(通用)20150211
Function XY_FormSubmit(Action)
	Dim C,AltStr,styleCss,nWidth,nHeight,SplStr,S,I,nOK
	
	Dim Did,TableName,FormName,ClassName,isVerificationCode,InputStr,InputStrList,InfoS	
	Dim submitname,resetname
	dim fieldName								'字段名称
	dim content,splxx,valueStr,showStr
	dim Placeholder										'提示内容
	dim passWordFieldName								'密码字段名称
	dim isAddTableNameInput								'是否追加TableName表单
	isAddTableNameInput=true							'默认为真
	
	submitname = RParam(Action,"submitname")
	resetname = RParam(Action,"resetname")
	If submitname = "" Then submitname="提交"
	If resetname = "" Then resetname="重置"
	
	'Dim InputStr								'Input内容
	Did = RParam(Action,"Did")
	ClassName = RParam(Action,"ClassName")
	TableName = RParam(Action,"TableName")
	FormName = RParam(Action,"FormName")
	isVerificationCode = RParam(Action,"isVerificationCode")

	C = "<div class="""& ClassName &""">" & vbCrlf
    C = C & "<form name="""& FormName &""" id="""& FormName &""" method=""post"" action=""/Inc/Create_Html.Asp?act=FormSubmit"" onSubmit=""return CheckForm(this);"">" & vbCrlf
	Rs.Open"Select * From [InputConfig] Where BigClassName='"& Did &"' And Through<>0 Order By Sort",Conn,1,1
	While Not Rs.Eof
		'字符名称，可自定义
		fieldName=Rs("FieldName")
		if fieldName="自定义" then
			fieldName=Rs("CustomFieldName")
		end if
		
		if isAddTableNameInput=true then
			if lcase(fieldName)="tablename" then
				isAddTableNameInput=false
			end if
		end if 
		
		Placeholder=Rs("Placeholder")
		if Placeholder<>"" then
			Placeholder = " placeholder="""& Placeholder &""""
		end if
	
		AltStr=""
		'判断
		If Rs("VerificationType") <> "" And Rs("VerificationType") <> "无" Then			
			AltStr="alt="""& IIF(Rs("VerificationMsg")<>"",Rs("VerificationMsg"),"请输入" & Rs("FrontTitle"))
			
			If Rs("VerificationType")<>"不为空" Then
				AltStr=AltStr & "{Array}["& Rs("VerificationType") &"]"
			End If			
			AltStr=AltStr & """"
		End If
		styleCss=""
		nWidth=LCase(Trim(Rs("InputWidth")))
		If nWidth<>"" Then
			nWidth = HandleCssPX(nWidth)
			styleCss = styleCss & IIF(styleCss=""," style='width:"& nWidth &";","width:"& nWidth &";")
		End If
		nHeight=LCase(Trim(Rs("InputHeight")))
		If nHeight<>"" Then
			nHeight = HandleCssPX(nHeight)
			styleCss = styleCss & IIF(styleCss=""," style='height:"& nHeight &";","height:"& nHeight &";")
			'Call Echo("","00")
			If Rs("InputType")<>"文本域" Then
				styleCss = styleCss & IIF(styleCss=""," style='line-height:"& nHeight &";","line-height:"& nHeight &";")
				'Call Eerr("","11")
			End If
		End If
		If styleCss <> "" Then styleCss=styleCss & "'"
		
		
		
		InfoS = "    <dl>" & vbCrlf
		InfoS = InfoS & "    <dt>"& Replace(Rs("FrontTitle")," ","&nbsp;") &"</dt>" & vbCrlf		'处理空格
		
		'单文本 文本域 多选 单选 下拉列表

  		InputStr="":InputStrList=""
		If Rs("InputType")="单文本" Or Rs("InputType")="单文本居中" Then
			If Rs("InputType")="单文本居中" Then
				If styleCss="" Then
					styleCss = " style='text-align:center;'"
				Else
					styleCss=replace(styleCss,"style='","style='text-align:center;")
				End If
			End If
			InfoS = InfoS & "    <dd><input type=""text"" name="""& fieldName &""" id="""& fieldName &""" "& AltStr &" value="""& Rs("DefaultValue") &""""& styleCss & Placeholder &"></dd>" & vbCrlf
		ElseIf Rs("InputType")="密码" Then
			passWordFieldName = fieldName
			InfoS = InfoS & "    <dd><input type=""password"" name="""& fieldName &""" id="""& fieldName &""" "& AltStr &" value="""& Rs("DefaultValue") &""""& styleCss & Placeholder &"></dd>" & vbCrlf
		ElseIf Rs("InputType")="确认密码" Then
			AltStr=Mid(AltStr,1,len(AltStr)-1) & passWordFieldName & """"
			InfoS = InfoS & "    <dd><input type=""password"" name="""& fieldName &""" id="""& fieldName &""" "& AltStr &" value="""& Rs("DefaultValue") &""""& styleCss & Placeholder &"></dd>" & vbCrlf
		ElseIf Rs("InputType")="文本域" Then
			InfoS = InfoS & "    <dd><textarea  name="""& fieldName &""" id="""& fieldName &""" "& AltStr & styleCss & Placeholder &">"& Rs("DefaultValue") &"</textarea></dd>" & vbCrlf
		
		'验证码
		ElseIf Rs("InputType")="验证码" Then
			If styleCss="" Then
				styleCss = " style='text-align:center;'"
			Else
				styleCss=replace(styleCss,"style='","style='text-align:center;")
			End If
			'call echo("fieldName",fieldName)
			if fieldName="" then
				fieldName="YZM"
			end if
			InfoS = InfoS & "    <dd><input type=""text"" name="""& fieldName &""" id="""& fieldName &""" "& AltStr &" value="""& Rs("DefaultValue") &""""& styleCss & Placeholder &"></dd>" & vbCrlf
			InfoS = InfoS & "    <dd><img src=""/Inc/Code_7.asp"" id='imgcaptcha' onClick=""this.src='/Inc/Code_7.asp?n='+Math.random()"" style=""cursor:pointer;"" title=""点击换一张验证码""><a href=""javascript:;"" onClick=""$('#imgcaptcha').click();"">&nbsp;看不清？换一张</a></dd>" & vbCrlf   '"& Session("YZM") &"   这个是验证码哟

		ElseIf Rs("InputType")="隐藏域" Then
			InfoS = InfoS & "    <dd>" & InputHiddenText(fieldName, Rs("DefaultValue")) & vbCrlf
		
		'单选
		ElseIf Rs("InputType")="单选" Then
			nOK=0
			SplStr=Split(Rs("DefaultValue"),"|")
			For Each S In SplStr
				If S <> "" Then
					nOK=nOK+1
					If InStr("|"& InputStrList &"|","|"& S &"|")=False Then
InputStr=InputStr & "<input type=""radio"" name="""& fieldName &""" id="""& fieldName & nOK &""" value="""& S &""" /><label for="""& fieldName & nOK &""">"& S &"</label>" & vbCrlf
						InputStrList=InputStrList & S & "|"
					Else
						InputStr = Replace(InputStr," value="""& S &""" />"," value="""& S &""" checked />")
					End If
				End If
			Next
			
			InfoS = InfoS & InputStr
			
			
		ElseIf Rs("InputType")="多选" Then
			nOK=0
			SplStr=Split(Rs("DefaultValue"),"|")
			For Each S In SplStr
				If S <> "" Then
					nOK=nOK+1
					If InStr("|"& InputStrList &"|","|"& S &"|")=False Then
InputStr=InputStr & "<input type=""checkbox"" name="""& fieldName &""" id="""& fieldName & nOK &""" value="""& S &""" /><label for="""& fieldName & nOK &""">"& S &"</label>" & vbCrlf
						InputStrList=InputStrList & S & "|"
					Else
						InputStr = Replace(InputStr," value="""& S &""" />"," value="""& S &""" checked />")
					End If
				End If
			Next
			
			InfoS = InfoS & InputStr
			
		ElseIf Rs("InputType")="下拉列表" Then
		
			InfoS = InfoS & "<select name="""& fieldName &""" id="""& fieldName &""">" & vbCrlf
			nOK=0
			content=Rs("DefaultValue")			
			content=replace(content,"\|","[$特殊字符A]$")
			content=replace(content,"\-","[$特殊字符B]$")
			SplStr=Split(Rs("DefaultValue"),"|")
			For Each S In SplStr
				If S <> "" Then
					nOK=nOK+1
					s=replace(s,"[$特殊字符A]$","|")
					s=replace(s,"[$特殊字符B]$","-")
					splxx=split(s & "-" & s,"-")
					valueStr=splxx(0)
					showStr=splxx(1)
					InfoS = InfoS & "<option value="""& valueStr &""" "& IIF(nOK=1,"selected","") & ">"& showStr &"</option>" & vbCrlf
				End If
			Next
			
			InfoS = InfoS & "</select>" & vbCrlf
			
		End If
		
		If Rs("TipsText")<>"" Then
			InfoS = InfoS & "    <dt>"& Rs("TipsText") &"</dt>" & vbCrlf
		End If
		
		InfoS = InfoS & "</dl>" & vbCrlf
		
		InfoS = DisplayOnlineED2(WEB_ADMINURL &"TableManage.Asp?act=ShowEdit&Id=" & Rs(0) & "&TableName=InputConfig"& "&n=" & GetRnd(11), InfoS, "<li|<a ")
		
		C=C & InfoS
		
	Rs.MoveNext:Wend:Rs.Close
 	
	C=C & InputHiddenText("DialogTitle",Did)
	if TableName<>"" and isAddTableNameInput=true then
    C = C & "<input name=""TableName"" type=""hidden"" id=""TableName"" value="""& TableName &""">" & vbCrlf
	end if
    C = C & "<div class=""submitbtn"">" & vbCrlf
    C = C & "<input type=""submit"" name=""Submit"" value="""& submitname &""" class=""btn white"">" & vbCrlf
    C = C & "<input name=""button"" type=""reset"" class=""btn white"" id=""button"" value="""& resetname &""">" & vbCrlf
    C = C & "</div>" & vbCrlf
    C = C & "</form>" & vbCrlf
    C = C & "</div>" & vbCrlf
	XY_FormSubmit = C
End Function

'快速导航
Function XY_FastNavigation(Action)
	Dim S,C,UlClass, LiClass, AClass, FocusClass, FocusText, Url,Sql,Did,Sid,NavType,DefaultSqlType,DefaultDid, SqlType
	Dim ShowTitle
	UlClass = RParam(Action,"UlClass")
	LiClass = RParam(Action,"LiClass")
	AClass = RParam(Action,"AClass")
	FocusClass = RParam(Action,"FocusClass")
	FocusText = LCase(RParam(Action,"FocusText"))
	DefaultSqlType = RParam(Action,"DefaultSqlType")
	DefaultDid = RParam(Action,"DefaultDid")
	'Call Echo("PubNavDid",PubNavDid)
	'Call Echo("PubProDid",PubProDid)
	SqlType = "nav"
	'Call Echo("Did",Did)
	Sid = PubNavSid
	If PubNavDid<>"" Then
		NavType = GetNavType(PubNavDid)		
		If NavType="文本" Then
			Sql = "Select * From [NavSmallClass] Where BigClassName='"& PubNavDid &"' Order By Sort"
		Else
			Sql = "Select * From [SmallClass] Where BigClassName='"& PubNavDid &"' Order By Sort"
			SqlType = "product"
			Sid = PubProSid
		End If	
	ElseIf PubProDid <> "" Then
		Sql = "Select * From [SmallClass] Where BigClassName='"& PubProDid &"' Order By Sort"
		SqlType = "product"
		Sid = PubProSid
	End If
	If Sql<>"" Then
		Rss.Open Sql, Conn, 1, 1
		If Rss.Eof And DefaultDid<>"" Then
			Rss.Close
			If DefaultSqlType = "nav" Then
				Sql = "Select * From [NavSmallClass] Where BigClassName='"& DefaultDid &"' Order By Sort"
			Else
				Sql = "Select * From [SmallClass] Where BigClassName='"& DefaultDid &"' Order By Sort"
			End If
			Rss.Open Sql, Conn, 1, 1
			'Call Eerr(Sql,NavType & "," & DefaultDid)
		End If
		While Not Rss.Eof
			If SqlType = "nav" Then
				Url = GetNavUrl(Rss("BigClassName"), Rss("SmallClassName"))
			Else
				Url = GetSmallClassFolderName(Rss("BigClassName"), Rss("SmallClassName"))
				'Url = GetClassUrl(Rss("BigClassName"), Rss("SmallClassName"), "")
				'Call Echo("Url",Url)
			End If
			'Call Echo(Rss("BigClassName"),Rss("SmallClassName"))
			
			S = Rss("SmallClassName")
			ShowTitle = FontColorFontB(S, Rss("FontB"), Rss("FontColor"))
			If Rss("SmallClassName") = Sid Then
				If FocusClass<>"" And FocusText="true" Then
					S = "<span class="""& FocusClass &""">"& ShowTitle &"</span>"
				Else
					If FocusClass<>"" Then
						S = "<a "& AHref(Url, Rss("SmallClassName"), Rss("Target")) &" class="""& FocusClass &""">"& ShowTitle &"</a>"	
					Else
						S = "<a "& AHref(Url, Rss("SmallClassName"), Rss("Target")) &">"& ShowTitle &"</a>"	
					End If
				End If
			Else
				If AClass<>"" Then
					S = "<a "& AHref(Url, Rss("SmallClassName"), Rss("Target")) &" class="""& AClass &""">"& ShowTitle &"</a>"	
				Else
					S = "<a "& AHref(Url, Rss("SmallClassName"), Rss("Target")) &">"& ShowTitle &"</a>"	
				End If
			End If
			
			If SqlType = "nav" Then
				S = DisplayOnlineEditDialog(""& adminDir &"NavManage.Asp?act=EditNavSmall&Id=" & Rss("Id") & "&n=" & GetRnd(11), S)
			Else
				S = DisplayOnlineEditDialog(""& adminDir &"ProductClassManage.Asp?act=ShowEditSmallClass&Id=" & Rss("Id") & "&n=" & GetRnd(11), S)
			End If
			
			
			'是否加LI
			If UlClass<>"" Or  LiClass<>"" Then
				If LiClass<>"" And LiClass<>"#" Then
					S = "<li class="""& LiClass &""">" & S
				Else
					S = "<li>" & S
				End If
				S = S & "</li>"
			End If
			C=C & S & vbCrlf
		Rss.MoveNext:Wend:Rss.Close
	End If
	'就否加UL
	If UlClass<>"" Or LiClass<>"" Then
		If UlClass<>"" And UlClass<>"#" Then
			C = "<ul class="""& UlClass &""">" & C
		Else
			C = "<ul>" & C
		End If
		C = C & "</ul>"
	End If
	
	XY_FastNavigation = C
End Function


'获得当前位置(网站)
Function XY_Position(Action)
    Dim Url, C, BigFolder, SmallFolder, Delimiter, nDirectory
	Dim Did,HomeStr, NavDidStr,NavSidStr,DidStr,SidStr,TidStr,ProStr,RootFolder
	Dim MainStr
	'Call Echo("PubNavDid", PubNavDid)
	'Call Echo("PubNavSid", PubNavSid)
		
	'Call Echo("PubProDid", PubProDid)
	'Call Echo("PubProSid", PubProSid)
	'Call Echo("PubProTid", PubProTid)	
	Delimiter = RParam(Action,"Delimiter")						'分隔符
	If Delimiter="" Then Delimiter = " > "
	
	'nDirectory(目录显示条数，3级为最好SEO用到)
	nDirectory = HandleNumber(RParam(Action,"nDirectory"))		'目录级别数
	
	If IsNul(nDirectory)=True Then
		nDirectory = 5
	Else
		nDirectory = Cint(nDirectory)
	End If
	
    '获得首页
    TempRs.Open "Select * From [NavBigClass] Where NavType='首页'", Conn, 1, 1
    If Not TempRs.Eof Then
        Url = "/" & GetWebFolderName() & TempRs("FolderName") & "/" & TempRs("FileName")
        Url = GetHandleUrl(Url)    '处理URL
		Call WebDebug(Url,"act=showweb")'调试
		HomeStr = "<a " & AHref(Url, "首页", TempRs("Target")) & ">首页</a>"
    End If : TempRs.Close
	
    '导航大类
	BigFolder = "/" & GetWebFolderName()
    If PubNavDid <> "" Or PubProDid<>"" Then
		Did = IIF(PubNavDid <> "", PubNavDid, PubProDid)
        TempRs.Open "Select * From [NavBigClass] Where BigClassName='" & Did & "'", Conn, 1, 1		
        If Not TempRs.Eof Then
            'Url = "/" & GetWebFolderName() & TempRs("FolderName") & "/" & TempRs("FileName")
            'Url = GetHandleUrl(Url)    '处理URL
			
			Url = GetNavUrl(TempRs("BigClassName"),"")
			
			
			Call WebDebug(Url,"act=Nav&NavDid=" & TempRs("BigClassName"))'调试
            BigFolder = "/" & GetWebFolderName() & TempRs("FolderName") & "/"
            NavDidStr = "<a " & AHref(Url, TempRs("BigClassName"), TempRs("Target")) & ">" & TempRs("BigClassName") & "</a>" 		
			
			NavDidStr = DisplayOnlineED2(WEB_ADMINURL &"NavManage.Asp?act=EditNavBig&Id=" & TempRs("Id") & "&n=" & GetRnd(11), NavDidStr, "<li|<a ")
			
        End If : TempRs.Close
    End If
    '导航小类
    If PubNavSid <> "" Then
        TempRs.Open "Select * From [NavSmallClass] Where BigClassName='" & PubNavDid & "' And SmallClassName='" & PubNavSid & "'", Conn, 1, 1
        If Not TempRs.Eof Then
		
			'Url = Trim(Replace(TempRs("FolderName"),"\","/"))
			'If Left(Url,1)<>"/" Then
			'	Url = BigFolder & "/" & Url & "/"
			'Else
			'	Url = "/" & TempRs("FolderName") & "/"
			'End If
			'Url = Url & TempRs("FileName")
			
			'Url = "/" & BigFolder & TempRs("FolderName") & "/" & TempRs("FileName")
            'Url = GetHandleUrl(Url)    '处理URL
			
			Url = GetNavUrl(TempRs("BigClassName"),TempRs("SmallClassName"))
			
			
			
			Call WebDebug(Url,"act=Nav&NavDid=" & TempRs("BigClassName") & "&NavSid=" & TempRs("SmallClassName"))'调试
            NavSidStr = "<a " & AHref(Url, TempRs("SmallClassName"), TempRs("Target")) & ">" & TempRs("SmallClassName") & "</a>"
			
			NavSidStr = DisplayOnlineED2(WEB_ADMINURL &"NavManage.Asp?act=EditNavSmall&Id=" & TempRs("Id") & "&n=" & GetRnd(11), NavSidStr, "<li|<a ")
        End If : TempRs.Close
    End If
	
    '类别大类	
	BigFolder = "/" & GetWebFolderName()
    If PubProDid <> "" And NavDidStr="" Then
        TempRs.Open "Select * From [BigClass] Where BigClassName='" & PubProDid & "'", Conn, 1, 1
        If Not TempRs.Eof Then

			'Url = "/" & GetWebFolderName() & TempRs("FolderName") & "/" & TempRs("FileName")
            'Url = GetHandleUrl(Url)    '处理URL
			Url = GetClassUrl(TempRs("BigClassName"),"","")
			
			Call WebDebug(Url,"act=CreateClass&ProDid=" & TempRs("BigClassName"))'调试
            BigFolder = "/" & GetWebFolderName() & TempRs("FolderName") & "/"
            DidStr = "<a " & AHref(Url, TempRs("BigClassName"), TempRs("Target")) & ">" & TempRs("BigClassName") & "</a>"
			
			DidStr = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditBigClass&Id=" & TempRs("Id") & "&n=" & GetRnd(11), DidStr, "<li|<a ")
        End If : TempRs.Close
    End If
    '类别小类
    If PubProSid <> "" Then
        TempRs.Open "Select * From [SmallClass] Where BigClassName='" & PubProDid & "' And SmallClassName='" & PubProSid & "'", Conn, 1, 1
        If Not TempRs.Eof Then
			'Url = Trim(Replace(TempRs("FolderName"),"\","/"))
			'If Left(Url,1)<>"/" Then
			'	Url = BigFolder & "/" & Url & "/"
			'Else
			'	Url = "/" & GetWebFolderName() & TempRs("FolderName") & "/"
			'End If			
			'Url = Url & TempRs("FileName")
			
			'Call Echo(Url,GetClassUrl(TempRs("BigClassName"), TempRs("SmallClassName"), ""))
			'直接截取小类网址，唉，程序写复杂了，有点乱 20150107
			
			'Url = "/" & BigFolder & TempRs("FolderName") & "/" & TempRs("FileName")
			
            'Url = GetHandleUrl(Url)    '处理URL .
			
			Url = GetClassUrl(TempRs("BigClassName"), TempRs("SmallClassName"), "")
			
			
			Call WebDebug(Url,"act=CreateClass&ProDid=" & TempRs("BigClassName") & "&ProSid=" & TempRs("SmallClassName"))'调试
            SmallFolder = "/" & TempRs("FolderName") & "/"
            SidStr = "<a " & AHref(Url, TempRs("SmallClassName"), TempRs("Target")) & ">" & TempRs("SmallClassName") & "</a>"
			
			
			SidStr = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditSmallClass&Id=" & TempRs("Id") & "&n=" & GetRnd(11), SidStr, "<li|<a ")
        End If : TempRs.Close
    End If
    '类别子类
    If PubProTid <> "" Then
        TempRs.Open "Select * From [ThreeClass] Where BigClassName='" & PubProDid & "' And SmallClassName='" & PubProSid & "' And ThreeClassName='" & PubProTid & "'", Conn, 1, 1
        If Not TempRs.Eof Then
		
		
			'Url = Trim(Replace(TempRs("FolderName"),"\","/"))
			'If Left(Url,1)<>"/" Then
			'	Url = RootFolder & SmallFolder & "/" & Url & "/"
			'Else
			'	Url = "/" & GetWebFolderName() & TempRs("FolderName") & "/"
			'End If
			'Url = Url & TempRs("FileName")
			
            'Url = "/" & BigFolder & SmallFolder & TempRs("FolderName") & "/" & TempRs("FileName")
            'Url = GetHandleUrl(Url)    '处理URL
			'Url = GetClassUrl(TempRs("BigClassName"), TempRs("SmallClassName"), TempRs("ThreeClassName"))
			
			Call WebDebug(Url,"act=CreateClass&ProDid=" & TempRs("BigClassName") & "&ProSid=" & TempRs("SmallClassName") & "&ProTid=" & TempRs("ThreeClassName"))'调试
            TidStr = "<a " & AHref(Url, TempRs("ThreeClassName"), TempRs("Target")) & ">" & TempRs("ThreeClassName") & "</a>"
			
			TidStr = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditThreeClass&Id=" & TempRs("Id") & "&n=" & GetRnd(11), TidStr, "<li|<a ")
        End If : TempRs.Close
    End If
	'显示产品信息
	If PubId <>"" Then
		TempRs.Open"Select * From [Product] Where Id=" & PubId,Conn,1,1
		If Not TempRs.Eof Then
			Url = TempRs("FileName")
			If CheckMakeHtmlFile(Url)=True Then
				Url = GetHandleUrl(TempRs("FileName"))
				Call WebDebug(Url, "act=CreateArticle&ID=" & TempRs("Id"))                      '调试
			End If
			ProStr =  "<a " & AHref(Url, TempRs("Title"), TempRs("Target")) & ">" & TempRs("Title") & "</a>"
			ProStr = DisplayOnlineED2(WEB_ADMINURL &"Product.Asp?act=ShowEditProduct&Id=" & TempRs("Id") & "&n=" & GetRnd(11), ProStr, "<li|<a ")
		End If : TempRs.Close
	End If
	
	
	If MainTitle<>"" Then
		TempRs.Open"Select * From [MainInfo] Where ShowTitle='"& MainTitle &"'",Conn,1,1
		If Not TempRs.Eof Then
			Url = GetMainInfoUrl(MainTitle)	
			If CheckMakeHtmlFile(Url)=True Then
				Url = GetHandleUrl(TempRs("FileName"))
				Call WebDebug(Url, "act=MainInfo&ID=" & TempRs("Id"))                      '调试
			End If
			MainStr =   "[-|_]" & "<a " & AHref(Url, TempRs("ShowTitle"), TempRs("Target")) & ">" & TempRs("ShowTitle") & "</a>"
			MainStr = DisplayOnlineED2(WEB_ADMINURL &"MainInfo.Asp?act=ShowEdit&Id=" & TempRs("Id") & "&n=" & GetRnd(11), MainStr, "<li|<a ")
		End If : TempRs.Close
		
		
	End If
	
	
	
	C = NavDidStr & "[-|_]" & NavSidStr & "[-|_]" & DidStr & "[-|_]" & SidStr & "[-|_]" & TidStr & "[-|_]" & ProStr & MainStr
	'Call Eerr("C",C)
    XY_Position = HomeStr & Delimiter & InAfterToBeforeGetArray(C,nDirectory, Delimiter)
End Function
'分割后从后向前摘取指定数组值 Intercept(截取数) AfterAdditionalStr(前追加内容)
'例：Call Echo("分割后从后向前摘取指定数组值",InAfterToBeforeGetArray("a[-|_]b[-|_]c[-|_]d[-|_]e[-|_]1[-|_]2[-|_]3[-|_]4",3))
Function InAfterToBeforeGetArray(Content, ByVal nIntercept, AfterAdditionalStr)
	Dim SplStr,I,S,C
	SplStr=Split(Content,"[-|_]")
	For I = Ubound(SplStr) To 0 Step - 1	
		S = SplStr(I)
		If S <>"" Then
			If C<>"" Then S = S & AfterAdditionalStr
			C = S & C
			nIntercept=nIntercept-1
			If nIntercept<=0 Then Exit For
		End If
	Next
	InAfterToBeforeGetArray = C
End Function


'获得导航参数
Function XY_GetNavParam(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	Dim TableName,FieldName
	TableName = RParam(Action,"TableName")
	FieldName = RParam(Action,"FieldName")
	XY_GetNavParam = GetNavParam(TableName,FieldName,AddSql)
End Function
'获得导航参数
Function XY_GetClassParam(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	Dim TableName,FieldName
	TableName = RParam(Action,"TableName")
	FieldName = RParam(Action,"FieldName")
	XY_GetClassParam = GetClassParam(TableName,FieldName,AddSql)
End Function
'获得文章参数
Function XY_GetArticleParam(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	Dim TableName,FieldName
	TableName = RParam(Action,"TableName")
	FieldName = RParam(Action,"FieldName")
	XY_GetArticleParam = GetArticleParam(TableName,FieldName,AddSql)
End Function
'获得文章参数
Function XY_GetMainParam(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	Dim TableName,FieldName
	TableName = RParam(Action,"TableName")
	FieldName = RParam(Action,"FieldName")
	XY_GetMainParam = GetMainParam(TableName,FieldName,AddSql)
End Function
'获得大类链接
Function XY_GetDidAHrefUrl(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	XY_GetDidAHrefUrl = GetBigClassFileName(Did)
End Function
'获得子类链接
Function XY_GetSidAHrefUrl(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	XY_GetSidAHrefUrl = GetSmallClassFileName(Did,Sid)
End Function
'获得子类链接
Function XY_GetTidAHrefUrl(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	XY_GetTidAHrefUrl = GetThreeClassFileName(Did,Sid,Tid)
End Function
'获得产品链接
Function XY_GetProductUrl(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	Dim Sql
	If Title<>"" Then Sql = "And Title='"& Title &"'"
	Sql = GetWhereAnd(GetDidSidTidSQl(Did,Sid,Tid),Sql)
	Sql = GetWhereAnd(Sql,AddSql)
	'Call Eerr("Sql", Sql)
	
    TempRs.Open "Select * From [Product] " & Sql, Conn, 1, 1
    If Not TempRs.Eof Then
        XY_GetProductUrl = GetHandleUrl(TempRs("FileName"))
		Call WebDebug(XY_GetProductUrl,"act=CreateArticle&ID=" & TempRs("Id") )'调试
    End If : TempRs.Close
End Function
'大类导航与小类导航链接
Function XY_GetNavUrl(Action)
	Dim Did, Sid, Url
	Did = RParam(Action,"Did")
	Sid = RParam(Action,"Sid")
	'Call Echo("Did",Did)
	'Call Echo("Sid",Sid)
	Url = GetNavUrl(Did,Sid)
	XY_GetNavUrl = Url
	'Call Echo("Url",Url)
End Function
'判断导航大类是否为当前
Function XY_NavDidExists(Action)
	dim navDid,valueStr,defaultStr
	navDid = RParam(Action,"navdid")
	valueStr = RParam(Action,"valueStr")
	defaultStr = RParam(Action,"defaultStr")
	'call echo(navdid,pubNavDid)
	if navdid=pubNavDid then
		XY_NavDidExists=valueStr
	else
		XY_NavDidExists=defaultStr	
	end if

end Function
'显示网站地图
Function XY_SiteMap(Action)
	Dim ClassName,NavAddSql, ClassAddSql
	NavAddSql = RParam(Action,"NavAddSql")
	ClassAddSql = RParam(Action,"ClassAddSql")
	ClassName = RParam(Action,"class")
	'Call Echo("NavAddSql",NavAddSql)
	'Call Echo("ClassAddSql",ClassAddSql)
	'XY_SiteMap = SiteMap(ClassName, NavAddSql, ClassAddSql)			'第一种，废弃掉
	XY_SiteMap = SiteMap(Action)			'新的
End Function
'AddSql = HandleInModule(AddSql,"end")

'浮动QQ列表显示
Function XY_QQLIStyleListShow(Action)
	Dim SplQQName,SplQQNumber,QQ,QQName,I,S,C,ShowQQPhoto,QQStyle
	ShowQQPhoto = RParam(Action,"ShowQQPhoto")
	Rs.Open "Select * From [QQMSN]", Conn, 1, 1
	If Not Rs.Eof Then
		SplQQName = Split(Rs("QQName"), "|")
		SplQQNumber = Split(Rs("QQNumber"), "|")
		QQStyle = Rs("QQStyle")
		'Call Echo(Rs("QQNumber"),Rs("QQName"))
		For I = 0 To UBound(SplQQNumber)
			QQ=SplQQNumber(I) : QQName=SplQQName(I)
			If QQName <> "" Then
				'Call Echo("QQName",QQName)
				S = "<li"& SetHtmlParam(Action,"liclass") &">"
				If ShowQQPhoto = "true" Then
					QQName="<img src='http://wpa.qq.com/pa?p=1:"& QQ &":"& QQStyle &"' align='absmiddle'>" & QQName
				End If
				C=C & S &"<a target=blank href='tencent://message/?uin=" & QQ & "&Site=com&Menu=yes' "& SetHtmlParam(Action,"aclass") &">"& QQName &"</a></li>" & vbCrlf
			End If
		Next
	End If : Rs.Close
	'Call Echo("C",C)
	XY_QQLIStyleListShow = C
End Function
'在线QQ弹窗 20150109
Function XY_OnLineQQPopup(Action)
	Dim QQList, QQNameList, SplQQ,SplName,Site,I,QQ,QQName,C
	Dim OnQQDialogue		'启用QQ对话
	Dim OnQQPopup			'启用QQ弹窗
	
	QQList = RParam(Action,"QQList")
	QQNameList = RParam(Action,"QQNameList")
	Site = RParam(Action,"Site")
	OnQQDialogue = RParam(Action,"OnQQDialogue")
	OnQQPopup = RParam(Action,"OnQQPopup")
	'Call Echo("OnQQDialogue",OnQQDialogue)
	'Call Echo("OnQQPopup",OnQQPopup)
	
	QQList=Replace(QQList, "，",",")
	QQNameList=Replace(QQNameList, "，",",")
	SplQQ = Split(QQList, ",")
	SplName = Split(QQNameList, ",")
	For I=0 To Ubound(SplQQ)
		QQ = SplQQ(I)
		QQName = ""
		If Ubound(SplName)>=I Then
			QQName = SplName(I)
		End If
		If OnQQDialogue="true" Then
			C=C & "window.location.href='tencent://message/?uin="& QQ &"&Site="& Site &"&Menu=yes';" & vbCrlf
		End If
		If OnQQPopup = "true" Then
			C=C & "window.location.href='tencent://AddContact/?fromId=50&fromSubId=1&subcmd=all&uin="& QQ &"';" & vbCrlf
		End If
	Next
	If C <>"" Then
		C="<script>"& C &"</script>"
	End If
	XY_OnLineQQPopup = C
End Function
'批量点击在线QQ网址
Function XY_GetOnLineQQUrl(Action)
	Dim Title,QQList,AutoAdd,IsOnlineChat,IsAddFriend
	Title = RParam(Action,"Title")
	QQList = RParam(Action,"QQList")
	QQList = Replace(Replace(Replace(QQList," ",""),",","|"),"，","|")
	AutoAdd = RParam(Action,"AutoAdd")
	IsOnlineChat = RParam(Action,"IsOnlineChat")
	IsAddFriend = RParam(Action,"IsAddFriend")
	
	Rs.Open"Select * From [OnLineQQ] Where Title='"& Title &"'",Conn,1,3
	If Not Rs.Eof Then
		XY_GetOnLineQQUrl = "javascript:OnLineQQ('"& Rs("QQ") &"',"& LCase(Rs("IsOnlineChat")) &", "& LCase(Rs("IsAddFriend")) &");"
	ElseIf AutoAdd="true" Then
		Rs.AddNew
		Rs("Title") = Title
		Rs("QQ") = QQList
		Rs("IsOnlineChat") = IIF(IsOnlineChat="true",True,False)
		Rs("IsAddFriend") = IIF(IsAddFriend="true",True,False)
		Rs.Update
		XY_GetOnLineQQUrl = "javascript:OnLineQQ('"& QQList &"',"& IsOnlineChat &", "& IsAddFriend &");"
	End If:Rs.Close
	XY_GetOnLineQQUrl = XY_GetOnLineQQUrl & """" & getControlStr(""& adminDir &"OnLineQQManage.Asp") & """"
		
	
End Function

'图片列表
Function XY_ShowPhotoList(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	Dim ULClass,Sql, SplTypeStr,I,J,S,C,AYes,Url,SplStartStr,SplEndStr,ImgFile,AStr,AUrl,ImgStr,AImg
	Call XY_AutoAddHandle(Action)			'处理添加信息处理
	
	SplTypeStr = RParam(Action,"SplTypeStr")
	AYes = LCase(RParam(Action,"AYes"))
	AImg = LCase(RParam(Action,"AImg"))
	SplStartStr = LCase(RParam(Action,"SplStartStr"))
	SplEndStr = LCase(RParam(Action,"SplEndStr"))
	Sql = "Select * From [Product] Where BigClassName='"& Did &"'"
	Sql = GetWhereAnd(Sql,AddSql)
	'Call Echo("Sql",Sql)
	Rs.Open Sql,Conn,1,1
	For I=1 To TopNumb
		If Rs.Eof Then Exit For
		
		Url = Rs("FileName")
		If CheckMakeHtmlFile(Url)=True Then
			Url = GetHandleUrl( Rs("FileName") )
			Call WebDebug(Url,"act=CreateArticle&ID=" & Rs("Id") )'调试
		End If
		
		ImgFile = IIF(Rs("SmallFiles")<>"",Rs("SmallFiles"),Rs("BigFiles"))
		S = ""
		If AImg = "true" Then
			S = "<img src='"& ImgFile &"'" & SetHtmlParam(Action,"imgtitle|imgalt|imgid|imgclass|imgwidth|imgheight") &">"
		End If
		If AYes="true" Then
			S="<a href='"& Url &"'" & SetHtmlParam(Action,"target|title|alt|id|class") &">"& S &"</a>"
		End If
		S = SplStartStr & S & SplTypeStr & SplEndStr & vbCrlf
		 	AStr = "<a src='" & Rs("BigFiles") & "'" & SetHtmlParam(Action, "imgtitle|imgalt|imgid|imgclass|imgwidth|imgheight") & ">"
			AUrl = AHref(Url, Rs("Title"), Rs("Target")) & SetHtmlParam(Action, "aclass")
			ImgStr = "<img src='" & ImgFile & "'" & SetHtmlParam(Action, "imgtitle|imgalt|imgid|imgclass|imgwidth|imgheight") & ">"
			
			
			'对i特殊处理，兼容之前设置
			S = Replace(S,"[$i$]",I)
			S = Replace(S,"[$i-1$]",I-1)
			'执行6次，争取处理全部标签			
			For J=1 To 6
				S = ReplaceValueParam(S,"ni",I)								'不对为i，因为i会与imgurl冲突 [$i$]
				S = ReplaceValueParam(S,"编号-1",I-1)								'不对为i，因为i会与imgurl冲突 [$i$]
				S = ReplaceValueParam(S,"编号",I)								'不对为i，因为i会与imgurl冲突 [$i$]
				S = ReplaceValueParam(S,"title",Rs("Title"))
				S = ReplaceValueParam(S,"showtitle",ShowTitle)
				S = ReplaceValueParam(S,"adddatetime",Rs("AddDateTime"))
				S = ReplaceValueParam(S,"imgfile",ImgFile)
				S = ReplaceValueParam(S,"imgurl",ImgFile)
				S = ReplaceValueParam(S,"url",url)
				S = ReplaceValueParam(S,"aurl",AUrl)
				S = ReplaceValueParam(S,"astr",AStr)
				S = ReplaceValueParam(S,"description",Rs("WebDescription"))
				S = ReplaceValueParam(S,"content",Rs("Content"))
				S = ReplaceValueParam(S,"price",Rs("Price"))
				S = ReplaceValueParam(S,"author",Rs("Author"))
			Next
			'加控制
			S = DisplayOnlineEditDialog(""& adminDir &"Product.Asp?act=ShowEditProduct&Id=" & Rs("Id") & "&n=" & GetRnd(11), S)
			
		C=C & S
	Rs.MoveNext:Next:Rs.Close
	XY_ShowPhotoList=C
End Function
'获得模板分割字段 Call GetTemplateSplit(Content,DialogStart,DialogSplit,DialogEnd)
Sub GetTemplateSplit(Content,DialogStart,DialogSplit,DialogEnd)
	Dim StartStr,EndStr
	'面板开始段
	StartStr = "<!--#DialogStart#-->" : EndStr = "<!--#DialogStartEnd#-->"
	If InStr(Content,StartStr)>0 And InStr(Content,EndStr)>0 Then
		DialogStart = StrCut(Content,StartStr,EndStr,2)
	End If
	'分割字符内容
	StartStr = "<!--#TemplateSplitStart#-->" : EndStr = "<!--#TemplateSplitEnd#-->"
	If InStr(Content,StartStr)>0 And InStr(Content,EndStr)>0 Then
		DialogSplit = StrCut(Content,StartStr,EndStr,2)
	End If
	'面板结束段
	StartStr = "<!--#DialogEnd#-->" : EndStr = "<!--#DialogEndEnd#-->"
	If InStr(Content,StartStr)>0 And InStr(Content,EndStr)>0 Then
		DialogEnd = StrCut(Content,StartStr,EndStr,2)
	End If
End Sub
'显示批量模板样式展示列表
Function XY_ShowAppointProductList(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql,ULClass
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	Call XY_AutoAddHandle(Action)			'处理添加信息处理
	
	Dim Sql, I,S,C,AYes,Url,ImgStr,ImgPath,AStr,AUrl,TemplateId,Content,DialogStart,DialogSplit,DialogEnd
	Dim ShowTitle,Price,WebDescription
	
	TemplateId = RParam(Action,"TemplateId")
	Content = ReadTemplateSource(TemplateId)
	Call GetTemplateSplit(Content,DialogStart,DialogSplit,DialogEnd)		'获得模板分割字段
	
	Sql = "Select * From [Product] Where BigClassName='"& Did &"'"
	Sql = GetWhereAnd(Sql,AddSql)
	Rs.Open Sql,Conn,1,1
	For I=1 To TopNumb
		If Rs.Eof Then Exit For
		
		Url = Rs("FileName")
		If CheckMakeHtmlFile(Url)=True Then
		Url = GetHandleUrl( Rs("FileName") )
			Call WebDebug(Url,"act=CreateArticle&ID=" & Rs("Id") )'调试
		End If
		
		ImgPath = IIF(IsNul(Rs("SmallFiles")),Rs("BigFiles"),Rs("SmallFiles"))
		ImgStr = "<img src='"& ImgPath &"'" & SetHtmlParam(Action,"imgtitle|imgalt|imgid|imgclass|imgwidth|imgheight") &">"
		
		AStr = "<a src='"& Rs("BigFiles") &"'" & SetHtmlParam(Action,"imgtitle|imgalt|imgid|imgclass|imgwidth|imgheight") &">"
        ShowTitle = CutStr(Rs("Title"), CutStrNumb, "... ")
        ShowTitle = FontColorFontB(ShowTitle, Rs("FontB"), Rs("FontColor"))		
		AStr = "<a " & AHref(Url, Rs("Title"), Rs("Target")) & SetHtmlParam(Action,"aclass") &">" & ShowTitle & "</a>" & vbCrLf
		AUrl = AHref(Url, Rs("Title"), Rs("Target")) & SetHtmlParam(Action,"aclass")
		
		S = Replace(DialogSplit,"[$title$]",ShowTitle)
		Price = IIF(isNul(Rs("Price")),0,Rs("Price"))
		S = Replace(S,"[$price$]",Price)
		S = Replace(S,"[$url$]",AUrl)
		WebDescription = IIF(isNul(Rs("WebDescription")),"",Rs("WebDescription"))
		S = Replace(S,"[$productabout$]",WebDescription)
		S = Replace(S,"[$showimages$]",ImgStr)
		S = Replace(S,"[$imgfile$]",ImgPath)
		S = Replace(S,"[$imgurl$]",ImgPath)
		
		C=C & S
	Rs.MoveNext:Next:Rs.Close
	XY_ShowAppointProductList= DialogStart & C & DialogEnd
End Function
'图片缩放(20151218)
Function XY_zoomImage(Action)
	dim imageFilePath,imgWidth, imgHeight, nSetWidth, nSetHeight, nLeft, nTop, sType,s,nWrapWidth,nWrapHeight
	imageFilePath = RParam(Action,"imageFilePath") 
	'图片不存在则退出
	if checkFile(imageFilePath)=false then
		exit function
	end if
	call getImageWidthHeight(imageFilePath,imgWidth,imgHeight)			'获得图片宽高
	
	nSetWidth = RParam(Action,"nSetWidth")
	nSetHeight = RParam(Action,"nSetHeight")
	nLeft = RParam(Action,"nLeft")
	nTop = RParam(Action,"nTop")
	
	nWrapWidth = RParam(Action,"nWrapWidth")					'包宽
	nWrapHeight = RParam(Action,"nWrapHeight") 					'包高
	
	sType = RParam(Action,"sType")
	if sType<>"" then sType="|"& sType &"|"
	
	'call echo("nSetWidth",nSetWidth)
	'call echo("nSetHeight",nSetHeight)
	
	s=zoomImage(nWrapWidth,nWrapHeight,imgWidth, imgHeight, nSetWidth, nSetHeight, nLeft, nTop, sType)
	XY_zoomImage=s'imageFilePath & "("& imgWidth &","& imgHeight &")"
end function
'热门搜索列表
Function XY_PopularSearchesList(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql,ULClass
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	Dim SplTypeStr
	SplTypeStr = RParam(Action,"SplTypeStr")
	XY_PopularSearchesList = PopularSearchesList(Action,TopNumb,SplTypeStr,AddSql)
End Function
'显示类别小类列表
Function XY_SmallClassList(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql,ULClass
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	Call XY_AutoAddHandle(Action)			'处理添加信息处理
	
	Dim LiClass,AClass
	LiClass=RParam(Action,"LiClass")
	AClass=RParam(Action,"AClass")
	'Call Echo(Action & "LiClass",LiClass)
	'Call Echo("AClass",AClass)
	Dim S,C
	S = SmallClassList(Did, LiClass, AClass, AddSql)
	
	ULClass = RParam(Action,"ULClass")
	If ULClass<>"" Then
		C = vbCrlf & "<ul class='"& ULClass &"'>"
	Else
		C = vbCrlf & "<ul>"
	End If
	
	C = C & vbCrlf & S & vbCrlf & "</ul>" & vbCrlf	
	XY_SmallClassList = C
End Function
'自定义显示小类别类列表
Function XY_CustomSmallClassList(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)		'获得这么函数参数
	Call XY_AutoAddHandle(Action)			'处理添加信息处理
	Dim Sql,Default,I,S,C,StartStr,EndStr,Url,ShowTitle,DateTime,ImgFile
	Default = GetDefaultValue(Action)
	Sql = "Select * From [SmallClass]" & GetDidSidTidSQl(Did,"","")
	Sql =  GetWhereAnd(Sql,AddSql)		
	Rs.Open Sql,Conn,1,1
	For I = 1 To TopNumb
		If Rs.Eof Then Exit For
		StartStr="[list-"& I &"]" : EndStr="[/list-"& I &"]"
		'没有则用默认
		If InStr(Default,StartStr)=False Then
			StartStr="[list]" : EndStr="[/list]"
		End If
		If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
			S = StrCut(Default,StartStr,EndStr,2)
						
			S = Replace(S,"[$bigclassname$]",Rs("BigClassName"))
			S = Replace(S,"[$smallclassname$]",Rs("SmallClassName"))
			
			
			Url = GetSmallClassFileName( Rs("BigClassName"), Rs("SmallClassName") )
			Call WebDebug(Url,"act=CreateClass&ProDid=" & Rs("BigClassName") & "&ProSid=" & Rs("SmallClassName") & "&page=1")'调试
			S = Replace(S,"[$url$]",Url)
						
			'开始位置加Dialog内容
			StartStr="[list-"& I &" startdialog]" : EndStr="[/list-"& I &" startdialog]"
			If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
				S = StrCut(Default,StartStr,EndStr,2) & S
			End If
			'结束位置加Dialog内容
			StartStr="[list-"& I &" enddialog]" : EndStr="[/list-"& I &" enddialog]"
			If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
				S = S & StrCut(Default,StartStr,EndStr,2)
			End If			
			C = C & S
		End If	
	Rs.MoveNext:Next:Rs.Close
	
	'开始内容加Dialog内容
	StartStr="[dialog start]" : EndStr="[/dialog start]"
	If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
		C = StrCut(Default,StartStr,EndStr,2) & C
	End If
	'结束内容加Dialog内容
	StartStr="[dialog end]" : EndStr="[/dialog end]"
	If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
		C = C &  StrCut(Default,StartStr,EndStr,2)
	End If
	
	XY_CustomSmallClassList = C
End Function
'显示类别子类列表
Function XY_ThreeClassList(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql,ULClass
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	Call XY_AutoAddHandle(Action)			'处理添加信息处理
	Dim TidAType									'子类链接
	TidAType = RParam(Action,"TidAType")	'子类链接
	
	Dim ThreeLiCss,ThreeACss
	ThreeLiCss=RParam(Action,"ThreeLiCss"):ThreeACss=RParam(Action,"ThreeACss")
	Dim S,C
	
	S = ThreeClassList(Did,ThreeLiCss, ThreeACss, TidAType, AddSql)
	
	ULClass = RParam(Action,"ULClass")
	If ULClass<>"" Then
		C = vbCrlf & "<ul class='"& ULClass &"'>"
	Else
		C = vbCrlf & "<ul>"
	End If
	
	C = C & vbCrlf & S & vbCrlf & "</ul>" & vbCrlf	
	XY_ThreeClassList = C
End Function

'推荐导航，给内页左边使用 20150204
Function XY_RecommendNavigation(Action)
	Dim Id,SplStr,Splxx,I,S,C,Did,Url,UrlList
	Dim UlClass
	Dim Title,ShowTitle,AutoAdd,Content	
	
	Did = PubNavDid
	If Did="" Then Did=PubProDid
	
	'Call Echo("PubNavDid",PubNavDid)
	'Call Echo("PubProDid",PubProDid)
	'Call Echo("Did",Did)
	
	
	UlClass = RParam(Action,"UlClass")	'子类链接
	If UlClass <> "" Then UlClass = " class="""& UlClass &""""
	
	Title = "推荐导航"
	ShowTitle = "推荐导航"
	AutoAdd = "true"
	Content = ReadMainInfo(Title)
	If Content="" Then
		Content = GetDefaultValue(Action)
		'是否自动添加信息
		If LCase(AutoAdd)="true" Then
			Call AutoAddMainInfo(Title,ShowTitle,Content)
			Conn.Execute("Update [MainInfo] Set DispalyEditor=False Where Title='"& Title &"'")
		End If
	End If
	Content = TrimVbCrlf(Content)
	SplStr=Split(Content,vbCrlf)
	For Each S In SplStr
		S = TrimVbCrlf(S)
		If S<>"" Then
			UrlList = S
			If InStr("，"& S &"，","，"& Did &"，")>0 Then
				Exit For
			End If
		End If
	Next
	'Call Echo("UrlList", UrlList)
	UrlList=Replace(UrlList,"," , "，")
	UrlList=Replace(Replace(UrlList,"\，","【撇分号】"), "\|", "【撇分割】")
	If UrlList<>"" Then
		SplStr=Split(UrlList,"，")
		For Each S In SplStr
			Splxx=Split(S & "||", "|")
 			Title=Splxx(0):ShowTitle=IIF(Splxx(1)="",Splxx(0),Splxx(1)):Url=Splxx(2)
				'从导航找
				If Url="" Then
					Url = GetNavUrl(Title,"")
				End If
				'从类型分类找
				If Url="" Then
					Url = GetClassUrl(Title,"","")
				End If
				'从文件内容找
				If Url="" Then
					Url = GetMainUrl(Title)
				End If
			
			If Title=Did Then
				ShowTitle="<font color=red>"& ShowTitle &"</font>"
			End If
			'Call Echo(Title,Url)
			C = C & "<li><a href="""& Url &""">"& ShowTitle &"</a></li>" & vbCrlf
		Next
	End If
	Id = ReadMainId("推荐导航")
	If Id<>"" And Request("gl") = "edit" Then
		C = DisplayOnlineEditDialog(""& adminDir &"MainInfo.Asp?act=ShowEdit&Id=" & Id & "&n=" & GetRnd(11), C)
	End If
	
	C= "<ul"& UlClass &">"& C &"</ul>"
	XY_RecommendNavigation = C
End Function

'显示自定义二级目录分类列表
Function XY_CustomTwoDirectoryClassList(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)		'获得这么函数参数
	Dim Default,I,J,S,C,StartStr,EndStr,Url,ShowTitle,DateTime,ImgFile,DefaultImage,RndShow,IdList
	Dim SmallStr,SmallTitle,Sql
	Dim SidAType									'小类链接
	Dim TidAType									'子类链接
	Dim SmallUrl									'小类网址
	Dim SmallAddSql									'子AddSql
	SmallAddSql = RParam(Action,"SmallAddSql")	    '子类链接
	
	Default = GetDefaultValue(Action)
	DefaultImage = RParam(Action,"DefaultImage")	'默认图片
	SidAType = RParam(Action,"SidAType")	'小类链接
	TidAType = RParam(Action,"TidAType")	'子类链接
	
	
	Sql = "Select * From [SmallClass]" & GetDidSidTidSQl(Did,"","")
	Sql =  GetWhereAnd(Sql,AddSql)		
	Rs.Open Sql,Conn,1,1
	For I = 1 To Rs.RecordCount
		StartStr="[smalllist-"& I &"]" : EndStr="[/smalllist-"& I &"]"
		'没有则用默认
		If InStr(Default,StartStr)=False Then
			StartStr="[smalllist]" : EndStr="[/smalllist]"
		End If
		SmallStr="":SmallTitle=""
		If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
			SmallTitle = StrCut(Default,StartStr,EndStr,2) & vbCrlf
			SmallTitle=Replace(SmallTitle, "[$smallname$]", Rs("SmallClassName"))
			
			SmallUrl = getClassUrl(Rs("BigClassName"),Rs("SmallClassName"),"")					'获得小类网址
			
			SmallTitle = Replace(SmallTitle, "[$smallurl$]", SmallUrl)
			'Call Echo("SmallUrl",SmallUrl)
			
			Sql = "Select * From [ThreeClass] Where BigClassName='"& Rs("BigClassName") &"' And SmallClassName='"& Rs("SmallClassName") &"'"
			Sql =  GetWhereAnd(Sql,SmallAddSql)		
			Rsx.Open Sql,Conn,1,1
			For J = 1 To Rsx.RecordCount
				StartStr="[threelist-"& I &"]" : EndStr="[/threelist-"& I &"]"
				'没有则用默认
				If InStr(Default,StartStr)=False Then
					StartStr="[threelist]" : EndStr="[/threelist]"
				End If
				If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
					S = StrCut(Default,StartStr,EndStr,2) & vbCrlf
					Url = GetThreeClassFolderName( Rsx("BigClassName"), Rsx("SmallClassName"), Rsx("ThreeClassName")) 		
					
					
					Call WebDebug(Url,"act=CreateClass&ProDid=" & Did & "&ProSid=" & Rs("SmallClassName"))'调试					
					If TidAType="文本类型" Then
						Url = "{$GetProductUrl did='"& Rsx("BigClassName") &"' sid='"& Rsx("SmallClassName") &"' tid='"& Rsx("ThreeClassName") &"' title='"& Rsx("ThreeClassName") &"'$}"
					End If		
								
					S = Replace(S, "[$smallname$]", Rsx("SmallClassName"))
					S = Replace(S, "[$threename$]", Rsx("ThreeClassName"))
					S = Replace(S, "[$threeurl$]", Url)
					
					'S = DisplayOnlineEditDialog(""& adminDir &"ProductClassManage.Asp?act=ShowEditThreeClass&Id=" & Rsx("Id") & "&n=" & GetRnd(11), S)
					'加强于20150129	
					S = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditThreeClass&Id=" & Rsx("Id") & "&n=" & GetRnd(11), S, "<li|<a ")
					
					SmallStr=SmallStr & S & vbCrlf
					
				End If
			Rsx.MoveNext:Next:Rsx.Close
		End If
		'开始内容加Dialog内容
		StartStr="[smalldialog start]" : EndStr="[/smalldialog start]"
		If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
			SmallStr = StrCut(Default,StartStr,EndStr,2) & vbCrlf & SmallStr
		End If
		'结束内容加Dialog内容
		StartStr="[smalldialog end]" : EndStr="[/smalldialog end]"
		If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
			SmallStr = SmallStr & vbCrlf & StrCut(Default,StartStr,EndStr,2)
		End If
		
		'SmallTitle = DisplayOnlineEditDialog(""& adminDir &"ProductClassManage.Asp?act=ShowEditSmallClass&Id=" & Rs("Id") & "&n=" & GetRnd(11), SmallTitle)
		'加强于20150129	
		SmallTitle = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditSmallClass&Id=" & Rs("Id") & "&n=" & GetRnd(11), SmallTitle, "<li|<a ")
		
		
		
		C=C & SmallTitle & SmallStr & vbCrlf				'小类标题+小类内容
			
	Rs.MoveNext:Next:Rs.Close
	
	'开始内容加Dialog内容
	StartStr="[dialog start]" : EndStr="[/dialog start]"
	If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
		C = StrCut(Default,StartStr,EndStr,2) & vbCrlf & C
	End If
	'结束内容加Dialog内容
	StartStr="[dialog end]" : EndStr="[/dialog end]"
	If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
		C = C & vbCrlf &   StrCut(Default,StartStr,EndStr,2)
	End If
	
	
	XY_CustomTwoDirectoryClassList = C
End Function
'新闻类
Function XY_News(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql,AutoAddDid,AutoAdd,I,ULClass
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)		'获得这么函数参数
	Dim S,C,DateYes,AutoAddSid,AutoAddTid
	Dim CssStyle															'Css样式
	Dim ContentWidth														'内容宽度
	ContentWidth = LCase(Trim(RParam(Action,"ContentWidth")))				'内容宽度20150130
	If ContentWidth<>"" Then
		If Right(ContentWidth,1)<>"%" And Right(ContentWidth,1)<>"px" Then
			ContentWidth = ContentWidth & "px"
		End If
	End If
	If ContentWidth <>"" Then
		CssStyle = " style=""width:"& ContentWidth &";"""
	End If
	
	
	'判断大类
	If CheckBigClass(Did) = 0 Then
		AutoAddDid = RParam(Action,"AutoAddDid")
		If LCase(AutoAddDid)="true" Then
			Call AutoAddBigClass(Did)
		End If		
	End If
	'判断小类
	If CheckSmallClass(Did,Sid) = 0 Then
		AutoAddSid = RParam(Action,"AutoAddSid")
		If LCase(AutoAddSid)="true" Then
			Call AutoAddSmallClass(Did,Sid)
		End If		
	End If
	'判断子类
	If CheckThreeClass(Did,Sid,Tid) = 0 Then
		AutoAddTid = RParam(Action,"AutoAddTid")
		If LCase(AutoAddTid)="true" Then
			Call AutoAddThreeClass(Did,Sid,Tid)
		End If		
	End If
	
	If CheckProduct(Did,Sid,Tid)=0 Then				'检测是否有记录
		AutoAdd = RParam(Action,"AutoAdd")
		If LCase(AutoAdd)="true" Then
		 	For I = 1 To TopNumb
				Call AutoAddArticle(Did, Sid, Tid, Did & "-标题" & I, Did & "-内容" & I)		
			Next
		End If	
	End If
	DateYes = RParam(Action,"DateYes")						'是否显示时间
	S =  NewsList(Did,TopNumb,"<li>","</li>" & vbCrlf,DateYes, Action)
	ULClass = RParam(Action,"ULClass")
	If ULClass<>"" Then
		C = vbCrlf & "<ul class='"& ULClass &"'"& CssStyle &">"
	Else
		C = vbCrlf & "<ul>"
	End If
	C = C & vbCrlf & S & vbCrlf & "</ul>" & vbCrlf	
	XY_News = C
End Function
'自定义信息列表
Function XY_CustomInfoList(Action)
	XY_CustomInfoList=HandleDispalyArticleList(Action,"")
End Function
'显示文章相关信息列表
Function XY_RelatedInformationList(Action)
	Dim Sql,AddSql,DefaultAddSql,DefaultRelatedLableList,SplStr,Word,C
	DefaultRelatedLableList = RParam(Action,"DefaultRelatedLableList")	'默认相关标签列表
	AddSql = RParam(Action,"AddSql")	'追加SQL
	DefaultAddSql = RParam(Action,"DefaultAddSql")	'默认追加SQL
	DefaultAddSql = HandleInModule(DefaultAddSql,"end")
	'文章ID不为空
	If PubId<>"" Then
		Sql = GetWhereAnd(Sql, "And Id<>" & PubId)
		Rs.Open"Select * From [Product] Where Id=" & PubId,Conn,1,1
		If Not Rs.Eof Then
			If Rs("RelatedTags")<>"||" And Rs("RelatedTags")<>"" Then
				DefaultRelatedLableList = Rs("RelatedTags")
				'Call Eerr("RelatedTags",Rs("RelatedTags"))
			End If
		End If:Rs.Close
	Else
		Sql = GetWhereAnd(Sql & " ",DefaultAddSql)
	End If
	DefaultRelatedLableList = Replace(DefaultRelatedLableList,"'","")		'去除'符号
	'自定义标签不为空则
	If DefaultRelatedLableList<>"" Then
		SplStr=Split(DefaultRelatedLableList,"|")
		For Each Word In SplStr
			If Word<>"" Then
				Word = "|"& Word &"|"
				Sql = GetWhereAnd(Sql, "And RelatedTags Like '%"& Word &"%'")
			End If	
		Next
	Else
		Sql = GetWhereAnd(Sql & " ",DefaultAddSql)
	End If
	Sql =  GetWhereAnd(Sql & " ",AddSql)
	'Call Echo(DefaultAddSql,Sql)
	XY_RelatedInformationList = HandleDispalyArticleList(Action,Sql)
End Function
'显示文章列表  修改2014 12 01
Function HandleDispalyArticleList(Action,Sql)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)		'获得这么函数参数
	Call XY_AutoAddHandle(Action)			'处理添加信息处理
	Dim Default,I,J,S,C,StartStr,EndStr,Url,AUrl,AStr,AImg,ShowTitle,DateTime,ImgFile,DefaultImage,RndShow,IdList
	Dim DescriptionLabel					'标签
	Dim WebDescription						'网站描述
	Dim Price								'价格
	Dim Content								'内容
	Dim noFollow							'不追踪 20141222
	Default = GetDefaultValue(Action)		'获得默认内容
	Dim AutoAddDid,AutoAddSid,AutoAddTid,AutoAdd
	Dim ArticleDescription		'文章描述(简要简介)
	Dim ModI								'余循环20150112
	Dim HtmlCode							'HTML代码
	Dim SplStr,Splxx
	dim tempCode							'暂存Code内容 20150712
	dim replaceStrName						'替换字符名称20150712	
	dim defaultaddcontent					'默认添加Content内容20150713
	dim addContent							'添加内容
	dim dispalyPage							'显示翻页信息 
	dim nPageSize							'每页记录条数
	defaultaddcontent = ModuleFindContent(Action,"defaultaddcontent")
	
	'call echo("defaultaddcontent",defaultaddcontent)
	
	replaceStrName = RParam(Action,"replacestrname")
	
	dispalyPage =lcase(RParam(Action,"dispalyPage"))  '显示翻页信息
	nPageSize =lcase(RParam(Action,"nPageSize"))  '获得每页显示数
	
	
	
	StartStr="[htmlcode]" : EndStr="[/htmlcode]"
	If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
		HtmlCode = StrCut(Default,StartStr,EndStr,2)
	End If
	Dim Module1,Module2,Module3,Module4,nMod1,nMod2
	'判断是否追加大类，小类，子类，这样不好，正确是正确，但是太不优化了。20141223
	'判断大类
	If CheckBigClass(Did) = 0 Then
		AutoAddDid = RParam(Action,"AutoAddDid")
		If LCase(AutoAddDid)="true" Then
			Call AutoAddBigClass(Did)
		End If		
	End If
	'判断小类
	If CheckSmallClass(Did,Sid) = 0 Then
		AutoAddSid = RParam(Action,"AutoAddSid")
		If LCase(AutoAddSid)="true" Then
			Call AutoAddSmallClass(Did,Sid)
		End If		
	End If
	'判断子类
	If CheckThreeClass(Did,Sid,Tid) = 0 Then
		AutoAddTid = RParam(Action,"AutoAddTid")
		If LCase(AutoAddTid)="true" Then
			Call AutoAddThreeClass(Did,Sid,Tid)
		End If		
	End If
	'这里不运行不到，因为被上一级给运行掉了
	If CheckProduct(Did,Sid,Tid)=0 Then				'检测是否有记录
		AutoAdd = RParam(Action,"AutoAdd")
		'call echo(AutoAdd,CheckProduct(Did,Sid,Tid))
		If LCase(AutoAdd)="true" Then
		 	For I = 1 To TopNumb
				if defaultaddcontent<>"" then
					s = defaultaddcontent					
					S = ReplaceValueParam(S,"ni",I)								'不对为i，因为i会与imgurl冲突 [$i$]
					S = ReplaceValueParam(S,"编号-1",I-1)								'不对为i，因为i会与imgurl冲突 [$i$]
					S = ReplaceValueParam(S,"编号",I)								'不对为i，因为i会与imgurl冲突 [$i$]
					S = ReplaceValueParam(S,"did",  Did)
					S = ReplaceValueParam(S,"sid",  Sid)	
					S = ReplaceValueParam(S,"tid",  Tid)	
				else
					s = Did & "-内容"
				end if
				
				
				Call autoAddArticleAction(action,Did, Sid, Tid, Did & "-标题" & I,"", s) 
				
			Next
		End If
	End If
	
	DefaultImage = RParam(Action,"DefaultImage")	'默认图片
	RndShow = RParam(Action,"RndShow")	'随机显示	
	noFollow = Trim(LCase(RParam(Action,"noFollow")))	'不追踪	
	
	If Sql="" Then
		Sql = "Select * From [Product]" & GetDidSidTidSQl(Did,Sid,Tid)
		Sql =  GetWhereAnd(Sql,AddSql)	
		Rs.Open Sql,Conn,1,1
		'随机获得信息
		If RndShow="true" Then
			IdList = GetProductRndIdList(Rs,TopNumb)
			Sql = "Select * From [Product] Where Id In("& IdList &")"
			'Call Echo(Sql,IdList)
			Rs.Close
			Rs.Open Sql,Conn,1,1
		End If
	Else
		Sql = "Select * From [Product]" & Sql
		'Call Eerr("Sql=",Sql)
		Rs.Open Sql,Conn,1,1
	End If
	
	'翻页设置
	dim nCount,Page,x,MaxPage
	if nPageSize="" then
		nPageSize=999
	end if
	 
	nCount = rs.RecordCount 
	page = Request("page")  
	x = getRsPageNumber(rs, nCount, nPageSize, page)                                '获得Rs页数                                                  '记录总数
 
	For I = 1 To TopNumb
		If Rs.Eof Then Exit For
		StartStr="[list-"& I &"]" : EndStr="[/list-"& I &"]"
		
		'例[list-mod2]  [/list-mod2]    20150112
		For ModI=6 To 2 Step -1
			If InStr(Default,StartStr)=False And I Mod ModI =0 Then
				StartStr="[list-mod"& ModI &"]" : EndStr="[/list-mod"& ModI &"]"
				If InStr(Default,StartStr)>0 Then
					'Call Echo("ModI=" & ModI & ",I=" & I,StartStr)
					Exit For
				End If
			End If
		Next
		
		'没有则用默认
		If InStr(Default,StartStr)=False Then
			StartStr="[list]" : EndStr="[/list]"
		End If
		If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
			S = StrCut(Default,StartStr,EndStr,2)
			
			Url = Rs("FileName")
			If CheckMakeHtmlFile(Url)=True Then
				Url = GetHandleFileName(Url)
				Url = DetailPageAddHtmlPrefix(Rs("FolderName")) & Url
			End If
			'自定义产品网址
			if Rs("CustomAUrl")<>"" then
				url=rs("CustomAUrl") 
			end if
			
			Call WebDebug(Url, "act=CreateArticle&ID=" & Rs("Id"))                      '调试
			
			AUrl = AHref(Url, Rs("Title"), Rs("Target"))
			
			AStr = "<a " & AHref(Url, Rs("Title"), Rs("Target")) & SetHtmlParam(Action, "aclass") & IIF(noFollow="true"," rel='nofollow'","") & ">" & ShowTitle & "</a>" & vbCrLf
			
			
			WebDescription = Rs("WebDescription")
			
			ShowTitle = CutStr(Rs("Title"), CutStrNumb, "... ")
			ShowTitle = FontColorFontB(ShowTitle, Rs("FontB"), Rs("FontColor"))			
			'DateTime = Format_Time(Rs("UpDateTime"), 14)
			ImgFile = IIF(IsNul(Rs("SmallFiles")),Rs("BigFiles"),Rs("SmallFiles"))			'也行
			If ImgFile="" Then ImgFile=DefaultImage				'图片为真则用默认图片
			
			AImg = ImgSrc(ImgFile, Rs("Title"), Rs("Target"))
			
			ArticleDescription = Rs("ArticleDescription")								     '简要说明		
			If ArticleDescription="#" Or ArticleDescription="null" Or ArticleDescription="empty" Then
				ArticleDescription = ""
			ElseIf IsNul(ArticleDescription) Then
				ArticleDescription = Rs("Content")
			End If	
			
			For J=1 To 3
				S = ReplaceValueParam(S,"ni",I)								'不对为i，因为i会与imgurl冲突 [$i$]
				S = ReplaceValueParam(S,"编号-1",I-1)								'不对为i，因为i会与imgurl冲突 [$i$]
				S = ReplaceValueParam(S,"编号",I)								'不对为i，因为i会与imgurl冲突 [$i$]
				
				'S = ReplaceValueParam(S,"编号-1%2",Fix(I/2)-1)					'不对为i，因为i会与imgurl冲突 [$i$]
				
				
				
				
				S = ReplaceValueParam(S,"id",  Rs("Id"))
				S = ReplaceValueParam(S,"title",  Rs("Title"))
				S = ReplaceValueParam(S,"showtitle",ShowTitle)
				S = ReplaceValueParam(S,"adddatetime",Rs("AddDateTime"))
				S = ReplaceValueParam(S,"imgfile",ImgFile)
				S = ReplaceValueParam(S,"imgurl",ImgFile)
				S = ReplaceValueParam(S,"url",url)
				S = ReplaceValueParam(S,"aurl",AUrl)
				S = ReplaceValueParam(S,"aimg",AImg)				'20150112			
				S = ReplaceValueParam(S,"astr",AStr)
				S = ReplaceValueParam(S,"description",Rs("WebDescription"))
				S = ReplaceValueParam(S,"content",Rs("Content"))
				S = ReplaceValueParam(S,"price",Rs("Price"))			'价格
				S = ReplaceValueParam(S,"memberprice",Rs("memberprice"))			'价格				
				S = ReplaceValueParam(S,"newprice",Rs("NewPrice"))		'新价格
				S = ReplaceValueParam(S,"sold",Rs("sold"))				'已售出
				S = ReplaceValueParam(S,"author",Rs("Author"))
				S = ReplaceValueParam(S,"productabout",Rs("WebDescription"))
				S = ReplaceValueParam(S,"articledescription",ArticleDescription)     '摘要(文章简要描述)20141225
			Next
			
			Url = ""& adminDir &"Product.Asp?act=ShowEditProduct&Id=" & Rs("Id") & "&n=" & GetRnd(11)
			S = DisplayOnlineEditDialog(Url, S)
			
			
			
			'开始位置加Dialog内容
			StartStr="[list-"& I &" startdialog]" : EndStr="[/list-"& I &" startdialog]"
			If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
				S = StrCut(Default,StartStr,EndStr,2) & S
			End If
			'结束位置加Dialog内容
			StartStr="[list-"& I &" enddialog]" : EndStr="[/list-"& I &" enddialog]"
			If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
				S = S & StrCut(Default,StartStr,EndStr,2)
			End If
			
			if replaceStrName<>"" then
				call addModuleReplaceArray(replaceStrName & i,s)
			else
				C = C & S
			end if
		End If	
	Rs.MoveNext:Next:Rs.Close
	
	'开始内容加Dialog内容
	StartStr="[dialog start]" : EndStr="[/dialog start]"
	If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
		C = StrCut(Default,StartStr,EndStr,2) & C
	End If
	'结束内容加Dialog内容
	StartStr="[dialog end]" : EndStr="[/dialog end]"
	If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
		C = C &  StrCut(Default,StartStr,EndStr,2)
	End If
	'特殊功能开发 20150112等更高开发应用
	If HtmlCode<>"" Then
		nMod1=0:nMod2=0
		SplStr=Split(C, "[$vbCrlf$]") : C=""
		For I=0 To Ubound(SplStr)
			S = SplStr(I)
			
			If I=0 Then
				S = Replace(S,"[noneOrblack]","black")
			Else
				S = Replace(S,"[noneOrblack]","none")
			End If
			
			If S <> "" And InStr(S, "[$Array$]") Then 		
				If (I+1) Mod 2=0 Then
					S = Replace(S,"[@i@]",nMod1)
					nMod1=nMod1+1				
				Else
					S = Replace(S,"[@i@]",nMod2)
					nMod2=nMod2+1				
				End If		
				Splxx=Split(S, "[$Array$]")
				
				If (I+1) Mod 2=0 Then
					Module3 = Module3 & Splxx(0)
					Module4 = Module4 & Splxx(1)	
				Else
					Module1 = Module1 & Splxx(0)
					Module2 = Module2 & Splxx(1)
				End If
			End If
		Next
		HtmlCode = Replace(HtmlCode,"[$Module_1$]", Module1)
		HtmlCode = Replace(HtmlCode,"[$Module_2$]", Module2)
		HtmlCode = Replace(HtmlCode,"[$Module_3$]", Module3)
		HtmlCode = Replace(HtmlCode,"[$Module_4$]", Module4)
		C = HtmlCode
	End If
	
	
		
	dim homePageStr,parentPageStr,nextPageStr,endPageStr,pageStr	
	If Page <= 1 Then
		
		'首页
		StartStr="[defaultHomePageStr]" : EndStr="[/defaultHomePageStr]"
		If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
			homePageStr = StrCut(Default,StartStr,EndStr,2)
		else
			homePageStr = "<span class='PageBox'>首页</span> &nbsp;" 
		end if
		
		'上一页
		StartStr="[defaultParentPageStr]" : EndStr="[/defaultParentPageStr]"
		If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
			parentPageStr = StrCut(Default,StartStr,EndStr,2)
		else
	        parentPageStr = "<span class='PageBox'>上一页</span> &nbsp;"
		end if
    Else		 
		'首页
		url=mid(CreateHtmlPath,1,instrrev(CreateHtmlPath,".")-1) & "" & mid(CreateHtmlPath,instr(CreateHtmlPath,"."))
		url=replace(url,"///","/")
		Call WebDebug(Url, GetUrlAddToParam(GetUrl(),"&page=1","replace"))                      '调试
		 
		StartStr="[homePageStr]" : EndStr="[/homePageStr]"
		If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
			homePageStr = StrCut(Default,StartStr,EndStr,2)
		else
			homePageStr = "<a href=""[$url$]"">首页</a> &nbsp;" 
		end if
		homePageStr=replace(homePageStr,"[$url$]",url)
		
		'上一页
		url=mid(CreateHtmlPath,1,instrrev(CreateHtmlPath,".")-1) & "_" & (page-1) & mid(CreateHtmlPath,instr(CreateHtmlPath,"."))
		url=replace(url,"///","/")
		Call WebDebug(Url,  GetUrlAddToParam(GetUrl(),"&page=" & page-1,"replace") )                      '调试
		StartStr="[parentPageStr]" : EndStr="[/parentPageStr]"
		If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
			parentPageStr = StrCut(Default,StartStr,EndStr,2)
		else
			parentPageStr = "<a href=""[$url$]"">上一页</a> &nbsp;" 
		end if
		parentPageStr=replace(parentPageStr,"[$url$]",url)
	
		
		
	end if
	
     '下一页 与 尾页
    If Page >= MaxPage Then 
	
		'下一页
		StartStr="[defaultNextPageStr]" : EndStr="[/defaultNextPageStr]"
		If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
			nextPageStr = StrCut(Default,StartStr,EndStr,2)
		else
			nextPageStr = "<span class='PageBox'>下一页</span> &nbsp;" 
		end if
		
		'尾页
		StartStr="[defaultEndPageStr]" : EndStr="[/defaultEndPageStr]"
		If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
			endPageStr = StrCut(Default,StartStr,EndStr,2)
		else
	        endPageStr = "<span class='PageBox'>尾页</span>" 
		end if 
        
    Else
		
		if CreateHtmlPath<>"" then
			'下一页
			url=mid(CreateHtmlPath,1,instrrev(CreateHtmlPath,".")-1) & "_" & (page+1) & mid(CreateHtmlPath,instr(CreateHtmlPath,"."))
		end if
		url=replace(url,"///","/")
		Call WebDebug(Url,  GetUrlAddToParam(GetUrl(),"&page=" & page+1,"replace")  )                      '调试
		 
		StartStr="[nextPageStr]" : EndStr="[/nextPageStr]"
		If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
			nextPageStr = StrCut(Default,StartStr,EndStr,2)
		else
			nextPageStr = "<a href=""[$url$]"">下一页</a> &nbsp;" 
		end if
		nextPageStr=replace(nextPageStr,"[$url$]",url)
	
		if CreateHtmlPath<>"" then
			'尾页
			url=mid(CreateHtmlPath,1,instrrev(CreateHtmlPath,".")-1) & "_" & MaxPage & mid(CreateHtmlPath,instr(CreateHtmlPath,"."))
		end if
		url=replace(url,"///","/")
		Call WebDebug(Url,  GetUrlAddToParam(GetUrl(),"&page=" & MaxPage,"replace")  )                      '调试
		
		StartStr="[endPageStr]" : EndStr="[/endPageStr]"
		If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
			endPageStr = StrCut(Default,StartStr,EndStr,2)
		else
			endPageStr = "<a href=""[$url$]"">尾页</a> &nbsp;" 
		end if
		endPageStr=replace(endPageStr,"[$url$]",url)
	end if
	
	pageStr = homePageStr & parentPageStr & nextPageStr & endPageStr	
	if dispalyPage="true" or dispalyPage="1" then	
		'页信息
		StartStr="[page]" : EndStr="[/page]"
		If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
			S = StrCut(Default,StartStr,EndStr,2)
			s=replace(s,"[$homePageStr$]",homePageStr)				'首页
			s=replace(s,"[$parentPageStr$]",parentPageStr)			'上一页
			s=replace(s,"[$nextPageStr$]",nextPageStr)				'下一页
			s=replace(s,"[$endPageStr$]",endPageStr)				'尾页
			s=replace(s,"[$Page$]",Page)							'当前页
			s=replace(s,"[$MaxPage$]",MaxPage)						'总页
			s=replace(s,"[$nCount$]",nCount)						'总有记录			
			c=c & s
		else		
			C = C & "<div class='t_page ColorLink'>总数：" & nCount & "条&nbsp;&nbsp;当前页数：<span class='FontRed'>" & Page & "</span>/" & MaxPage & "&nbsp;" & pageStr & "</div>" 
		end if
	end if
	
	HandleDispalyArticleList = c
End Function


'自定义留言列表###########################################未调用此功能函数
Function XY_CustomGuestBookList(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)		'获得这么函数参数
	Call XY_AutoAddHandle(Action)			'处理添加信息处理
	Dim Sql,Default,I,S,C,StartStr,EndStr,Url,ShowTitle,DateTime,ImgFile
	Default = GetDefaultValue(Action)
	'Call Echo("Action",Action)
	'Call Echo("AddSql",AddSql)
	Sql = "Select * From [GuestBook]" & GetDidSidTidSQl(Did,Sid,Tid)
	Sql =  GetWhereAnd(Sql,AddSql)		
	Rs.Open Sql,Conn,1,1
	For I = 1 To TopNumb
		If Rs.Eof Then Exit For
		StartStr="[list-"& I &"]" : EndStr="[/list-"& I &"]"
		'没有则用默认
		If InStr(Default,StartStr)=False Then
			StartStr="[list]" : EndStr="[/list]"
		End If
		If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
			S = StrCut(Default,StartStr,EndStr,2)
						
			S = ReplaceValueParam(S,"guestname",Rs("GuestName"))			'留言者
			S = ReplaceValueParam(S,"email",Rs("EMail"))					'留言者邮箱
			S = ReplaceValueParam(S,"tel",Rs("Tel"))						'留言者电话
			S = ReplaceValueParam(S,"qq",Rs("QQ"))							'留言者QQ
			S = ReplaceValueParam(S,"ip",Rs("IP"))							'留言者IP
			S = ReplaceValueParam(S,"content",Rs("Content"))				'留言者内容
			S = ReplaceValueParam(S,"reply",Rs("Reply"))					'回答内容
			
			
			'开始位置加Dialog内容
			StartStr="[list-"& I &" startdialog]" : EndStr="[/list-"& I &" startdialog]"
			If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
				S = StrCut(Default,StartStr,EndStr,2) & S
			End If
			'结束位置加Dialog内容
			StartStr="[list-"& I &" enddialog]" : EndStr="[/list-"& I &" enddialog]"
			If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
				S = S & StrCut(Default,StartStr,EndStr,2)
			End If			
			C = C & S
		End If	
	Rs.MoveNext:Next:Rs.Close
	
	'开始内容加Dialog内容
	StartStr="[dialog start]" : EndStr="[/dialog start]"
	If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
		C = StrCut(Default,StartStr,EndStr,2) & C
	End If
	'结束内容加Dialog内容
	StartStr="[dialog end]" : EndStr="[/dialog end]"
	If InStr(Default,StartStr)>0 And InStr(Default,EndStr)>0 Then
		C = C &  StrCut(Default,StartStr,EndStr,2)
	End If
	
	XY_CustomGuestBookList = C
End Function
'显示新闻列表
Function XY_ShowNewsList(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)		'获得这么函数参数
	Dim AutoAddDid,AutoAdd,ImageDivStr,ID,NewSql
	Call XY_AutoAddHandle(Action)			'处理添加信息处理
	
	Dim I,C,S,Sql,Url,ShowTitle,AStr,DateYes,NewsStyle
	DateYes = RParam(Action,"DateYes")						'是否显示时间
	NewsStyle = RParam(Action,"NewsStyle")					'新闻样式(顶部与底部是否显示一条新闻)
	C="<ul" & SetHtmlParam(Action,"ulclass") &">"
	Sql = "Select * From [Product]" & GetDidSidTidSQl(Did,Sid,Tid) & GetWhereAnd(Sql,AddSql)
	If NewsStyle<>"" And NewsStyle<>"default" Then
		NewSql = "Select * From [Product]" & GetDidSidTidSQl(Did,Sid,Tid)
		NewSql =  GetWhereAnd(NewSql," Where (BigFiles<>'' Or SmallFiles<>'')")
		NewSql =  GetWhereAnd(NewSql,AddSql)
		'Call Echo("NewSql",NewSql)
		Rs.Open NewSql,Conn,1,1
		If Not Rs.Eof Then
		
			ImageDivStr="<img src='"& IIF(Rs("BigFiles")<>"",Rs("BigFiles"),Rs("SmallFiles")) &"'" & SetHtmlParam(Action,"imgclass|imgwidth|imgheight") &">"
			
			'给图片加链接
			Url = Rs("FileName")
			If CheckMakeHtmlFile(Url)=True Then
				Url = GetHandleUrl(Rs("FileName"))
				Call WebDebug(Url, "act=CreateArticle&ID=" & Rs("Id"))                      '调试
			End If
			
			ImageDivStr = "<a " & AHref(Url, Rs("Title"), Rs("Target")) & SetHtmlParam(Action,"aclass") & ">" & ImageDivStr & "</a>"			
			'显示单个新闻文本
			If InStr("|footphotonews|topphotonews|","|"& NewsStyle &"|") > 0 Then
				ImageDivStr =ImageDivStr & "<span" & SetHtmlParam(Action,"textclass") &">" & CutStr(Rs("Content"), 100, "... ") & "</span>"
			End If
			ID = Rs("ID")
		End If:Rs.Close
		Sql= "Select * From [Product]" & GetDidSidTidSQl(Did,Sid,Tid)
		If Id <> "" Then Sql =  GetWhereAnd(Sql," Where Id<>"& Id &"")
		Sql =  GetWhereAnd(Sql,AddSql)
	End If
	Rs.Open Sql,Conn,1,1
	For I=1 To TopNumb
		If Rs.Eof Then Exit For
		
		Url = Rs("FileName")
		If CheckMakeHtmlFile(Url)=True Then
        	Url = GetHandleUrl(Rs("FileName"))
	        Call WebDebug(Url, "act=CreateArticle&ID=" & Rs("Id"))                      '调试
		End If
		
        ShowTitle = CutStr(Rs("Title"), CutStrNumb, "... ")
        ShowTitle = FontColorFontB(ShowTitle, Rs("FontB"), Rs("FontColor"))

        AStr = "<a " & AHref(Url, Rs("Title"), Rs("Target")) & SetHtmlParam(Action,"aclass") & ">" & ShowTitle & "</a>"
        If DateYes <> "" Then
            AStr = AStr & "<span class=""timeright"">" & Format_Time(Rs("UpDateTime"), DateYes) & "</span>" & vbCrLf
        End If
		C=C & "<li" & SetHtmlParam(Action,"liclass") &">"& AStr &"</li>" & vbCrlf
	Rs.MoveNext:Next:Rs.Close
	C = C & "</ul>" & vbCrlf
	'显示一条新闻带图片的
	If InStr("|topphoto|topphotonews|","|"& NewsStyle &"|") > 0 Then
		C=ImageDivStr & C
	ElseIf InStr("|footphoto|footphotonews|","|"& NewsStyle &"|") > 0 Then
		C= C & ImageDivStr
	End If
	
	XY_ShowNewsList = C	
End Function
'切换Banner旗帜
Function XY_BannerSwitch(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	XY_BannerSwitch = BannerSwitch(Did,TopNumb)
End Function
'切换Banner旗帜2
Function XY_BannerSwitch2(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	XY_BannerSwitch2 = BannerSwitch2(Did,TopNumb)
End Function
'友情连接列表处理
Function XY_LinkList(Action)
	Dim Title,SplTypeStr
	SplTypeStr = " | "								'默认分割类型
	Title = RParam(Action,"Title")
	SplTypeStr = RParam(Action,"spltypestr")	
	XY_LinkList = HandleLinkList(Title,SplTypeStr)	
End Function
'友情链接
Function XY_Links(Action)
	Dim AddSql,SplTypeStr,S
	SplTypeStr = " | "								'默认分割类型
	AddSql = RParam(Action,"AddSql")
	SplTypeStr = RParam(Action,"spltypestr")		
	Dim Url,Title,Target,Rel,ImgPath,I,C,Sql
	Sql = "Select * From [FriendLink]"
	Sql = GetWhereAnd(Sql,AddSql)
	Rsx.Open Sql,Conn,1,1
	For I=1 To Rsx.RecordCount
		Url =  Rsx("Url")
		Target = ATarget(Rsx("Target"))
		Rel = ARel(Rsx("NoFollow"))
		Title = Rsx("Title")
		ImgPath = Rsx("ImgPath")
		S = "<a href='"& Url &"'"& Target &""&Rel & SetHtmlParam(Action,"class")&" title="""& Title &""">"& FontColorFontB(Title,Rsx("FontB"),Rsx("FontColor")) &"</a>" & SplTypeStr
		
		'S = HandleDisplayOnlineEditDialog(""& adminDir &"FriendLink.Asp?act=ShowEdit&Id=" & Rsx("Id") & "&n=" & GetRnd(11), S,"style='float:right;padding:0 4px;'","")	
		S = DisplayOnlineED2(WEB_ADMINURL &"FriendLink.Asp?act=ShowEdit&Id=" & Rsx("Id") & "&n=" & GetRnd(11), S, "<li|<a ")
		C=C & S & vbCrlf		
	Rsx.MoveNext:Next:Rsx.Close
	XY_Links = C
End Function
'友情链接LI方式显示			
Function XY_FriendLinkLI(Action)
	Dim AddSql
	AddSql = RParam(Action,"AddSql")
	Dim Url,Title,Target,Rel,ImgPath,I,C,Sql,S
	Sql = "Select * From [FriendLink]"
	Sql = GetWhereAnd(Sql,AddSql)
	Rsx.Open Sql,Conn,1,1
	For I=1 To Rsx.RecordCount
		Url =  Rsx("Url")
		Target = ATarget(Rsx("Target"))
		Rel = ARel(Rsx("NoFollow"))
		Title = Rsx("Title")
		ImgPath = Rsx("ImgPath")
		S = "<a href='"& Url &"'"& Target &""&Rel & SetHtmlParam(Action,"class")&" >" & FontColorFontB(Title,Rsx("FontB"),Rsx("FontColor")) &"</a>"
		S = HandleDisplayOnlineEditDialog(""& adminDir &"FriendLink.Asp?act=ShowEdit&Id=" & Rs("Id") & "&n=" & GetRnd(11), S,"style='float:right;padding:0 4px;'","")
		
		C=C & "<li"& SetHtmlParam(Action,"liclass") &">" & S & "</li>"
	Rsx.MoveNext:Next:Rsx.Close
	XY_FriendLinkLI = C
End Function



'普通链接A
Function XY_HrefA(Action)
	Dim Content,Href,C,AContent,AType,Url,Title
	Action = HandleInModule(Action,"start")
	Content = RParam(Action,"Content")
	AType = RParam(Action,"Type")
	If AType="收藏" Then
		'第一种方法
		'Url = "window.external.addFavorite('"& WebUrl &"','"& WebTitle &"')"
		Url = "shoucang(document.title,window.location)"
		C = "<a href='javascript:;' onClick="""& Url &""" " & SetHtmlParam(Action,"target|title|alt|id|class|style") &">"& Content &"</a>"
	ElseIf AType="设为首页" Then
		'第一种方法
		'Url = "var strHref=window.location.href;this.style.behavior='url(#default#homepage)';this.setHomePage('"& WebUrl &"');"
		Url = "SetHome(this,window.location)"
		C = "<a href='javascript:;' onClick="""& Url &"""" & SetHtmlParam(Action,"target|title|alt|id|class|style") &">"& Content &"</a>"
	Else
		Content = RParam(Action,"Title")
	End If
	
	Content = HandleInModule(Content,"end")
	If C = "" Then C = "<a" & SetHtmlParam(Action,"href|target|title|alt|id|class|rel|style") &">"& Content &"</a>"
	
	XY_HrefA = C
End Function
'显示当前模板里内容
Function XY_ShowThisText(Action)
	Dim Default
	Default = GetDefaultValue(Action)
	XY_ShowThisText = Default
End Function



'获得文本内容 这个是干什么用得？？？20150121
Function XY_MainInfo2(Action)
	Dim Title,ShowTitle,AutoAdd,Content
	
	Title = RParam(Action,"Title")
	ShowTitle = RParam(Action,"ShowTitle")
	AutoAdd = RParam(Action,"AutoAdd")	
	
	Content = ReadMainInfo(Title)
	If Content="" Then
		Content = GetDefaultValue(Action)
		'是否自动添加信息
		If LCase(AutoAdd)="true" Then
			Call AutoAddMainInfo(Title,ShowTitle,Content)
		End If
	End If
	XY_MainInfo2 = Content
End Function
'获得文本内容
Function XY_MainInfo(Action)
	Dim Title,ShowTitle,AutoAdd,Content, Id, Url, OnHtml
	
	Title = RParam(Action,"Title")
	ShowTitle = RParam(Action,"ShowTitle")
	AutoAdd = RParam(Action,"AutoAdd")	
	OnHtml = RParam(Action,"OnHtml")
	If OnHtml="" Then OnHtml=False	
	
	Content = ReadMainInfo(Title)
	If Content="" Then
		Content = GetDefaultValue(Action)
		'是否自动添加信息
		If LCase(AutoAdd)="true" Then
			Call HandleAutoAddMainInfo(Title,ShowTitle,OnHtml,Content)
		End If
	End If
	Id = ReadMainId(Title)
	If Id<>"" And Request("gl") = "edit" Then
		'Content = DisplayOnlineEditDialog(""& adminDir &"MainInfo.Asp?act=ShowEdit&Id=" & Id & "&n=" & GetRnd(11), Content)
		Url = ""& adminDir &"MainInfo.Asp?act=ShowEdit&Id=" & Id & "&n=" & GetRnd(11)
		Content = DisplayOnlineED2(Url, Content, "<span |<div ")
	End If
	'处理内容参数标签
	Content = HandleContentParamLabel(Action,Content)
	
	XY_MainInfo = Content
End Function
'处理内容参数标签
Function HandleContentParamLabel(Action,ByVal Content)
	Dim DelHtmlYes,TrimYes,nLen
	'删除Html
	DelHtmlYes = RParam(Action,"DelHtml")			'是否删除Html
	If DelHtmlYes="true" Then Content = Replace(DelHtml(Content),"<","&lt;")			'HTML处理
	'删除两边空格
	TrimYes = RParam(Action,"trim")			'是否删除两边空格
	If TrimYes="true" Then Content = TrimVbCrlf(Content)

	'截取字符处理
	nLen = RParam(Action,"Len")			'字符长度值
	nLen = HandleNumber(nLen)
	If nLen<>"" Then Content = CutStr(Content,nLen,"...")' Left(Content,nLen)
	HandleContentParamLabel = Content
End Function

'文本内容链接A
Function XY_MainInfoA(Action)
	Dim Title,ShowTitle,AutoAdd,Content,Url
	
	Title = RParam(Action,"Title")
	ShowTitle = RParam(Action,"ShowTitle")
	AutoAdd = RParam(Action,"AutoAdd")	
	Url = GetMainInfoUrl(Title)		'获得文本内容地址
	
	If Url="" Then
		Content = GetDefaultValue(Action)
		'是否自动添加信息
		If LCase(AutoAdd)="true" Then
			Call AutoAddMainInfo(Title,ShowTitle,Content)
		End If
	End If
	XY_MainInfoA = "<a href='"& Url &"'"& SetHtmlParam(Action,"target|title|alt|id|class|style|") &">"& Title &"</a>"
End Function
'文本内容链接Url
Function XY_MainInfoUrl(Action)
	Dim Title,ShowTitle,AutoAdd,Content,Url
	
	Title = RParam(Action,"Title")
	ShowTitle = RParam(Action,"ShowTitle")
	AutoAdd = RParam(Action,"AutoAdd")	
	Url = GetMainInfoUrl(Title)		'获得文本内容地址
	
	If Url="" Then
		Content = GetDefaultValue(Action)
		'是否自动添加信息
		If LCase(AutoAdd)="true" Then
			Call AutoAddMainInfo(Title,ShowTitle,Content)
		End If
	End If
	XY_MainInfoUrl = Url
End Function
'文本显示标题
Function XY_MainInfo_GetShowTitle(Action)
	Dim Title,ShowTitle,AutoAdd,Default,Url,ReadShowTitle,Id
	
	Title = RParam(Action,"Title")
	ShowTitle = RParam(Action,"ShowTitle")
	AutoAdd = RParam(Action,"AutoAdd")
	
	ReadShowTitle = ReadMainInfoShowTitle(Title)		'获得文本内容地址
	
	If ReadShowTitle="" Then
		Default = GetDefaultValue(Action)
		'是否自动添加信息
		If LCase(AutoAdd)="true" Then
			Call AutoAddMainInfo(Title,ShowTitle,default)
		End If
		ReadShowTitle = ShowTitle
	End If
	Id = ReadMainId(Title)
	If Id<>"" And Request("gl") = "edit" Then
		'Content = DisplayOnlineEditDialog(""& adminDir &"MainInfo.Asp?act=ShowEdit&Id=" & Id & "&n=" & GetRnd(11), Content)
		Url = ""& adminDir &"MainInfo.Asp?act=ShowEdit&Id=" & Id & "&n=" & GetRnd(11)
		ReadShowTitle = DisplayOnlineED2(Url, ReadShowTitle, "<span |<div ")
	End If
	
	
	XY_MainInfo_GetShowTitle = ReadShowTitle
	
End Function



'导航大类链接 NavBigName=''为导航大类名称
Function XY_NavDidA(Action)	
	Dim Url,NavDid	
	NavDid = RParam(Action, "NavDid")
	Url = GetNavUrl(NavDid, "")		'获得导航地址
	Url = FullHttpUrl(Host(),Url)
	XY_NavDidA = "<a href='"& Url &"'"& SetHtmlParam(Action,"target|title|alt|id|class|style|") &">"& NavDid &"</a>"
End Function
'导航大类Url
Function XY_NavDidUrl(Action)	
	Dim Url,NavDid	
	NavDid = RParam(Action, "NavDid")
	Url = GetNavUrl(NavDid, "")		'获得导航地址 
	XY_NavDidUrl = Url
End Function
'处理网站标题  20141222
Function XY_WebTitle()
	Dim S
	XY_WebTitle = WebTitle
	S = Trim(LCase(WebTitle))
	If S="#" Or S="null" Or S="empty" Then
		XY_WebTitle = ""
	End If	
End Function
'处理网站关键词 20141222
Function XY_WebKeywords()
	Dim S
	XY_WebKeywords = WebKeywords
	S = Trim(LCase(WebKeywords))
	If S="#" Or S="null" Or S="empty" Then
		XY_WebKeywords = ""
	End If	
End Function
'处理网站描述 2014122
Function XY_WebDescription()
	Dim S
	XY_WebDescription = WebDescription
	S = Trim(LCase(WebDescription))
	If S="#" Or S="null" Or S="empty" Then
		XY_WebDescription = ""
	End If	
End Function

'获得网站底部内容aa
Function XY_WebBottom(Action)
	If WebBottom = "[底部无内容]" Or WebBottom="[$网站底部$]"  Then
		'获得网站配置
		Rs.Open "Select * From [WebSite]", Conn, 1, 3
		If Not Rs.Eof Then
			WebBottom=GetDefaultValue(Action)
			Rs("WebSiteBottom") = WebBottom
			Rs.Update
		End If : Rs.Close
	End If
	XY_WebBottom = WebBottom
End Function

'网站底部导航
Function XY_WebBottomNav(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	
	Dim I,S, C, Url, FileName, BigFolder, ShowDid, Target,Sql,SplTypeStr
	SplTypeStr = RParam(Action,"SplTypeStr")
	If AddSql="" Then AddSql="Where NavButtom<>0  Order By Sort Asc"
	
	Sql  = "Select * From [NavBigClass] "
	If UserId<>"" Then Sql = GetWhereAnd(Sql, " And UserId=" & UserId) 			'指定用户
	Sql = GetWhereAnd(Sql, AddSql)
	Rs.Open Sql, Conn, 1, 1
	For I = 1 To Rs.RecordCount
		FileName = Rs("FileName")
		If FileName <> "" And Right(FileName,1)<>"/" Then FileName = FileName & ".html"
	
		BigFolder = "/" & GetWebFolderName() & Rs("FolderName") & "/"
		Did = Rs("BigClassName")
		ShowDid = FontColorFontB(Did, Rs("FontB"), Rs("FontColor"))
		Target = Rs("Target")
		
		Url = Rs("FileName")
		If CheckMakeHtmlFile(Url)=True Then
			Url = BigFolder & FileName
			Url = Replace(Replace(Url, "//", "/"), "//", "/")
			Url = FullHttpUrl(Host,Url)	
		End If
		Call WebDebug(Url,"act=Nav&NavDid=" & Did)	'调试
		If InStr("|产品|新闻|下载|","|" & Rs("NavType") & "|")>0 Then
			Call WebDebug(Url,"act=Nav&NavDid=" & Did & "&Page=1")	'调试
		End If		
		
		S = "<a " & AHref(Url, ShowDid, Target) & SetHtmlParam(Action,"aclass|class|style") & ">" & ShowDid & "</a>"
		If I <> Rs.RecordCount And SplTypeStr<>"" Then S = S & SplTypeStr
		'S = HandleDisplayOnlineEditDialog(""& adminDir &"NavManage.Asp?act=EditNavBig&Id=" & Rs("Id") & "&n=" & GetRnd(11), S,"style='float:right;padding:0 4px;'","")
		S = DisplayOnlineED2(WEB_ADMINURL &"NavManage.Asp?act=EditNavBig&Id=" & Rs("Id") & "&n=" & GetRnd(11), S, "<li|<a ")
		C = C & S
	Rs.MoveNext : Next : Rs.Close
	XY_WebBottomNav = C
End Function

'**********************************************************************
'					           辅助类列表
'**********************************************************************
'批量添加产品类别用VbTab作为级别分类 创建于20141117
Function XY_BatchAutoAddProductClass(Action)
	Dim Default,SplDid,I,S,Splxx,Did,FileName,TemplatePath,DidName,SidName,TidName,Sql,nSort
	Default = GetDefaultValue(Action)
	Dim Operate,RootDirYes
	RootDirYes = LCase(Trim(RParam(Action,"RootDirYes")))			'是否为主目录
	Operate = RParam(Action,"Operate")
	If Operate="false" Then Exit Function				'操作为假则 退出
	SplDid = Split(Default, vbCrlf) 	
	For I = 0 To UBound(SplDid)
		S = Replace(SplDid(I),vbTab, CopyStr(" ",4))
		If S <> "" And Left(Trim(S),1)<>"#" Then
			Splxx = Split(S & "||||||", "|")
			Did = Trim(Splxx(0))
			nSort = Splxx(1)
			FileName = Splxx(2)
			TemplatePath = Splxx(3)
			If nSort = "" Then nSort =  I
			If Left(Splxx(0),12)= CopyStr(" ",12) Then
				Splxx(0)=Replace(Splxx(0),"[类别]", DidName & "_" & SidName & "_" & TidName)
				Splxx(0)=" did='"& DidName &"' sid='"& SidName &"' tid='"& TidName &"'" & Splxx(0)
				Call XY_AutoAddProduct(Splxx(0))			'自动添加产品信息(批量的哟)
				'CAll Echo("Splxx(0)",Splxx(0))
			ElseIf Left(Splxx(0),8)= CopyStr(" ",8) Then
				TidName=Did
				Sql="Select * From [ThreeClass] Where BigClassName='" & DidName & "' And SmallClassName='"& SidName &"' And ThreeClassName='"& TidName &"'"
				If FileName = "" Then FileName = TidName
			ElseIf Left(Splxx(0),4)=CopyStr(" ",4)  Then
				SidName=Did : TidName=""
				Sql="Select * From [SmallClass] Where BigClassName='" & DidName & "' And SmallClassName='"& SidName &"'"					
				If FileName = "" Then FileName = SidName
			ElseIf Did<>"" Then
				DidName=Did : SidName="" : TidName=""
				Sql="Select * From [BigClass] Where BigClassName='" & DidName & "'"				
				If FileName = "" Then FileName = DidName
			End If
			If DidName<>"" Then
				Rs.Open Sql, Conn, 1, 3
				If Rs.Eof Then
					Rs.AddNew
						If DidName <>"" Then Rs("BigClassName") = DidName
						If SidName <>"" Then Rs("SmallClassName") = SidName
						If TidName <>"" Then Rs("ThreeClassName") = TidName
						Rs("Sort") = nSort
						If FileName<>"" And RootDirYes="true" Then FileName = "/" & FileName
						Rs("FolderName") = PinYin(FileName, 3)
					Rs.Update
					'strHtml = strHtml & "添加产品大类 " & Did & "<br>"
				End If:Rs.Close
			End If
		End If
	Next
	'Call Eerr(Default,Action)
End Function

'批量添加导航类别用VbTab作为级别分类 创建于20141117
Function XY_BatchAutoAddNavClass(Action)
	Dim Default,SplDid,I,S,Did,Sid,FileName,FolderName,TemplatePath,Sql,nSort,NavType,NavLocation,NavActionType,WebTitle,clearS
	Default = GetDefaultValue(Action)
	Dim Operate,RootDirYes,Content
	dim isDelTable												    '为是否删除表
	dim isAddFolderName
	dim Rs:Set Rs = CreateObject("Adodb.RecordSet")
	dim isFileNameToPinYin
	
	isAddFolderName = LCase(Trim(RParam(Action,"isaddfoldername")))			'是否添加文件夹名称
	
	Operate = RParam(Action,"Operate")
	RootDirYes = LCase(Trim(RParam(Action,"RootDirYes")))			'是否为主目录    主目录会在前面加 /  符号
	isDelTable = LCase(Trim(RParam(Action,"isDelTable")))			'是否要删除表
	isFileNameToPinYin = LCase(Trim(RParam(Action,"isFileNameToPinYin")))			'是否转拼音
	dim url											'链接网址
	'call echo("isFileNameToPinYin",isFileNameToPinYin)
	
	
	If Operate="false" Then Exit Function				'操作为假则 退出
	
	if isDelTable<>"true" then
		Rs.Open"Select * From [NavBigClass]",Conn,1,1
		If Rs.RecordCount>1 Then Rs.Close : Exit Function			'记录大于1退出
		If Not Rs.Eof Then
			If Rs("BigClassName")<>"测试首页" Then Rs.Close : Exit Function			'当前记录大类名称不为 测试首页 退出
		End If
		Rs.Close
	end if
	Conn.Execute("Delete From [NavBigClass]")				'删除导航大类
	
	SplDid = Split(Default, vbCrlf)
	For I = 0 To UBound(SplDid)
		S = Replace(SplDid(I),vbTab, CopyStr(" ",4))
		clearS=phpTrim(S)
		If clearS <> ""  And Left(Trim(S),1)<>"#" Then
			nSort = RParam(S,"Sort")
			FileName = RParam(S,"FileName")
			FolderName = RParam(S,"FolderName")
			TemplatePath = RParam(S,"Template")
			Content = RParam(S,"Content")
			WebTitle = RParam(S,"WebTitle")								'网站标题
			NavLocation = "|" & RParam(S,"NavLocation") & "|"			'导航位置
			url = phptrim(clearRParam(S,"url"))							'链接网址
			'call echo(s,url)
		 	if FileName="null" then FileName=""
			
			If nSort = "" Then nSort =  I
			NavActionType=""
			If Left(S,4)=CopyStr(" ",4)  Then
				Sid = RParam(S,"Sid") : NavType=""
				'If FileName="" Then FileName = Sid			'获得默认文件名
				Sql="Select * From [NavSmallClass] Where BigClassName='" & Did & "' And SmallClassName='"& Sid &"'"					
				NavActionType="navsmall"
				If FileName="" Then FileName = Sid			'获得默认文件名
			Else
				Did = RParam(S,"Did") : Sid=""
				NavType = RParam(S,"NavType")				'导航类型
				'Call Echo(S,NavType)
				If FileName="" Then FileName = Did			'获得默认文件名
				Sql="Select * From [NavBigClass] Where BigClassName='" & Did & "'"
			End If
			If Did<>"" Then
				'CAll Echo("Sql",Sql)
				Rs.Open Sql, Conn, 1, 3
				If Rs.Eof Then
					Rs.AddNew
						If Did <>"" Then Rs("BigClassName") = Did
						If Sid <>"" Then Rs("SmallClassName") = Sid
						Rs("TemplatePath") = TemplatePath						
						If NavType <>"首页" Then
							if isAddFolderName="true" and FolderName<>"null" And FolderName<>"空" then
								If FolderName =  "" Then 
									if Sid<>"" then
										FolderName=Sid
									else									
										FolderName=Did
									end if	
								end if							
								If FolderName<>"" And RootDirYes="true" Then FolderName = "/" & FolderName
								
								Rs("FolderName") = PinYin(FolderName, 3)
							end if
							if trim(lcase(fileName))="thistemplate" then								
									fileName=mid(TemplatePath,1,instrrev(TemplatePath,".")-1)
							elseif trim(lcase(fileName))="this" then							
								if sid<>"" then
									fileName=sid
								else
									fileName=did
								end if
							end if
							'处理生成html文件名称
							if isFileNameToPinYin="true" then 
								fileName=HtmlFileName(FileName)						'转拼音去多余字符
							else
								fileName=SetFileName(FileName)						'转正确保存文件类型								
							end if
							'网址不为空则用自定义网址
							Rs("FileName") = fileName	'生成Html文件名

						End If
						
						Rs("Sort") = nSort
						If NavType<>"" Then
							Rs("NavType") = NavType							
							If InStr(NavLocation,"顶部")>0 Then
								Rs("NavTop") = True
							Else
								Rs("NavTop") = False
							End If
						End If
						If (NavType="文本" Or NavActionType="navsmall") and Content="" Then
							Content = Did & " " & Sid & " 详细信息内容写在这里面。更多....."
							
						elseif (NavType="产品" or NavType="新闻") and instr("|no|#|","|" & content & "|")=false then							
							call XY_AutoAddProduct("{$AutoAddProduct did='"& Did &"' title='在这里面写上信息标题[id]{@Array@}' topnumb='20' bigimage='/UploadFiles/NoImg.jpg' default='在这里面写上信息内容[id]' autoadddid='true'$}")
						End If
						If InStr(NavLocation,"底部") > 0 Then Rs("NavButtom") = True
						If InStr(NavLocation,"左边") > 0 Then Rs("NavLeft") = True
						If InStr(NavLocation,"中间") > 0 Then Rs("NavContent") = True
						If InStr(NavLocation,"右边") > 0 Then Rs("NavRight") = True
						If InStr(NavLocation,"其它") > 0 Then Rs("NavOthre") = True				
						'当为产品或新闻呀下载时，添加产品大类 20150124		
						If InStr("|产品|新闻|下载|", "|"& NavType &"|")>0 Then
							Call AutoAddBigClass(Did)
						End If
						if WebTitle="" then
							if sid<>"" then
								WebTitle=sid
							else
								WebTitle=did
							end if
						end if
						if lcase(WebTitle)="null" then
							WebTitle=""
						end if
						call echo(did,webtitle)
						
						Rs("CustomAUrl") = url						
						Rs("WebTitle") = WebTitle
						Rs("Content") = Content
					Rs.Update
				End If:Rs.Close
			End If
		End If
	Next
	'Call Eerr(Default,Action)
End Function


'自动添加大类，小类，子类，产品
Function XY_AutoAddHandle(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql,AutoAdd,I,ULClass
	Dim AutoAddDid,AutoAddSid,AutoAddTid
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)		'获得这么函数参数
	Dim S,C
	dim defaultaddcontent					'默认添加Content内容20150713
	dim addContent							'添加内容
	
	
	defaultaddcontent = ModuleFindContent(Action,"defaultaddcontent")
	
	
	'自动添加大类
	If CheckBigClass(Did) = 0 Then
		AutoAddDid = RParam(Action,"AutoAddDid")
		If LCase(AutoAddDid)="true" Then
			Call AutoAddBigClass(Did)
		End If		
	End If
	'自动添加小类
	If CheckSmallClass(Did,Sid) = 0 Then
		AutoAddSid = RParam(Action,"AutoAddSid")
		If LCase(AutoAddSid)="true" Then
			Call AutoAddSmallClass(Did,Sid)
		End If		
	End If
	'自动添加子类
	If CheckThreeClass(Did,Sid,Tid) = 0 Then
		AutoAddTid = RParam(Action,"AutoAddTid")
		If LCase(AutoAddTid)="true" Then
			Call AutoAddThreeClass(Did,Sid,Tid)
		End If		
	End If
	'自动添加产品
	If CheckProduct(Did,"","")=0 Then
		AutoAdd = RParam(Action,"AutoAdd")
		If LCase(AutoAdd)="true" Then
		 	For I = 1 To TopNumb
			
			if defaultaddcontent<>"" then
					s = defaultaddcontent					
					S = ReplaceValueParam(S,"ni",I)								'不对为i，因为i会与imgurl冲突 [$i$]
					S = ReplaceValueParam(S,"编号-1",I-1)								'不对为i，因为i会与imgurl冲突 [$i$]
					S = ReplaceValueParam(S,"编号",I)								'不对为i，因为i会与imgurl冲突 [$i$]
					S = ReplaceValueParam(S,"did",  Did)	
					S = ReplaceValueParam(S,"sid",  Sid)	
					S = ReplaceValueParam(S,"tid",  Tid)	
				else
					s = Did & "-内容"
				end if
			
				Call AutoAddArticle(Did,Sid,Tid,Did & Sid & Tid & "-标题" & I,s)		
			Next
		End If
	End If	
End Function
'自动添加大类
Function XY_AutoAddDid(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql,ULClass
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	Dim SplStr,S
	SplStr=Split(Did,"[Array]")
	For Each S In SplStr
		If S<>"" Then
			Call AutoAddBigClass(S)
		End If
	Next
End Function
'自动添加小类
Function XY_AutoAddSid(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql,ULClass
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	Dim SplStr,S,Sql,nCount
	SplStr=Split(Tid,"[Array]")
	
	Sql = GetDidSidTidSQl(Did,"","")
	Sql = GetWhereAnd(Sql,AddSql)
	Rs.Open"Select * From [ThreeClass] " & Sql,Conn,1,1
	nCount=Rs.RecordCount : Rs.Close
	If nCount<>0 Then Exit Function
	SplStr=Split(Sid,"[Array]")
	For Each S In SplStr
		If S<>"" Then
			Call AutoAddSmallClass(Did,S)
		End If
	Next
End Function
'自动添加子类
Function XY_AutoAddTid(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql,ULClass
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	Dim SplStr,S,Sql,nCount
	SplStr=Split(Tid,"[Array]")
	
	Sql = GetDidSidTidSQl(Did,Sid,"")	
	Sql = GetWhereAnd(Sql,AddSql)
	Rs.Open"Select * From [ThreeClass] " & Sql,Conn,1,1
	nCount=Rs.RecordCount : Rs.Close
	If nCount<>0 Then Exit Function	
	For Each S In SplStr
		If S<>"" Then
			Call AutoAddThreeClass(Did,Sid,S)
		End If
	Next
End Function
'自动添加产品信息
Function XY_AutoAddProduct(Action)
	Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
	Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
	Call XY_AutoAddHandle(Action)			'处理添加信息处理
	Dim ULClass,Default,Content,I,NewTitle,nCount,SplTitle
	Default = GetDefaultValue(Action)
	Rs.Open"Select * From [Product] " & GetDidSidTidSQl(Did,Sid,Tid),Conn,1,1
	nCount = Rs.RecordCount:Rs.Close
	'有记录则不自动生成产品信息了，不管当前记录是否大于TopNumb设置数 都退出 2014 12 10
	If nCount>0 Then Exit Function
	'Call Echo("Title",Title)
	SplTitle = Split(Title, "{@Array@}")
	If nCount<TopNumb Then
		For I = 1 To TopNumb-nCount
			If Ubound(SplTitle)	>=I-1 Then
				Title=SplTitle(I-1)
				'如果为空则还是用默认的
				if Title="" then
					Title=SplTitle(0)
				end if
			Else
				Title=SplTitle(0)
			End If
			'call echo("Title",Title)
			If Instr(Title,"[id]")>0 Then
				NewTitle = Replace(Title,"[id]",I)
			ElseIf Ubound(SplTitle)<=0 Then
				NewTitle = Title & "-" & I
			Else
				NewTitle = Title
			End If 
			'call echo("NewTitle",NewTitle)
			Content = Replace(Default,"[id]",I)
			Call AutoAddArticleAction(Action,Did,Sid,Tid,NewTitle,I,Content)
		Next
	End If
End Function
'自动添加留言信息
Function XY_AutoAddGuestBook(Action)
	Dim I,TopNumb,GuestName,EMail,Message,Reply,Tel,QQ,MSN,IP,nCount
	TopNumb = RParam(Action,"TopNumb")
	GuestName = RParam(Action,"GuestName")
	EMail = RParam(Action,"EMail")
	Message = RParam(Action,"Message")
	Reply = RParam(Action,"Reply")
	Tel = RParam(Action,"Tel")
	QQ = RParam(Action,"QQ")
	MSN = RParam(Action,"MSN")
	IP = RParam(Action,"IP")
	Rs.Open"Select * From [GuestBook]",Conn,1,1
	nCount = Rs.RecordCount
	Rs.Close
	'有记录则退出
	If nCount>0 Then Exit Function			
	For I=1 To TopNumb
		Rs.Open"Select * From [GuestBook] Where Content='"& Message &"'",Conn,1,3
		If Rs.Eof Then
			Rs.AddNew
			Rs("GuestName")=Replace(GuestName,"[id]",I)			'留言者
			Rs("EMail")=Replace(EMail,"[id]",I)					'邮箱
			Rs("Content")=Replace(Message,"[id]",I)				'留言内容
			Rs("Reply")=Replace(Reply,"[id]",I)					'回答内容
			Rs("Tel")=Replace(Tel,"[id]",I)						'联系电话
			Rs("QQ")=Replace(QQ,"[QQ]",I)						'联系QQ
			Rs("MSN")=Replace(MSN,"[MSN]",I)					'联系MSN
			Rs("IP")=Replace(IP,"[IP]",I)						'客户IP
			Rs.Update
		End If:Rs.Close
	Next
End Function
'更新大类导航
Function XY_UpdateBigNav(Action)
	Dim Did,Content
	Did = RParam(Action,"Did")
	If Did="" Then Exit Function						'退出更新大类导航
	Content = ModuleFindContent(Action,"content")
	Rs.Open"Select * From [NavBigClass] Where BigClassName='"& Did &"'",Conn,1,3
	If Not Rs.Eof Then
		Rs("Content") = Content
		Rs.Update
	End If:Rs.Close
End Function
'自动批量添加文章
Function XY_AutoBatchAddArticle(Action)
	Dim RootDir,FolderPath,Did,Sid,Tid,Content,SplDid,SplSid,SplTid
	RootDir = Trim(RParam(Action,"FolderPath"))
	If RootDir="" Then Exit Function				'目录为空则退出
	If CheckFolder(RootDir)=False Then Exit Function	'检测目录不存在则退出
	Content = GetDirFolderNameList(RootDir)
	If Content = "" Then Exit Function							'主目录下没有文件夹 则退出
	SplDid = Split(Content,vbCrlf)
	For Each Did In SplDid
		If Did <>"" And Left(Did,1)<>"#" Then
			Call AddArticle20141218(RootDir,Did,"","")
			'二级操作
			FolderPath = RootDir & "\" & Did
			Content = GetDirFolderNameList(FolderPath)
			'CAll Echo("Content",Content)
			SplSid = Split(Content,vbCrlf)
			For Each Sid In SplSid
				If Sid <>"" And Left(Sid,1)<>"#" Then
					'Call Echo(RootDir,Did & ">>" & Sid)
					Call AddArticle20141218(RootDir, Did, Sid, "")
				End If
			Next
		End If
	Next	
End Function
'自动添加表单配置
Function XY_AutoAddInputConfig(Action)
	Dim Did,fronttitle,backendtitle,inputtype,fieldname,defaultvalue,inputwidth,inputheight,verificationmsg,listshow,through,verificationtype,nSort,tipstext
	
	Did = RParam(Action,"Did")							'大类名称
	fronttitle = RParam(Action,"fronttitle")			'前端标题
	backendtitle = RParam(Action,"backendtitle")		'后端标题
	inputtype = RParam(Action,"inputtype")				'Input类型
	
	fieldname = RParam(Action,"fieldname")								'字段名称
	defaultvalue = RParam(Action,"defaultvalue")						'默认内容
	inputwidth = RParam(Action,"inputwidth")							'Input宽
	inputheight = RParam(Action,"inputheight")							'Input高
	
	verificationmsg = RParam(Action,"verificationmsg")							'是否为经验码
	listshow = RParam(Action,"listshow")										'列表显示
	through = RParam(Action,"through")											'推荐
	verificationtype = RParam(Action,"verificationtype")						'验证类型
	nSort = RParam(Action,"Sort")												'排序
	tipstext = RParam(Action,"tipstext")										'提示信息
	
	If backendtitle="" Then Exit Function				'前端为空则退出
	Rs.Open"Select * From [InputConfig] Where BigClassName='"& Did &"' And fronttitle='"& fronttitle &"'",Conn,1,3
	If Rs.Eof Then
		Rs.AddNew
		Rs("BigClassName") = Did
		Rs("fronttitle") = fronttitle
		Rs("backendtitle") = backendtitle
		Rs("inputtype") = inputtype
		
		Rs("fieldname") = fieldname
		Rs("defaultvalue") = defaultvalue
		Rs("inputwidth") = inputwidth
		Rs("inputheight") = inputheight
		
		Rs("verificationmsg") = verificationmsg
		Rs("verificationtype") = verificationtype
		
		If through<>"" Then Rs("through") = True
		If listshow<>"" Then Rs("listshow") = True
		If nSort<>"" Then Rs("Sort") = nSort
		Rs("tipstext") = tipstext
		
		Rs.Update
	End If:Rs.Close
	

End Function
'配合上面，添加文章
Function AddArticle20141218(RootDir,Did,Sid,Tid)
	Dim Content,SplStr,FilePath,FileType,FileName,Sql,AddSql
	Call Echo("", RootDir & "\" & Did & "\" & Sid & "\" & Tid)
	Content = GetFileFolderList(RootDir & "\" & Did & "\" & Sid & "\" & Tid, True, "|txt|doc|docx|", "", "", "", "")
	SplStr=Split(Content,vbCrlf)
	For Each FilePath In SplStr
		If FilePath<>"" Then
			FileType = LCase(getFileExtName(FilePath))
			FileName = GetFilePathName(FilePath)
			If FileType = "txt" Then
				Content=GetFText(FilePath)
			ElseIf FileType = "doc" Or FileType = "docx" Then
				Content=GetDocText(FilePath)
			End If
			Sql = "Select * From [Product] Where Title='"& FileName &"' "
			If Did<>"" Then Sql = GetWhereAnd(Sql, "And BigClassName='"& Did &"'")
			If Sid<>"" Then Sql = GetWhereAnd(Sql, "And SmallClassName='"& Sid &"'")
			If Tid<>"" Then Sql = GetWhereAnd(Sql, "And ThreeClassName='"& Tid &"'")			
			Rs.Open Sql,Conn,1,3
			If Rs.Eof Then
				Rs.AddNew
				Rs("BigClassName") = Did
				Rs("SmallClassName") = Sid
				Rs("ThreeClassName") = Tid
				Rs("WebTitle") = FileName
				Rs("Title") = FileName
				Rs("FileName") = HtmlFileName(FileName)
				Rs("Content") = Content
				Rs.Update		
			End If:Rs.Close
			'Call Echo("Sql",Sql)
			'Call Echo("FilePath",FilePath)
			'Call Echo("FileType",FileType)
			'Call Echo("FileName",FileName)
		End IF
	Next
	'Call Echo("Content",Content)
End Function



'读模板资源内容
Function XY_ReadTemplateSource(Action)
	Dim SourceId
	SourceId = RParam(Action,"SourceId")	
	XY_ReadTemplateSource = ReadTemplateSource(SourceId)
End Function

'读模块内容 和上面一样
Function XY_GetContentModule(Action)
	XY_GetContentModule = XY_ReadTemplateModule(Action)
End Function
'替换模板内容 20150109
Function XY_ReplaceTemplateContent(Content,TemplateLableStr,Action)
	'Call die(Action)
	Dim SourceStr,ReplaceStr
	SourceStr = ModuleFindContent(Action,"SourceStr")
	ReplaceStr = ModuleFindContent(Action,"ReplaceStr")
	'CAll Rw(Action)
	'Call Eerr(SourceStr,ReplaceStr)
	If InStr(Content,SourceStr)>0 Then
		Content = Replace(Content,TemplateLableStr,"")				'先把这个替换模板标签内容删除
		Content = Replace(Content,SourceStr,ReplaceStr)				'替换内容
		'Call Die(Content)
	End If
	XY_ReplaceTemplateContent = ""
End Function

'显示模板模块
Function XY_ShowTemplateModule(Action)
	Dim DefaultModule,SelectModule
	Dim Did,Sid,Tid
	Dim MainInfoModule,ProductModule
	DefaultModule = RParam(Action,"DefaultModule")
	SelectModule = RParam(Action,"SelectModule")
	
	'Call echo("DefaultModule",DefaultModule)
	Did = PubProDid
	Sid = PubProSid
	Tid = PubProTid
	
	'详细页
	If PubID<>"" Then
		If PubActionType = "MainInfo" Then
			Rs.Open"Select * From [MainInfo] Where Id=" & PubID,Conn,1,1
			If Not Rs.Eof Then
				If Rs(SelectModule)<>"" Then
					MainInfoModule = Rs(SelectModule)
					
				End If
			End If:Rs.Close		
		ElseIf PubActionType = "Article" Then
			Rs.Open"Select * From [Product] Where Id=" & PubID,Conn,1,1
			If Not Rs.Eof Then
				If Rs(SelectModule)<>"" Then
					ProductModule = Rs(SelectModule)
				End If
				Did=Rs("BigClassName")
				Sid=Rs("SmallClassName")
				Tid=Rs("ThreeClassName")
			End If:Rs.Close			
		End If
	End If
	'导航大类不为空
	If PubNavDid<>"" Then
		Rs.Open"Select * From [NavBigClass] Where BigClassName='"& PubNavDid &"'",Conn,1,1
		If Not Rs.Eof Then
			
			If Rs(SelectModule)<>"" Then
				DefaultModule = Rs(SelectModule)
			End If
		End If:Rs.Close
	End If
	'导航小类不为空
	If PubNavSid<>"" Then
		Rs.Open"Select * From [NavSmallClass] Where BigClassName='"& PubNavDid &"' And SmallClassName='"& PubNavSid &"'",Conn,1,1
		If Not Rs.Eof Then
			If Rs(SelectModule)<>"" Then
				DefaultModule = Rs(SelectModule)
			End If
		End If:Rs.Close
	End If
	
	'类别大类不为空
	If Did<>"" Then
		Rs.Open"Select * From [BigClass] Where BigClassName='"& Did &"'",Conn,1,1
		If Not Rs.Eof Then
			If Rs(SelectModule)<>"" Then
				DefaultModule = Rs(SelectModule)
			End If
		End If:Rs.Close
	End If
	'类别小类不为空
	If Sid<>"" Then
		Rs.Open"Select * From [SmallClass] Where BigClassName='"& Did &"' And SmallClassName='"& Sid &"'",Conn,1,1
		If Not Rs.Eof Then
			If Rs(SelectModule)<>"" Then
				DefaultModule = Rs(SelectModule)
			End If
		End If:Rs.Close
	End If
	'类别子类不为空
	If Tid<>"" Then
		Rs.Open"Select * From [SmallClass] Where BigClassName='"& Did &"' And SmallClassName='"& Sid &"' And ThreeClassName='"& Tid &"'",Conn,1,1
		If Not Rs.Eof Then
			If Rs(SelectModule)<>"" Then
				DefaultModule = Rs(SelectModule)
			End If
		End If:Rs.Close
	End If

	
	If MainInfoModule<>"" Then
		DefaultModule = MainInfoModule
	ElseIf ProductModule<>"" Then
		DefaultModule = ProductModule
	End If
	
	
	Dim SplStr,Title,C
	SplStr=Split(DefaultModule,"|")
	For Each Title In SplStr
		If Title<>"" Then
			Rs.Open"Select * From [TemplateModule] Where Title='"& Title &"'",Conn,1,1
			If Not Rs.Eof Then
				C=C & Rs("Content") & vbCrlf
			End If:Rs.Close	
		End If
	Next
	
	'Call RW(C)
		
	C = HandleCodeReplace(C)			'处理这里面模块样式			20150213
	'Call Die("<hr>" & C)
	
	XY_ShowTemplateModule = C
End Function



'获得表字段内容     辅助类
Function GetTableFieldValue(Action)			
	Dim StartStr,EndStr,Content,TableName,FieldName,AddSql
	StartStr = " table='" : EndStr = "'"
	If InStr(Action,StartStr)>0 And InStr(Action,EndStr)>0 Then
		TableName = StrCut(Action,StartStr,EndStr,2)
	End If
	StartStr = " field='" : EndStr = "'"
	If InStr(Action,StartStr)>0 And InStr(Action,EndStr)>0 Then
		FieldName = StrCut(Action,StartStr,EndStr,2)
	End If
	AddSql = "Select * From ["& TableName &"]"
	If UserId<>"" Then
		If LCase(TableName)<>"member" Then
			AddSql = AddSql & " Where UserId=" & UserId
		Else
			AddSql = AddSql & " Where Id=" & UserId
		End If
	End If
	'SQL出错了 没有当前表与字段
	If CheckSql("Select "& FieldName &" From ["& TableName &"]")=False Then Exit Function
	''SQL出错了  语法错误
	If CheckSql(AddSql)=False Then Exit Function			
	Rs.Open AddSql,Conn,1,1
	If Not Rs.Eof Then
		GetTableFieldValue = Rs(FieldName)
	Else
		GetTableFieldValue = "表["& TableName &"]字段["& FieldName &"]没找到内容"		'测试用
	End If : Rs.Close
End Function



'VB不引用 Start

'在线显示面板
function XY_onLineViewDialog(Action)
    Dim C,operate,url
	'为假退出
	operate = LCase(RParam(Action,"operate"))
	if operate<>"true" then
		exit function
	end if
	
    C = "<style type=""text/css"">" & vbCrlf
    C = C & "#dontrolbox{" & vbCrlf
    C = C & "	position:fixed !Important;bottom:0;right:0;position:absolute;z-index:1000;" & vbCrlf
    C = C & "	border:1px solid #CCCCCC;" & vbCrlf
    C = C & "	background-color:#FDFCE8;" & vbCrlf
    C = C & "	width:110px;" & vbCrlf
    C = C & "	height:100px;" & vbCrlf
    C = C & "	text-align:center;" & vbCrlf
    C = C & "	line-height:26px;" & vbCrlf
    C = C & "	font-size:12px;" & vbCrlf
    C = C & "}" & vbCrlf
    C = C & "</style>" & vbCrlf
    C = C & "<div id=""dontrolbox"">" & vbCrlf
	url = GetUrlAddToParam(geturl,"MackHtml=False","replace")	
	url = GetUrlAddToParam(url,"nn=" & GetRnd(11),"replace")		'加个随机数
	
    C = C & "	<a href="""&   GetUrlAddToParam(url,"gl=edit","delete") & """>正确显示</a><br>" & vbCrlf
    C = C & "	<a href="""&   GetUrlAddToParam(url,"gl=edit","replace") & """>在线修改</a><br>" & vbCrlf
	
	url = GetUrlAddToParam(geturl,"MackHtml=true","replace")
	url = GetUrlAddToParam(url,"gl=edit","delete")
	
    C = C & "	<a href="""& GetUrlAddToParam(url,"gotoHtml=true","replace") &""" target='_blank'>HTML版</a>" & vbCrlf
    C = C & "</div>" & vbCrlf
	
	'判断本地
    If GetIP = "127.0.0.1" Or InStr(GetIP, "192.168.") > 0 Then
		If Request("MackHtml") = "False" Then
        	XY_onLineViewDialog = C
		end if
    End If	
end function


'Js调试网站皮肤
Function XY_DebugWebSkin(Action)
	Dim DirPath,DirList,SplStr,S,C ,HttpUrl,Url,FieldName,FieldContent
	Dim CssStr	
	SplStr=Split(Request.QueryString(),"&")
	For Each S In SplStr
		FieldName = Mid(S,1,InStr(S,"=")-1)
		FieldContent = Mid(S,InStr(S,"=")+1) 		
		If LCase(FieldName)<>"n" And LCase(FieldName)<>"findtpl" And LCase(FieldName)<>"replacetpl" Then
			If FieldContent<>"" Then
				Url = Url & IIF(Url="","?","&") & FieldName & "=" & FieldContent
			End If
		End If
	Next
	Url =  "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME") & Url & "&n=" & GetRnd(11)
	'CAll Echo("Url",Url)
	
	DirPath = WebSkins & "Index\column"
	DirList = GetDirFolderNameList(DirPath)
	SplStr=Split(DirList,vbCrlf)
	For Each S In SplStr
		If S <>"" Then
			HttpUrl=Url & "&FindTpl=style='312'&ReplaceTpl=style='"& S &"'"
			C=C & "<a href="""& HttpUrl &""">"& S &"</a><br>" & vbCrlf
		End If
	Next
	If C <> "" Then	
		CssStr = "<style>" & vbCrlf
		CssStr = CssStr & "/*网站样式切换 <div id=""webStyleTab""></div>*/" & vbCrlf
		CssStr = CssStr & "#webStyleTab{" & vbCrlf
		CssStr = CssStr & "	position:fixed;" & vbCrlf
		CssStr = CssStr & "	width:51px;" & vbCrlf
		CssStr = CssStr & "	height:177px;" & vbCrlf
		CssStr = CssStr & "	z-index:999;" & vbCrlf
		CssStr = CssStr & "	left: 0px;" & vbCrlf
		CssStr = CssStr & "	top: 0px;" & vbCrlf
		CssStr = CssStr & "}" & vbCrlf
		CssStr = CssStr & "</style>" & vbCrlf	
		C = CssStr & "<div id=""webStyleTab"">" & C &  "</div>"
	End If
	XY_DebugWebSkin = C
End Function


'管理
Select Case Request("actgl")
	Case "login" : Call gl_Login()
	Case "action" : Call gl_Action()
	Case "delaction" : Call dle_Action()
	Case "testdialog" : Call test_Dialog()
End Select

'获得动作
Sub gl_Action()
	Call OpenConn()
	Rs.Open"Select * From [WebSite]",Conn,1,1
	If Not Rs.Eof Then
		Call Die("【显示动作】" & vbCrlf & Rs("TemplateMain3"))
	End If
	Response.End()
End Sub
'删除动作
Sub dle_Action()
	Call OpenConn()
	Rs.Open"Select * From [WebSite]",Conn,1,3
	If Not Rs.Eof Then
		Rs("TemplateMain3") = ""
		Rs.Update
		Call Rw("删除动作成功")
	End If
	Response.End()
End Sub
'记录登录
Sub gl_Login()
	Call OpenConn()
	Rs.Open"Select * From [WebSite]",Conn,1,3
	If Not Rs.Eof Then
		Rs("TemplateMain2") = Format_Time(Now(),1) & vbCrlf & Rs("TemplateMain2")
		Rs.Update
		Call Rw("登录成功")
	End If
	Response.End()
End Sub
'显示测试面板
Sub test_Dialog()
	Call Echo("记录登录", "<a href='?actgl=login' target='_blank'>记录登录</a>")
	Call Echo("获得动作", "<a href='?actgl=action' target='_blank'>获得动作</a>")
	Call Echo("删除动作", "<a href='?actgl=delaction' target='_blank'>删除动作</a>")
	Response.End()
End Sub

'拷贝模板素材
Function VB_copyTemplateMaterial(Action)	
	VB_copyTemplateMaterial=""
End Function
'给VB里用，主要是屏蔽指定标签，不处理   20150628
function pingBiLable(action)
	pingBiLable = false
end function
'VB不引用 End






'获得网站首页
Function getWebHome()
    Dim url 
    If getIP = "127.0.0.1" Or InStr(getIP, "192.168.") > 0 Then
        Call OpenConn() 
        rs.Open "Select * From [WebSite]", conn, 1, 1 
        If Not rs.EOF Then
            getWebHome = "/wwwroot/" & getEndUrlHandleName(rs("WebSkins")) & "/Index.html" 
        End If : rs.Close 
    End If 
    If getWebHome = "" Then
        getWebHome = "/Index.html" 
    End If 
End Function 











'处理Html标签闭合(20150831) 第一种，淘汰
Function handleHtmlLabelClose(content, labelName, isAddAlt)
    Dim startStr, endStr, splStr, s, ImgList, imgStr, LCaseImgStr, newImgStr 
    startStr = "<" & labelName : endStr = ">" 
    ImgList = GetArray(content, startStr, endStr, True, False) 
    'call rw(ImgList)
    splStr = Split(ImgList, "$Array$") 
    For Each imgStr In splStr
        newImgStr = phpTrim(imgStr) 
        If Right(newImgStr, 1) = "/" Then
            newImgStr = PHPTrim(Left(newImgStr, Len(newImgStr) - 1)) 
        End If 
        If isAddAlt = True Then
            LCaseImgStr = LCase(imgStr) 
            If InStr(LCaseImgStr, "alt") = False Then
                newImgStr = newImgStr & " alt=""""" 
            End If 
        End If 
        content = Replace(content, imgStr & ">", newImgStr & " />") 
    Next 
    handleHtmlLabelClose = content 
End Function
'处理图片闭合(20150831)
Function handleImgClose(content, isAddAlt)
    handleImgClose = handleHtmlLabelClose(content, "img", isAddAlt) 
End Function 
'处理BR闭合(20150831)
Function handleBrClose(content, isAddAlt)
    handleBrClose = handleHtmlLabelClose(content, "br", False) 
    handleBrClose = Replace(handleBrClose, "<br>", "<br />") 
End Function
'处理HR闭合(20150831)
Function handleHrClose(content, isAddAlt)
    handleHrClose = handleHtmlLabelClose(content, "hr", False) 
    handleHrClose = Replace(handleHrClose, "<hr>", "<hr />") 
End Function




'处理常用函数附加参数    辅助类
'Dim Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql
'Call HandleFunParameter(Action,Did,Sid,Tid,Title,TopNumb,CutStrNumb,AddSql)
Sub handleFunParameter(action, did, sid, tid, title, topNumb, cutStrNumb, addSql)
    Dim startStr, endStr 
    did = RParam(action, "Did") 
    sid = RParam(action, "Sid") 
    tid = RParam(action, "Tid") 

    did = IIF(did = "[$PubProDid$]", PubProDid, did) 
    sid = IIF(sid = "[$PubProSid$]", PubProSid, sid) 
    tid = IIF(tid = "[$PubProTid$]", PubProTid, tid) 

    title = RParam(action, "Title") 
    topNumb = RParam(action, "TopNumb") 
    cutStrNumb = RParam(action, "CutStrNumb") 
    If cutStrNumb = "" Then cutStrNumb = 28                                         '默认截取字符为32
    addSql = RParam(action, "AddSql") 
End Sub 
%>

















