<%
'======================================
'数据库表操作
'======================================
 

'给表配置
Function ShowClassName(Table, Content, AscDescStr,PrintStr, TDWidth)
	Dim SplStr,Splxx,S,C,N,I
	Dim IDStr,QQStr,QQNameStr,HandleStr,VisitorsListStr,SpacePurviewStr,OnLineTimeStr,QQSexStr,QQAgeStr,AgeStr
	Dim BigClassNameStr,SmallClassNameStr,CountNumbStr,QQNumbStr
	Dim CategoryStr,TitleStr,BlogidStr,MessageContentStr, FileNameStr,FolderNameStr, DirPathStr
	Dim PhotoPreStr,PhotoNameStr,BigPhotoUrlStr,PhotoIDStr,UpDateTimeStr,PhotoUrlStr,ContentStr,MessageQQStr,MessageUpDateTimeStr
	Dim AddTimeStr,SearchValueStr,SearchTitleStr,SearchContentStr,SearchTimeStr,SearchSortStr,WebSiteWebHttpUrlStr
	Dim UserNameStr,PriceStr,FarePriceStr,ProvinceStr,CityStr,DescriptionStr,ReplyNumbStr,ReplyNameStr,QQTouXiangStr
	Dim ZipSizeStr,WindowsStr,SortStr,PhotoShowTypeStr,WorkTitleStr,CompanyStr,DiQuStr,WorkEXPStr,QQBeiZhuNameStr
	Dim AuthorStr,RenQiStr,WebTitleStr,WebUrlStr,DayNumbStr,CiStr,ShiStr,PoJieStr,PoJiePhotoStr,DownNumbStr,ENStr,CNStr
	Dim HttpUrlStr,PlayTimeStr,PlayNumbStr,PostingsTimeStr,NavNameStr,ThreeClassNameStr,ClickNumbStr,MutualFriendStr,VisitorsNumbStr
	Dim SupportNumbStr,CreateDateStr,ThroughStr,SearchWordStr,XiangGuanWordStr,AdoptionRateStr,BrowseNumbStr,EmailStr,IPStr
	Dim AddDateTimeStr,FilePathStr,MyMoneyStr,ReplyTimeStr,AskTypeStr,AskIdStr,WebSiteStr,PageNumberStr,RankingStr,IncludedTimeStr
	Dim SearchUrlStr,CateStr,PubTimeStr
	IDStr = "选择|42|,"
	CateStr = "分类|42|Cate|分类重复,"
	PubTimeStr = "发布时间|100|PubTime|日志发布时间重复,"
	QQStr = "QQ号码|70|QQ|QQ号码重复|QQ号码搜索,"
	QQNameStr = "QQ名称|200|QQName|QQ名称重复|QQ名称搜索,"
	HandleStr = "操作||详细|||修改,"
	VisitorsListStr = "访客列表|80|VisitorsList||访客列表搜索,"
	VisitorsNumbStr = "访客数量|80|VisitorsNumb||访客数量搜索,"
	
	SpacePurviewStr = "空间权限|80|SpacePurview|空间权限重复|空间权限搜索,"
	QQSexStr = "QQ性别|60|QQSex|QQ性别重复|QQ性别搜索,"
	QQAgeStr = "QQ年龄|60|QQAge|QQ年龄重复|QQ年龄搜索,"
	AgeStr = "年龄|60|Age|年龄重复|年龄搜索,"
	'关系 
	BigClassNameStr = "大类名称|100|BigClassName|大类名称重复|大类名称搜索," 
	SmallClassNameStr = "小类名称|100|SmallClassName|小类名称重复|小类名称搜索," 
	ThreeClassNameStr = "子类名称|100|ThreeClassName|子类名称重复|子类名称搜索," 
	
	SearchWordStr = "搜索词|200|BigClassName|搜索词重复|搜索词搜索," 
	XiangGuanWordStr = "相关词|200|SmallClassName|相关词重复|相关词搜索,"  
	
	CountNumbStr = "好友度|60|CountNumb," 											'注意这里 下面 CountNumbStr总数给代替了
	QQNumbStr = "存在数|60|QQNumb," 
	'日志
	OnLineTimeStr = "登录时间|116|OnLineTime|登录时间重复|登录时间搜索," 
	CategoryStr = "日志类型|80|Category|日志类型重复|日志类型搜索,"
	TitleStr = "标题|204|Title|标题重复|标题搜索," 
	BlogidStr = "日志ID|70|Blogid|日志ID重复,"
	'相册
	PhotoPreStr = "缩览图|100|PhotoPre,"
	PhotoNameStr = "图片名称|100|PhotoName|图片名称重复|图片名称搜索,"
	BigPhotoUrlStr = "图片缩览图|100|BigPhotoUrl||图片缩览图搜索,"
	PhotoIDStr = "图片ID|100|PhotoID|图片ID重复|图片ID搜索,"	
	MessageQQStr = "留言者QQ|70|MessageQQ|留言者QQ重复|留言者QQ搜索," 
	MessageUpDateTimeStr = "留言时间|120|UpDateTime|留言时间重复|留言时间搜索,"
	PhotoUrlStr = "留言缩览图|100|PhotoUrl,"
	MessageContentStr = "留言内容||Content||留言内容搜索,"
	PhotoShowTypeStr = "存在大小图|80|PhotoShowType|存在大小图重复|存在大小图搜索,"
	'相同	
	CountNumbStr = "总数|60|CountNumb," 
	AddTimeStr = "添加时间|120|AddTime|添加时间重复|添加时间搜索," 
	'搜索引擎
	SearchValueStr = "搜索值|100|SearchValue|搜索值重复|搜索值搜索,"
	SearchTitleStr = "搜索标题|200|SearchTitle|搜索标题重复|搜索标题搜索,"
	SearchContentStr = "搜索内容|200|SearchContent||搜索内容搜索,"
	SearchTimeStr = "搜索时间|80|SearchTime|搜索时间重复|搜索时间搜索," 
	SearchSortStr = "搜索排序|80|SearchSort|搜索排序重复|搜索排序搜索,"
	WebSiteWebHttpUrlStr = "域名/网址|200|域名/网址|域名重复|网址搜索/域名搜索,"
	'淘宝
	UserNameStr = "用户|100|UserName|用户重复|用户搜索,"
	PriceStr = "价格|60|Price|价格重复|价格搜索,"
	FarePriceStr = "运费|60|FarePrice|运费重复|运费搜索,"
	ProvinceStr = "省份|40|Province|省份重复|省份搜索,"
	CityStr = "城市|40|City|城市重复|城市搜索,"
	DescriptionStr = "描述值|60|Description|描述值重复|描述值搜索,"
	ReplyNumbStr = "评论数|60|ReplyNumb|评论数重复|评论数搜索,"
	ReplyNameStr = "评论者|60|ReplyName|评论者重复|评论者搜索,"
	ClickNumbStr = "点击数|60|ClickNumb|点击数重复|点击数搜索,"	
	'产品
	ContentStr = "备注|200|Content||备注搜索,"	'这样注意200可去掉，让它自动
	UpDateTimeStr = "上传时间|80|UpDateTime|上传时间重复|上传时间搜索,"
	ZipSizeStr = "文件大小|60|ZipSize|文件大小重复,"
	WindowsStr = "运行平台|80|Windows|运行平台重复|运行平台搜索,"
	DownNumbStr = "下载数|80|DownNumb|下载数重复|下载数搜索,"
	'Words
	SortStr = "排序|60|Sort|排序重复|排序搜索,"
	'招聘
	WorkTitleStr = "招聘职位|100|WorkTitle|招聘职位重复|招聘职位搜索,"
	CompanyStr = "公司名称|100|Company|公司名称重复|公司名称搜索,"
	DiQuStr = "地区|40|DiQu|地区重复|地区搜索,"
	WorkEXPStr = "工作经验|90|WorkEXP|工作经验重复|工作经验搜索,"
	'QQ备注
	QQBeiZhuNameStr = "QQ备注名|90|QQBeiZhuName|QQ备注名重复|QQ备注名搜索,"
	QQTouXiangStr = "QQ头像|100|QQ头像||,"
	'Flash163SWF
	AuthorStr = "作者|90|Author|作者重复|作者搜索,"
	RenQiStr = "人气|90|RenQi|人气重复|人气搜索,"
	'Web
	WebTitleStr = "网站标题|200|WebTitle|网站标题重复|网站标题搜索,"
	WebUrlStr = "网站地址|200|WebUrl|网站地址重复|网站地址搜索,"
	DayNumbStr = "星期数|200|DayNumb|星期数重复|星期数搜索,"
	'ce
	CiStr = "单词|100|Ci|单词重复|单词搜索,"
	ShiStr = "英文解释||Shi|英文解释重复|英文解释搜索,"
	'QQ图片破解
	PoJieStr = "破解值|80|PoJie|破解值重复|破解值搜索,"
	PoJiePhotoStr = "破解图片|80|PoJiePhoto,"
	'CnEnWords
	ENStr = "英文|120|EN|英文重复|英文搜索,"
	CNStr = "中文|200|CN|中文重复|中文搜索,"
	'文件（夹）目录
	FileNameStr = "文件|150|FileName|文件重复|文件搜索,"
	FolderNameStr = "文件夹|120|FolderName|文件夹重复|文件夹搜索,"
	DirPathStr = "目录路径||DirPath|目录路径重复|目录路径搜索," 
	'视频
	HttpUrlStr = "网址|200|HttpUrl|网址重复|网址搜索,"
	PlayTimeStr = "播放时间|80|PlayTime|播放时间重复|播放时间搜索,"
	PlayNumbStr = "播放次数|80|PlayNumb|播放次数重复|播放次数搜索,"
	PostingsTimeStr = "发布时间|100|PostingsTime|发布时间重复|发布时间搜索,"
	'其它
	NavNameStr = "导航名称|100|cname|导航名称重复|导航名称搜索,"
	MutualFriendStr = "共同好友|60|MutualFriend|共同好友重复|共同好友搜索,"
	SupportNumbStr = "支持数|80|SupportNumb|支持数重复|支持数搜索,"
	CreateDateStr = "提交日期|80|CreateDate|提交日期重复|提交日期搜索,"
	ThroughStr = "是否通过|80|Through|是否通过重复|是否通过搜索,"
	AdoptionRateStr = "采纳率|80|AdoptionRate|采纳率重复|采纳率搜索,"
	BrowseNumbStr = "浏览次数|80|BrowseNumb|浏览次数重复|浏览次数搜索,"
	EmailStr = "邮箱|120|Email|邮箱重复|邮箱搜索,"
	IPStr = "IP|120|IP|IP重复|IP搜索,"
	AddDateTimeStr = "添加时间|120|AddDateTime|添加时间重复|添加时间搜索,"
	FilePathStr = "文件路径|120|FilePath|文件路径重复|文件路径搜索,"
	MyMoneyStr ="用户积分|120|MyMoney|用户积分重复|用户积分搜索,"
	ReplyTimeStr = "提问时间|120|ReplyTime|提问时间重复|提问时间搜索,"
	AskTypeStr = "问问类型|120|AskType|问问类型重复|问问类型搜索,"
	AskIdStr = "问问ID|120|AskId|问问ID重复|问问ID搜索,"
	WebSiteStr = "域名|120|WebSite|域名重复|域名搜索,"
	PageNumberStr = "页数|120|PageNumber|页数重|复页数搜索,"
	RankingStr = "排名|120|Ranking|排名重复|排名搜索,"
	IncludedTimeStr = "收录时间|120|IncludedTime|收录时间重复|收录时间搜索,"
	SearchUrlStr = "搜索引擎网址|120|SearchUrl|搜索引擎网址重复|搜索引擎网址搜索,"
	
	'【产品表】
	If CheckStr(Table, "^Product$") Then
		Content = IDStr & BigClassNameStr & SmallClassNameStr & ThreeClassNameStr & TitleStr & ContentStr  & UpDateTimeStr & HandleStr
		AscDescStr = "点击数"
		PrintStr = "标题|内容"
	'【QQ名称表】
	ElseIf CheckStr(Table, "^QQNameTable$") Then
		Content = IDStr & QQStr & QQNameStr & ContentStr & HandleStr
		AscDescStr = "QQ号码|QQ名称"
		PrintStr = "QQ号码|QQ名称|QQ名称和号码"
	'【QQ备注表】
	ElseIf CheckStr(Table, "^QQBeiZhu$") Then
		Content = IDStr & QQStr & QQBeiZhuNameStr & HandleStr
		AscDescStr = "QQ号码|QQ名称"
		PrintStr = "QQ号码|QQ备注名|QQ名称和号码"
	'【QQ表】
	ElseIf CheckStr(Table, "^QQ$") Then
		If Rq("Style") = "2" Then
			Content = IDStr & QQStr & QQNameStr & QQSexStr & AgeStr & UpDateTimeStr & QQAgeStr & HandleStr
			AscDescStr = "QQ号码|QQ名称|QQ性别|年龄|上传时间|QQ年龄"
			PrintStr = "QQ号码|QQ访客|QQ和访客"
		Else
			Content = IDStr & BigClassNameStr & QQStr & QQNameStr & VisitorsListStr & SpacePurviewStr & OnLineTimeStr & MutualFriendStr & VisitorsNumbStr & HandleStr
			AscDescStr = "QQ号码|QQ名称|空间权限|登录时间|访客数量"
			PrintStr = "QQ号码|QQ访客|QQ和访客"
		End If
	'【关系QQ表】
	ElseIf CheckStr(Table, "^GuanXiQQ$") Then
		Content = IDStr & BigClassNameStr & SmallClassNameStr & QQStr & QQNameStr & CountNumbStr & QQNumbStr & HandleStr
		AscDescStr = "好友度|存在数|好友度＋存在数|[联合查询]"
		PrintStr = "QQ号码|QQ名称|QQ名称和号码"
	'【日志表】
	ElseIf CheckStr(Table, "^DailyRcord$") Then
		Content = IDStr & QQStr & CategoryStr & TitleStr & ReplyNumbStr & BlogidStr & HandleStr
		AscDescStr = "QQ号码|评论数"
		PrintStr = "QQ号码|日志标题|日志地址"
	'【QQ相册类型表】
	ElseIf CheckStr(Table, "^QQPhotoClass$") Then
		Content = IDStr & BigClassNameStr & QQStr & PhotoPreStr & PhotoShowTypeStr & HandleStr
		AscDescStr = "QQ号码"
		PrintStr = "大类名称|QQ号码|缩览图"
	'【QQ相册信息表】
	ElseIf CheckStr(Table, "^QQPhotoProduct$") Then
		Content = IDStr & QQStr & PhotoNameStr & BigPhotoUrlStr & PhotoIDStr & PhotoShowTypeStr & HandleStr
		AscDescStr = "QQ号码"
		PrintStr = "QQ号码|相册小图|相册大图"
	'【QQ相册留言表】
	ElseIf CheckStr(Table, "^QQPhotoMessage$") Then
		Content = IDStr & MessageQQStr &  QQStr & QQNameStr  & MessageUpDateTimeStr & PhotoUrlStr & PhotoShowTypeStr & MessageContentStr
		AscDescStr = "QQ号码|留言时间"
		PrintStr = "QQ号码|留言图片|留言内容"
	'【QQ在线表】
	ElseIf CheckStr(Table, "^QQOnLine") Then
		Content = IDStr &  QQStr & OnLineTimeStr & HandleStr
		AscDescStr = "QQ号码|登录时间"
		PrintStr = "QQ号码"
	'【QQ日志】
	ElseIf CheckStr(Table, "^QQJournal") Then
		Content = IDStr &  QQStr & CateStr &  BlogIdStr & TitleStr & PubTimeStr & HandleStr
		AscDescStr = "QQ号码|登录时间"
		PrintStr = "QQ号码"
	'【相同表】
	ElseIf CheckStr(Table, "^WithTable$") Then
		Content = IDStr & BigClassNameStr & SmallClassNameStr &  TitleStr & CountNumbStr & HandleStr
		AscDescStr = "总数"
		PrintStr = "标题|分析出图"
	'【QQ暂存表】
	ElseIf CheckStr(Table, "^QQTemp$") Then
		Content = IDStr &  QQStr & AddTimeStr & HandleStr
		AscDescStr = "QQ号码|添加时间"
		PrintStr = "QQ号码"
	'【搜索引擎表】
	ElseIf CheckStr(Table, "^SearchEngine$") Then
		Content = IDStr &  SearchValueStr & SearchTitleStr & SearchTimeStr & SearchSortStr & SearchContentStr & WebSiteWebHttpUrlStr & HandleStr
		AscDescStr = "QQ号码|添加时间"
		PrintStr = "搜索标题|搜索网址|域名"
	'【淘宝表】
	ElseIf CheckStr(Table, "^TaoBaoProduct$") Then
		Content = IDStr &  SearchValueStr & UserNameStr & PriceStr & FarePriceStr & ProvinceStr & CityStr & DescriptionStr & ReplyNumbStr
		AscDescStr = "价格|运费|描述值|评论"
		PrintStr = "用户|标题|价格"
	'【文档表】
	ElseIf CheckStr(Table, "^Word$") Then
		Content = IDStr &  TitleStr & AddTimeStr & SortStr & HandleStr
		AscDescStr = "添加时间|排序"
		PrintStr = "标题"
	'【51招聘表】
	ElseIf CheckStr(Table, "^JOB_51Job$") Then
		Content = IDStr & SearchValueStr & WorkTitleStr & CompanyStr & CityStr & DiQuStr & UpDateTimeStr & WorkEXPStr  & HandleStr
		AscDescStr = "上传时间|排序"
		PrintStr = "招聘职位|公司名称|职位简介|招聘职位和职位简介"
	'【？？？表】
	ElseIf CheckStr(Table, "^QQTouXiang$") Then
		Content = IDStr & QQTouXiangStr & QQStr & PhotoShowTypeStr & HandleStr
		AscDescStr = "添加时间|排序"
		PrintStr = "标题"
	'【大活动漫表】
	ElseIf CheckStr(Table, "^Flash163SWF$") Then
		Content = IDStr & TitleStr & AuthorStr & RenQiStr & PhotoShowTypeStr & HandleStr
		AscDescStr = "标题|人气|作者"
		PrintStr = "标题|图片地址|SWF地址|下载地址"
	'【QQ访客表】
	ElseIf CheckStr(Table, "^QQVisitorsStat$") Then
		Content = IDStr & QQStr & QQNameStr & CountNumbStr & HandleStr
		AscDescStr = "QQ号码|QQ名称|总数"
		PrintStr = "QQ号码|QQ名称|QQ名称和号码"
	'【暂存网址表】
	ElseIf CheckStr(Table, "^TempUrl$") Then
		Content = IDStr & FolderNameStr & FileNameStr & DirPathStr & HandleStr
		AscDescStr = "文件|文件夹|目录路径"
		PrintStr = "文件|文件夹|目录路径"
	'【视频表】
	ElseIf CheckStr(Table, "^Video$") Then
		Content = IDStr & TitleStr & HttpUrlStr & PlayTimeStr & PlayNumbStr & PostingsTimeStr & AuthorStr & HandleStr
		AscDescStr = "播放时间|播放次数|发布时间"
		PrintStr = "视频图片|视频标题|视频网址|视频标题加网址" 
	'【Web表】外部MDB表需要
	ElseIf CheckStr(Table, "^Web$") Then
		Content = IDStr & WebTitleStr & WebUrlStr & DayNumbStr & HandleStr
		AscDescStr = "星期数"
		PrintStr = "网站标题|网站地址|网站标题加地址"
	'【单词表】外部MDB表需要
	ElseIf CheckStr(Table, "^ce$") Then
		Content = IDStr & CiStr & ShiStr & HandleStr
		AscDescStr = "单词"
		PrintStr = "单词|英文解释|单词加英文解释"
	'【验证码表】外部MDB表需要
	ElseIf CheckStr(Table, "^QQYingZheng$") Then
		Content = IDStr & TitleStr & PoJieStr & PoJiePhotoStr & HandleStr
		AscDescStr = "破解值"
		PrintStr = "破解值"
	'【中英表】外部MDB表需要
	ElseIf CheckStr(Table, "^CnEnWords$") Then
		Content = IDStr & ENStr & CNStr & SortStr & HandleStr
		AscDescStr = "排序"
		PrintStr = "中英对照"
	'【类表】外部MDB表需要
	ElseIf CheckStr(Table, "^class$") Then
		Content = IDStr & NavNameStr & HttpUrlStr & HandleStr
		AscDescStr = "导航名称"
		PrintStr = "导航名称"
	'【类表】QQ名称表
	ElseIf CheckStr(Table, "^qqname$") Then
		Content = IDStr & QQStr & QQNameStr & HandleStr
		AscDescStr = "QQ号码|QQ名称"
		PrintStr = "QQ号码|QQ名称"		
	'【类表】百度知道表
	ElseIf CheckStr(Table, "^baiduzhidao$") Then
		Content = IDStr & BigClassNameStr & SmallClassNameStr & TitleStr & HttpUrlStr & AuthorStr & ReplyNumbStr & ReplyNameStr & AdoptionRateStr & SupportNumbStr & CreateDateStr & BrowseNumbStr & ThroughStr & HandleStr
		AscDescStr = "提交日期|评论数|浏览次数|是否通过" 
		PrintStr = "|"
	'【类表】百度知道相关词表
	ElseIf CheckStr(Table, "^BaiduZhidaoWords$") Then
		Content = IDStr & SearchWordStr & XiangGuanWordStr & HandleStr
		AscDescStr = "|"
		PrintStr = "|"		
	'【类表】数据统计表
	ElseIf CheckStr(Table, "^StatisticalData$") Then
		Content = IDStr & BigClassNameStr & TitleStr & CountNumbStr & HandleStr
		AscDescStr = "总数"
		PrintStr = "|"
	'【类表】软件任务表
	ElseIf CheckStr(Table, "^SoftTask$") Then
		Content = IDStr & BigClassNameStr & EmailStr & ContentStr & UpDateTimeStr & FilePathStr & ThroughStr & HandleStr
		AscDescStr = ""
		PrintStr = "|"		
	'【类表】会员表
	ElseIf CheckStr(Table, "^Member$") Then
		Content = IDStr & UserNameStr & MyMoneyStr & EmailStr & IPStr & HandleStr
		AscDescStr = ""
		PrintStr = "|"
	'【类表】会员日志表
	ElseIf CheckStr(Table, "^UserLog$") Then
		Content = IDStr & BigClassNameStr & ContentStr & IPStr & AddDateTimeStr & HandleStr
		AscDescStr = ""
		PrintStr = "|"
	'【类表】反馈表
	ElseIf CheckStr(Table, "^Feedback$") Then
		Content = IDStr & EmailStr & ContentStr & IPStr & AddDateTimeStr & HandleStr
		AscDescStr = ""
		PrintStr = "|"
	'【类表】关键词表
	ElseIf CheckStr(Table, "^Words$") Then
		Content = IDStr & BigClassNameStr & EmailStr & TitleStr & IPStr & AddDateTimeStr & HandleStr
		AscDescStr = ""
		PrintStr = "|"
	'【类表】问问表
	ElseIf CheckStr(Table, "^Ask$") Then
		Content = IDStr & AskTypeStr & BigClassNameStr & AskIdStr & TitleStr & ReplyNumbStr & ReplyTimeStr & HandleStr
		AscDescStr = "提问时间"
		PrintStr = "|"
	'【类表】搜索引擎收录词
	ElseIf CheckStr(Table, "^SearchEngineInfoList$") Then
		Content = IDStr & BigClassNameStr & WebSiteStr & TitleStr & PageNumberStr & RankingStr & IncludedTimeStr & SearchUrlStr & HandleStr
		AscDescStr = ""
		PrintStr = "|" 
		
	'未配置
	Else
		Call Eerr("【"& Table &"】表","未配置，已停止。注意：【请到 Fun_Table.Asp 页找到 ShowClassName 函数修改】")
		 
	End If
	'循环获得标题列表
	SplStr=Split(Content, ",")
	N=1
	For Each S In SplStr
		If S <> "" Then 
			Splxx=Split(S & "|||", "|")
			TDWidth(N) = Splxx(1)
			C = C & "<td width="""& Splxx(1) &""" align=""center"" bgcolor=""#FFFFFF"" class=""B14"">"& Splxx(0) &"</td>" & vbCrlf
			N = N +1
		End If
	Next
	ShowClassName = C
End Function


'搜索类型 加 多个查询 Or 与 And
Function HandleSearchSQL(SearchType,SearchValue)
	Dim AddSQL,SeectField,SvalueType
	'判断是否为等于搜索
	If Left(SearchValue,1) = """" And Right(SearchValue,1) = """" Then
		SearchValue = Left(SearchValue, Len(SearchValue)-1)
		SearchValue = Right(SearchValue, Len(SearchValue)-1)
		SvalueType = "真"
	End If
	If SearchValue <> "" Then
		If SearchType = "QQ号码搜索" Then
			SeectField = "QQ"
		ElseIf SearchType = "QQ名称搜索" Then
			SeectField = "QQName"
		ElseIf SearchType = "访客列表搜索" Then
			SeectField = "VisitorsList"
		ElseIf SearchType = "大类名称搜索" Then
			SeectField = "BigClassName"
		ElseIf SearchType = "小类名称搜索" Then
			SeectField = "SmallClassName"
		ElseIf SearchType = "图片名称搜索" Then
			SeectField = "PhotoName"
		ElseIf SearchType = "留言者QQ搜索" Then
			SeectField = "MessageQQ"
		ElseIf SearchType = "登录时间搜索" Then
			SeectField = "OnLineTime"
		ElseIf SearchType = "标题搜索" Then
			SeectField = "Title"
		ElseIf SearchType = "添加时间搜索" Then
			SeectField = "AddTime"
		ElseIf SearchType = "搜索值搜索" Then
			SeectField = "SearchValue"
		ElseIf SearchType = "搜索标题搜索" Then
			SeectField = "SearchTitle" 
		ElseIf SearchType = "搜索时间搜索" Then
			SeectField = "SearchTime" 
		ElseIf SearchType = "搜索排序搜索" Then
			SeectField = "SearchSort" 
		ElseIf SearchType = "搜索内容搜索" Then
			SeectField = "SearchContent" 
		ElseIf SearchType = "搜索网址搜索" Then
			SeectField = "SearchHttpUrl" 
		ElseIf SearchType = "网址搜索" Then
			SeectField = "HttpUrl" 
		ElseIf SearchType = "域名搜索" Then
			SeectField = "WebSite" 
		ElseIf SearchType = "备注搜索" Then
			SeectField = "Content" 
		ElseIf SearchType = "上传时间搜索" Then
			SeectField = "UpDateTime"
		ElseIf SearchType = "运行平台搜索" Then
			SeectField = "Windows"
		ElseIf SearchType = "存在大小图搜索" Then
			SeectField = "PhotoShowType"
		ElseIf SearchType = "图片ID搜索" Then
			SeectField = "PhotoID"
		ElseIf SearchType = "公司名称搜索" Then
			SeectField = "Company"
		ElseIf SearchType = "招聘职位搜索" Then
			SeectField = "WorkTitle"
		ElseIf SearchType = "网站标题搜索" Then
			SeectField = "WebTitle"
		ElseIf SearchType = "网站地址搜索" Then
			SeectField = "WebUrl"
		ElseIf SearchType = "星期数搜索" Then
			SeectField = "DayNumb"
		ElseIf SearchType = "单词搜索" Then
			SeectField = "Ci"
		ElseIf SearchType = "英文解释搜索" Then
			SeectField = "Shi"
		ElseIf SearchType = "图片缩览图搜索" Then
			SeectField = "BigPhotoUrl"
		ElseIf SearchType = "QQ性别搜索" Then
			SeectField = "QQSex"
		ElseIf SearchType = "QQ年龄搜索" Then
			SeectField = "QQAge"
		ElseIf SearchType = "破解值搜索" Then
			SeectField = "PoJie"
		ElseIf SearchType = "下载数搜索" Then
			SeectField = "DownNumb"
		ElseIf SearchType = "评论数搜索" Then
			SeectField = "ReplyNumb"
		ElseIf SearchType = "日志类型搜索" Then
			SeectField = "Category"
		ElseIf SearchType = "英文搜索" Then
			SeectField = "EN"
		ElseIf SearchType = "中文搜索" Then
			SeectField = "CN"
		ElseIf SearchType = "排序搜索" Then
			SeectField = "Sort"
		ElseIf SearchType = "文件搜索" Then
			SeectField = "FileName"
		ElseIf SearchType = "文件夹搜索" Then
			SeectField = "FolderName"
		ElseIf SearchType = "目录路径搜索" Then
			SeectField = "DirPath"
		ElseIf SearchType = "播放时间搜索" Then
			SeectField = "PlayTime"
		ElseIf SearchType = "播放次数搜索" Then
			SeectField = "PlayNumb"
		ElseIf SearchType = "发布时间搜索" Then
			SeectField = "PostingsTime"
		ElseIf SearchType = "作者搜索" Then
			SeectField = "Author"
		ElseIf SearchType = "搜索词搜索" Then
			SeectField = "BigClassName"
		ElseIf SearchType = "相关词搜索" Then
			SeectField = "SmallClassName"
		ElseIf SearchType = "采纳率搜索" Then
			SeectField = "AdoptionRate"
		ElseIf SearchType = "支持数搜索" Then
			SeectField = "SupportNumb"
		ElseIf SearchType = "提交日期搜索" Then
			SeectField = "CreateDate"
		ElseIf SearchType = "浏览次数搜索" Then
			SeectField = "BrowseNumb"
		ElseIf SearchType = "是否通过搜索" Then
			SeectField = "Through"
		ElseIf SearchType = "提问时间搜索" Then
			SeectField = "ReplyTime"
		ElseIf SearchType = "问问ID搜索" Then
			SeectField = "AskId"
	 
			
			
		ElseIf SearchType <> "" Then
			Call Echo(SearchType, "搜索不存在")
		End If
		
		
		'获得SQL
		If SeectField<>"" Then
			'为SqlServer数据库 字段二有需要转换为字符类型
			If MDBPath = "" Then
				'添加于2013,10,24  让字段为数字类型转为字符类型，要不然搜索有问题，但不是数字类型不要用，否则就搜索不到全部内容
				Call OpenConn()
				If Conn.Execute("Select "& SeectField &" From ["& Table & "]")(0).Type = 3 Then
					SeectField = "Convert(VarChar, "& SeectField &")"
				End If
			End If 
			AddSql = "Where "& SeectField &" Like '%"& SearchValue &"%'"
			'高级搜索
			If SvalueType = "真" Then
				AddSql = "Where "& SeectField &" = '"& SearchValue &"'"
			End If
	 	End If
		'多个查询 Or 与 And
		AddSql = OrAndSearch(AddSql,SeectField,SearchValue)
	End If
	HandleSearchSQL = AddSQL
End Function

'处理排序
Function HandleAscDesc(AddSql, AscDescType)
	If AscDescType = "" Then
		HandleAscDesc = AddSql
		Exit Function
	End If
	
	If AscDescType = "QQ号码正排序" Then
		AddSql = AddSql & "   Order By Convert(Money, QQ) Asc"	
	ElseIf AscDescType = "QQ号码倒排序" Then	
		AddSql = AddSql & "   Order By Convert(Money, QQ) Desc"
		
	ElseIf AscDescType = "QQ名称正排序" Then
		AddSql = AddSql & "   Order By QQName Asc"	
	ElseIf AscDescType = "QQ名称倒排序" Then	
		AddSql = AddSql & "   Order By QQName Desc"
	
	ElseIf AscDescType = "备注为空" Then
		If InStr(AddSql, "Where ") > 0 Then AddSql = AddSql & "   And" Else AddSql = AddSql & "   Where"
		AddSql = AddSql & " Content Is Null"	
	ElseIf AscDescType = "备注不为空" Then
		If InStr(AddSql, "Where ") > 0 Then AddSql = AddSql & "   And" Else AddSql = AddSql & "   Where"
		AddSql = AddSql & " Content Is Not Null"	
		
	ElseIf AscDescType = "空间权限正排序" Then

		AddSql = AddSql & "   Order By SpacePurview Asc"	
	ElseIf AscDescType = "空间权限倒排序" Then	
		AddSql = AddSql & "   Order By SpacePurview Desc"
		
	ElseIf AscDescType = "登录时间正排序" Then
		AddSql = AddSql & "   Order By OnLineTime Asc"	
	ElseIf AscDescType = "登录时间倒排序" Then	
		AddSql = AddSql & "   Order By OnLineTime Desc"
		
	ElseIf AscDescType = "好友度正排序" Then
		AddSql = AddSql & "   Order By CountNumb"
	ElseIf AscDescType = "好友度倒排序" Then
		AddSql = AddSql & "   Order By CountNumb Desc"
		
	ElseIf AscDescType = "存在数正排序" Then
		AddSql = AddSql & "   Order By QQNumb"
	ElseIf AscDescType = "存在数倒排序" Then
		AddSql = AddSql & "   Order By QQNumb Desc"
		
	ElseIf AscDescType = "好友度＋存在数正排序" Then
		AddSql = AddSql & "   Order By CountNumb+QQNumb"
	ElseIf AscDescType = "好友度＋存在数倒排序" Then
		AddSql = AddSql & "   Order By CountNumb+QQNumb Desc"
		
	ElseIf AscDescType = "评论数正排序" Then
		AddSql = AddSql & "   Order By ReplyNumb Asc"	
	ElseIf AscDescType = "评论数倒排序" Then	
		AddSql = AddSql & "   Order By ReplyNumb Desc"
		
	ElseIf AscDescType = "留言时间正排序" Then
		AddSql = AddSql & "   Order By UpDateTime Asc"	
	ElseIf AscDescType = "留言时间倒排序" Then	
		AddSql = AddSql & "   Order By UpDateTime Desc"
		
	ElseIf AscDescType = "总数正排序" Then
		AddSql = AddSql & "   Order By CountNumb Asc"	
	ElseIf AscDescType = "总数倒排序" Then	
		AddSql = AddSql & "   Order By CountNumb Desc"
		
	ElseIf AscDescType = "添加时间正排序" Then
		AddSql = AddSql & "   Order By AddTime Asc"	
	ElseIf AscDescType = "添加时间倒排序" Then	
		AddSql = AddSql & "   Order By AddTime Desc"
		
	ElseIf AscDescType = "价格正排序" Then
		AddSql = AddSql & "   Order By Price Asc"	
	ElseIf AscDescType = "价格倒排序" Then	
		AddSql = AddSql & "   Order By Price Desc"
		
	ElseIf AscDescType = "运费正排序" Then
		AddSql = AddSql & "   Order By FarePrice Asc"	
	ElseIf AscDescType = "运费倒排序" Then	
		AddSql = AddSql & "   Order By FarePrice Desc"
		
	ElseIf AscDescType = "运费正排序" Then
		AddSql = AddSql & "   Order By FarePrice Asc"	
	ElseIf AscDescType = "运费倒排序" Then	
		AddSql = AddSql & "   Order By FarePrice Desc"
		
	ElseIf AscDescType = "描述值正排序" Then
		AddSql = AddSql & "   Order By Description Asc"	
	ElseIf AscDescType = "描述值倒排序" Then	
		AddSql = AddSql & "   Order By Description Desc"
		
	ElseIf AscDescType = "评论正排序" Then
		AddSql = AddSql & "   Order By ReplyNum Asc"	
	ElseIf AscDescType = "评论倒排序" Then	
		AddSql = AddSql & "   Order By ReplyNum Desc"
		
	ElseIf AscDescType = "上传时间正排序" Then
		AddSql = AddSql & "   Order By UpDateTime Asc"	
	ElseIf AscDescType = "上传时间倒排序" Then	
		AddSql = AddSql & "   Order By UpDateTime Desc"
		
	ElseIf AscDescType = "文件大小正排序" Then
		AddSql = AddSql & "   Order By ZipSize Asc"	
	ElseIf AscDescType = "文件大小倒排序" Then	
		AddSql = AddSql & "   Order By ZipSize Desc"
		
	ElseIf AscDescType = "运行平台正排序" Then
		AddSql = AddSql & "   Order By Windows Asc"	
	ElseIf AscDescType = "运行平台倒排序" Then	
		AddSql = AddSql & "   Order By Windows Desc"
		
	ElseIf AscDescType = "排序正排序" Then
		AddSql = AddSql & "   Order By Sort Asc"	
	ElseIf AscDescType = "排序倒排序" Then	
		AddSql = AddSql & "   Order By Sort Desc"
		
	ElseIf AscDescType = "标题正排序" Then
		AddSql = AddSql & "   Order By Title Asc"	
	ElseIf AscDescType = "标题倒排序" Then	
		AddSql = AddSql & "   Order By Title Desc"		
		
	ElseIf AscDescType = "人气正排序" Then
		AddSql = AddSql & "   Order By RenQi Asc"	
	ElseIf AscDescType = "人气倒排序" Then	
		AddSql = AddSql & "   Order By RenQi Desc"
		
	ElseIf AscDescType = "作者正排序" Then
		AddSql = AddSql & "   Order By Author Asc"	
	ElseIf AscDescType = "作者倒排序" Then	
		AddSql = AddSql & "   Order By Author Desc"
		
		
	ElseIf AscDescType = "单词正排序" Then
		AddSql = AddSql & "   Order By Ci Asc"	
	ElseIf AscDescType = "单词倒排序" Then	
		AddSql = AddSql & "   Order By Ci Desc"
		
	ElseIf AscDescType = "破解值正排序" Then
		AddSql = AddSql & "   Order By PoJie Asc"	
	ElseIf AscDescType = "破解值倒排序" Then	
		AddSql = AddSql & "   Order By PoJie Desc" 
		
	ElseIf AscDescType = "QQ性别正排序" Then
		AddSql = AddSql & "   Order By QQSex Asc"	
	ElseIf AscDescType = "QQ性别倒排序" Then	
		AddSql = AddSql & "   Order By QQSex Desc" 
		
	ElseIf AscDescType = "QQ年龄正排序" Then
		AddSql = AddSql & "   Order By QQAge Asc"	
	ElseIf AscDescType = "QQ年龄倒排序" Then	
		AddSql = AddSql & "   Order By QQAge Desc" 
		
	ElseIf AscDescType = "年龄正排序" Then
		AddSql = AddSql & "   Order By Age Asc"	
	ElseIf AscDescType = "年龄倒排序" Then	
		AddSql = AddSql & "   Order By Age Desc" 
		
		
	ElseIf AscDescType = "下载数正排序" Then
		AddSql = AddSql & "   Order By DownNumb Asc"	
	ElseIf AscDescType = "下载数倒排序" Then	
		AddSql = AddSql & "   Order By DownNumb Desc" 
		
	ElseIf AscDescType = "文件正排序" Then
		AddSql = AddSql & "   Order By FileName Asc"	
	ElseIf AscDescType = "文件倒排序" Then	
		AddSql = AddSql & "   Order By FileName Desc" 
		
	ElseIf AscDescType = "文件夹正排序" Then
		AddSql = AddSql & "   Order By FolderName Asc"	
	ElseIf AscDescType = "文件夹倒排序" Then	
		AddSql = AddSql & "   Order By FolderName Desc" 
		
	ElseIf AscDescType = "目录路径正排序" Then
		AddSql = AddSql & "   Order By DirPath Asc"	
	ElseIf AscDescType = "目录路径倒排序" Then	
		AddSql = AddSql & "   Order By DirPath Desc" 
		
	ElseIf AscDescType = "播放时间正排序" Then
		AddSql = AddSql & "   Order By PlayTime Asc"	
	ElseIf AscDescType = "播放时间倒排序" Then	
		AddSql = AddSql & "   Order By PlayTime Desc" 
		
	ElseIf AscDescType = "播放次数正排序" Then
		AddSql = AddSql & "   Order By PlayNumb Asc"	
	ElseIf AscDescType = "播放次数倒排序" Then	
		AddSql = AddSql & "   Order By PlayNumb Desc" 
		
	ElseIf AscDescType = "发布时间正排序" Then
		AddSql = AddSql & "   Order By PostingsTime Asc"	
	ElseIf AscDescType = "发布时间倒排序" Then	
		AddSql = AddSql & "   Order By PostingsTime Desc" 
		
	ElseIf AscDescType = "点击数正排序" Then
		AddSql = AddSql & "   Order By ClickNumb Asc"	
	ElseIf AscDescType = "点击数倒排序" Then	
		AddSql = AddSql & "   Order By ClickNumb Desc" 
		
	ElseIf AscDescType = "访客数量正排序" Then
		AddSql = AddSql & "   Order By VisitorsNumb Asc"	
	ElseIf AscDescType = "访客数量倒排序" Then	
		AddSql = AddSql & "   Order By VisitorsNumb Desc" 
		
	ElseIf AscDescType = "提交日期正排序" Then
		AddSql = AddSql & "   Order By CreateDate Asc"	
	ElseIf AscDescType = "提交日期倒排序" Then	
		AddSql = AddSql & "   Order By CreateDate Desc" 
		
	ElseIf AscDescType = "浏览次数正排序" Then
		AddSql = AddSql & "   Order By BrowseNumb Asc"	
	ElseIf AscDescType = "浏览次数倒排序" Then	
		AddSql = AddSql & "   Order By BrowseNumb Desc" 
		 
	ElseIf AscDescType = "是否通过正排序" Then
		AddSql = AddSql & "   Order By Through Asc"	
	ElseIf AscDescType = "是否通过倒排序" Then	
		AddSql = AddSql & "   Order By Through Desc" 
		
	ElseIf AscDescType = "提问时间正排序" Then
		AddSql = AddSql & "   Order By ReplyTime Asc"	
	ElseIf AscDescType = "提问时间倒排序" Then	
		AddSql = AddSql & "   Order By ReplyTime Desc" 
		 

		
		
		 
		
	ElseIf AscDescType = "联合查询" Then
		Call LianHeSearch()
	Else
		Call Echo(AscDescType, "排序不存在")
	End If
	
	
	HandleAscDesc = AddSql
End Function
'联合查询 【GuanXiQQ.Asp表】
Function LianHeSearch()
	Dim SearchType, SearchValue
	SearchType = Request("SearchType")
	SearchValue = Request("SearchValue")
	'联合查询，某个QQ不存在，提示操作
	If SearchType = "" And InStr(SearchValue,"|")>0 Then
		CAll Echo("", "联合")
		Dim SplStr,QQIN,QQ,YesNo,C,N
		YesNo = True
		SplStr=Split(SearchValue, "|")
		For Each QQ In SplStr
			If QQ<> "" Then
				If QQIN = "" Then
					QQIN = "'"& QQ &"'"
				Else
					QQIN = QQIN & ",'"& QQ &"'"
				End If
				Rs.Open"Select BigClassName From [GuanXiQQ] Where BigClassName = '"& QQ &"'",Conn,1,1
				If Rs.Eof Then
					C = C & "<div align=center><iframe width='800' height=100 frameborder=0 scrolling=no src='Print.Asp?act=处理分析QQ&QQ="&QQ&"'></iframe></div>"
					N=N+1
					YesNo = False
				End If:Rs.Close
			End If
		Next
		
		If YesNo = False Then
			C = C & "<div align=center style='color:red;'>提示：有未处理的关系QQ【<b>"& N &"</b>】个</div>" 
			Call Rw( C ) 
			Response.End()
		End If
		Conn.Execute("Delete From [GuanXiQQ] Where BigClassName='联合查询'")
		Rs.Open"Select * From [GuanXiQQ] Where BigClassName In("& QQIN &")",Conn,1,1
		While Not Rs.Eof
			'合
			Rsx.Open"Select * From [GuanXiQQ] Where BigClassName='联合查询' And QQ='"& Rs("QQ") &"'",Conn,1,3
			If Rsx.Eof Then 
				Rsx.AddNew
				Rsx("BigClassName") = "联合查询"
				Rsx("SearchList") = SearchValue
				Rsx("SmallClassName") = Rs("BigClassName")
				Rsx("QQ") = Rs("QQ")
				Rsx("QQName") = Rs("QQName")
				Rsx("CountNumb") = Rs("CountNumb") + Rs("QQNumb")
			Else
				Rsx("CountNumb") = Rsx("CountNumb") + Rs("CountNumb") + Rs("QQNumb")
				Rsx("QQNumb") = Rsx("QQNumb") +1
				'累加
				If InStr("|"& Rsx("SmallClassName") &"|", "|"& Rs("QQ") &"|") = False Then
					Rsx("SmallClassName") = Rsx("SmallClassName") & "|" & Rs("BigClassName")
				End If
			End If
			Rsx.Update:Rsx.Close	
		Rs.MoveNext:Wend:Rs.Close
		'跳转到 联合查询
		CAll Eerr("哈哈","哈哈")
		'Response.Redirect("GuanXiQQ.Asp?act=ShowGuanXiQQList&BigClassName=联合查询&AscDescType=两者相加倒排序&CountNumb=1")
		Response.End()
	End If
	'LianHeSearch = AddSql
End Function
'处理重复
Function HandleReset(AddSql, AType, Table)
	If AType = "" Then
		HandleReset = AddSql
		Exit Function
	End If
	
	If AType = "QQ号码重复" Then
		AddSql = "Select QQ From ["& Table &"] Group By QQ Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "QQ名称重复" Then
		AddSql = "Select QQName  From ["& Table &"] Group By QQName Having(Count(*))>" & Cint(Rq("N")) 
	
	ElseIf AType = "空间权限重复" Then
		AddSql = "Select SpacePurview From ["& Table &"] Group By SpacePurview Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "登录时间重复" Then
		AddSql = "Select OnLineTime  From ["& Table &"] Group By OnLineTime Having(Count(*))>" & Cint(Rq("N")) 
	
	ElseIf AType = "大类名称重复" Then
		AddSql = "Select BigClassName From ["& Table &"] Group By BigClassName Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "小类名称重复" Then
		AddSql = "Select SmallClassName  From ["& Table &"] Group By SmallClassName Having(Count(*))>" & Cint(Rq("N")) 
	
	ElseIf AType = "日志类型重复" Then
		AddSql = "Select QQcategory From ["& Table &"] Group By QQcategory Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "标题重复" Then
		AddSql = "Select Title  From ["& Table &"] Group By Title Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "评论数重复" Then
		AddSql = "Select ReplyNumb  From ["& Table &"] Group By ReplyNumb Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "日志ID重复" Then
		AddSql = "Select Blogid  From ["& Table &"] Group By Blogid Having(Count(*))>" & Cint(Rq("N")) 
		
	ElseIf AType = "添加时间重复" Then
		AddSql = "Select AddTime  From ["& Table &"] Group By AddTime Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "域名重复" Then
		AddSql = "Select WebSite  From ["& Table &"] Group By WebSite Having(Count(*))>" & Cint(Rq("N"))  
		
	ElseIf AType = "搜索值重复" Then
		AddSql = "Select SearchValue  From ["& Table &"] Group By SearchValue Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "用户重复" Then
		AddSql = "Select UserName  From ["& Table &"] Group By UserName Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "价格重复" Then
		AddSql = "Select Price  From ["& Table &"] Group By Price Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "运费重复" Then
		AddSql = "Select FarePrice  From ["& Table &"] Group By FarePrice Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "省份重复" Then
		AddSql = "Select Province  From ["& Table &"] Group By Province Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "城市重复" Then
		AddSql = "Select City  From ["& Table &"] Group By City Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "描述值重复" Then
		AddSql = "Select Description  From ["& Table &"] Group By Description Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "评论重复" Then
		AddSql = "Select ReplyNum  From ["& Table &"] Group By ReplyNum Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "QQ备注名重复" Then
		AddSql = "Select QQBeiZhuName  From ["& Table &"] Group By QQBeiZhuName Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "公司名称重复" Then	 
		AddSql = "Select Company  From ["& Table &"] Group By Company Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "招聘职位重复" Then	 
		AddSql = "Select WorkTitle  From ["& Table &"] Group By WorkTitle Having(Count(*))>" & Cint(Rq("N"))  
		
		
	ElseIf AType = "作者重复" Then	 
		AddSql = "Select Author  From ["& Table &"] Group By Author Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "人气重复" Then	 
		AddSql = "Select RenQi  From ["& Table &"] Group By RenQi Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "存在大小图重复" Then	 
		AddSql = "Select PhotoShowType  From ["& Table &"] Group By PhotoShowType Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "网站标题重复" Then	 
		AddSql = "Select WebTitle  From ["& Table &"] Group By WebTitle Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "网站地址重复" Then	 
		AddSql = "Select WebUrl  From ["& Table &"] Group By WebUrl Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "星期数重复" Then	 
		AddSql = "Select DayNumb  From ["& Table &"] Group By DayNumb Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "上传时间重复" Then	 
		AddSql = "Select UpDateTime  From ["& Table &"] Group By UpDateTime Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "QQ年龄重复" Then	 
		AddSql = "Select QQAge  From ["& Table &"] Group By QQAge Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "QQ性别重复" Then	 
		AddSql = "Select QQSex  From ["& Table &"] Group By QQSex Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "年龄重复" Then	 
		AddSql = "Select Age  From ["& Table &"] Group By Age Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "运行平台重复" Then	 
		AddSql = "Select Windows  From ["& Table &"] Group By Windows Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "下载数重复" Then	 
		AddSql = "Select DownNumb  From ["& Table &"] Group By DownNumb Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "文件大小重复" Then	 
		AddSql = "Select ZipSize  From ["& Table &"] Group By ZipSize Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "文件重复" Then	 
		AddSql = "Select FileName  From ["& Table &"] Group By FileName Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "文件夹重复" Then	 
		AddSql = "Select FolderName  From ["& Table &"] Group By FolderName Having(Count(*))>" & Cint(Rq("N"))
	ElseIf AType = "目录路径重复" Then	 
		AddSql = "Select DirPath  From ["& Table &"] Group By DirPath Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "网址重复" Then	 
		AddSql = "Select HttpUrl  From ["& Table &"] Group By HttpUrl Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "播放时间重复" Then	 
		AddSql = "Select PlayTime  From ["& Table &"] Group By PlayTime Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "播放次数重复" Then	 
		AddSql = "Select PlayNumb  From ["& Table &"] Group By PlayNumb Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "发布时间重复" Then	 
		AddSql = "Select PostingsTime  From ["& Table &"] Group By PostingsTime Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "子类名称重复" Then	 
		AddSql = "Select ThreeClassName  From ["& Table &"] Group By ThreeClassName Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "点击数重复" Then	 
		AddSql = "Select ClickNumb  From ["& Table &"] Group By ClickNumb Having(Count(*))>" & Cint(Rq("N"))  
		
	ElseIf AType = "搜索词重复" Then	 
		AddSql = "Select BigClassName  From ["& Table &"] Group By BigClassName Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "分类重复" Then	 
		AddSql = "Select Cate  From ["& Table &"] Group By Cate Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "日志发布时间重复" Then	 
		AddSql = "Select pubTime  From ["& Table &"] Group By pubTime Having(Count(*))>" & Cint(Rq("N"))  

		
		
	'外部MDB
	ElseIf AType = "单词重复" Then	 
		AddSql = "Select Ci  From ["& Table &"] Group By Ci Having(Count(*))>" & Cint(Rq("N"))  
	'外部MDB
	ElseIf AType = "英文解释重复" Then	 
		AddSql = "Select Shi  From ["& Table &"] Group By Shi Having(Count(*))>" & Cint(Rq("N"))  
	'外部MDB
	ElseIf AType = "英文解释重复" Then	 
		AddSql = "Select Shi  From ["& Table &"] Group By Shi Having(Count(*))>" & Cint(Rq("N"))  
		


	Else
		Call Echo(AType, "重复不存在")
	End If
	
	HandleReset = AddSql
End Function

'自定义SQL
Function SelfSql(AddSql, SearchValue, Table)
	Dim SQL 
	 
	SearchValue = RegExp_Replace(SearchValue, "SQL *=", "Sql ") 
	If CheckStr(SearchValue, "SQL ") = True Then
		SQL = Trim(SearchValue)
		SQL = Trim( Right(SQL, Len(SQL) - 4) ) 
		If CheckStr(SQL, "Delete ") = True Then
			SQL = SQL & " &nbsp; <a href=""?act=ClearTable&MDBPath="&Request("MDBPath")&"&SQL="&SQL&""" target=""_blank"" class=""ARed"" onClick=""return Confirm();"">删除记录</a>"
			CAll Eerr("删除很危险，请慎重", SQL) 
		ElseIf CheckStr(SQL, "Update ") = True Then
			CAll Eerr("更新表操作", SQL)
		ElseIf CheckStr(SQL, "Select ") = False Then
			SQL = "Select * From ["& Table &"] " & SQL 
		End If
		AddSql = SQL
	End If
	SelfSql = AddSql
End Function

'替换多个或单个搜索值
Function ReplaceSearchValue(Content, SearchValue)
	Dim SplStr,S
	If InStr(SearchValue," Or ")>0 Then
		SplStr = Split(SearchValue, " Or ")
		For Each S In SplStr
			If S <> "" Then
				Content = Replace(Content, S,"<font color=red><b>" & S & "</b></font>")
			End If
		Next
	ElseIf InStr(SearchValue," And ")>0 Then
		SplStr = Split(SearchValue, " And ")
		For Each S In SplStr
			If S <> "" Then
				Content = Replace(Content, S,"<font color=red><b>" & S & "</b></font>")
			End If
		Next
	ElseIf Content <> "" Then
		Content = Replace(Content, SearchValue,"<font color=red><b>" & SearchValue & "</b></font>")
	End If
	ReplaceSearchValue = Content
End Function

'显示复制内容列表
Sub ShowCopyList(Title,OKNumb,StrList, StartTime)
	If StartTime <> "" Then StartTime = "用时：" & PrintTimeValue(DateDiff("s", StartTime, Now()) ) & "，&nbsp;"
	Call Echo("【"& Title &"】", "共有【"& OKNumb &"】条，&nbsp;" & StartTime & "<a href=javascript:CopyTEXT('CopyContent');>复制</a>")
	CAll Rw("<div style='margin-left:220px;'><textarea name='CopyContent' id='CopyContent' style='width:800px;height:400px;'>"& StrList &"</textarea></div>")
End Sub
%>










