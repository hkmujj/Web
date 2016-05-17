<% 
'注：代码要重新写
'产品展示列表 (2013,12,19)
'ActionType为动作类型
'wd为搜索内容
Function ProductList(Did, Sid, Tid, ActionType, ClassName, wd, ByVal Page)
	'call echo("Page",Page)
    Dim AddSql, Sql, I, J, X, nCount, nPageSize, MaxPage, C, Url, Title, ShowTitle, Target, PageStr, SearchEcho, SplStr, SplType
    Dim HomePage, UPPage, NextPage, EndPage, iStartPage, iEndPage, sHtml 
    Dim Action, Query, A, Temp, DefaultUrl, CutNumb, ImgUrl, ClassStr 
    Dim Content, DialogStart, DialogSplit, DialogEnd, S, ImgFile, ImgStr, AStr,  AUrl, Price
	Dim ConfigContent,ConfigValue,StartStr,EndStr,TempStartStr,TempEndStr,TagName
	Dim ProductModuleTypeList,NewsModuleTypeList,DownloadModuleTypeList		'产品,新闻,下载,模块类型列表
	Dim Rs
	Set Rs = CreateObject("Adodb.RecordSet")
	Dim PageParam				'翻页参数 
	Dim ModuleName				'模块名称
	Dim TitleLen				'标题长度
	Dim NewFile					'新文件
	Dim ArticleDescription		'文章描述(简要简介)
	Dim ModI					'I求余
	Dim AImg					'A图片链接
	
	NewFile=CreateHtmlPath	
	'功能：/Index.Html 变成这样 /Index[id].Html                     20141230
	If InStr(CreateHtmlPath,".")>0 Then
		NewFile=Mid(CreateHtmlPath,1,InStr(CreateHtmlPath,".")-1)	
		NewFile = Host() & NewFile & "[id]" & Mid(CreateHtmlPath,InStr(CreateHtmlPath,"."))
	End If	
	NewFile = HandleHttpUrl(NewFile)
	
	
	'Call Eerr(NewFile, CreateHtmlPath)
	
	
	'Call Echo("Did",Did)
	'处理页参数
	If Did <> "" Then
		PageParam=PageParam & "&ProDid=" & Did
	End If
	If Sid <> "" Then
		PageParam=PageParam & "&ProSid=" & Sid
	End If
	If Tid <> "" Then
		PageParam=PageParam & "&ProTid=" & Tid
	End If
	If Did="" Then Did = "产品"
    '让网站导航点击也能进入 产品类与新闻类
    If WebActionType = "Nav" Then
		'这个是为了让产品展示或新闻展示显示出来，因为导航也可以显示图片与新闻类的
        Did = WebNavDid : Sid = WebNavSid       								
    End If
	'为搜索
    If ActionType = "Search" Then
        AddSql = ""
        If InStr(ClassName, "{分}") > 0 Then
            SplStr = Split(ClassName, "{分}") 
            If SplStr(2) <> "" Then
                AddSql = " And BigClassName='" & SplStr(0) & "' And SmallClassName='" & SplStr(1) & "' And ThreeClassName='" & SplStr(2) & "'" 
            ElseIf SplStr(1) <> "" Then
                AddSql = " And BigClassName='" & SplStr(0) & "' And SmallClassName='" & SplStr(1) & "'" 
            ElseIf SplStr(0) <> "" Then
                AddSql = " And BigClassName='" & SplStr(0) & "'" 
            End If
        End If 
        Sql = "Select * From [Product] Where Title Like '%" & wd & "%'" & AddSql 
        Action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME") 
        Query = Split(Request.ServerVariables("QUERY_STRING"), "&") 
        For Each X In Query
            A = Split(X, "=") 
            If StrComp(A(0), "page", vbTextCompare) <> 0 Then
                Temp = Temp & A(0) & "=" & A(1) & "&" 
            End If
        Next 
        DefaultUrl = Action & "?" & Temp & "page=" 
    ElseIf Tid <> "" Then
        Sql = "Select * From [Product] Where BigClassName='" & Did & "' And SmallClassName='" & Sid & "' And ThreeClassName='" & Tid & "'" 
    ElseIf Sid <> "" Then
        Sql = "Select * From [Product] Where BigClassName='" & Did & "' And SmallClassName='" & Sid & "'" 
    Else
        Sql = "Select * From [Product] Where BigClassName='" & Did & "'" 
    End If 

    '判断是否有图片，如果没有图片自动为新闻类
    WebNavType = ""

    SearchEcho = "对不起，没有找到目标！关键词尽量要精简！" 

	Dim TemplateSourceID				'模板资源ID
	TemplateSourceID = GetArticleListStyle(Did,Sid,Tid)
	Content = ReadArticleListStyleSource(TemplateSourceID)
	
	
	'每页显示条数
	nPageSize = GetArticlePageSize(Did,Sid,Tid)
	
	'测试区
	
	'处理配置信息
	dim pageConfigContent,dispalyPageInfo,temp_nPageSize
	startStr="<!--#config "
	endStr="#-->"
	pageConfigContent=GetStrCut(Content,startStr,endStr,2)
	dispalyPageInfo=lcase(RParam(pageConfigContent,"dispalyPageInfo"))
	temp_nPageSize=HandleNumber(RParam(pageConfigContent,"nPageSize"))
	if temp_nPageSize<>"" then nPageSize=cint(temp_nPageSize)
	'call echo("dispalyPageInfo",dispalyPageInfo)
	'call echo("nPageSize",nPageSize)
	
	
	If IsNul(TitleLen)=True Then TitleLen = 32								'标题长度为空则显示默认
	If IsNul(nPageSize)=True Then nPageSize = 3								'每页显示条数为空则显示默认
	'获得模板分割字段
    Call GetTemplateSplit(Content, DialogStart, DialogSplit, DialogEnd)
	 
	If DialogSplit="" Then
		ProductList =  "出错了2，S=" & S & "<hr>，里面没有Did=【" & Did & "】这个参数。" & TemplateSourceID
		Exit Function
	End If 
	
	'Call Echo("AddSql",AddSql)
	Sql = GetWhereAnd(Sql,AddSql)
	Sql = GetWhereAnd(Sql,GetProductListSort(Did,Sid,Tid))		'获得产品列表排序 20150124
	'Call Echo("SQl",SQl)
    Rs.Open Sql, Conn, 1, 1
    If Not Rs.EOF Then 
        SearchEcho = "搜索结果：符合查询条件的记录 <strong>" & Rs.RecordCount & " </strong>条，以下为您列出" 
        Rs.PageSize = nPageSize         '每页记录条数 可自定义
        nCount = Rs.RecordCount                                                         '记录总数
        nPageSize = Rs.PageSize 
        MaxPage = Rs.PageCount 
        If Not IsNumeric(Page) Or Page = "" Then
            Page = 1 
        Else
            Page = CInt(Page) 
        End If 
        If Page < 1 Then
            Page = 1 
        ElseIf Page > MaxPage Then
            Page = MaxPage 
        End If 
        Rs.AbsolutePage = Page 
        If Page = MaxPage Then
            X = nCount - (MaxPage - 1) * nPageSize 
        Else
            X = nPageSize 
        End If 
        For I = 1 To X	
			
			Url = Rs("FileName")
			If CheckRemoteUrl(Url)=False Then		
				Url = GetHandleUrl(Rs("FileName")) 
				Call WebDebug(Url, "act=CreateArticle&ID=" & Rs("Id"))                      '调试
			End If
			
			ArticleDescription = Rs("ArticleDescription")								     '简要说明		
			If ArticleDescription="#" Or ArticleDescription="null" Or ArticleDescription="empty" Then
				ArticleDescription = ""
			ElseIf IsNul(ArticleDescription) Then
				ArticleDescription = Rs("Content") 
			End If	
			
            ImgFile = IIF(IsNul(Rs("SmallFiles")), Rs("BigFiles"), Rs("SmallFiles"))
			ImgFile = ReplaceGlobleLable(ImgFile)
			
			If ImgFile="" Then
				ImgFile = "/UploadFiles/NoImg.jpg"
			End If
			'Call Echo("ImgFile",ImgFile)
			
            ImgStr = "<img src='" & ImgFile & "'" & SetHtmlParam(Action, "imgtitle|imgalt|imgid|imgclass|imgwidth|imgheight") & ">" 
			AImg = ImgSrc(ImgFile, Rs("Title"), Rs("Target"))
			
            AStr = "<a src='" & Rs("BigFiles") & "'" & SetHtmlParam(Action, "imgtitle|imgalt|imgid|imgclass|imgwidth|imgheight") & ">" 
            ShowTitle = CutStr(Rs("Title"), TitleLen, "... ") 
            ShowTitle = FontColorFontB(ShowTitle, Rs("FontB"), Rs("FontColor")) 
            AStr = "<a " & AHref(Url, Rs("Title"), Rs("Target")) & SetHtmlParam(Action, "aclass") & ">" & ShowTitle & "</a>" & vbCrLf 
            AUrl = AHref(Url, Rs("Title"), Rs("Target")) & SetHtmlParam(Action, "aclass") 
            If wd <> "" Then
                ShowTitle = Replace(ShowTitle, wd, "<font color=red><b>" & wd & "</b></font>") 
            End If
			
			
			
			StartStr="[list-"& I &"]" : EndStr="[/list-"& I &"]"			
			'例[list-mod2]  [/list-mod2]    20150112
			For ModI=6 To 2 Step -1 
				If InStr(DialogSplit,StartStr)=False And I Mod ModI =0 Then 
					StartStr="[list-mod"& ModI &"]" : EndStr="[/list-mod"& ModI &"]"
					If InStr(DialogSplit,StartStr)>0 Then
						'Call Echo("ModI=" & ModI & ",I=" & I,StartStr)
						Exit For
					End If
				End If
			Next 			
			'没有则用默认
			If InStr(DialogSplit,StartStr)=False Then
				StartStr="[list]" : EndStr="[/list]"
			End If
			If InStr(DialogSplit,StartStr) And InStr(DialogSplit,EndStr) Then
				S = StrCut(DialogSplit,StartStr,EndStr,2)
			Else
				S = DialogSplit
			End If
			
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
				S = ReplaceValueParam(S,"img",ImgFile)
				S = ReplaceValueParam(S,"url",url)
				S = ReplaceValueParam(S,"href",url)
				S = ReplaceValueParam(S,"aurl",AUrl)
				S = ReplaceValueParam(S,"astr",AStr)
				S = ReplaceValueParam(S,"aimg",AImg)				'20150112			
				S = ReplaceValueParam(S,"description",Rs("WebDescription"))
				S = ReplaceValueParam(S,"content",Rs("Content"))
				S = ReplaceValueParam(S,"price",Rs("Price"))
				S = ReplaceValueParam(S,"author",Rs("Author"))
				S = ReplaceValueParam(S,"articledescription",ArticleDescription)     '摘要(文章简要描述)20141225
								
			Next

			'添加在线修改
			S = DisplayOnlineEditDialog(""& adminDir &"Product.Asp?act=ShowEditProduct&id=" & Rs("Id") & "&n=" & GetRnd(11), S)
            C = C & S 

        Rs.MoveNext : Next 
    End If : Rs.Close

	
	
    If Page <= 1 Then
        HomePage = "<span class='PageBox'>首页</span>" 
        UPPage = "<span class='PageBox'>上一页</span>" 
    Else
        If DefaultUrl <> "" Then		
			Url = DefaultUrl & "1"
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=1")                      '调试			
            HomePage = "<a href='" & Url & "'>首页</a>" 
			
			Url = DefaultUrl &(Page - 1)
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & (Page - 1))                      '调试	
            UPPage = "<a href='" & Url & "'>上一页</a>" 
			
        Else
			'Url = CreateHtmlPath
			Url = Replace(NewFile,"[id]", "")
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=1")                      '调试		
            HomePage = "<a href='" & Url & "'>首页</a>" 
			
			'Url = Replace(CreateHtmlPath, ".Html", Page - 1 & ".Html")			
			Url = Replace(NewFile,"[id]", Page-1)
			'链接到首页
			If Page-1=1 Then
				Url = Replace(NewFile,"[id]", "")
			End If
			
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & Page - 1)                      '调试		
            UPPage = "<a href='" & Url & "'>上一页</a>" 
        End If 
    End If 
    '下一页 与 尾页
    If Page >= MaxPage Then
        NextPage = "<span class='PageBox'>下一页</span>" 
        EndPage = "<span class='PageBox'>尾页</span>" 
    Else
        If DefaultUrl <> "" Then
			Url = DefaultUrl &(Page + 1)
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & (Page+1))                      '调试
            NextPage = "<a href='" & Url & "'>下一页</a>" 
			
			Url = DefaultUrl & MaxPage
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & MaxPage)                      '调试
            EndPage = "<a href='" & Url & "'>尾页</a>" 
        Else
			'Url = Replace(CreateHtmlPath, ".Html", Page + 1 & ".Html")
			Url = Replace(NewFile,"[id]", Page+1)
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & (Page+1))                      '调试
            NextPage = "<a href='" & Url & "'>下一页</a>" 
			
			'Url = Replace(CreateHtmlPath, ".Html", MaxPage & ".Html")
			Url = Replace(NewFile,"[id]", MaxPage)
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & MaxPage)                      '调试
            EndPage = "<a href='" & Url & "'>尾页</a>" 
        End If 
    End If 
    '循环显示页
    iStartPage = Page - 2
    iEndPage = Page + 2 
    If iStartPage <= 0 Then
        iEndPage = iEndPage + iStartPage * - 1 + 1 
        iStartPage = 1 
    End If 
    If iEndPage > MaxPage Then
        iStartPage = iStartPage + (MaxPage - iEndPage) 
        iEndPage = MaxPage 
    End If 
    'Call Echo(iStartPage,iEndPage)
    For I = iStartPage To iEndPage
        If I >= 1 And I <= MaxPage Then
            If I = Page Then
                PageStr = PageStr & "<span class='PageBox'>" & I & "</span>" 
            Else
                If DefaultUrl <> "" Then
					Url = DefaultUrl & I & "'" & ClassStr
					Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & I)                      '调试	
                    PageStr = PageStr & "&nbsp;<a href='" & Url & ">" & I & "</a>" 
                Else
                    'Url = IIF(I <> 1, Replace(CreateHtmlPath, ".Html", I & ".Html"), CreateHtmlPath) 
					Url = Replace(NewFile,"[id]", IIF(I<>1,I,""))
					Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & I)                      '调试	
                    PageStr = PageStr & "&nbsp;<a href='" & Url & "'" & ClassStr & ">" & I & "</a>" 
                End If 

            End If 
        End If 
    Next 
    C = DialogStart & C & DialogEnd                                                 '面板前面与面板后面加上
	'是否显示翻页信息
	if dispalyPageInfo<>"false" and dispalyPageInfo<>"0" then
	    C = C & "<div class='t_page ColorLink'>总数：" & nCount & "条&nbsp;&nbsp;当前页数：<span class='FontRed'>" & Page & "</span>/" & MaxPage & "&nbsp;" & HomePage & "&nbsp;&nbsp;" & UPPage & "&nbsp;" & PageStr & "&nbsp;&nbsp;" & NextPage & "&nbsp;&nbsp;" & EndPage & "</div>" 
	end if

    CreatePage = MaxPage                                                            '需要生成的页数
    '显示搜索回显内容
    If wd <> "" Then
        C = "<div class='SearchEcho'>" & SearchEcho & "</div>" & vbCrLf & C 
    End If 
    ProductList = C  
End Function
'获得产品列表排序 20150124
Function GetProductListSort(Did,Sid,Tid)
	Dim AddSortType									'添加排序类型 20150124
	Dim OrderSortType								'顺序排序类型 20150124
	Dim TimeSortType								'时间排序类型 20150124
	Dim Sql
	
	'产品大类
	If Did<>"" Then
		Rs.Open"Select * From [BigClass] Where BigClassName='"& Did &"'",Conn,1,1
		If Not Rs.Eof Then
			AddSortType = Rs("AddSortType")
			OrderSortType = Rs("OrderSortType")
			TimeSortType = Rs("TimeSortType")
		End If:Rs.Close
	End If
	'产品小类
	If Sid<>"" Then
		Rs.Open"Select * From [SmallClass] Where BigClassName='"& Did &"' And SmallClassName='"& Sid &"'",Conn,1,1
		If Not Rs.Eof Then
			'添加排序类型
			If Rs("AddSortType")="无" Then
				AddSortType=""
			ElseIf Rs("AddSortType")="" Then
				AddSortType = Rs("AddSortType")
			End If
			'顺序排序类型
			If Rs("OrderSortType")="无" Then
				OrderSortType=""
			ElseIf Rs("OrderSortType")="" Then
				OrderSortType = Rs("OrderSortType")
			End If
			'时间排序类型
			If Rs("TimeSortType")="无" Then
				TimeSortType=""
			ElseIf Rs("TimeSortType")="" Then
				TimeSortType = Rs("TimeSortType")
			End If
		End If:Rs.Close
	End If
	'产品子类
	If Sid<>"" Then
		Rs.Open"Select * From [ThreeClass] Where BigClassName='"& Did &"' And SmallClassName='"& Sid &"' And ThreeClassName='"& Tid &"'",Conn,1,1
		If Not Rs.Eof Then
			'添加排序类型
			If Rs("AddSortType")="无" Then
				AddSortType=""
			ElseIf Rs("AddSortType")="" Then
				AddSortType = Rs("AddSortType")
			End If
			'顺序排序类型
			If Rs("OrderSortType")="无" Then
				OrderSortType=""
			ElseIf Rs("OrderSortType")="" Then
				OrderSortType = Rs("OrderSortType")
			End If
			'时间排序类型
			If Rs("TimeSortType")="无" Then
				TimeSortType=""
			ElseIf Rs("TimeSortType")="" Then
				TimeSortType = Rs("TimeSortType")
			End If
		End If:Rs.Close
	End If
	
	If AddSortType = "添加正排序" Then
		If Sql<>"" Then Sql = Sql & ","
		Sql=Sql & "Id Asc"
	ElseIf AddSortType = "添加倒排序" Then
		If Sql<>"" Then Sql = Sql & ","
		Sql=Sql & "Id  Desc"	
	End If
	
	If OrderSortType = "顺序正排序" Then
		If Sql<>"" Then Sql = Sql & ","
		Sql=Sql & "Sort Asc"
	ElseIf OrderSortType = "顺序倒排序" Then
		If Sql<>"" Then Sql = Sql & ","
		Sql=Sql & "Sort Desc"	
	End If
	
	If TimeSortType = "时间正排序" Then
		If Sql<>"" Then Sql = Sql & ","
		Sql=Sql & "AddDateTime Asc"
	ElseIf TimeSortType = "时间倒排序" Then
		If Sql<>"" Then Sql = Sql & ","
		Sql=Sql & "AddDateTime Desc"	
	End If
	If Sql<>"" Then Sql = " Order By " & Sql
	
	GetProductListSort = Sql
	If 1=2 Then
		Call Echo("Did",Did)
		Call Echo("Sid",Sid)
		Call Echo("Tid",Tid)
		Call Echo("AddSortType",AddSortType)
		Call Echo("OrderSortType",OrderSortType)
		Call Echo("TimeSortType",TimeSortType)
		Call Echo("Sql",Sql)
	End If
End Function





%> 











