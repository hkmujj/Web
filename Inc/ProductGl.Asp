<% 
'ע������Ҫ����д
'��Ʒչʾ�б� (2013,12,19)
'ActionTypeΪ��������
'wdΪ��������
Function ProductList(Did, Sid, Tid, ActionType, ClassName, wd, ByVal Page)
	'call echo("Page",Page)
    Dim AddSql, Sql, I, J, X, nCount, nPageSize, MaxPage, C, Url, Title, ShowTitle, Target, PageStr, SearchEcho, SplStr, SplType
    Dim HomePage, UPPage, NextPage, EndPage, iStartPage, iEndPage, sHtml 
    Dim Action, Query, A, Temp, DefaultUrl, CutNumb, ImgUrl, ClassStr 
    Dim Content, DialogStart, DialogSplit, DialogEnd, S, ImgFile, ImgStr, AStr,  AUrl, Price
	Dim ConfigContent,ConfigValue,StartStr,EndStr,TempStartStr,TempEndStr,TagName
	Dim ProductModuleTypeList,NewsModuleTypeList,DownloadModuleTypeList		'��Ʒ,����,����,ģ�������б�
	Dim Rs
	Set Rs = CreateObject("Adodb.RecordSet")
	Dim PageParam				'��ҳ���� 
	Dim ModuleName				'ģ������
	Dim TitleLen				'���ⳤ��
	Dim NewFile					'���ļ�
	Dim ArticleDescription		'��������(��Ҫ���)
	Dim ModI					'I����
	Dim AImg					'AͼƬ����
	
	NewFile=CreateHtmlPath	
	'���ܣ�/Index.Html ������� /Index[id].Html                     20141230
	If InStr(CreateHtmlPath,".")>0 Then
		NewFile=Mid(CreateHtmlPath,1,InStr(CreateHtmlPath,".")-1)	
		NewFile = Host() & NewFile & "[id]" & Mid(CreateHtmlPath,InStr(CreateHtmlPath,"."))
	End If	
	NewFile = HandleHttpUrl(NewFile)
	
	
	'Call Eerr(NewFile, CreateHtmlPath)
	
	
	'Call Echo("Did",Did)
	'����ҳ����
	If Did <> "" Then
		PageParam=PageParam & "&ProDid=" & Did
	End If
	If Sid <> "" Then
		PageParam=PageParam & "&ProSid=" & Sid
	End If
	If Tid <> "" Then
		PageParam=PageParam & "&ProTid=" & Tid
	End If
	If Did="" Then Did = "��Ʒ"
    '����վ�������Ҳ�ܽ��� ��Ʒ����������
    If WebActionType = "Nav" Then
		'�����Ϊ���ò�Ʒչʾ������չʾ��ʾ��������Ϊ����Ҳ������ʾͼƬ���������
        Did = WebNavDid : Sid = WebNavSid       								
    End If
	'Ϊ����
    If ActionType = "Search" Then
        AddSql = ""
        If InStr(ClassName, "{��}") > 0 Then
            SplStr = Split(ClassName, "{��}") 
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

    '�ж��Ƿ���ͼƬ�����û��ͼƬ�Զ�Ϊ������
    WebNavType = ""

    SearchEcho = "�Բ���û���ҵ�Ŀ�꣡�ؼ��ʾ���Ҫ����" 

	Dim TemplateSourceID				'ģ����ԴID
	TemplateSourceID = GetArticleListStyle(Did,Sid,Tid)
	Content = ReadArticleListStyleSource(TemplateSourceID)
	
	
	'ÿҳ��ʾ����
	nPageSize = GetArticlePageSize(Did,Sid,Tid)
	
	'������
	
	'����������Ϣ
	dim pageConfigContent,dispalyPageInfo,temp_nPageSize
	startStr="<!--#config "
	endStr="#-->"
	pageConfigContent=GetStrCut(Content,startStr,endStr,2)
	dispalyPageInfo=lcase(RParam(pageConfigContent,"dispalyPageInfo"))
	temp_nPageSize=HandleNumber(RParam(pageConfigContent,"nPageSize"))
	if temp_nPageSize<>"" then nPageSize=cint(temp_nPageSize)
	'call echo("dispalyPageInfo",dispalyPageInfo)
	'call echo("nPageSize",nPageSize)
	
	
	If IsNul(TitleLen)=True Then TitleLen = 32								'���ⳤ��Ϊ������ʾĬ��
	If IsNul(nPageSize)=True Then nPageSize = 3								'ÿҳ��ʾ����Ϊ������ʾĬ��
	'���ģ��ָ��ֶ�
    Call GetTemplateSplit(Content, DialogStart, DialogSplit, DialogEnd)
	 
	If DialogSplit="" Then
		ProductList =  "������2��S=" & S & "<hr>������û��Did=��" & Did & "�����������" & TemplateSourceID
		Exit Function
	End If 
	
	'Call Echo("AddSql",AddSql)
	Sql = GetWhereAnd(Sql,AddSql)
	Sql = GetWhereAnd(Sql,GetProductListSort(Did,Sid,Tid))		'��ò�Ʒ�б����� 20150124
	'Call Echo("SQl",SQl)
    Rs.Open Sql, Conn, 1, 1
    If Not Rs.EOF Then 
        SearchEcho = "������������ϲ�ѯ�����ļ�¼ <strong>" & Rs.RecordCount & " </strong>��������Ϊ���г�" 
        Rs.PageSize = nPageSize         'ÿҳ��¼���� ���Զ���
        nCount = Rs.RecordCount                                                         '��¼����
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
				Call WebDebug(Url, "act=CreateArticle&ID=" & Rs("Id"))                      '����
			End If
			
			ArticleDescription = Rs("ArticleDescription")								     '��Ҫ˵��		
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
			'��[list-mod2]  [/list-mod2]    20150112
			For ModI=6 To 2 Step -1 
				If InStr(DialogSplit,StartStr)=False And I Mod ModI =0 Then 
					StartStr="[list-mod"& ModI &"]" : EndStr="[/list-mod"& ModI &"]"
					If InStr(DialogSplit,StartStr)>0 Then
						'Call Echo("ModI=" & ModI & ",I=" & I,StartStr)
						Exit For
					End If
				End If
			Next 			
			'û������Ĭ��
			If InStr(DialogSplit,StartStr)=False Then
				StartStr="[list]" : EndStr="[/list]"
			End If
			If InStr(DialogSplit,StartStr) And InStr(DialogSplit,EndStr) Then
				S = StrCut(DialogSplit,StartStr,EndStr,2)
			Else
				S = DialogSplit
			End If
			
			'��i���⴦��������֮ǰ����
			S = Replace(S,"[$i$]",I) 
			S = Replace(S,"[$i-1$]",I-1) 
			'ִ��6�Σ���ȡ����ȫ����ǩ			
			For J=1 To 6
				S = ReplaceValueParam(S,"ni",I)								'����Ϊi����Ϊi����imgurl��ͻ [$i$]
				S = ReplaceValueParam(S,"���-1",I-1)								'����Ϊi����Ϊi����imgurl��ͻ [$i$]
				S = ReplaceValueParam(S,"���",I)								'����Ϊi����Ϊi����imgurl��ͻ [$i$]
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
				S = ReplaceValueParam(S,"articledescription",ArticleDescription)     'ժҪ(���¼�Ҫ����)20141225
								
			Next

			'���������޸�
			S = DisplayOnlineEditDialog(""& adminDir &"Product.Asp?act=ShowEditProduct&id=" & Rs("Id") & "&n=" & GetRnd(11), S)
            C = C & S 

        Rs.MoveNext : Next 
    End If : Rs.Close

	
	
    If Page <= 1 Then
        HomePage = "<span class='PageBox'>��ҳ</span>" 
        UPPage = "<span class='PageBox'>��һҳ</span>" 
    Else
        If DefaultUrl <> "" Then		
			Url = DefaultUrl & "1"
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=1")                      '����			
            HomePage = "<a href='" & Url & "'>��ҳ</a>" 
			
			Url = DefaultUrl &(Page - 1)
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & (Page - 1))                      '����	
            UPPage = "<a href='" & Url & "'>��һҳ</a>" 
			
        Else
			'Url = CreateHtmlPath
			Url = Replace(NewFile,"[id]", "")
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=1")                      '����		
            HomePage = "<a href='" & Url & "'>��ҳ</a>" 
			
			'Url = Replace(CreateHtmlPath, ".Html", Page - 1 & ".Html")			
			Url = Replace(NewFile,"[id]", Page-1)
			'���ӵ���ҳ
			If Page-1=1 Then
				Url = Replace(NewFile,"[id]", "")
			End If
			
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & Page - 1)                      '����		
            UPPage = "<a href='" & Url & "'>��һҳ</a>" 
        End If 
    End If 
    '��һҳ �� βҳ
    If Page >= MaxPage Then
        NextPage = "<span class='PageBox'>��һҳ</span>" 
        EndPage = "<span class='PageBox'>βҳ</span>" 
    Else
        If DefaultUrl <> "" Then
			Url = DefaultUrl &(Page + 1)
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & (Page+1))                      '����
            NextPage = "<a href='" & Url & "'>��һҳ</a>" 
			
			Url = DefaultUrl & MaxPage
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & MaxPage)                      '����
            EndPage = "<a href='" & Url & "'>βҳ</a>" 
        Else
			'Url = Replace(CreateHtmlPath, ".Html", Page + 1 & ".Html")
			Url = Replace(NewFile,"[id]", Page+1)
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & (Page+1))                      '����
            NextPage = "<a href='" & Url & "'>��һҳ</a>" 
			
			'Url = Replace(CreateHtmlPath, ".Html", MaxPage & ".Html")
			Url = Replace(NewFile,"[id]", MaxPage)
			Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & MaxPage)                      '����
            EndPage = "<a href='" & Url & "'>βҳ</a>" 
        End If 
    End If 
    'ѭ����ʾҳ
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
					Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & I)                      '����	
                    PageStr = PageStr & "&nbsp;<a href='" & Url & ">" & I & "</a>" 
                Else
                    'Url = IIF(I <> 1, Replace(CreateHtmlPath, ".Html", I & ".Html"), CreateHtmlPath) 
					Url = Replace(NewFile,"[id]", IIF(I<>1,I,""))
					Call WebDebug(Url, "act=CreateClass"& PageParam &"&Page=" & I)                      '����	
                    PageStr = PageStr & "&nbsp;<a href='" & Url & "'" & ClassStr & ">" & I & "</a>" 
                End If 

            End If 
        End If 
    Next 
    C = DialogStart & C & DialogEnd                                                 '���ǰ�������������
	'�Ƿ���ʾ��ҳ��Ϣ
	if dispalyPageInfo<>"false" and dispalyPageInfo<>"0" then
	    C = C & "<div class='t_page ColorLink'>������" & nCount & "��&nbsp;&nbsp;��ǰҳ����<span class='FontRed'>" & Page & "</span>/" & MaxPage & "&nbsp;" & HomePage & "&nbsp;&nbsp;" & UPPage & "&nbsp;" & PageStr & "&nbsp;&nbsp;" & NextPage & "&nbsp;&nbsp;" & EndPage & "</div>" 
	end if

    CreatePage = MaxPage                                                            '��Ҫ���ɵ�ҳ��
    '��ʾ������������
    If wd <> "" Then
        C = "<div class='SearchEcho'>" & SearchEcho & "</div>" & vbCrLf & C 
    End If 
    ProductList = C  
End Function
'��ò�Ʒ�б����� 20150124
Function GetProductListSort(Did,Sid,Tid)
	Dim AddSortType									'������������ 20150124
	Dim OrderSortType								'˳���������� 20150124
	Dim TimeSortType								'ʱ���������� 20150124
	Dim Sql
	
	'��Ʒ����
	If Did<>"" Then
		Rs.Open"Select * From [BigClass] Where BigClassName='"& Did &"'",Conn,1,1
		If Not Rs.Eof Then
			AddSortType = Rs("AddSortType")
			OrderSortType = Rs("OrderSortType")
			TimeSortType = Rs("TimeSortType")
		End If:Rs.Close
	End If
	'��ƷС��
	If Sid<>"" Then
		Rs.Open"Select * From [SmallClass] Where BigClassName='"& Did &"' And SmallClassName='"& Sid &"'",Conn,1,1
		If Not Rs.Eof Then
			'������������
			If Rs("AddSortType")="��" Then
				AddSortType=""
			ElseIf Rs("AddSortType")="" Then
				AddSortType = Rs("AddSortType")
			End If
			'˳����������
			If Rs("OrderSortType")="��" Then
				OrderSortType=""
			ElseIf Rs("OrderSortType")="" Then
				OrderSortType = Rs("OrderSortType")
			End If
			'ʱ����������
			If Rs("TimeSortType")="��" Then
				TimeSortType=""
			ElseIf Rs("TimeSortType")="" Then
				TimeSortType = Rs("TimeSortType")
			End If
		End If:Rs.Close
	End If
	'��Ʒ����
	If Sid<>"" Then
		Rs.Open"Select * From [ThreeClass] Where BigClassName='"& Did &"' And SmallClassName='"& Sid &"' And ThreeClassName='"& Tid &"'",Conn,1,1
		If Not Rs.Eof Then
			'������������
			If Rs("AddSortType")="��" Then
				AddSortType=""
			ElseIf Rs("AddSortType")="" Then
				AddSortType = Rs("AddSortType")
			End If
			'˳����������
			If Rs("OrderSortType")="��" Then
				OrderSortType=""
			ElseIf Rs("OrderSortType")="" Then
				OrderSortType = Rs("OrderSortType")
			End If
			'ʱ����������
			If Rs("TimeSortType")="��" Then
				TimeSortType=""
			ElseIf Rs("TimeSortType")="" Then
				TimeSortType = Rs("TimeSortType")
			End If
		End If:Rs.Close
	End If
	
	If AddSortType = "����������" Then
		If Sql<>"" Then Sql = Sql & ","
		Sql=Sql & "Id Asc"
	ElseIf AddSortType = "���ӵ�����" Then
		If Sql<>"" Then Sql = Sql & ","
		Sql=Sql & "Id  Desc"	
	End If
	
	If OrderSortType = "˳��������" Then
		If Sql<>"" Then Sql = Sql & ","
		Sql=Sql & "Sort Asc"
	ElseIf OrderSortType = "˳������" Then
		If Sql<>"" Then Sql = Sql & ","
		Sql=Sql & "Sort Desc"	
	End If
	
	If TimeSortType = "ʱ��������" Then
		If Sql<>"" Then Sql = Sql & ","
		Sql=Sql & "AddDateTime Asc"
	ElseIf TimeSortType = "ʱ�䵹����" Then
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










