<%
'======================================
'���ݿ������
'======================================
 

'��������
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
	IDStr = "ѡ��|42|,"
	CateStr = "����|42|Cate|�����ظ�,"
	PubTimeStr = "����ʱ��|100|PubTime|��־����ʱ���ظ�,"
	QQStr = "QQ����|70|QQ|QQ�����ظ�|QQ��������,"
	QQNameStr = "QQ����|200|QQName|QQ�����ظ�|QQ��������,"
	HandleStr = "����||��ϸ|||�޸�,"
	VisitorsListStr = "�ÿ��б�|80|VisitorsList||�ÿ��б�����,"
	VisitorsNumbStr = "�ÿ�����|80|VisitorsNumb||�ÿ���������,"
	
	SpacePurviewStr = "�ռ�Ȩ��|80|SpacePurview|�ռ�Ȩ���ظ�|�ռ�Ȩ������,"
	QQSexStr = "QQ�Ա�|60|QQSex|QQ�Ա��ظ�|QQ�Ա�����,"
	QQAgeStr = "QQ����|60|QQAge|QQ�����ظ�|QQ��������,"
	AgeStr = "����|60|Age|�����ظ�|��������,"
	'��ϵ 
	BigClassNameStr = "��������|100|BigClassName|���������ظ�|������������," 
	SmallClassNameStr = "С������|100|SmallClassName|С�������ظ�|С����������," 
	ThreeClassNameStr = "��������|100|ThreeClassName|���������ظ�|������������," 
	
	SearchWordStr = "������|200|BigClassName|�������ظ�|����������," 
	XiangGuanWordStr = "��ش�|200|SmallClassName|��ش��ظ�|��ش�����,"  
	
	CountNumbStr = "���Ѷ�|60|CountNumb," 											'ע������ ���� CountNumbStr������������
	QQNumbStr = "������|60|QQNumb," 
	'��־
	OnLineTimeStr = "��¼ʱ��|116|OnLineTime|��¼ʱ���ظ�|��¼ʱ������," 
	CategoryStr = "��־����|80|Category|��־�����ظ�|��־��������,"
	TitleStr = "����|204|Title|�����ظ�|��������," 
	BlogidStr = "��־ID|70|Blogid|��־ID�ظ�,"
	'���
	PhotoPreStr = "����ͼ|100|PhotoPre,"
	PhotoNameStr = "ͼƬ����|100|PhotoName|ͼƬ�����ظ�|ͼƬ��������,"
	BigPhotoUrlStr = "ͼƬ����ͼ|100|BigPhotoUrl||ͼƬ����ͼ����,"
	PhotoIDStr = "ͼƬID|100|PhotoID|ͼƬID�ظ�|ͼƬID����,"	
	MessageQQStr = "������QQ|70|MessageQQ|������QQ�ظ�|������QQ����," 
	MessageUpDateTimeStr = "����ʱ��|120|UpDateTime|����ʱ���ظ�|����ʱ������,"
	PhotoUrlStr = "��������ͼ|100|PhotoUrl,"
	MessageContentStr = "��������||Content||������������,"
	PhotoShowTypeStr = "���ڴ�Сͼ|80|PhotoShowType|���ڴ�Сͼ�ظ�|���ڴ�Сͼ����,"
	'��ͬ	
	CountNumbStr = "����|60|CountNumb," 
	AddTimeStr = "����ʱ��|120|AddTime|����ʱ���ظ�|����ʱ������," 
	'��������
	SearchValueStr = "����ֵ|100|SearchValue|����ֵ�ظ�|����ֵ����,"
	SearchTitleStr = "��������|200|SearchTitle|���������ظ�|������������,"
	SearchContentStr = "��������|200|SearchContent||������������,"
	SearchTimeStr = "����ʱ��|80|SearchTime|����ʱ���ظ�|����ʱ������," 
	SearchSortStr = "��������|80|SearchSort|���������ظ�|������������,"
	WebSiteWebHttpUrlStr = "����/��ַ|200|����/��ַ|�����ظ�|��ַ����/��������,"
	'�Ա�
	UserNameStr = "�û�|100|UserName|�û��ظ�|�û�����,"
	PriceStr = "�۸�|60|Price|�۸��ظ�|�۸�����,"
	FarePriceStr = "�˷�|60|FarePrice|�˷��ظ�|�˷�����,"
	ProvinceStr = "ʡ��|40|Province|ʡ���ظ�|ʡ������,"
	CityStr = "����|40|City|�����ظ�|��������,"
	DescriptionStr = "����ֵ|60|Description|����ֵ�ظ�|����ֵ����,"
	ReplyNumbStr = "������|60|ReplyNumb|�������ظ�|����������,"
	ReplyNameStr = "������|60|ReplyName|�������ظ�|����������,"
	ClickNumbStr = "�����|60|ClickNumb|������ظ�|���������,"	
	'��Ʒ
	ContentStr = "��ע|200|Content||��ע����,"	'����ע��200��ȥ���������Զ�
	UpDateTimeStr = "�ϴ�ʱ��|80|UpDateTime|�ϴ�ʱ���ظ�|�ϴ�ʱ������,"
	ZipSizeStr = "�ļ���С|60|ZipSize|�ļ���С�ظ�,"
	WindowsStr = "����ƽ̨|80|Windows|����ƽ̨�ظ�|����ƽ̨����,"
	DownNumbStr = "������|80|DownNumb|�������ظ�|����������,"
	'Words
	SortStr = "����|60|Sort|�����ظ�|��������,"
	'��Ƹ
	WorkTitleStr = "��Ƹְλ|100|WorkTitle|��Ƹְλ�ظ�|��Ƹְλ����,"
	CompanyStr = "��˾����|100|Company|��˾�����ظ�|��˾��������,"
	DiQuStr = "����|40|DiQu|�����ظ�|��������,"
	WorkEXPStr = "��������|90|WorkEXP|���������ظ�|������������,"
	'QQ��ע
	QQBeiZhuNameStr = "QQ��ע��|90|QQBeiZhuName|QQ��ע���ظ�|QQ��ע������,"
	QQTouXiangStr = "QQͷ��|100|QQͷ��||,"
	'Flash163SWF
	AuthorStr = "����|90|Author|�����ظ�|��������,"
	RenQiStr = "����|90|RenQi|�����ظ�|��������,"
	'Web
	WebTitleStr = "��վ����|200|WebTitle|��վ�����ظ�|��վ��������,"
	WebUrlStr = "��վ��ַ|200|WebUrl|��վ��ַ�ظ�|��վ��ַ����,"
	DayNumbStr = "������|200|DayNumb|�������ظ�|����������,"
	'ce
	CiStr = "����|100|Ci|�����ظ�|��������,"
	ShiStr = "Ӣ�Ľ���||Shi|Ӣ�Ľ����ظ�|Ӣ�Ľ�������,"
	'QQͼƬ�ƽ�
	PoJieStr = "�ƽ�ֵ|80|PoJie|�ƽ�ֵ�ظ�|�ƽ�ֵ����,"
	PoJiePhotoStr = "�ƽ�ͼƬ|80|PoJiePhoto,"
	'CnEnWords
	ENStr = "Ӣ��|120|EN|Ӣ���ظ�|Ӣ������,"
	CNStr = "����|200|CN|�����ظ�|��������,"
	'�ļ����У�Ŀ¼
	FileNameStr = "�ļ�|150|FileName|�ļ��ظ�|�ļ�����,"
	FolderNameStr = "�ļ���|120|FolderName|�ļ����ظ�|�ļ�������,"
	DirPathStr = "Ŀ¼·��||DirPath|Ŀ¼·���ظ�|Ŀ¼·������," 
	'��Ƶ
	HttpUrlStr = "��ַ|200|HttpUrl|��ַ�ظ�|��ַ����,"
	PlayTimeStr = "����ʱ��|80|PlayTime|����ʱ���ظ�|����ʱ������,"
	PlayNumbStr = "���Ŵ���|80|PlayNumb|���Ŵ����ظ�|���Ŵ�������,"
	PostingsTimeStr = "����ʱ��|100|PostingsTime|����ʱ���ظ�|����ʱ������,"
	'����
	NavNameStr = "��������|100|cname|���������ظ�|������������,"
	MutualFriendStr = "��ͬ����|60|MutualFriend|��ͬ�����ظ�|��ͬ��������,"
	SupportNumbStr = "֧����|80|SupportNumb|֧�����ظ�|֧��������,"
	CreateDateStr = "�ύ����|80|CreateDate|�ύ�����ظ�|�ύ��������,"
	ThroughStr = "�Ƿ�ͨ��|80|Through|�Ƿ�ͨ���ظ�|�Ƿ�ͨ������,"
	AdoptionRateStr = "������|80|AdoptionRate|�������ظ�|����������,"
	BrowseNumbStr = "�������|80|BrowseNumb|��������ظ�|�����������,"
	EmailStr = "����|120|Email|�����ظ�|��������,"
	IPStr = "IP|120|IP|IP�ظ�|IP����,"
	AddDateTimeStr = "����ʱ��|120|AddDateTime|����ʱ���ظ�|����ʱ������,"
	FilePathStr = "�ļ�·��|120|FilePath|�ļ�·���ظ�|�ļ�·������,"
	MyMoneyStr ="�û�����|120|MyMoney|�û������ظ�|�û���������,"
	ReplyTimeStr = "����ʱ��|120|ReplyTime|����ʱ���ظ�|����ʱ������,"
	AskTypeStr = "��������|120|AskType|���������ظ�|������������,"
	AskIdStr = "����ID|120|AskId|����ID�ظ�|����ID����,"
	WebSiteStr = "����|120|WebSite|�����ظ�|��������,"
	PageNumberStr = "ҳ��|120|PageNumber|ҳ����|��ҳ������,"
	RankingStr = "����|120|Ranking|�����ظ�|��������,"
	IncludedTimeStr = "��¼ʱ��|120|IncludedTime|��¼ʱ���ظ�|��¼ʱ������,"
	SearchUrlStr = "����������ַ|120|SearchUrl|����������ַ�ظ�|����������ַ����,"
	
	'����Ʒ����
	If CheckStr(Table, "^Product$") Then
		Content = IDStr & BigClassNameStr & SmallClassNameStr & ThreeClassNameStr & TitleStr & ContentStr  & UpDateTimeStr & HandleStr
		AscDescStr = "�����"
		PrintStr = "����|����"
	'��QQ���Ʊ���
	ElseIf CheckStr(Table, "^QQNameTable$") Then
		Content = IDStr & QQStr & QQNameStr & ContentStr & HandleStr
		AscDescStr = "QQ����|QQ����"
		PrintStr = "QQ����|QQ����|QQ���ƺͺ���"
	'��QQ��ע����
	ElseIf CheckStr(Table, "^QQBeiZhu$") Then
		Content = IDStr & QQStr & QQBeiZhuNameStr & HandleStr
		AscDescStr = "QQ����|QQ����"
		PrintStr = "QQ����|QQ��ע��|QQ���ƺͺ���"
	'��QQ����
	ElseIf CheckStr(Table, "^QQ$") Then
		If Rq("Style") = "2" Then
			Content = IDStr & QQStr & QQNameStr & QQSexStr & AgeStr & UpDateTimeStr & QQAgeStr & HandleStr
			AscDescStr = "QQ����|QQ����|QQ�Ա�|����|�ϴ�ʱ��|QQ����"
			PrintStr = "QQ����|QQ�ÿ�|QQ�ͷÿ�"
		Else
			Content = IDStr & BigClassNameStr & QQStr & QQNameStr & VisitorsListStr & SpacePurviewStr & OnLineTimeStr & MutualFriendStr & VisitorsNumbStr & HandleStr
			AscDescStr = "QQ����|QQ����|�ռ�Ȩ��|��¼ʱ��|�ÿ�����"
			PrintStr = "QQ����|QQ�ÿ�|QQ�ͷÿ�"
		End If
	'����ϵQQ����
	ElseIf CheckStr(Table, "^GuanXiQQ$") Then
		Content = IDStr & BigClassNameStr & SmallClassNameStr & QQStr & QQNameStr & CountNumbStr & QQNumbStr & HandleStr
		AscDescStr = "���Ѷ�|������|���Ѷȣ�������|[���ϲ�ѯ]"
		PrintStr = "QQ����|QQ����|QQ���ƺͺ���"
	'����־����
	ElseIf CheckStr(Table, "^DailyRcord$") Then
		Content = IDStr & QQStr & CategoryStr & TitleStr & ReplyNumbStr & BlogidStr & HandleStr
		AscDescStr = "QQ����|������"
		PrintStr = "QQ����|��־����|��־��ַ"
	'��QQ������ͱ���
	ElseIf CheckStr(Table, "^QQPhotoClass$") Then
		Content = IDStr & BigClassNameStr & QQStr & PhotoPreStr & PhotoShowTypeStr & HandleStr
		AscDescStr = "QQ����"
		PrintStr = "��������|QQ����|����ͼ"
	'��QQ�����Ϣ����
	ElseIf CheckStr(Table, "^QQPhotoProduct$") Then
		Content = IDStr & QQStr & PhotoNameStr & BigPhotoUrlStr & PhotoIDStr & PhotoShowTypeStr & HandleStr
		AscDescStr = "QQ����"
		PrintStr = "QQ����|���Сͼ|����ͼ"
	'��QQ������Ա���
	ElseIf CheckStr(Table, "^QQPhotoMessage$") Then
		Content = IDStr & MessageQQStr &  QQStr & QQNameStr  & MessageUpDateTimeStr & PhotoUrlStr & PhotoShowTypeStr & MessageContentStr
		AscDescStr = "QQ����|����ʱ��"
		PrintStr = "QQ����|����ͼƬ|��������"
	'��QQ���߱���
	ElseIf CheckStr(Table, "^QQOnLine") Then
		Content = IDStr &  QQStr & OnLineTimeStr & HandleStr
		AscDescStr = "QQ����|��¼ʱ��"
		PrintStr = "QQ����"
	'��QQ��־��
	ElseIf CheckStr(Table, "^QQJournal") Then
		Content = IDStr &  QQStr & CateStr &  BlogIdStr & TitleStr & PubTimeStr & HandleStr
		AscDescStr = "QQ����|��¼ʱ��"
		PrintStr = "QQ����"
	'����ͬ����
	ElseIf CheckStr(Table, "^WithTable$") Then
		Content = IDStr & BigClassNameStr & SmallClassNameStr &  TitleStr & CountNumbStr & HandleStr
		AscDescStr = "����"
		PrintStr = "����|������ͼ"
	'��QQ�ݴ����
	ElseIf CheckStr(Table, "^QQTemp$") Then
		Content = IDStr &  QQStr & AddTimeStr & HandleStr
		AscDescStr = "QQ����|����ʱ��"
		PrintStr = "QQ����"
	'�������������
	ElseIf CheckStr(Table, "^SearchEngine$") Then
		Content = IDStr &  SearchValueStr & SearchTitleStr & SearchTimeStr & SearchSortStr & SearchContentStr & WebSiteWebHttpUrlStr & HandleStr
		AscDescStr = "QQ����|����ʱ��"
		PrintStr = "��������|������ַ|����"
	'���Ա�����
	ElseIf CheckStr(Table, "^TaoBaoProduct$") Then
		Content = IDStr &  SearchValueStr & UserNameStr & PriceStr & FarePriceStr & ProvinceStr & CityStr & DescriptionStr & ReplyNumbStr
		AscDescStr = "�۸�|�˷�|����ֵ|����"
		PrintStr = "�û�|����|�۸�"
	'���ĵ�����
	ElseIf CheckStr(Table, "^Word$") Then
		Content = IDStr &  TitleStr & AddTimeStr & SortStr & HandleStr
		AscDescStr = "����ʱ��|����"
		PrintStr = "����"
	'��51��Ƹ����
	ElseIf CheckStr(Table, "^JOB_51Job$") Then
		Content = IDStr & SearchValueStr & WorkTitleStr & CompanyStr & CityStr & DiQuStr & UpDateTimeStr & WorkEXPStr  & HandleStr
		AscDescStr = "�ϴ�ʱ��|����"
		PrintStr = "��Ƹְλ|��˾����|ְλ���|��Ƹְλ��ְλ���"
	'������������
	ElseIf CheckStr(Table, "^QQTouXiang$") Then
		Content = IDStr & QQTouXiangStr & QQStr & PhotoShowTypeStr & HandleStr
		AscDescStr = "����ʱ��|����"
		PrintStr = "����"
	'����������
	ElseIf CheckStr(Table, "^Flash163SWF$") Then
		Content = IDStr & TitleStr & AuthorStr & RenQiStr & PhotoShowTypeStr & HandleStr
		AscDescStr = "����|����|����"
		PrintStr = "����|ͼƬ��ַ|SWF��ַ|���ص�ַ"
	'��QQ�ÿͱ���
	ElseIf CheckStr(Table, "^QQVisitorsStat$") Then
		Content = IDStr & QQStr & QQNameStr & CountNumbStr & HandleStr
		AscDescStr = "QQ����|QQ����|����"
		PrintStr = "QQ����|QQ����|QQ���ƺͺ���"
	'���ݴ���ַ����
	ElseIf CheckStr(Table, "^TempUrl$") Then
		Content = IDStr & FolderNameStr & FileNameStr & DirPathStr & HandleStr
		AscDescStr = "�ļ�|�ļ���|Ŀ¼·��"
		PrintStr = "�ļ�|�ļ���|Ŀ¼·��"
	'����Ƶ����
	ElseIf CheckStr(Table, "^Video$") Then
		Content = IDStr & TitleStr & HttpUrlStr & PlayTimeStr & PlayNumbStr & PostingsTimeStr & AuthorStr & HandleStr
		AscDescStr = "����ʱ��|���Ŵ���|����ʱ��"
		PrintStr = "��ƵͼƬ|��Ƶ����|��Ƶ��ַ|��Ƶ�������ַ" 
	'��Web�����ⲿMDB����Ҫ
	ElseIf CheckStr(Table, "^Web$") Then
		Content = IDStr & WebTitleStr & WebUrlStr & DayNumbStr & HandleStr
		AscDescStr = "������"
		PrintStr = "��վ����|��վ��ַ|��վ����ӵ�ַ"
	'�����ʱ����ⲿMDB����Ҫ
	ElseIf CheckStr(Table, "^ce$") Then
		Content = IDStr & CiStr & ShiStr & HandleStr
		AscDescStr = "����"
		PrintStr = "����|Ӣ�Ľ���|���ʼ�Ӣ�Ľ���"
	'����֤������ⲿMDB����Ҫ
	ElseIf CheckStr(Table, "^QQYingZheng$") Then
		Content = IDStr & TitleStr & PoJieStr & PoJiePhotoStr & HandleStr
		AscDescStr = "�ƽ�ֵ"
		PrintStr = "�ƽ�ֵ"
	'����Ӣ�����ⲿMDB����Ҫ
	ElseIf CheckStr(Table, "^CnEnWords$") Then
		Content = IDStr & ENStr & CNStr & SortStr & HandleStr
		AscDescStr = "����"
		PrintStr = "��Ӣ����"
	'��������ⲿMDB����Ҫ
	ElseIf CheckStr(Table, "^class$") Then
		Content = IDStr & NavNameStr & HttpUrlStr & HandleStr
		AscDescStr = "��������"
		PrintStr = "��������"
	'�������QQ���Ʊ�
	ElseIf CheckStr(Table, "^qqname$") Then
		Content = IDStr & QQStr & QQNameStr & HandleStr
		AscDescStr = "QQ����|QQ����"
		PrintStr = "QQ����|QQ����"		
	'��������ٶ�֪����
	ElseIf CheckStr(Table, "^baiduzhidao$") Then
		Content = IDStr & BigClassNameStr & SmallClassNameStr & TitleStr & HttpUrlStr & AuthorStr & ReplyNumbStr & ReplyNameStr & AdoptionRateStr & SupportNumbStr & CreateDateStr & BrowseNumbStr & ThroughStr & HandleStr
		AscDescStr = "�ύ����|������|�������|�Ƿ�ͨ��" 
		PrintStr = "|"
	'��������ٶ�֪����شʱ�
	ElseIf CheckStr(Table, "^BaiduZhidaoWords$") Then
		Content = IDStr & SearchWordStr & XiangGuanWordStr & HandleStr
		AscDescStr = "|"
		PrintStr = "|"		
	'�����������ͳ�Ʊ�
	ElseIf CheckStr(Table, "^StatisticalData$") Then
		Content = IDStr & BigClassNameStr & TitleStr & CountNumbStr & HandleStr
		AscDescStr = "����"
		PrintStr = "|"
	'����������������
	ElseIf CheckStr(Table, "^SoftTask$") Then
		Content = IDStr & BigClassNameStr & EmailStr & ContentStr & UpDateTimeStr & FilePathStr & ThroughStr & HandleStr
		AscDescStr = ""
		PrintStr = "|"		
	'���������Ա��
	ElseIf CheckStr(Table, "^Member$") Then
		Content = IDStr & UserNameStr & MyMoneyStr & EmailStr & IPStr & HandleStr
		AscDescStr = ""
		PrintStr = "|"
	'���������Ա��־��
	ElseIf CheckStr(Table, "^UserLog$") Then
		Content = IDStr & BigClassNameStr & ContentStr & IPStr & AddDateTimeStr & HandleStr
		AscDescStr = ""
		PrintStr = "|"
	'�������������
	ElseIf CheckStr(Table, "^Feedback$") Then
		Content = IDStr & EmailStr & ContentStr & IPStr & AddDateTimeStr & HandleStr
		AscDescStr = ""
		PrintStr = "|"
	'��������ؼ��ʱ�
	ElseIf CheckStr(Table, "^Words$") Then
		Content = IDStr & BigClassNameStr & EmailStr & TitleStr & IPStr & AddDateTimeStr & HandleStr
		AscDescStr = ""
		PrintStr = "|"
	'����������ʱ�
	ElseIf CheckStr(Table, "^Ask$") Then
		Content = IDStr & AskTypeStr & BigClassNameStr & AskIdStr & TitleStr & ReplyNumbStr & ReplyTimeStr & HandleStr
		AscDescStr = "����ʱ��"
		PrintStr = "|"
	'�����������������¼��
	ElseIf CheckStr(Table, "^SearchEngineInfoList$") Then
		Content = IDStr & BigClassNameStr & WebSiteStr & TitleStr & PageNumberStr & RankingStr & IncludedTimeStr & SearchUrlStr & HandleStr
		AscDescStr = ""
		PrintStr = "|" 
		
	'δ����
	Else
		Call Eerr("��"& Table &"����","δ���ã���ֹͣ��ע�⣺���뵽 Fun_Table.Asp ҳ�ҵ� ShowClassName �����޸ġ�")
		 
	End If
	'ѭ����ñ����б�
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


'�������� �� �����ѯ Or �� And
Function HandleSearchSQL(SearchType,SearchValue)
	Dim AddSQL,SeectField,SvalueType
	'�ж��Ƿ�Ϊ��������
	If Left(SearchValue,1) = """" And Right(SearchValue,1) = """" Then
		SearchValue = Left(SearchValue, Len(SearchValue)-1)
		SearchValue = Right(SearchValue, Len(SearchValue)-1)
		SvalueType = "��"
	End If
	If SearchValue <> "" Then
		If SearchType = "QQ��������" Then
			SeectField = "QQ"
		ElseIf SearchType = "QQ��������" Then
			SeectField = "QQName"
		ElseIf SearchType = "�ÿ��б�����" Then
			SeectField = "VisitorsList"
		ElseIf SearchType = "������������" Then
			SeectField = "BigClassName"
		ElseIf SearchType = "С����������" Then
			SeectField = "SmallClassName"
		ElseIf SearchType = "ͼƬ��������" Then
			SeectField = "PhotoName"
		ElseIf SearchType = "������QQ����" Then
			SeectField = "MessageQQ"
		ElseIf SearchType = "��¼ʱ������" Then
			SeectField = "OnLineTime"
		ElseIf SearchType = "��������" Then
			SeectField = "Title"
		ElseIf SearchType = "����ʱ������" Then
			SeectField = "AddTime"
		ElseIf SearchType = "����ֵ����" Then
			SeectField = "SearchValue"
		ElseIf SearchType = "������������" Then
			SeectField = "SearchTitle" 
		ElseIf SearchType = "����ʱ������" Then
			SeectField = "SearchTime" 
		ElseIf SearchType = "������������" Then
			SeectField = "SearchSort" 
		ElseIf SearchType = "������������" Then
			SeectField = "SearchContent" 
		ElseIf SearchType = "������ַ����" Then
			SeectField = "SearchHttpUrl" 
		ElseIf SearchType = "��ַ����" Then
			SeectField = "HttpUrl" 
		ElseIf SearchType = "��������" Then
			SeectField = "WebSite" 
		ElseIf SearchType = "��ע����" Then
			SeectField = "Content" 
		ElseIf SearchType = "�ϴ�ʱ������" Then
			SeectField = "UpDateTime"
		ElseIf SearchType = "����ƽ̨����" Then
			SeectField = "Windows"
		ElseIf SearchType = "���ڴ�Сͼ����" Then
			SeectField = "PhotoShowType"
		ElseIf SearchType = "ͼƬID����" Then
			SeectField = "PhotoID"
		ElseIf SearchType = "��˾��������" Then
			SeectField = "Company"
		ElseIf SearchType = "��Ƹְλ����" Then
			SeectField = "WorkTitle"
		ElseIf SearchType = "��վ��������" Then
			SeectField = "WebTitle"
		ElseIf SearchType = "��վ��ַ����" Then
			SeectField = "WebUrl"
		ElseIf SearchType = "����������" Then
			SeectField = "DayNumb"
		ElseIf SearchType = "��������" Then
			SeectField = "Ci"
		ElseIf SearchType = "Ӣ�Ľ�������" Then
			SeectField = "Shi"
		ElseIf SearchType = "ͼƬ����ͼ����" Then
			SeectField = "BigPhotoUrl"
		ElseIf SearchType = "QQ�Ա�����" Then
			SeectField = "QQSex"
		ElseIf SearchType = "QQ��������" Then
			SeectField = "QQAge"
		ElseIf SearchType = "�ƽ�ֵ����" Then
			SeectField = "PoJie"
		ElseIf SearchType = "����������" Then
			SeectField = "DownNumb"
		ElseIf SearchType = "����������" Then
			SeectField = "ReplyNumb"
		ElseIf SearchType = "��־��������" Then
			SeectField = "Category"
		ElseIf SearchType = "Ӣ������" Then
			SeectField = "EN"
		ElseIf SearchType = "��������" Then
			SeectField = "CN"
		ElseIf SearchType = "��������" Then
			SeectField = "Sort"
		ElseIf SearchType = "�ļ�����" Then
			SeectField = "FileName"
		ElseIf SearchType = "�ļ�������" Then
			SeectField = "FolderName"
		ElseIf SearchType = "Ŀ¼·������" Then
			SeectField = "DirPath"
		ElseIf SearchType = "����ʱ������" Then
			SeectField = "PlayTime"
		ElseIf SearchType = "���Ŵ�������" Then
			SeectField = "PlayNumb"
		ElseIf SearchType = "����ʱ������" Then
			SeectField = "PostingsTime"
		ElseIf SearchType = "��������" Then
			SeectField = "Author"
		ElseIf SearchType = "����������" Then
			SeectField = "BigClassName"
		ElseIf SearchType = "��ش�����" Then
			SeectField = "SmallClassName"
		ElseIf SearchType = "����������" Then
			SeectField = "AdoptionRate"
		ElseIf SearchType = "֧��������" Then
			SeectField = "SupportNumb"
		ElseIf SearchType = "�ύ��������" Then
			SeectField = "CreateDate"
		ElseIf SearchType = "�����������" Then
			SeectField = "BrowseNumb"
		ElseIf SearchType = "�Ƿ�ͨ������" Then
			SeectField = "Through"
		ElseIf SearchType = "����ʱ������" Then
			SeectField = "ReplyTime"
		ElseIf SearchType = "����ID����" Then
			SeectField = "AskId"
	 
			
			
		ElseIf SearchType <> "" Then
			Call Echo(SearchType, "����������")
		End If
		
		
		'���SQL
		If SeectField<>"" Then
			'ΪSqlServer���ݿ� �ֶζ�����Ҫת��Ϊ�ַ�����
			If MDBPath = "" Then
				'������2013,10,24  ���ֶ�Ϊ��������תΪ�ַ����ͣ�Ҫ��Ȼ���������⣬�������������Ͳ�Ҫ�ã��������������ȫ������
				Call OpenConn()
				If Conn.Execute("Select "& SeectField &" From ["& Table & "]")(0).Type = 3 Then
					SeectField = "Convert(VarChar, "& SeectField &")"
				End If
			End If 
			AddSql = "Where "& SeectField &" Like '%"& SearchValue &"%'"
			'�߼�����
			If SvalueType = "��" Then
				AddSql = "Where "& SeectField &" = '"& SearchValue &"'"
			End If
	 	End If
		'�����ѯ Or �� And
		AddSql = OrAndSearch(AddSql,SeectField,SearchValue)
	End If
	HandleSearchSQL = AddSQL
End Function

'��������
Function HandleAscDesc(AddSql, AscDescType)
	If AscDescType = "" Then
		HandleAscDesc = AddSql
		Exit Function
	End If
	
	If AscDescType = "QQ����������" Then
		AddSql = AddSql & "   Order By Convert(Money, QQ) Asc"	
	ElseIf AscDescType = "QQ���뵹����" Then	
		AddSql = AddSql & "   Order By Convert(Money, QQ) Desc"
		
	ElseIf AscDescType = "QQ����������" Then
		AddSql = AddSql & "   Order By QQName Asc"	
	ElseIf AscDescType = "QQ���Ƶ�����" Then	
		AddSql = AddSql & "   Order By QQName Desc"
	
	ElseIf AscDescType = "��עΪ��" Then
		If InStr(AddSql, "Where ") > 0 Then AddSql = AddSql & "   And" Else AddSql = AddSql & "   Where"
		AddSql = AddSql & " Content Is Null"	
	ElseIf AscDescType = "��ע��Ϊ��" Then
		If InStr(AddSql, "Where ") > 0 Then AddSql = AddSql & "   And" Else AddSql = AddSql & "   Where"
		AddSql = AddSql & " Content Is Not Null"	
		
	ElseIf AscDescType = "�ռ�Ȩ��������" Then

		AddSql = AddSql & "   Order By SpacePurview Asc"	
	ElseIf AscDescType = "�ռ�Ȩ�޵�����" Then	
		AddSql = AddSql & "   Order By SpacePurview Desc"
		
	ElseIf AscDescType = "��¼ʱ��������" Then
		AddSql = AddSql & "   Order By OnLineTime Asc"	
	ElseIf AscDescType = "��¼ʱ�䵹����" Then	
		AddSql = AddSql & "   Order By OnLineTime Desc"
		
	ElseIf AscDescType = "���Ѷ�������" Then
		AddSql = AddSql & "   Order By CountNumb"
	ElseIf AscDescType = "���Ѷȵ�����" Then
		AddSql = AddSql & "   Order By CountNumb Desc"
		
	ElseIf AscDescType = "������������" Then
		AddSql = AddSql & "   Order By QQNumb"
	ElseIf AscDescType = "������������" Then
		AddSql = AddSql & "   Order By QQNumb Desc"
		
	ElseIf AscDescType = "���Ѷȣ�������������" Then
		AddSql = AddSql & "   Order By CountNumb+QQNumb"
	ElseIf AscDescType = "���Ѷȣ�������������" Then
		AddSql = AddSql & "   Order By CountNumb+QQNumb Desc"
		
	ElseIf AscDescType = "������������" Then
		AddSql = AddSql & "   Order By ReplyNumb Asc"	
	ElseIf AscDescType = "������������" Then	
		AddSql = AddSql & "   Order By ReplyNumb Desc"
		
	ElseIf AscDescType = "����ʱ��������" Then
		AddSql = AddSql & "   Order By UpDateTime Asc"	
	ElseIf AscDescType = "����ʱ�䵹����" Then	
		AddSql = AddSql & "   Order By UpDateTime Desc"
		
	ElseIf AscDescType = "����������" Then
		AddSql = AddSql & "   Order By CountNumb Asc"	
	ElseIf AscDescType = "����������" Then	
		AddSql = AddSql & "   Order By CountNumb Desc"
		
	ElseIf AscDescType = "����ʱ��������" Then
		AddSql = AddSql & "   Order By AddTime Asc"	
	ElseIf AscDescType = "����ʱ�䵹����" Then	
		AddSql = AddSql & "   Order By AddTime Desc"
		
	ElseIf AscDescType = "�۸�������" Then
		AddSql = AddSql & "   Order By Price Asc"	
	ElseIf AscDescType = "�۸�����" Then	
		AddSql = AddSql & "   Order By Price Desc"
		
	ElseIf AscDescType = "�˷�������" Then
		AddSql = AddSql & "   Order By FarePrice Asc"	
	ElseIf AscDescType = "�˷ѵ�����" Then	
		AddSql = AddSql & "   Order By FarePrice Desc"
		
	ElseIf AscDescType = "�˷�������" Then
		AddSql = AddSql & "   Order By FarePrice Asc"	
	ElseIf AscDescType = "�˷ѵ�����" Then	
		AddSql = AddSql & "   Order By FarePrice Desc"
		
	ElseIf AscDescType = "����ֵ������" Then
		AddSql = AddSql & "   Order By Description Asc"	
	ElseIf AscDescType = "����ֵ������" Then	
		AddSql = AddSql & "   Order By Description Desc"
		
	ElseIf AscDescType = "����������" Then
		AddSql = AddSql & "   Order By ReplyNum Asc"	
	ElseIf AscDescType = "���۵�����" Then	
		AddSql = AddSql & "   Order By ReplyNum Desc"
		
	ElseIf AscDescType = "�ϴ�ʱ��������" Then
		AddSql = AddSql & "   Order By UpDateTime Asc"	
	ElseIf AscDescType = "�ϴ�ʱ�䵹����" Then	
		AddSql = AddSql & "   Order By UpDateTime Desc"
		
	ElseIf AscDescType = "�ļ���С������" Then
		AddSql = AddSql & "   Order By ZipSize Asc"	
	ElseIf AscDescType = "�ļ���С������" Then	
		AddSql = AddSql & "   Order By ZipSize Desc"
		
	ElseIf AscDescType = "����ƽ̨������" Then
		AddSql = AddSql & "   Order By Windows Asc"	
	ElseIf AscDescType = "����ƽ̨������" Then	
		AddSql = AddSql & "   Order By Windows Desc"
		
	ElseIf AscDescType = "����������" Then
		AddSql = AddSql & "   Order By Sort Asc"	
	ElseIf AscDescType = "��������" Then	
		AddSql = AddSql & "   Order By Sort Desc"
		
	ElseIf AscDescType = "����������" Then
		AddSql = AddSql & "   Order By Title Asc"	
	ElseIf AscDescType = "���⵹����" Then	
		AddSql = AddSql & "   Order By Title Desc"		
		
	ElseIf AscDescType = "����������" Then
		AddSql = AddSql & "   Order By RenQi Asc"	
	ElseIf AscDescType = "����������" Then	
		AddSql = AddSql & "   Order By RenQi Desc"
		
	ElseIf AscDescType = "����������" Then
		AddSql = AddSql & "   Order By Author Asc"	
	ElseIf AscDescType = "���ߵ�����" Then	
		AddSql = AddSql & "   Order By Author Desc"
		
		
	ElseIf AscDescType = "����������" Then
		AddSql = AddSql & "   Order By Ci Asc"	
	ElseIf AscDescType = "���ʵ�����" Then	
		AddSql = AddSql & "   Order By Ci Desc"
		
	ElseIf AscDescType = "�ƽ�ֵ������" Then
		AddSql = AddSql & "   Order By PoJie Asc"	
	ElseIf AscDescType = "�ƽ�ֵ������" Then	
		AddSql = AddSql & "   Order By PoJie Desc" 
		
	ElseIf AscDescType = "QQ�Ա�������" Then
		AddSql = AddSql & "   Order By QQSex Asc"	
	ElseIf AscDescType = "QQ�Ա�����" Then	
		AddSql = AddSql & "   Order By QQSex Desc" 
		
	ElseIf AscDescType = "QQ����������" Then
		AddSql = AddSql & "   Order By QQAge Asc"	
	ElseIf AscDescType = "QQ���䵹����" Then	
		AddSql = AddSql & "   Order By QQAge Desc" 
		
	ElseIf AscDescType = "����������" Then
		AddSql = AddSql & "   Order By Age Asc"	
	ElseIf AscDescType = "���䵹����" Then	
		AddSql = AddSql & "   Order By Age Desc" 
		
		
	ElseIf AscDescType = "������������" Then
		AddSql = AddSql & "   Order By DownNumb Asc"	
	ElseIf AscDescType = "������������" Then	
		AddSql = AddSql & "   Order By DownNumb Desc" 
		
	ElseIf AscDescType = "�ļ�������" Then
		AddSql = AddSql & "   Order By FileName Asc"	
	ElseIf AscDescType = "�ļ�������" Then	
		AddSql = AddSql & "   Order By FileName Desc" 
		
	ElseIf AscDescType = "�ļ���������" Then
		AddSql = AddSql & "   Order By FolderName Asc"	
	ElseIf AscDescType = "�ļ��е�����" Then	
		AddSql = AddSql & "   Order By FolderName Desc" 
		
	ElseIf AscDescType = "Ŀ¼·��������" Then
		AddSql = AddSql & "   Order By DirPath Asc"	
	ElseIf AscDescType = "Ŀ¼·��������" Then	
		AddSql = AddSql & "   Order By DirPath Desc" 
		
	ElseIf AscDescType = "����ʱ��������" Then
		AddSql = AddSql & "   Order By PlayTime Asc"	
	ElseIf AscDescType = "����ʱ�䵹����" Then	
		AddSql = AddSql & "   Order By PlayTime Desc" 
		
	ElseIf AscDescType = "���Ŵ���������" Then
		AddSql = AddSql & "   Order By PlayNumb Asc"	
	ElseIf AscDescType = "���Ŵ���������" Then	
		AddSql = AddSql & "   Order By PlayNumb Desc" 
		
	ElseIf AscDescType = "����ʱ��������" Then
		AddSql = AddSql & "   Order By PostingsTime Asc"	
	ElseIf AscDescType = "����ʱ�䵹����" Then	
		AddSql = AddSql & "   Order By PostingsTime Desc" 
		
	ElseIf AscDescType = "�����������" Then
		AddSql = AddSql & "   Order By ClickNumb Asc"	
	ElseIf AscDescType = "�����������" Then	
		AddSql = AddSql & "   Order By ClickNumb Desc" 
		
	ElseIf AscDescType = "�ÿ�����������" Then
		AddSql = AddSql & "   Order By VisitorsNumb Asc"	
	ElseIf AscDescType = "�ÿ�����������" Then	
		AddSql = AddSql & "   Order By VisitorsNumb Desc" 
		
	ElseIf AscDescType = "�ύ����������" Then
		AddSql = AddSql & "   Order By CreateDate Asc"	
	ElseIf AscDescType = "�ύ���ڵ�����" Then	
		AddSql = AddSql & "   Order By CreateDate Desc" 
		
	ElseIf AscDescType = "�������������" Then
		AddSql = AddSql & "   Order By BrowseNumb Asc"	
	ElseIf AscDescType = "�������������" Then	
		AddSql = AddSql & "   Order By BrowseNumb Desc" 
		 
	ElseIf AscDescType = "�Ƿ�ͨ��������" Then
		AddSql = AddSql & "   Order By Through Asc"	
	ElseIf AscDescType = "�Ƿ�ͨ��������" Then	
		AddSql = AddSql & "   Order By Through Desc" 
		
	ElseIf AscDescType = "����ʱ��������" Then
		AddSql = AddSql & "   Order By ReplyTime Asc"	
	ElseIf AscDescType = "����ʱ�䵹����" Then	
		AddSql = AddSql & "   Order By ReplyTime Desc" 
		 

		
		
		 
		
	ElseIf AscDescType = "���ϲ�ѯ" Then
		Call LianHeSearch()
	Else
		Call Echo(AscDescType, "���򲻴���")
	End If
	
	
	HandleAscDesc = AddSql
End Function
'���ϲ�ѯ ��GuanXiQQ.Asp����
Function LianHeSearch()
	Dim SearchType, SearchValue
	SearchType = Request("SearchType")
	SearchValue = Request("SearchValue")
	'���ϲ�ѯ��ĳ��QQ�����ڣ���ʾ����
	If SearchType = "" And InStr(SearchValue,"|")>0 Then
		CAll Echo("", "����")
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
					C = C & "<div align=center><iframe width='800' height=100 frameborder=0 scrolling=no src='Print.Asp?act=��������QQ&QQ="&QQ&"'></iframe></div>"
					N=N+1
					YesNo = False
				End If:Rs.Close
			End If
		Next
		
		If YesNo = False Then
			C = C & "<div align=center style='color:red;'>��ʾ����δ�����Ĺ�ϵQQ��<b>"& N &"</b>����</div>" 
			Call Rw( C ) 
			Response.End()
		End If
		Conn.Execute("Delete From [GuanXiQQ] Where BigClassName='���ϲ�ѯ'")
		Rs.Open"Select * From [GuanXiQQ] Where BigClassName In("& QQIN &")",Conn,1,1
		While Not Rs.Eof
			'��
			Rsx.Open"Select * From [GuanXiQQ] Where BigClassName='���ϲ�ѯ' And QQ='"& Rs("QQ") &"'",Conn,1,3
			If Rsx.Eof Then 
				Rsx.AddNew
				Rsx("BigClassName") = "���ϲ�ѯ"
				Rsx("SearchList") = SearchValue
				Rsx("SmallClassName") = Rs("BigClassName")
				Rsx("QQ") = Rs("QQ")
				Rsx("QQName") = Rs("QQName")
				Rsx("CountNumb") = Rs("CountNumb") + Rs("QQNumb")
			Else
				Rsx("CountNumb") = Rsx("CountNumb") + Rs("CountNumb") + Rs("QQNumb")
				Rsx("QQNumb") = Rsx("QQNumb") +1
				'�ۼ�
				If InStr("|"& Rsx("SmallClassName") &"|", "|"& Rs("QQ") &"|") = False Then
					Rsx("SmallClassName") = Rsx("SmallClassName") & "|" & Rs("BigClassName")
				End If
			End If
			Rsx.Update:Rsx.Close	
		Rs.MoveNext:Wend:Rs.Close
		'��ת�� ���ϲ�ѯ
		CAll Eerr("����","����")
		'Response.Redirect("GuanXiQQ.Asp?act=ShowGuanXiQQList&BigClassName=���ϲ�ѯ&AscDescType=������ӵ�����&CountNumb=1")
		Response.End()
	End If
	'LianHeSearch = AddSql
End Function
'�����ظ�
Function HandleReset(AddSql, AType, Table)
	If AType = "" Then
		HandleReset = AddSql
		Exit Function
	End If
	
	If AType = "QQ�����ظ�" Then
		AddSql = "Select QQ From ["& Table &"] Group By QQ Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "QQ�����ظ�" Then
		AddSql = "Select QQName  From ["& Table &"] Group By QQName Having(Count(*))>" & Cint(Rq("N")) 
	
	ElseIf AType = "�ռ�Ȩ���ظ�" Then
		AddSql = "Select SpacePurview From ["& Table &"] Group By SpacePurview Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "��¼ʱ���ظ�" Then
		AddSql = "Select OnLineTime  From ["& Table &"] Group By OnLineTime Having(Count(*))>" & Cint(Rq("N")) 
	
	ElseIf AType = "���������ظ�" Then
		AddSql = "Select BigClassName From ["& Table &"] Group By BigClassName Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "С�������ظ�" Then
		AddSql = "Select SmallClassName  From ["& Table &"] Group By SmallClassName Having(Count(*))>" & Cint(Rq("N")) 
	
	ElseIf AType = "��־�����ظ�" Then
		AddSql = "Select QQcategory From ["& Table &"] Group By QQcategory Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "�����ظ�" Then
		AddSql = "Select Title  From ["& Table &"] Group By Title Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "�������ظ�" Then
		AddSql = "Select ReplyNumb  From ["& Table &"] Group By ReplyNumb Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "��־ID�ظ�" Then
		AddSql = "Select Blogid  From ["& Table &"] Group By Blogid Having(Count(*))>" & Cint(Rq("N")) 
		
	ElseIf AType = "����ʱ���ظ�" Then
		AddSql = "Select AddTime  From ["& Table &"] Group By AddTime Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "�����ظ�" Then
		AddSql = "Select WebSite  From ["& Table &"] Group By WebSite Having(Count(*))>" & Cint(Rq("N"))  
		
	ElseIf AType = "����ֵ�ظ�" Then
		AddSql = "Select SearchValue  From ["& Table &"] Group By SearchValue Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "�û��ظ�" Then
		AddSql = "Select UserName  From ["& Table &"] Group By UserName Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "�۸��ظ�" Then
		AddSql = "Select Price  From ["& Table &"] Group By Price Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "�˷��ظ�" Then
		AddSql = "Select FarePrice  From ["& Table &"] Group By FarePrice Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "ʡ���ظ�" Then
		AddSql = "Select Province  From ["& Table &"] Group By Province Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "�����ظ�" Then
		AddSql = "Select City  From ["& Table &"] Group By City Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "����ֵ�ظ�" Then
		AddSql = "Select Description  From ["& Table &"] Group By Description Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "�����ظ�" Then
		AddSql = "Select ReplyNum  From ["& Table &"] Group By ReplyNum Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "QQ��ע���ظ�" Then
		AddSql = "Select QQBeiZhuName  From ["& Table &"] Group By QQBeiZhuName Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "��˾�����ظ�" Then	 
		AddSql = "Select Company  From ["& Table &"] Group By Company Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "��Ƹְλ�ظ�" Then	 
		AddSql = "Select WorkTitle  From ["& Table &"] Group By WorkTitle Having(Count(*))>" & Cint(Rq("N"))  
		
		
	ElseIf AType = "�����ظ�" Then	 
		AddSql = "Select Author  From ["& Table &"] Group By Author Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "�����ظ�" Then	 
		AddSql = "Select RenQi  From ["& Table &"] Group By RenQi Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "���ڴ�Сͼ�ظ�" Then	 
		AddSql = "Select PhotoShowType  From ["& Table &"] Group By PhotoShowType Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "��վ�����ظ�" Then	 
		AddSql = "Select WebTitle  From ["& Table &"] Group By WebTitle Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "��վ��ַ�ظ�" Then	 
		AddSql = "Select WebUrl  From ["& Table &"] Group By WebUrl Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "�������ظ�" Then	 
		AddSql = "Select DayNumb  From ["& Table &"] Group By DayNumb Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "�ϴ�ʱ���ظ�" Then	 
		AddSql = "Select UpDateTime  From ["& Table &"] Group By UpDateTime Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "QQ�����ظ�" Then	 
		AddSql = "Select QQAge  From ["& Table &"] Group By QQAge Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "QQ�Ա��ظ�" Then	 
		AddSql = "Select QQSex  From ["& Table &"] Group By QQSex Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "�����ظ�" Then	 
		AddSql = "Select Age  From ["& Table &"] Group By Age Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "����ƽ̨�ظ�" Then	 
		AddSql = "Select Windows  From ["& Table &"] Group By Windows Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "�������ظ�" Then	 
		AddSql = "Select DownNumb  From ["& Table &"] Group By DownNumb Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "�ļ���С�ظ�" Then	 
		AddSql = "Select ZipSize  From ["& Table &"] Group By ZipSize Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "�ļ��ظ�" Then	 
		AddSql = "Select FileName  From ["& Table &"] Group By FileName Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "�ļ����ظ�" Then	 
		AddSql = "Select FolderName  From ["& Table &"] Group By FolderName Having(Count(*))>" & Cint(Rq("N"))
	ElseIf AType = "Ŀ¼·���ظ�" Then	 
		AddSql = "Select DirPath  From ["& Table &"] Group By DirPath Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "��ַ�ظ�" Then	 
		AddSql = "Select HttpUrl  From ["& Table &"] Group By HttpUrl Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "����ʱ���ظ�" Then	 
		AddSql = "Select PlayTime  From ["& Table &"] Group By PlayTime Having(Count(*))>" & Cint(Rq("N")) 
	ElseIf AType = "���Ŵ����ظ�" Then	 
		AddSql = "Select PlayNumb  From ["& Table &"] Group By PlayNumb Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "����ʱ���ظ�" Then	 
		AddSql = "Select PostingsTime  From ["& Table &"] Group By PostingsTime Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "���������ظ�" Then	 
		AddSql = "Select ThreeClassName  From ["& Table &"] Group By ThreeClassName Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "������ظ�" Then	 
		AddSql = "Select ClickNumb  From ["& Table &"] Group By ClickNumb Having(Count(*))>" & Cint(Rq("N"))  
		
	ElseIf AType = "�������ظ�" Then	 
		AddSql = "Select BigClassName  From ["& Table &"] Group By BigClassName Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "�����ظ�" Then	 
		AddSql = "Select Cate  From ["& Table &"] Group By Cate Having(Count(*))>" & Cint(Rq("N"))  
	ElseIf AType = "��־����ʱ���ظ�" Then	 
		AddSql = "Select pubTime  From ["& Table &"] Group By pubTime Having(Count(*))>" & Cint(Rq("N"))  

		
		
	'�ⲿMDB
	ElseIf AType = "�����ظ�" Then	 
		AddSql = "Select Ci  From ["& Table &"] Group By Ci Having(Count(*))>" & Cint(Rq("N"))  
	'�ⲿMDB
	ElseIf AType = "Ӣ�Ľ����ظ�" Then	 
		AddSql = "Select Shi  From ["& Table &"] Group By Shi Having(Count(*))>" & Cint(Rq("N"))  
	'�ⲿMDB
	ElseIf AType = "Ӣ�Ľ����ظ�" Then	 
		AddSql = "Select Shi  From ["& Table &"] Group By Shi Having(Count(*))>" & Cint(Rq("N"))  
		


	Else
		Call Echo(AType, "�ظ�������")
	End If
	
	HandleReset = AddSql
End Function

'�Զ���SQL
Function SelfSql(AddSql, SearchValue, Table)
	Dim SQL 
	 
	SearchValue = RegExp_Replace(SearchValue, "SQL *=", "Sql ") 
	If CheckStr(SearchValue, "SQL ") = True Then
		SQL = Trim(SearchValue)
		SQL = Trim( Right(SQL, Len(SQL) - 4) ) 
		If CheckStr(SQL, "Delete ") = True Then
			SQL = SQL & " &nbsp; <a href=""?act=ClearTable&MDBPath="&Request("MDBPath")&"&SQL="&SQL&""" target=""_blank"" class=""ARed"" onClick=""return Confirm();"">ɾ����¼</a>"
			CAll Eerr("ɾ����Σ�գ�������", SQL) 
		ElseIf CheckStr(SQL, "Update ") = True Then
			CAll Eerr("���±�����", SQL)
		ElseIf CheckStr(SQL, "Select ") = False Then
			SQL = "Select * From ["& Table &"] " & SQL 
		End If
		AddSql = SQL
	End If
	SelfSql = AddSql
End Function

'�滻����򵥸�����ֵ
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

'��ʾ���������б�
Sub ShowCopyList(Title,OKNumb,StrList, StartTime)
	If StartTime <> "" Then StartTime = "��ʱ��" & PrintTimeValue(DateDiff("s", StartTime, Now()) ) & "��&nbsp;"
	Call Echo("��"& Title &"��", "���С�"& OKNumb &"������&nbsp;" & StartTime & "<a href=javascript:CopyTEXT('CopyContent');>����</a>")
	CAll Rw("<div style='margin-left:220px;'><textarea name='CopyContent' id='CopyContent' style='width:800px;height:400px;'>"& StrList &"</textarea></div>")
End Sub
%>









