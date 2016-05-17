<%
'�����վ������Ϣ
Sub GetWebConfig(UserId)
	'�򿪹����û�������
	If DefaultWebTitle <> "" Then		
		WebTitle = DefaultWebTitle
		WebKeywords = DefaultWebKeywords
		WebDescription = DefaultWebDescription
		Exit Sub
	End If
    '�����վ����
    Rs.Open "Select * From [WebSite] Where UserId="&UserId, Conn, 1, 1 
    If Not Rs.Eof Then
        WebTemplate = Rs("WebTemplate") 
        WebSkins = Rs("WebSkins") 		
		
		WebTitle = Rs("WebTitle") 
        WebKeywords = Rs("WebKeywords") 
        WebDescription = Rs("WebDescription") 
		WebBottom = Rs("WebSiteBottom")
    End If : Rs.Close 
End Sub
'��õ������ģ������ 
Function GetNavTemplateText(NavDid, NavSid, NavType,UserID)
    Dim AddSql, TemplateFileName, FileName, Url,NavBigFolderName
    If NavSid <> "" Then
        AddSql = " Where BigClassName='" & NavDid & "' And SmallClassName='"& NavSid &"'" 
	ElseIf NavDid <> "" Then
        AddSql = " Where BigClassName='" & NavDid & "'" 
    Else
        AddSql = " Where NavType='" & NavType & "'" 
    End If

	'��������
	If NavDid <> "" Then
		TempRs.Open"Select * From [NavBigClass] Where BigClassName='"& NavDid &"' And UserId="&UserID,Conn,1,1
		If Not TempRs.Eof Then			
			NavBigFolderName = "/" & TempRs("FolderName") & "/"
			WebNavType = TempRs("NavType")		'�����վ��������
		End If:TempRs.Close
	End If
	'����С��
	If NavSid <> "" Then
		TempRs.Open"Select * From [NavSmallClass]  " & AddSql & " And UserId="&UserID,Conn,1,1
		If Not TempRs.Eof Then
			'�Զ���ģ��
			If Trim(TempRs("TemplatePath"))<>"" Then
				TemplateFileName = TempRs("TemplatePath") 
			Else
				TemplateFileName = "Main_Model.Html"
			End If
			FileName = GetHandleFileName(TempRs("FileName"))
			'����һ���ļ���(�����ļ���·��)
			Call CreateDirFolder("/" & NavBigFolderName & TempRs("FolderName"))
			CreateHTMLPath = "/" & NavBigFolderName & TempRs("FolderName") & "/" & FileName
			CreateHTMLPath = GetHandleUrl(CreateHTMLPath)	'����URL   
			
			MainContent = TempRs("Content")		'����ı�������

			'��ʾ��ƪ���µ���վ����ؼ������� 
			WebTitle = IIf(TempRs("WebTitle")<>"",TempRs("WebTitle"),WebTitle)
			WebKeywords = IIf(TempRs("WebKeywords")<>"",TempRs("WebKeywords"),WebKeywords)
			WebDescription = IIf(TempRs("WebDescription")<>"",TempRs("WebDescription"),WebDescription) 
		
		End If:TempRs.Close
		GetNavTemplateText = TemplateFileName:Exit Function
	End If
	'��������
    Rs.Open "Select * From [NavBigClass] " & AddSql & "And UserId="&UserID, Conn, 1, 1 
    If Not Rs.Eof Then
		'�Զ���ģ��
		If Trim(Rs("TemplatePath"))<>"" Then
            TemplateFileName = Rs("TemplatePath") 
        End If
		'����ļ�����
		FileName = GetHandleFileName(Rs("FileName"))
        '����һ���ļ���(�����ļ���·��)
        Call CreateDirFolder("/" & Rs("FolderName")) 
        CreateHTMLPath = "/" & Rs("FolderName") & "/" & FileName  
		CreateHTMLPath = GetHandleUrl(CreateHTMLPath)	'����URL   
		 
		MainContent = Rs("Content")		'����ı�������

		'��ʾ��ƪ���µ���վ����ؼ������� 
		WebTitle = IIf(Rs("WebTitle")<>"",Rs("WebTitle"),WebTitle)
		WebKeywords = IIf(Rs("WebKeywords")<>"",Rs("WebKeywords"),WebKeywords)
		WebDescription = IIf(Rs("WebDescription")<>"",Rs("WebDescription"),WebDescription) 
	Else
		'û�д˵��� ��ΪĬ��ģ��
		TemplateFileName = "Index_Model.Html"
		Call Echo("û�е������ࡾ"& NavDid &"���˵���ģ��ΪĬ��",TemplateFileName)
    End If : Rs.Close
    GetNavTemplateText = TemplateFileName
End Function 
'�����Ϣ���ģ������ 
Function GetClassTemplateText(ProDid,ProSid,ProTid,UserID) 
    Dim Sql, TemplateFileName, FileName, Url,RootFolder
    If ProTid <> "" Then
        Sql = "Select * From [ThreeClass] Where BigClassName='" & ProDid & "' And SmallClassName='"& ProSid &"' And ThreeClassName='"& ProTid &"'" 
	ElseIf ProSid <> "" Then
        Sql = "Select * From [SmallClass] Where BigClassName='" & ProDid & "' And SmallClassName='"& ProSid &"'" 
    Else
        Sql = "Select * From [BigClass] Where BigClassName='" & ProDid & "'"
    End If
	'������
	If ProDid <> "" Then
		TempRs.Open"Select * From [BigClass] Where BigClassName='"& ProDid &"' And UserID=" & UserID,Conn,1,1
		If Not TempRs.Eof Then			
			RootFolder = "/" & TempRs("FolderName") & "/"
		End If:TempRs.Close
	End If
	'���С��
	If ProDid <> "" Then
		TempRs.Open"Select * From [SmallClass] Where BigClassName='"& ProDid &"' And SmallClassName='"& ProSid &"' And UserID=" & UserID,Conn,1,1
		If Not TempRs.Eof Then			
			RootFolder = RootFolder & "/" & TempRs("FolderName") & "/"
		End If:TempRs.Close
	End If
	'�������
	If ProTid <> "" Then
		TempRs.Open"Select * From [ThreeClass] Where BigClassName='"& ProDid &"' And SmallClassName='"& ProSid &"' And ThreeClassName='"& ProTid &"' And UserID=" & UserID,Conn,1,1
		If Not TempRs.Eof Then			
			RootFolder = RootFolder & "/" & TempRs("FolderName") & "/" 
		End If:TempRs.Close
	End If
	
	TempRs.Open Sql,Conn,1,1
	If Not TempRs.Eof Then
		'�Զ���ģ��
		If Trim(TempRs("TemplatePath"))<>"" Then
			TemplateFileName = TempRs("TemplatePath") 
		Else
			TemplateFileName = "Template_Main.html"			 
		End If
		FileName = GetHandleFileName(TempRs("FileName"))
		'����һ���ļ���(�����ļ���·��)
		Call CreateDirFolder("/" & RootFolder )
		CreateHTMLPath = "/" & RootFolder & "/" & FileName
		CreateHTMLPath = GetHandleUrl(CreateHTMLPath)	'����URL   
		
		'MainContent = TempRs("Content")		'����ı�������
		If TempRs("WebTitle") <> "" Then WebTitle = TempRs("WebTitle")
		If TempRs("WebKeywords") <> "" Then WebKeywords = TempRs("WebKeywords")
		If TempRs("WebDescription") <> "" Then WebDescription = TempRs("WebDescription")
	End If:TempRs.Close
	 
	GetClassTemplateText = TemplateFileName
End Function
'�ı���MainInfo��
Function GetMainInfo(Did,UserID)
	Dim Url,C
	Rs.Open"Select * From [MainInfo] Where Title='"& Did &"' And Recommend=1 And UserID="& UserID &" Order By Sort,ID",Conn,1,1
	If Not Rs.Eof Then  
		 GetMainInfo = Rs("Content")
	End If : Rs.Close
End Function 
 '��Ʒ�����б�չʾ��SmallClass��
Function ProductNavList(Did, ProSid, ProTid)
	Dim SplStr, I, S, C,Sid,ShowSid,BigFileName,Target,Url,Title,ShowTid,Tid 
	Rs.Open"Select * From [SmallClass] Where BigClassName='"& Did &"'  And UserID="& UserID & " Order By Sort",Conn,1,1
	While Not Rs.Eof
		Sid = Rs("SmallClassName")
		ShowSid = FontColorFontB(Sid,Rs("FontB"),Rs("FontColor")) 	'�����ı���ɫ��Ӵ�	
		BigFileName = GetSmallClassFileName( Rs("BigClassName"), Rs("SmallClassName") )
		Call WebDebug(BigFileName,"act=CreateClass&ProDid=" & Did & "&ProSid=" & Rs("SmallClassName"))'����
		Target = ATarget(Rs("Target"))
		
		C=C & "        <div class='ClassSmall'><A "& AHref(BigFileName,"",Target) & StyleClass("ClassSid") &">"& ShowSid &"</A></div>" & vbCrlf
 
		C=C & "<div class='clear10'></div>"
	Rs.MoveNext:Wend:Rs.Close
	ProductNavList = C
End Function 


'����һƪ����
Sub CreateArticle(Id)
	Dim FileName,ImgUrl
	Rs.Open"Select * From [Product] Where Id=" & Id,Conn,1,1
	If Not Rs.Eof Then
		TempDid = Rs("BigclassName")
		TempSid = Rs("SmallclassName")
		TempTid = Rs("ThreeclassName")		
		MainTitle = Rs("Title")					'��Ϣ��������
		ImgUrl = ""
		If IsNul(Rs("SmallFiles"))=False Then	
			If IsNul(Rs("SmallFiles")) Then ImgUrl="Images/error.gif" Else ImgUrl=Rs("SmallFiles")	
		ElseIf IsNul(Rs("BigFiles"))=False Then	
			If IsNul(Rs("BigFiles")) Then ImgUrl="Images/error.gif" Else ImgUrl=Rs("BigFiles")
		End If 
		MainContent = "<div class='ArticleTitle'>"& Rs("Title") &"</div>"
		
		If ImgUrl <> "" And len(ImgUrl)>5  Then
			If CheckFile(ImgUrl)=True Then
				MainContent = MainContent & "<div style='text-align:center; border:1px double #CCCCCC; padding:4px;margin-bottom:10px;'><img src='"& ImgUrl &"' width='"& HandleShowImg(ImgUrl,740,0,"��") &"'></div>"
			End If
		End If 
		MainContent = MainContent & Rs("Content") 
		
		
		MainContent = MainContent & "<hr>" & UpPage(Rs(0)) & DownPage(Rs(0)) 
		 
		'�Զ���ģ��
		If Trim(Rs("TemplatePath"))<>"" Then
			TemplatePath = WebTemplate & Rs("TemplatePath") 
		Else
			TemplatePath = WebTemplate & "Template_Main.Html"
		End If 
		FileName = GetHandleFileName(Rs("FileName"))
		'����һ���ļ���(�����ļ���·��)
		Call CreateDirFolder("/Html")
		CreateHTMLPath = "/Html/" & Rs("FolderName") & "/" & FileName
		CreateHTMLPath = GetHandleUrl(CreateHTMLPath)	'����URL 
		
		'��ʾ��ƪ���µ���վ����ؼ�������
		WebTitle = IIf(Rs("Title")<>"",Rs("Title"),WebTitle)
		WebTitle = IIf(Rs("WebTitle")<>"",Rs("WebTitle"),WebTitle)
		WebKeywords = IIf(Rs("WebKeywords")<>"",Rs("WebKeywords"),WebKeywords)
		WebDescription = IIf(Rs("WebDescription")<>"",Rs("WebDescription"),WebDescription) 
	End If:Rs.Close 
End Sub
%>