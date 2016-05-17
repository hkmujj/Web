<%
'获得网站配置信息
Sub GetWebConfig(UserId)
	'打开过就用缓冲内容
	If DefaultWebTitle <> "" Then		
		WebTitle = DefaultWebTitle
		WebKeywords = DefaultWebKeywords
		WebDescription = DefaultWebDescription
		Exit Sub
	End If
    '获得网站配置
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
'获得导航类别模板内容 
Function GetNavTemplateText(NavDid, NavSid, NavType,UserID)
    Dim AddSql, TemplateFileName, FileName, Url,NavBigFolderName
    If NavSid <> "" Then
        AddSql = " Where BigClassName='" & NavDid & "' And SmallClassName='"& NavSid &"'" 
	ElseIf NavDid <> "" Then
        AddSql = " Where BigClassName='" & NavDid & "'" 
    Else
        AddSql = " Where NavType='" & NavType & "'" 
    End If

	'导航大类
	If NavDid <> "" Then
		TempRs.Open"Select * From [NavBigClass] Where BigClassName='"& NavDid &"' And UserId="&UserID,Conn,1,1
		If Not TempRs.Eof Then			
			NavBigFolderName = "/" & TempRs("FolderName") & "/"
			WebNavType = TempRs("NavType")		'获得网站导航类型
		End If:TempRs.Close
	End If
	'导航小类
	If NavSid <> "" Then
		TempRs.Open"Select * From [NavSmallClass]  " & AddSql & " And UserId="&UserID,Conn,1,1
		If Not TempRs.Eof Then
			'自定义模板
			If Trim(TempRs("TemplatePath"))<>"" Then
				TemplateFileName = TempRs("TemplatePath") 
			Else
				TemplateFileName = "Main_Model.Html"
			End If
			FileName = GetHandleFileName(TempRs("FileName"))
			'创建一组文件夹(导航文件夹路径)
			Call CreateDirFolder("/" & NavBigFolderName & TempRs("FolderName"))
			CreateHTMLPath = "/" & NavBigFolderName & TempRs("FolderName") & "/" & FileName
			CreateHTMLPath = GetHandleUrl(CreateHTMLPath)	'处理URL   
			
			MainContent = TempRs("Content")		'获得文本中内容

			'显示这篇文章的网站标题关键词描述 
			WebTitle = IIf(TempRs("WebTitle")<>"",TempRs("WebTitle"),WebTitle)
			WebKeywords = IIf(TempRs("WebKeywords")<>"",TempRs("WebKeywords"),WebKeywords)
			WebDescription = IIf(TempRs("WebDescription")<>"",TempRs("WebDescription"),WebDescription) 
		
		End If:TempRs.Close
		GetNavTemplateText = TemplateFileName:Exit Function
	End If
	'导航大类
    Rs.Open "Select * From [NavBigClass] " & AddSql & "And UserId="&UserID, Conn, 1, 1 
    If Not Rs.Eof Then
		'自定义模板
		If Trim(Rs("TemplatePath"))<>"" Then
            TemplateFileName = Rs("TemplatePath") 
        End If
		'获得文件名称
		FileName = GetHandleFileName(Rs("FileName"))
        '创建一组文件夹(导航文件夹路径)
        Call CreateDirFolder("/" & Rs("FolderName")) 
        CreateHTMLPath = "/" & Rs("FolderName") & "/" & FileName  
		CreateHTMLPath = GetHandleUrl(CreateHTMLPath)	'处理URL   
		 
		MainContent = Rs("Content")		'获得文本中内容

		'显示这篇文章的网站标题关键词描述 
		WebTitle = IIf(Rs("WebTitle")<>"",Rs("WebTitle"),WebTitle)
		WebKeywords = IIf(Rs("WebKeywords")<>"",Rs("WebKeywords"),WebKeywords)
		WebDescription = IIf(Rs("WebDescription")<>"",Rs("WebDescription"),WebDescription) 
	Else
		'没有此导航 则为默认模板
		TemplateFileName = "Index_Model.Html"
		Call Echo("没有导航大类【"& NavDid &"】此导航模板为默认",TemplateFileName)
    End If : Rs.Close
    GetNavTemplateText = TemplateFileName
End Function 
'获得信息类别模板内容 
Function GetClassTemplateText(ProDid,ProSid,ProTid,UserID) 
    Dim Sql, TemplateFileName, FileName, Url,RootFolder
    If ProTid <> "" Then
        Sql = "Select * From [ThreeClass] Where BigClassName='" & ProDid & "' And SmallClassName='"& ProSid &"' And ThreeClassName='"& ProTid &"'" 
	ElseIf ProSid <> "" Then
        Sql = "Select * From [SmallClass] Where BigClassName='" & ProDid & "' And SmallClassName='"& ProSid &"'" 
    Else
        Sql = "Select * From [BigClass] Where BigClassName='" & ProDid & "'"
    End If
	'类别大类
	If ProDid <> "" Then
		TempRs.Open"Select * From [BigClass] Where BigClassName='"& ProDid &"' And UserID=" & UserID,Conn,1,1
		If Not TempRs.Eof Then			
			RootFolder = "/" & TempRs("FolderName") & "/"
		End If:TempRs.Close
	End If
	'类别小类
	If ProDid <> "" Then
		TempRs.Open"Select * From [SmallClass] Where BigClassName='"& ProDid &"' And SmallClassName='"& ProSid &"' And UserID=" & UserID,Conn,1,1
		If Not TempRs.Eof Then			
			RootFolder = RootFolder & "/" & TempRs("FolderName") & "/"
		End If:TempRs.Close
	End If
	'类别子类
	If ProTid <> "" Then
		TempRs.Open"Select * From [ThreeClass] Where BigClassName='"& ProDid &"' And SmallClassName='"& ProSid &"' And ThreeClassName='"& ProTid &"' And UserID=" & UserID,Conn,1,1
		If Not TempRs.Eof Then			
			RootFolder = RootFolder & "/" & TempRs("FolderName") & "/" 
		End If:TempRs.Close
	End If
	
	TempRs.Open Sql,Conn,1,1
	If Not TempRs.Eof Then
		'自定义模板
		If Trim(TempRs("TemplatePath"))<>"" Then
			TemplateFileName = TempRs("TemplatePath") 
		Else
			TemplateFileName = "Template_Main.html"			 
		End If
		FileName = GetHandleFileName(TempRs("FileName"))
		'创建一组文件夹(导航文件夹路径)
		Call CreateDirFolder("/" & RootFolder )
		CreateHTMLPath = "/" & RootFolder & "/" & FileName
		CreateHTMLPath = GetHandleUrl(CreateHTMLPath)	'处理URL   
		
		'MainContent = TempRs("Content")		'获得文本中内容
		If TempRs("WebTitle") <> "" Then WebTitle = TempRs("WebTitle")
		If TempRs("WebKeywords") <> "" Then WebKeywords = TempRs("WebKeywords")
		If TempRs("WebDescription") <> "" Then WebDescription = TempRs("WebDescription")
	End If:TempRs.Close
	 
	GetClassTemplateText = TemplateFileName
End Function
'文本【MainInfo】
Function GetMainInfo(Did,UserID)
	Dim Url,C
	Rs.Open"Select * From [MainInfo] Where Title='"& Did &"' And Recommend=1 And UserID="& UserID &" Order By Sort,ID",Conn,1,1
	If Not Rs.Eof Then  
		 GetMainInfo = Rs("Content")
	End If : Rs.Close
End Function 
 '产品类型列表展示【SmallClass】
Function ProductNavList(Did, ProSid, ProTid)
	Dim SplStr, I, S, C,Sid,ShowSid,BigFileName,Target,Url,Title,ShowTid,Tid 
	Rs.Open"Select * From [SmallClass] Where BigClassName='"& Did &"'  And UserID="& UserID & " Order By Sort",Conn,1,1
	While Not Rs.Eof
		Sid = Rs("SmallClassName")
		ShowSid = FontColorFontB(Sid,Rs("FontB"),Rs("FontColor")) 	'处理文本颜色与加粗	
		BigFileName = GetSmallClassFileName( Rs("BigClassName"), Rs("SmallClassName") )
		Call WebDebug(BigFileName,"act=CreateClass&ProDid=" & Did & "&ProSid=" & Rs("SmallClassName"))'调试
		Target = ATarget(Rs("Target"))
		
		C=C & "        <div class='ClassSmall'><A "& AHref(BigFileName,"",Target) & StyleClass("ClassSid") &">"& ShowSid &"</A></div>" & vbCrlf
 
		C=C & "<div class='clear10'></div>"
	Rs.MoveNext:Wend:Rs.Close
	ProductNavList = C
End Function 


'创建一篇文章
Sub CreateArticle(Id)
	Dim FileName,ImgUrl
	Rs.Open"Select * From [Product] Where Id=" & Id,Conn,1,1
	If Not Rs.Eof Then
		TempDid = Rs("BigclassName")
		TempSid = Rs("SmallclassName")
		TempTid = Rs("ThreeclassName")		
		MainTitle = Rs("Title")					'信息标题名称
		ImgUrl = ""
		If IsNul(Rs("SmallFiles"))=False Then	
			If IsNul(Rs("SmallFiles")) Then ImgUrl="Images/error.gif" Else ImgUrl=Rs("SmallFiles")	
		ElseIf IsNul(Rs("BigFiles"))=False Then	
			If IsNul(Rs("BigFiles")) Then ImgUrl="Images/error.gif" Else ImgUrl=Rs("BigFiles")
		End If 
		MainContent = "<div class='ArticleTitle'>"& Rs("Title") &"</div>"
		
		If ImgUrl <> "" And len(ImgUrl)>5  Then
			If CheckFile(ImgUrl)=True Then
				MainContent = MainContent & "<div style='text-align:center; border:1px double #CCCCCC; padding:4px;margin-bottom:10px;'><img src='"& ImgUrl &"' width='"& HandleShowImg(ImgUrl,740,0,"宽") &"'></div>"
			End If
		End If 
		MainContent = MainContent & Rs("Content") 
		
		
		MainContent = MainContent & "<hr>" & UpPage(Rs(0)) & DownPage(Rs(0)) 
		 
		'自定义模板
		If Trim(Rs("TemplatePath"))<>"" Then
			TemplatePath = WebTemplate & Rs("TemplatePath") 
		Else
			TemplatePath = WebTemplate & "Template_Main.Html"
		End If 
		FileName = GetHandleFileName(Rs("FileName"))
		'创建一组文件夹(导航文件夹路径)
		Call CreateDirFolder("/Html")
		CreateHTMLPath = "/Html/" & Rs("FolderName") & "/" & FileName
		CreateHTMLPath = GetHandleUrl(CreateHTMLPath)	'处理URL 
		
		'显示这篇文章的网站标题关键词描述
		WebTitle = IIf(Rs("Title")<>"",Rs("Title"),WebTitle)
		WebTitle = IIf(Rs("WebTitle")<>"",Rs("WebTitle"),WebTitle)
		WebKeywords = IIf(Rs("WebKeywords")<>"",Rs("WebKeywords"),WebKeywords)
		WebDescription = IIf(Rs("WebDescription")<>"",Rs("WebDescription"),WebDescription) 
	End If:Rs.Close 
End Sub
%>