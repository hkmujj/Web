<!--#Include File="Config.Asp"-->
<% 
'选择动作
Select Case Request("act")
	Case "ViewInfo" : ViewInfo()			'查看信息
	Case "ViewList" : ViewList()			'查看列表
	Case "ViewList2" : ViewList2()			'查看列表2
	Case "test" : mytest()					'测试
End Select
'测试
Sub mytest()
	Dim UserName,UserName2,PassWord
	UserName=Rf("username")
	UserName2=Rq("username") 
	PassWord=Rf("userpass")
	Rw("UserName=" & UserName & vbCrlf)
	Rw("UserName2=" & UserName2 & vbCrlf)
	Rw("PassWord=" & PassWord)
End Sub

'查看列表
Sub ViewList2()
	Dim IPList,nIPCount,Mobile,MobileList,nMobileType,nMobile,nComputer,S,I,IP
	Response.Addheader "Content-Type","text/html; charset=utf-8"
	Dim nArticleCount,nAddArticle,C,Path,Content,SplStr
	 
	Call OpenConn()
	nArticleCount=Conn.Execute("Select Count(*) From [Product]")(0)
	
	For I=0 To 50
		nAddArticle=Conn.Execute("Select Count(*) From [Product] Where UpDateTime=#"& Date()-I &"#")(0)
		C = C &  "时间(" & Format_Time(Now-I,2) & ")" & vbCrlf & "新添文章("& nAddArticle &"/" & nArticleCount & ")" & vbCrlf
		Content = GetFText("\Admin\Data\Stat\" & Format_Time(Now()-I,2) & ".txt")
		SplStr=Split(Content,"当前：")
		nMobile=0			'默认手机浏览数
		nComputer=0			'默认电脑浏览数
		nIPCount = 0		'独立IP总数
		nMobileType = 0		'手机类型数
		
		For Each S In SplStr
			If InStr(S,"IP:")>0 Then
				If InStr(S,vbCrlf & "IE: U" & vbCrlf)>0 Then
					nMobile = nMobile + 1
					Mobile = Mid(S,InStr(S,"XP=")+3)
					Mobile = Mid(Mobile,1,InStr(Mobile,vbCrlf)-2)
					If InStr(vbCrlf & MobileList & vbCrlf, vbCrlf & Mobile & vbCrlf)=0 Then
						MobileList = MobileList & Mobile & vbCrlf
						nMobileType = nMobileType + 1
					End If
				Else
					nComputer = nComputer + 1
				End If
				IP = Mid(S,InStr(S,"IP:")+3)
				IP = Mid(IP,1,InStr(IP,vbCrlf)-2)
				If InStr(vbCrlf & IPList & vbCrlf, vbCrlf & IP & vbCrlf)=0 Then
					IPList = IPList & IP & vbCrlf
					nIPCount = nIPCount + 1
				End If
			End If
		Next
		
		C=C & "LogDate="& Format_Time(Now()-I,2) &"" & vbCrlf 		'日志时期
		C=C & "AddArticle="& nAddArticle &"" & vbCrlf 				'添加文章数
		C=C & "ArticleSum="& nArticleCount &"" & vbCrlf 			'文章总数
		C=C & "WebBrowseSum="& (nComputer + nMobile) &"" & vbCrlf 	'网页浏览总数
		C=C & "ComputerSum="& nComputer &"" & vbCrlf 				'电脑浏览数
		C=C & "MobileSum="& nMobile &"" & vbCrlf 					'手机浏览数
		C=C & "IndependentIP="& nIPCount &"" & vbCrlf 				'独立IP数
		C=C & "IndependentMobileType="& nMobileType &"" & vbCrlf 	'独立手机类型数
		C=C & "MobileTypeList="& MobileList &"" & vbCrlf 			'独立手机类列表		
		C=C & "-----------" & vbCrlf
	Next
	Response.Write(C)
End Sub
'查看列表
Sub ViewList()
	Dim IPList,nIPCount,Mobile,MobileList,nMobileType,nMobile,nComputer,S,I,IP
	Response.Addheader "Content-Type","text/html; charset=utf-8"
	Dim nArticleCount,nAddArticle,C,Path,Content,SplStr
	Call OpenConn()
	nArticleCount=Conn.Execute("Select Count(*) From [Product]")(0)

	
	For I=0 To 50
		nAddArticle=Conn.Execute("Select Count(*) From [Product] Where UpDateTime=#"& Date()-I &"#")(0)
		C = C &  "时间(" & Format_Time(Now-I,2) & ")" & vbCrlf & "新添文章("& nAddArticle &"/" & nArticleCount & ")" & vbCrlf
		Content = GetFText("\Admin\Data\Stat\" & Format_Time(Now()-I,2) & ".txt")
		SplStr=Split(Content,"当前：")
		nMobile=0			'默认手机浏览数
		nComputer=0			'默认电脑浏览数
		nIPCount = 0		'独立IP总数
		nMobileType = 0		'手机类型数
		For Each S In SplStr
			If InStr(S,"IP:")>0 Then
				If InStr(S,vbCrlf & "IE: U" & vbCrlf)>0 Then
					nMobile = nMobile + 1
					Mobile = Mid(S,InStr(S,"XP=")+3)
					Mobile = Mid(Mobile,1,InStr(Mobile,vbCrlf)-2)
					If InStr(vbCrlf & MobileList & vbCrlf, vbCrlf & Mobile & vbCrlf)=0 Then
						MobileList = MobileList & Mobile & vbCrlf
						nMobileType = nMobileType + 1
					End If
				Else
					nComputer = nComputer + 1
				End If
				IP = Mid(S,InStr(S,"IP:")+3)
				IP = Mid(IP,1,InStr(IP,vbCrlf)-2)
				If InStr(vbCrlf & IPList & vbCrlf, vbCrlf & IP & vbCrlf)=0 Then
					IPList = IPList & IP & vbCrlf
					nIPCount = nIPCount + 1
				End If
			End If
		Next
		C=C & "网页被浏览("& nComputer + nMobile &")次" & vbCrlf
		C=C &"IE浏览("& nComputer &")次" & vbCrlf
		C=C & "手机浏览("& nMobile &")次" & vbCrlf
		C=C & "独立IP数为("& nIPCount &")个" & vbCrlf 
		C=C & "-----------" & vbCrlf
	Next
	Response.Write(C)
End Sub

'查看信息
Sub ViewInfo()
	Dim IPList,nIPCount,Mobile,MobileList,nMobileType,nMobile,nComputer,S,I,IP
	Response.Addheader "Content-Type","text/html; charset=utf-8"
	Dim nArticleCount,nAddArticle,C,Path,Content,SplStr
	Call OpenConn()
	nArticleCount=Conn.Execute("Select Count(*) From [Product]")(0)
	nAddArticle=Conn.Execute("Select Count(*) From [Product] Where UpDateTime>=#"& Date() &"#")(0)
	

	
	C = "时间(" & Format_Time(Now,2) & ")" & vbCrlf & "新添文章("& nAddArticle &"/" & nArticleCount & ")" & vbCrlf
	Content = GetFText("\Admin\Data\Stat\" & Format_Time(Now(),2) & ".txt")
	SplStr=Split(Content,"当前：")
	nMobile=0			'默认手机浏览数
	nComputer=0			'默认电脑浏览数
	nIPCount = 0		'独立IP总数
	nMobileType = 0		'手机类型数
	For Each S In SplStr
		If InStr(S,"IP:")>0 Then
			If InStr(S,vbCrlf & "IE: U" & vbCrlf)>0 Then
				nMobile = nMobile + 1
				Mobile = Mid(S,InStr(S,"XP=")+3)
				Mobile = Mid(Mobile,1,InStr(Mobile,vbCrlf)-2)
				If InStr(vbCrlf & MobileList & vbCrlf, vbCrlf & Mobile & vbCrlf)=0 Then
					MobileList = MobileList & Mobile & vbCrlf
					nMobileType = nMobileType + 1
				End If
			Else
				nComputer = nComputer + 1
			End If
			IP = Mid(S,InStr(S,"IP:")+3)
			IP = Mid(IP,1,InStr(IP,vbCrlf)-2)
			If InStr(vbCrlf & IPList & vbCrlf, vbCrlf & IP & vbCrlf)=0 Then
				IPList = IPList & IP & vbCrlf
				nIPCount = nIPCount + 1
			End If
		End If
	Next
	C=C & "网页被浏览("& nComputer + nMobile &")次" & vbCrlf
	C=C &"IE浏览("& nComputer &")次" & vbCrlf
	C=C & "手机浏览("& nMobile &")次" & vbCrlf
	C=C & "独立IP数为("& nIPCount &")个" & vbCrlf
	C=C & "不同手机类型数为("& nMobileType &")个" & vbCrlf
	C=C & "不同手机列表："& vbCrlf &  MobileList
	Response.Write(C)
End Sub
%>