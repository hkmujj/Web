<%
'搜索Search

'产品搜索记录列表 PopularSearchesList(Action,10," | ","")
Function PopularSearchesList(Action,TopNumb,SplTypeStr,AddSql)
    Dim SplStr, I, S, C,Sel,Url,Title,ShowTitle,wd  
  	If AddSql = "" Then AddSql = " Where Through<>0 Order By Sort Desc"
	Rs.Open"Select * From [SearchCount] " & AddSql,Conn,1,1
	For I=1 To TopNumb
		If Rs.Eof Then Exit For
			Title = Rs("Title")
			
			Url = "/Inc/Create_Html.Asp?act=Search&MackHtml=False&wd=" & Title
			If CheckMakeHtmlFile(Rs("FileName"))=False Then
				Url = Rs("FileName")
			End If  
			
			ShowTitle = Title
			ShowTitle = FontColorFontB(ShowTitle,Rs("FontB"),Rs("FontColor")) 	'处理文本颜色与加粗 
			If Title = wd Then		
				S = TextFontB(Title,True)
			Else 
				S = "<a href='"& Url &"'"& SetHtmlParam(Action,"target|title|alt|id|class|style") &">"& ShowTitle &"</a>"
			End If
			
			'S = HandleDisplayOnlineEditDialog(""& adminDir &"SearchGl.Asp?act=ShowEdit&Id=" & Rs("Id") & "&n=" & GetRnd(11), S,"style='float:right;padding:0 4px;'","")
			S = DisplayOnlineED2(WEB_ADMINURL &"SearchGl.Asp?act=ShowEdit&Id=" & Rs("Id") & "&n=" & GetRnd(11), S, "<li|<a ")
			C=C & S
			If I+1<TopNumb And I+1<=Rs.RecordCount Then C=C & SplTypeStr & vbCrlf
	Rs.MoveNext:Next:Rs.Close
    PopularSearchesList=C
End Function 


%> 