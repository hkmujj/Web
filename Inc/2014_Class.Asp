<%
'类型类处理

 '显示类别小类
Function SmallClassList(Did, SmallLiCss, SmallACss, AddSql)
	Dim ShowSmallLiCss,C,S,Sid,ShowSid,BigFileName,Target,Url
	If SmallLiCss<>"" Then
		ShowSmallLiCss = " class='"& SmallLiCss &"'"
	End If 
	Rs.Open"Select * From [SmallClass] Where BigClassName='"& Did &"' Order By Sort",Conn,1,1
	While Not Rs.Eof
		Sid = Rs("SmallClassName")
		ShowSid = FontColorFontB(Sid,Rs("FontB"),Rs("FontColor")) 	'处理文本颜色与加粗	
		Url = GetSmallClassFileName( Rs("BigClassName"), Rs("SmallClassName") )
		Call WebDebug(Url,"act=CreateClass&ProDid=" & Did & "&ProSid=" & Rs("SmallClassName"))'调试
		Target = ATarget(Rs("Target"))
		
		S = "        <li"& ShowSmallLiCss &"><a "& AHref(Url,ShowSid,Target) & StyleClass(SmallACss) &">"& ShowSid &"</a></li>" & vbCrlf
		S = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditSmallClass&Id=" & Rs("Id") & "&n=" & GetRnd(11), S, "<a ")
		C=C & S
		
	Rs.MoveNext:Wend:Rs.Close
	SmallClassList = C
End Function 
 '显示类别子类
Function ThreeClassList(Did, ThreeLiCss, ThreeACss, TidAType, AddSql)
	Dim ShowThreeLiCss,C,S,Sid,ShowSid,BigFileName,Target,Url
	If ThreeLiCss<>"" Then
		ShowThreeLiCss = " class='"& ThreeLiCss &"'"
	End If 
	Rs.Open"Select * From [ThreeClass] Where BigClassName='"& Did &"' Order By Sort",Conn,1,1
	While Not Rs.Eof
		Sid = Rs("ThreeClassName")
		ShowSid = FontColorFontB(Sid,Rs("FontB"),Rs("FontColor")) 	'处理文本颜色与加粗	
		
		Url = GetThreeClassFileName( Rs("BigClassName"), Rs("SmallClassName") , Rs("ThreeClassName") )
		Call WebDebug(Url,"act=CreateClass&ProDid=" & Did & "&ProSid=" & Rs("SmallClassName") & "&ProTid=" & Rs("ThreeClassName"))'调试

		If TidAType="文本类型" Then
			Url = "{$GetProductUrl did='"& Rs("BigClassName") &"' sid='"& Rs("SmallClassName") &"' tid='"& Rs("ThreeClassName") &"' title='"& Rs("ThreeClassName") &"'$}" 
		End If
		
		Target = ATarget(Rs("Target"))				
		'Call Echo("Url",Url)
		S = "        <li"& ShowThreeLiCss &"><a "& AHref(Url,ShowSid,Target) & StyleClass(ThreeACss) &">"& ShowSid &"</a></li>" & vbCrlf	 
		S = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditThreeClass&Id=" & Rs("Id") & "&n=" & GetRnd(11), S, "<a ")
		'CAll Die(S)
		C=C & S
	Rs.MoveNext:Wend:Rs.Close
	ThreeClassList = C
End Function
'获得产品类别参数
Function GetClassParam(ByVal TableName,ByVal FieldName,ByVal AddSql)
	Dim Sql,FileName,Url,BigFileName,DebUrl
	Sql = "Select * From ["& TableName &"] " & AddSql
	TempRs2.Open Sql,Conn,1,1
	If Not TempRs2.Eof Then
		'根据名称获得网址
		If FieldName = "[url]" then
			If LCase(TableName) = "bigclass" Then
				BigFileName = GetBigClassFileName(TempRs2("BigClassName"))
				DebUrl = "act=CreateClass&ProDid=" & TempRs2("BigClassName")
			ElseIf LCase(TableName) = "smallclass" Then
				BigFileName = GetSmallClassFileName( TempRs2("BigClassName"), TempRs2("SmallClassName") )
				DebUrl = "act=CreateClass&ProDid=" & TempRs2("BigClassName") & "&ProSid=" & TempRs2("SmallClassName")
			ElseIf LCase(TableName) = "threeclass" Then
				BigFileName = GetThreeClassFileName( TempRs2("BigClassName"), TempRs2("SmallClassName"),TempRs2("ThreeClassName"))
				DebUrl = "act=CreateClass&ProDid=" & TempRs2("BigClassName") & "&ProSid=" & TempRs2("SmallClassName") & "&ProTid=" & TempRs2("ThreeClassName")
			End If			
			BigFileName = FullHttpUrl(Host,BigFileName)
			Call WebDebug(BigFileName, DebUrl)'调试
			GetClassParam = BigFileName
		ElseIf FieldName<>"" Then
			GetClassParam = TempRs2(FieldName)
		End If
	End If:TempRs2.Close
End Function

%>