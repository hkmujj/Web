<%
'添加于2014,04,02

Class Maiside 
	 '产品类型列表展示【SmallClass】
	Function LeftClassList(Did, ProSid, ProTid, AddSql)
		Dim SplStr, I, S, C,Sid,ShowSid,BigFileName,Target,Url,Title,ShowTid,Tid 
		Rs.Open"Select * From [SmallClass] Where BigClassName='"& Did &"' Order By Sort",Conn,1,1
		While Not Rs.Eof
			Sid = Rs("SmallClassName")
			ShowSid = FontColorFontB(Sid,Rs("FontB"),Rs("FontColor")) 	'处理文本颜色与加粗	
			BigFileName = GetSmallClassFileName( Rs("BigClassName"), Rs("SmallClassName") )
			Call WebDebug(BigFileName,"act=CreateClass&ProDid=" & Did & "&ProSid=" & Rs("SmallClassName"))'调试
			Target = ATarget(Rs("Target"))
			
			C=C & "        <div class='ClassSmall'><A "& AHref(BigFileName,"",Target) & StyleClass("ClassSid") &">"& ShowSid &"</A></div>" & vbCrlf
			Rsx.Open"Select * From [ThreeClass] Where BigClassName='"& Did &"' And SmallClassName='"& Rs("SmallClassName") &"' Order By Sort,ID Asc",Conn,1,1
			If Not Rsx.Eof Then				
				For I=1 To Rsx.RecordCount
					 
					If Sid="汉堡加盟菜单" Then
						'链接类别
						Url =  GetThreeClassFileName(Rsx("BigClassName"),Rsx("SmallClassName"),Rsx("ThreeClassName"))
					Else
						'获得产品信息URL
						Url = GetProductUrl("Where BigClassName='"& Rsx("BigClassName") &"' And SmallClassName='"& Rsx("SmallClassName") &"' And ThreeClassName='"& Rsx("ThreeClassName") &"'")
					End If
					
					
					
					Target = ATarget(Rsx("Target"))
					Tid = Rsx("ThreeClassName") 
					ShowTid = FontColorFontB(Tid,Rsx("FontB"),Rsx("FontColor")) 	'处理文本颜色与加粗	
					If Tid = ProTid Then
						ShowTid="<font color=red>"& ShowTid  &"</font>"
					End If
					C=C & "                <div class='ClassThree'><a "& AHref(Url,"",Target) & StyleClass("ClassTid") &">"& ShowTid &"</a></div>" & vbCrlf
				Rsx.MoveNext:Next
			End If:Rsx.Close
			C=C & "<div class='clear10'></div>"
		Rs.MoveNext:Wend:Rs.Close
		LeftClassList = C
	End Function 

	
	

End Class
%>