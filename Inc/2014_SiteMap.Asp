<%
'地图

'网页地图新版
Function SiteMap(Action)
    Dim I, J, P, C, Url, Target, Did, ShowDid, Sid, ShowSid, Tid, ShowTid, AddSql 
    Dim Tstr, TempC 
    Dim ShieldNavDidList, ShieldNavSidList, ShieldDidList, ShieldSidList, ShieldTidList 
    ShieldNavDidList = RParam(Action, "ShieldNavDidList") 
    ShieldNavSidList = RParam(Action, "ShieldNavSidList") 
    ShieldDidList = RParam(Action, "ShieldDidList") 
    ShieldSidList = RParam(Action, "ShieldSidList") 
    ShieldTidList = RParam(Action, "ShieldTidList") 
    'Call Echo("ShieldNavSidList",ShieldNavSidList)
    C = C & vbCrLf & "<!--SiteMap start-->" & vbCrLf 
    C = C & "<div class=""sitemap_con"">" & vbCrLf 
    '导航类别
    Rs.Open "Select * From [NavBigClass] Order By Sort", Conn, 1, 1 
    For I = 1 To Rs.RecordCount
        Did = Rs("BigClassName") 
        If InStr("|" & ShieldNavDidList & "|", "|" & Did & "|") = False Then
            Url = GetNavUrl(Rs("BigClassName"), "") 
            ShowDid = FontColorFontB(Did, Rs("FontB"), Rs("FontColor")) 
            Target = Rs("Target") : Target = "_blank"                                       '为新页打开

            C = C & "    <h4><a " & AHref(Url, Did, Target) & ">" & ShowDid & "</a></h4>" & vbCrLf 
            If Rs("NavType") = "文本" Then
                Rss.Open "Select * From [NavSmallClass] Where BigClassName='" & Rs("BigClassName") & "' Order By Sort", Conn, 1, 1 
                If Not Rss.EOF Then
                    TempC = "" 
                    For J = 1 To Rss.RecordCount
                        Sid = Rss("SmallClassName") 
                        If InStr("|" & ShieldNavSidList & "|", "|" & Sid & "|") = False Then
                            Url = GetNavUrl(Rss("BigClassName"), Rss("SmallClassName")) 
                            ShowSid = FontColorFontB(Sid, Rss("FontB"), Rss("FontColor")) 
                            Target = Rss("Target") : Target = "_blank"                                      '为新页打开
                            TempC = TempC & CopyStr(" ", 16) & "<li><a " & AHref(Url, "", Target) & ">" & ShowSid & "</a></li>" & vbCrLf 
                        End If 
                    Rss.MoveNext : Next 
                    '列表不为空则显示
                    If TempC <> "" Then
                        C = C & CopyStr(" ", 8) & "<div class=""group"">" & vbCrLf & CopyStr(" ", 12) & "<ul>" & vbCrLf 
                        C = C & TempC 
                        C = C & CopyStr(" ", 12) & "</ul>" & vbCrLf & CopyStr(" ", 12) & "<div class=""clear""></div>" & vbCrLf 
                        C = C & CopyStr(" ", 8) & "</div>" & vbCrLf 
                    End If 

                End If : Rss.Close 
            End If 
        End If 
    Rs.MoveNext : Next : Rs.Close 
    '导航类别
    Rs.Open "Select * From [BigClass] Order By Sort", Conn, 1, 1 
    For I = 1 To Rs.RecordCount
        Did = Rs("BigClassName") 
        If InStr("|" & ShieldDidList & "|", "|" & Did & "|") = False Then
            Url = GetBigClassUrl(Rs("BigClassName")) 
            ShowDid = FontColorFontB(Did, Rs("FontB"), Rs("FontColor")) 
            Target = Rs("Target") : Target = "_blank"                                       '为新页打开

            C = C & "    <h4><a " & AHref(Url, Did, Target) & ">" & ShowDid & "</a></h4>" & vbCrLf 
            Rss.Open "Select * From [SmallClass] Where BigClassName='" & Rs("BigClassName") & "' Order By Sort", Conn, 1, 1 
            If Not Rss.EOF Then
                C = C & CopyStr(" ", 8) & "<div class=""group"">" & vbCrLf & CopyStr(" ", 12) & "<ul>" & vbCrLf 
                For J = 1 To Rss.RecordCount
                    Sid = Rss("SmallClassName") 
                    If InStr("|" & ShieldSidList & "|", "|" & Sid & "|") = False Then
                        Url = GetSmallClassUrl(Rss("BigClassName"), Rss("SmallClassName")) 
                        ShowSid = FontColorFontB(Sid, Rss("FontB"), Rss("FontColor")) 
                        Target = Rss("Target") : Target = "_blank"                                      '为新页打开

                        C = C & CopyStr(" ", 16) & "<li><a " & AHref(Url, "", Target) & ">" & ShowSid & "</a>" 

                        AddSql = " Where BigClassName='" & Rss("BigClassName") & "' And  SmallClassName='" & Rss("SmallClassName") & "'" 
                        Rsx.Open "Select * From [ThreeClass]" & AddSql & " Order By Sort", Conn, 1, 1 
                        Tstr = "" 
                        For P = 1 To Rsx.RecordCount
                            Tid = Rsx("ThreeClassName")                                                     '为新页打开
                            '判断是否屏蔽子类
                            If InStr("|" & ShieldTidList & "|", "|" & Tid & "|") = False Then
                                Url = GetThreeClassUrl(Did, Sid, Tid) 
                                ShowTid = FontColorFontB(Tid, Rsx("FontB"), Rsx("FontColor")) 
                                Target = Rsx("Target") : Target = "_blank" 
                                Tstr = Tstr & CopyStr(" ", 16) & "<li><a " & AHref(Url, "", Target) & ">" & ShowTid & "</a></li>" & vbCrLf 
                            End If 
                        Rsx.MoveNext : Next : Rsx.Close 
                        If Tstr = "" Then
                            C = C & "</li>" & vbCrLf 
                        Else
                            Tstr = vbCrLf & CopyStr(" ", 16) & "<ul>" & vbCrLf & Tstr & CopyStr(" ", 16) & "</ul>" 
                            C = C & Tstr & CopyStr(" ", 16) & "</li>" & vbCrLf 
                        End If 
                    End If 
                Rss.MoveNext : Next 
                C = C & CopyStr(" ", 12) & "</ul>" & vbCrLf & CopyStr(" ", 12) & "<div class=""clear""></div>" & vbCrLf 
                C = C & CopyStr(" ", 8) & "</div>" & vbCrLf 
            End If : Rss.Close 
        End If 
    Rs.MoveNext : Next : Rs.Close 
    C = C & "</div>" & vbCrLf & "<!--SiteMap end-->" & vbCrLf & vbCrLf 
    SiteMap = C 
End Function

'网页地图
Function SiteMap2(ClassName,NavAddSql, ClassAddSql)
    Dim SplStr, I, J, P,S, C,Url, FileName,BigFolder,SmallFolder,Target,Did,ShowDid,Sid,ShowSid,Tid,ShowTid,MenuYes,MenuList
    C=C & vbCrlf & "<!--SiteMap start-->" & vbCrlf
    C=C & "<div class="""& ClassName &""">" & vbCrlf		
    C=C & "<ul>" & vbCrlf	
    Rs.Open "Select * From [NavBigClass] " & NavAddSql, Conn, 1, 1
	For I = 1 To Rs.RecordCount
        FileName = Rs("FileName") 
        If FileName <> "" Then FileName = FileName & ".html"
		BigFolder = "/" & Rs("FolderName") & "/"
		Did = Rs("BigClassName")
		
		Url = Rs("FileName")
		If CheckRemoteUrl(Url)=False Then			
			Url = BigFolder & FileName
			Url = Replace(Replace(Url, "//", "/"), "//", "/")
			Call WebDebug(Url,"act=Nav&NavDid=" & Did)'调试
		End If
		
		ShowDid = FontColorFontB(Did,Rs("FontB"),Rs("FontColor"))
		Target = Rs("Target") : Target="_blank"
		
		MenuYes = False
		C=C & "	<li><a "& AHref(Url,Did,Target) &">"& ShowDid &"</a>"
		If Rs("NavType") = "文本" Then
			Rss.Open "Select * From [NavSmallClass] Where BigClassName='"& Rs("BigClassName") &"'", Conn, 1, 1
			If Not Rss.Eof Then
				C=C & vbCrlf & CopyStr(" ",8) & "<ul>"& vbCrlf & CopyStr(" ",12) & "<li>"
				For J = 1 To Rss.RecordCount
					FileName = Rss("FileName") 
					If FileName <> "" Then FileName = FileName & ".html" 
					Sid = Rss("SmallClassName")
					
					
					Url = Rss("FileName")
					If CheckRemoteUrl(Url)=False Then	
						Url = BigFolder & FileName
						Url = BigFolder & "/" & Rss("FolderName") & "/" & FileName
						Url = Replace(Replace(Url, "//", "/"), "//", "/")
						Call WebDebug(Url,"act=Nav&NavDid=" & Did & "&NavSid=" & Sid)'调试
					End If
					
					ShowSid = FontColorFontB(Sid,Rss("FontB"),Rss("FontColor")) 
					Target = Rss("Target") : Target="_blank"
					C = C & "<a "& AHref(Url,"",Target) &">" & ShowSid & "</a>" & vbCrlf
				Rss.MoveNext: Next
				C=C & CopyStr(" ",12) & "</li>"& vbCrlf & CopyStr(" ",8) &"</ul>" & vbCrlf
				MenuYes = True
			End If : Rss.Close
		End If
		If MenuYes=True Then
			C=C & CopyStr(" ",4) & "</li><div class=""clear""></div>" & vbCrlf
		Else
			C=C & "</li>" & vbCrlf
		End If
	Rs.MoveNext : Next : Rs.Close
	'类型
    Rs.Open "Select * From [BigClass] " & ClassAddSql, Conn, 1, 1 
	For I = 1 To Rs.RecordCount
        FileName = Rs("FileName") 
        If FileName <> "" Then FileName = FileName & ".html" 
		BigFolder = "/" & Rs("FolderName") & "/"
		Did = Rs("BigClassName")

		Url = Rs("FileName")
		If CheckRemoteUrl(Url)=False Then
			Url = BigFolder & FileName 
			Url = Replace(Replace(Url, "//", "/"), "//", "/")
			Call WebDebug(Url,"act=CreateClass&ProDid=" & Did)'调试
		End If
		
		ShowDid = FontColorFontB(Did,Rs("FontB"),Rs("FontColor")) 
		Target = Rs("Target") : Target="_blank"
		MenuYes = False
		C=C & CopyStr(" ",4) & "<li><a "& AHref(Url,Did,Target) &">"& ShowDid &"</a>" 
		Rss.Open "Select * From [SmallClass] Where BigClassName='"& Rs("BigClassName") &"'", Conn, 1, 1
		If Not Rss.Eof Then			
			C=C & vbCrlf & CopyStr(" ",8) & "<ul>" & vbCrlf
			For J = 1 To Rss.RecordCount
				FileName = Rss("FileName") 
				If FileName <> "" Then FileName = FileName & ".html" 
				SmallFolder = "/" & Rss("FolderName") & "/" 
				Sid = Rss("SmallClassName")

				Url = Rss("FileName")
				If CheckRemoteUrl(Url)=False Then
					Url = BigFolder & "/" & Rss("FolderName") & "/" & FileName 
					Url = Replace(Replace(Url, "//", "/"), "//", "/")
					Call WebDebug(Url,"act=CreateClass&ProDid=" & Did & "&ProSid=" &Sid)'调试
				End If
				
				ShowSid = FontColorFontB(Sid,Rss("FontB"),Rss("FontColor")) 
				Target = Rss("Target") : Target="_blank"
				C = C & CopyStr(" ",12) & "<li><a "& AHref(Url,"",Target) &">" & ShowSid & "</a>" 
				Rsx.Open"Select * From [ThreeClass] Where BigClassName='"& Did &"' And SmallClassName='"& Sid &"' Order By Sort,ID Asc",Conn,1,1
				MenuList = ""
				For P = 1 To Rsx.RecordCount				
					FileName = Rsx("FileName") 
					If FileName <> "" Then FileName = FileName & ".html"  
					Tid = Rsx("ThreeClassName")

					Url = Rsx("FileName")
					If CheckRemoteUrl(Url)=False Then
						Url = BigFolder & SmallFolder & "/" & Rsx("FolderName") & "/" & FileName 
						Url = Replace(Replace(Url, "//", "/"), "//", "/")
						Call WebDebug(Url,"act=CreateClass&ProDid=" & Did & "&ProSid=" & Sid & "&ProTid=" & Tid)'调试
					End If
					
					ShowTid = FontColorFontB(Tid,Rsx("FontB"),Rsx("FontColor")) 
					Target = Rsx("Target") : Target="_blank"
					MenuList = MenuList & "<a "& AHref(Url,"",Target) &">" & ShowTid & "</a>"					 
				Rsx.MoveNext : Next : Rsx.Close 
				If MenuList <> "" Then C=C & "("& MenuList &")"
				C=C & "</li>" & vbCrlf
			Rss.MoveNext: Next
			C=C & CopyStr(" ",8) & "</ul>" & vbCrlf
			MenuYes=True
		End If : Rss.Close
		If MenuYes=True Then
			C=C & CopyStr(" ",4) & "</li><div class=""clear""></div>" & vbCrlf
		Else
			C=C & "</li>" & vbCrlf
		End If 
	Rs.MoveNext : Next : Rs.Close 
		
    C=C & "</ul>" & vbCrlf	
    C=C & "</div>" & vbCrlf
    C=C & "<!--SiteMap end-->" & vbCrlf & vbCrlf
    SiteMap2=C
End Function 
%>