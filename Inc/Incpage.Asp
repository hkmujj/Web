<% 
'dim nCount,nPageSize,maxpage,page,x,i,PageControl 
Function PageControl(nCount, nPageCount, Page, Table_style, Font_style)
    '生成上一页下一页链接
    Dim Query, A, X, Temp, Action, I, C 
    Action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME") 
    Query = Split(Request.ServerVariables("QUERY_STRING"), "&") 
    For Each X In Query
        A = Split(X, "=") 
        If StrComp(A(0), "page", vbTextCompare) <> 0 Then
            Temp = Temp & A(0) & "=" & A(1) & "&" 
        End If 
    Next 
	
    C = C & "<table width=100% border=0 cellpadding=0 cellspacing=0 class=hui>" & vbCrlf 
    C = C & "<TD align=center height=35 class='font'>" & vbCrlf 
    C = C & Font_style & vbCrlf 
    If Page <= 1 Then
        C = C & "首 页 " & vbCrlf 
        C = C & "上一页 " & vbCrlf 
    Else
        C = C & "<A HREF=" & Action & "?" & Temp & "Page=1 class='huilianjie'>首 页</A> " & vbCrlf 
        C = C & "<A HREF=" & Action & "?" & Temp & "Page=" &(Page - 1) & " class='huilianjie'>上一页</A> " & vbCrlf 
    End If 
    If Page >= nPageCount Then
        C = C & "下一页 " & vbCrlf 
        C = C & "尾 页 " & vbCrlf 
    Else
        C = C & "<A HREF=" & Action & "?" & Temp & "Page=" &(Page + 1) & " class='huilianjie'>下一页</A> " & vbCrlf 
        C = C & "<A HREF=" & Action & "?" & Temp & "Page=" & nPageCount & " class='huilianjie'>尾 页</A> " & vbCrlf 
    End If 
    C = C & " 页次：<b><font color=red>" & Page & "</font>/" & nPageCount & "</b>页" & vbCrlf 
    C = C & " 共有<b>" & nCount & "</b>个记录" & vbCrlf 
    If nPageCount < 10 Then
        C = C & "<SELECT onchange='location.href(this.options[this.selectedIndex].value)'>" 
        For I = 1 To nPageCount
            If Page = I Then
                C = C & "<option value=?" & Temp & "page=" & I & " selected> 第" & I & "页</option>" 
            Else
                C = C & "<option value=?" & Temp & "page=" & I & "> 第" & I & "页</option>" 
            End If 
        Next 
        C = C & "</select>" 
    Else
		Temp = Replace(Temp,"\","\\")
        C = C & "<input name=""GotoPage"" value=""" & Page & """ type=""text"" id=""GotoPage"" size=""5"" style=""text-align:center;"" /> &nbsp;" 
        C = C & "<input type=""button"" name=""button"" id=""button"" value=""跳转"" onclick=""location.href('?" & Temp & "Page=' + document.all.GotoPage.value)"" />" 
    End If 
    C = C & "<!--用时--></TD>" & vbCrlf 
    C = C & "</TR>" & vbCrlf 
    C = C & "</table>" & vbCrlf 
    PageControl = C 
End Function 

'页控制  记录总数  每页显示数  当前面 (2015117)   webPageControl(59,12,1,"http://www.baidu.com")
function webPageControl(nRecrodCount, nPageSize, nPage, configPageUrl, action)
	dim s,c,nCountPage,i,nDisplay,nDispalyOK,nTemp
	dim previousPage,nextPage						'定义上一页，下一页
	dim isDisplayTip		'是否显示提示翻页信息	
	isDisplayTip=true
	
	dim sPageStart,sPageEnd,sHomePage,sHomePageFocus,sUpPage,sUpPageFocus,sNextPage,sNextPageFocus,sForPage,sForPageFocus,sTailPage,sTailPageFocus
	if action<>"" then
		sPageStart=getStrCut(action, "[sPageStart]", "[/sPageStart]", 2)			'页头部分
		sPageEnd=getStrCut(action, "[sPageEnd]", "[/sPageEnd]", 2)					'页尾部分
		sHomePage=getStrCut(action, "[sHomePage]", "[/sHomePage]", 2)	'首页
		sHomePageFocus=getStrCut(action, "[sHomePageFocus]", "[/sHomePageFocus]", 2)		'首页交点
		sUpPage=getStrCut(action, "[sUpPage]", "[/sUpPage]", 2)						'上一页
		sUpPageFocus=getStrCut(action, "[sUpPageFocus]", "[/sUpPageFocus]", 2)				'上一页交点
		sNextPage=getStrCut(action, "[sNextPage]", "[/sNextPage]", 2)				'下一页
		sNextPageFocus=getStrCut(action, "[sNextPageFocus]", "[/sNextPageFocus]", 2)			'下一页交点
		sForPage=getStrCut(action, "[sForPage]", "[/sForPage]", 2)							'循环页
		sForPageFocus=getStrCut(action, "[sForPageFocus]", "[/sForPageFocus]", 2)				'循环页交点
		sTailPage=getStrCut(action, "[sTailPage]", "[/sTailPage]", 2)							'最后页
		sTailPageFocus=getStrCut(action, "[sTailPageFocus]", "[/sTailPageFocus]", 2)			'最后页交点
	end if
	'页头部分
	if sPageStart="" then
		sPageStart="<ul class=""pagecontrolwrap"">" & vbcrlf & "<li class=""pageinfo"">共[$nRecrodCount$]条 [$nPage$]/[$nCountPage$]页</li>" & vbcrlf
	end if
	'页尾部分
	if sPageEnd="" then
		sPageEnd="</ul><div class=""clear""></div>" & vbcrlf
	end if
	'首页
	if sHomePage="" then
		sHomePage="<li class=""pageli"">首页</li>" & vbcrlf
	end if
	'首页交点
	if sHomePageFocus="" then
		sHomePageFocus="<li><a href=""[$url$]"">首页</a></li>" & vbcrlf
	end if
	'上一页
	if sUpPage="" then
		sUpPage="<li class=""pageli"">上一页</li>" & vbcrlf
	end if
	'上一页交点
	if sUpPageFocus="" then
		sUpPageFocus="<li><a href=""[$url$]"">上一页</a></li>" & vbcrlf
	end if
	'下一页
	if sNextPage="" then
		sNextPage="<li class=""pageli"">下一页</li>" & vbcrlf
	end if
	'下一页交点
	if sNextPageFocus="" then
		sNextPageFocus="<li><a href=""[$url$]"">下一页</a></li>" & vbcrlf
	end if
	'循环页
	if sForPage="" then
		sForPage="<li><a href=""[$url$]"">[$i$]</a></li>" & vbcrlf
	end if
	'循环页交点
	if sForPageFocus="" then
		sForPageFocus="<li class=""pagefocus"">[$i$]</li>" & vbcrlf
	end if
	'最后页
	if sTailPage="" then
		sTailPage="<li class=""pageli"">末页</li>" & vbcrlf
	end if
	'最后页交点
	if sTailPageFocus="" then
		sTailPageFocus="<li><a href=""[$url$]"">末页</a></li>" & vbcrlf 
	end if
	
	'配置页为空则
    if configPageUrl="" then
        configPageUrl=GetUrlAddToParam( GetUrl(),"?page=[id]","replace")
	end if
	
	nDisplay=6			'显示数
	nDispalyOK=0		'显示成功数
	nPage=HandleNumberType(nPage)
	if nPage="" then 
		nPage=1
	else
		nPage=Cint(nPage)
	end if
	'获得总页数
	nCountPage=GetCountPage(nRecrodCount,nPageSize)
	
	
	
	previousPage=nPage-1
	nextPage=nPage+1
	
	'处理上一页
	if previousPage<=0 then
		previousPage=""
	end if
	'处理下一页
	if nextPage>nCountPage then
		nextPage=""
	end if 
	
	'页开始
	c=sPageStart 
	'首页
	if nPage>1 then
		c=c & replace(sHomePageFocus,"[$url$]",replace(configPageUrl,"[id]",""))
	elseif isDisplayTip=true then
		c=c & sHomePage
	end if
	'上一页
	if previousPage<>"" then
		nTemp=previousPage
		if previousPage<=1 then
			nTemp=""
		end if
		c=c & replace(sUpPageFocus,"[$url$]",replace(configPageUrl,"[id]",nTemp))
	elseif isDisplayTip=true then
		c=c & sUpPage
	end if
 
 	
	dim n
	'call echo(npage,ncountpage)
	n = (nPage-3)
	'call echo("n",n)
	
	'翻页循环
	for i = n to nCountPage
		if i>=1 then
			nDispalyOK=nDispalyOK+1
			'call echo(i,nPage)
			if i = nPage then
				c=c & replace(sForPageFocus,"[$i$]",i)
			else
				nTemp=i
				if i<=1 then
					nTemp=""
				end if
				c=c & replace(replace(sForPage,"[$url$]",replace(configPageUrl,"[id]",nTemp)),"[$i$]",i)
			end if
			if nDispalyOK>nDisplay then
				exit for
			end if
		end if
	next 
	'下一页
	if nCountPage>nPage then
		c=c & replace(sNextPageFocus,"[$url$]",replace(configPageUrl,"[id]",nextPage))
	elseif isDisplayTip=true then
		c=c & sNextPage
	end if
	'末页
	if nCountPage>nPage then
		c=c & replace(sTailPageFocus,"[$url$]",replace(configPageUrl,"[id]",nCountPage))
	elseif isDisplayTip=true then
		c=c & sTailPage
	end if 
	
	c=c & sPageEnd
	
	
	c = replaceValueParam(c, "nRecrodCount", nRecrodCount)
	c = replaceValueParam(c, "nPage", nPage)
	if nCountPage="0" then
		nCountPage=1
	end if
	c = replaceValueParam(c, "nCountPage", nCountPage)
		
	webPageControl=c  & vbcrlf
end function


'获得Rs页数
Function getRsPageNumber(rs, nCount, nPageSize, page)
	dim maxPage
    Dim x 
	x=0
    rs.PageSize = nPageSize                                                         '每页记录条数 可自定义
    maxPage = rs.PageCount
	if maxPage>0 then 
		If Not IsNumeric(page) Or page = "" Then
			page = 1 
		Else
			page = CInt(page) 
		End If 
		If page < 1 Then
			page = 1 
		ElseIf page > maxPage Then
			page = maxPage 
		End If 
		rs.AbsolutePage = page
		If page = maxPage Then
			x = nCount - (maxPage - 1) * nPageSize 
		Else
			x = nPageSize 
		End If 
	end if
    getRsPageNumber = x 
End Function
%> 

