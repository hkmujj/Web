<% 

'新的截取字符20160216
Function newGetStrCut(content, title)
    Dim s 
	'这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409 
	if instr(content,vbcrlf)=false then  
		content=replace(content,chr(10),vbcrlf)  
	end if
    If InStr(content, "【/" & title & "】") > 0 Then
        s = ADSql(phptrim(getStrCut(content, "【" & title & "】", "【/" & title & "】", 0))) 
    Else
        s = ADSql(phptrim(getStrCut(content, "【" & title & "】", vbCrLf, 0))) 
    End If 
    newGetStrCut = s 
End Function 


'重置数据库数据
Sub resetAccessData()

	call handlePower("恢复模板数据")						'管理权限处理
	
    Call OpenConn() 
    Dim splStr, i, s, columnname, title, nCount, webdataDir 
    webdataDir = Request("webdataDir") 
    If webdataDir <> "" Then
        If checkFolder(webdataDir) = False Then
            Call eerr("网站数据目录不存在，恢复默认数据未成功", webdataDir) 
        End If 
    Else
        webdataDir = "/Data/WebData/" 
    End If 

    Call echo("提示", "恢复数据完成") 
    Call rw("<hr><a href='../" & EDITORTYPE & "web." & EDITORTYPE & "' target='_blank'>进入首页</a> | <a href=""?"" target='_blank'>进入后台</a>") 

    Dim content, filePath, parentid, author, adddatetime, fileName, bodycontent, webtitle, webkeywords, webdescription, sortrank, labletitle, target 
    Dim websitebottom, webtemplate, webimages, webcss, webjs, flags, websiteurl, splxx, columntype, relatedtags, npagesize, customaurl, nofollow 
    Dim templatepath, isthrough,titlecolor
    Dim showreason, ncomputersearch, nmobliesearch, ncountsearch, ndegree           '竞价表
    Dim displaytitle, aboutcontent, isonhtml                                  '单页表
    Dim columnenname                                                                '导航表
    Dim smallimage, bigImage, bannerimage                                           '文章表
	dim httpurl,price,morepageurl,charset,thispage,countpage,bigclassname
 
    '网站配置
    content = getftext(webdataDir & "/website.txt") 
	'这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409 
	if instr(content,vbcrlf)=false then  
		content=replace(content,chr(10),vbcrlf)  
	end if
    If content <> "" Then
        webtitle = newGetStrCut(content, "webtitle") 
        webkeywords = newGetStrCut(content, "webkeywords") 
        webdescription = newGetStrCut(content, "webdescription") 
        websitebottom = newGetStrCut(content, "websitebottom") 
        webtemplate = newGetStrCut(content, "webtemplate") 
        webimages = newGetStrCut(content, "webimages") 
        webcss = newGetStrCut(content, "webcss") 
        webjs = newGetStrCut(content, "webjs") 
        flags = newGetStrCut(content, "flags") 
        websiteurl = newGetStrCut(content, "websiteurl")  

        If getRecordCount(db_PREFIX & "website", "") = 0 Then
            conn.Execute("insert into " & db_PREFIX & "website(webtitle) values('测试')") 
        End If 

        conn.Execute("update " & db_PREFIX & "website  set webtitle='" & webtitle & "',webkeywords='" & webkeywords & "',webdescription='" & webdescription & "',websitebottom='" & websitebottom & "',webtemplate='" & webtemplate & "',webimages='" & webimages & "',webcss='" & webcss & "',webjs='" & webjs & "',flags='" & flags & "',websiteurl='" & websiteurl & "'") 
    End If 

    '导航
    conn.Execute("delete from " & db_PREFIX & "webcolumn") 
    content = getDirTxtList(webdataDir & "/webcolumn/") 
	content=contentNameSort(content,"")
    splStr = Split(content, vbCrLf) 
    Call hr() 
    For Each filePath In splStr
        fileName = getfilename(filePath) 
        If filePath <> "" And InStr("_#", Left(fileName, 1)) = False Then
            Call echo("导航", filePath) 
            content = getftext(filePath) 
			'这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409 
			if instr(content,vbcrlf)=false then  
				content=replace(content,chr(10),vbcrlf)  
			end if
            splxx = Split(content, vbCrLf & "-------------------------------") 
            For Each s In splxx
                If InStr(s, "【webtitle】") > 0 Then 
					s=s & vbcrlf 
                    webtitle = newGetStrCut(s, "webtitle") 
                    webkeywords = newGetStrCut(s, "webkeywords") 
                    webdescription = newGetStrCut(s, "webdescription") 
					customaurl= newGetStrCut(s, "customaurl") 

                    sortrank = newGetStrCut(s, "sortrank") 
                    If sortrank = "" Then sortrank = 0 
                    fileName = newGetStrCut(s, "filename") 
                    columnname = newGetStrCut(s, "columnname") 
                    columnenname = newGetStrCut(s, "columnenname") 
                    columntype = newGetStrCut(s, "columntype") 
                    flags = newGetStrCut(s, "flags") 
                    parentid = newGetStrCut(s, "parentid") 
					 
                    parentid = phptrim(getColumnId(parentid))  										'可根据栏目名称找到对应ID   不存在为-1
					'call echo("parentid",parentid)
                    labletitle = newGetStrCut(s, "labletitle") 
                    '每页显示条数
                    npagesize = newGetStrCut(s, "npagesize") 
                    If npagesize = "" Then npagesize = 10                                           '默认分页数为10条

                    target = newGetStrCut(s, "target") 

                    smallimage = newGetStrCut(s, "smallimage") 
                    bigImage = newGetStrCut(s, "bigImage") 
                    bannerimage = newGetStrCut(s, "bannerimage") 

                    templatepath = newGetStrCut(s, "templatepath") 


                    bodycontent = newGetStrCut(s, "bodycontent") 
                    bodycontent = contentTranscoding(bodycontent) 
                    '是否启用生成html
                    isonhtml = newGetStrCut(s, "isonhtml") 
                    If isonhtml = "0" Or LCase(isonhtml) = "false" Then
                        isonhtml = 0 
                    Else
                        isonhtml = 1 
                    End If 
                    '是否为nofollow
                    nofollow = newGetStrCut(s, "nofollow") 
                    If nofollow = "1" Or LCase(nofollow) = "true" Then
                        nofollow = 1 
                    Else
                        nofollow = 0 
                    End If 
					'call echo(columnname,nofollow)


                    aboutcontent = newGetStrCut(s, "aboutcontent") 
                    aboutcontent = contentTranscoding(aboutcontent) 

                    bodycontent = newGetStrCut(s, "bodycontent") 
                    bodycontent = contentTranscoding(bodycontent) 

                    conn.Execute("insert into " & db_PREFIX & "webcolumn (webtitle,webkeywords,webdescription,columnname,columnenname,columntype,sortrank,filename,customaurl,flags,parentid,labletitle,aboutcontent,bodycontent,npagesize,isonhtml,nofollow,target,smallimage,bigImage,bannerimage,templatepath) values('" & webtitle & "','" & webkeywords & "','" & webdescription & "','" & columnname & "','" & columnenname & "','" & columntype & "'," & sortrank & ",'" & fileName & "','"& customaurl &"','" & flags & "'," & parentid & ",'" & labletitle & "','" & aboutcontent & "','" & bodycontent & "'," & npagesize & "," & isonhtml & "," & nofollow & ",'" & target & "','" & smallimage & "','" & bigImage & "','" & bannerimage & "','" & templatepath & "')") 
                End If 
            Next 
        End If 
    Next 

    '文章
    conn.Execute("delete from " & db_PREFIX & "articledetail") 
    content = getDirAllFileList(webdataDir & "/articledetail/","txt") 
	content=contentNameSort(content,"") 
    splStr = Split(content, vbCrLf) 
    Call hr() 
    For Each filePath In splStr
        fileName = getfilename(filePath) 
        If filePath <> "" And InStr("_#", Left(fileName, 1)) = False Then
            Call echo("文章", filePath) 
            content = getftext(filePath) 
			'这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409 
			if instr(content,vbcrlf)=false then  
				content=replace(content,chr(10),vbcrlf)  
			end if
            splxx = Split(content, vbCrLf & "-------------------------------") 
            For Each s In splxx
                If InStr(s, "【title】") > 0 Then 
                    s = s & vbCrLf 
                    parentid = newGetStrCut(s, "parentid") 
                    parentid = getColumnId(parentid) 
                    title = newGetStrCut(s, "title") 
                    titlecolor = newGetStrCut(s, "titlecolor") 
                    webtitle = newGetStrCut(s, "webtitle") 
                    webkeywords = newGetStrCut(s, "webkeywords") 
                    webdescription = newGetStrCut(s, "webdescription") 


                    author = newGetStrCut(s, "author") 
                    sortrank = newGetStrCut(s, "sortrank") 
                    If sortrank = "" Then sortrank = 0 
                    adddatetime = newGetStrCut(s, "adddatetime") 
                    fileName = newGetStrCut(s, "filename") 
                    templatepath = newGetStrCut(s, "templatepath") 
                    flags = newGetStrCut(s, "flags") 
                    relatedtags = newGetStrCut(s, "relatedtags") 

                    customaurl = newGetStrCut(s, "customaurl") 
                    target = newGetStrCut(s, "target") 


                    smallimage = newGetStrCut(s, "smallimage") 
                    bigImage = newGetStrCut(s, "bigImage") 
                    bannerimage = newGetStrCut(s, "bannerimage") 
                    labletitle = newGetStrCut(s, "labletitle") 
				 
                    aboutcontent = newGetStrCut(s, "aboutcontent") 
                    aboutcontent = contentTranscoding(aboutcontent) 

                    bodycontent = newGetStrCut(s, "bodycontent") 
                    bodycontent = contentTranscoding(bodycontent) 
                    '是否启用生成html
                    isonhtml = newGetStrCut(s, "isonhtml") 
                    If isonhtml = "0" Or LCase(isonhtml) = "false" Then
                        isonhtml = 0 
                    Else
                        isonhtml = 1 
                    End If 
                    '是否为nofollow
                    nofollow = newGetStrCut(s, "nofollow") 
                    If nofollow = "1" Or LCase(nofollow) = "true" Then
                        nofollow = 1 
                    Else
                        nofollow = 0 
                    End If
					
					'价格
                    price = getDianNumb(newGetStrCut(s, "price"))
					if price="" then
						price=0
					end if					
                    conn.Execute("insert into " & db_PREFIX & "articledetail (parentid,title,titlecolor,webtitle,webkeywords,webdescription,author,sortrank,adddatetime,filename,flags,relatedtags,aboutcontent,bodycontent,updatetime,isonhtml,customaurl,nofollow,target,smallimage,bigImage,bannerimage,templatepath,labletitle,price) values(" & parentid & ",'" & title & "','"& titlecolor &"','" & webtitle & "','" & webkeywords & "','" & webdescription & "','" & author & "'," & sortrank & ",'" & adddatetime & "','" & fileName & "','" & flags & "','" & relatedtags & "','"& aboutcontent &"','" & bodycontent & "','" & Now() & "'," & isonhtml & ",'" & customaurl & "'," & nofollow & ",'" & target & "','" & smallimage & "','" & bigImage & "','" & bannerimage & "','" & templatepath & "','"& labletitle &"',"& price &")") 
                End If 
            Next 
        End If 
    Next 

    '单页
    conn.Execute("delete from " & db_PREFIX & "OnePage") 
    content = getDirTxtList(webdataDir & "/OnePage/") 
	content=contentNameSort(content,"")
    splStr = Split(content, vbCrLf) 
    Call hr() 
    For Each filePath In splStr
        fileName = getfilename(filePath) 
        If filePath <> "" And InStr("_#", Left(fileName, 1)) = False Then
            Call echo("单页", filePath) 
            content = getftext(filePath) 
			'这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409 
			if instr(content,vbcrlf)=false then  
				content=replace(content,chr(10),vbcrlf)  
			end if
            splxx = Split(content, vbCrLf & "-------------------------------") 
            For Each s In splxx
                If InStr(s, "【webkeywords】") > 0 Then 
                    s = s & vbCrLf 
                    title = newGetStrCut(s, "title") 
                    displaytitle = newGetStrCut(s, "displaytitle") 
                    webtitle = newGetStrCut(s, "webtitle") 
                    webkeywords = newGetStrCut(s, "webkeywords") 
                    webdescription = newGetStrCut(s, "webdescription") 



                    adddatetime = newGetStrCut(s, "adddatetime") 
                    fileName = newGetStrCut(s, "filename") 

                    aboutcontent = newGetStrCut(s, "aboutcontent") 

                    aboutcontent = contentTranscoding(aboutcontent) 
                    target = newGetStrCut(s, "target") 
                    templatepath = newGetStrCut(s, "templatepath") 

                    bodycontent = newGetStrCut(s, "bodycontent") 
                    bodycontent = contentTranscoding(bodycontent) 
                    '是否启用生成html
                    isonhtml = newGetStrCut(s, "isonhtml") 
                    If isonhtml = "0" Or LCase(isonhtml) = "false" Then
                        isonhtml = 0 
                    Else
                        isonhtml = 1 
                    End If 
                    '是否为nofollow
                    nofollow = newGetStrCut(s, "nofollow") 
                    If nofollow = "1" Or LCase(nofollow) = "true" Then
                        nofollow = 1 
                    Else
                        nofollow = 0 
                    End If 


                    conn.Execute("insert into " & db_PREFIX & "onepage (title,displaytitle,webtitle,webkeywords,webdescription,adddatetime,filename,isonhtml,aboutcontent,bodycontent,nofollow,target,templatepath) values('" & title & "','" & displaytitle & "','" & webtitle & "','" & webkeywords & "','" & webdescription & "','" & adddatetime & "','" & fileName & "'," & isonhtml & ",'" & aboutcontent & "','" & bodycontent & "'," & nofollow & ",'" & target & "','" & templatepath & "')") 
                End If 
            Next 
        End If 
    Next 

    '竞价
    conn.Execute("delete from " & db_PREFIX & "Bidding") 
    content = getDirTxtList(webdataDir & "/Bidding/") 
	content=contentNameSort(content,"")
    splStr = Split(content, vbCrLf) 
    Call hr() 
    For Each filePath In splStr
        fileName = getfilename(filePath) 
        If filePath <> "" And InStr("_#", Left(fileName, 1)) = False Then
            Call echo("竞价", filePath) 
            content = getftext(filePath) 
			'这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409 
			if instr(content,vbcrlf)=false then  
				content=replace(content,chr(10),vbcrlf)  
			end if
            splxx = Split(content, vbCrLf & "-------------------------------") 
            For Each s In splxx
                If InStr(s, "【webkeywords】") > 0 Then 
					s=s & vbcrlf 
                    webkeywords = newGetStrCut(s, "webkeywords") 
                    showreason = newGetStrCut(s, "showreason") 
                    ncomputersearch = newGetStrCut(s, "ncomputersearch") 
                    nmobliesearch = newGetStrCut(s, "nmobliesearch") 
                    ncountsearch = newGetStrCut(s, "ncountsearch") 
                    ndegree = newGetStrCut(s, "ndegree") 
                    ndegree = getnumber(ndegree) 
                    If ndegree = "" Then
                        ndegree = 0 
                    End If 
                    conn.Execute("insert into " & db_PREFIX & "Bidding (webkeywords,showreason,ncomputersearch,nmobliesearch,ndegree) values('" & webkeywords & "','" & showreason & "'," & ncomputersearch & "," & nmobliesearch & "," & ndegree & ")") 
                End If 
            Next 
        End If 
    Next 

    '搜索统计
    conn.Execute("delete from " & db_PREFIX & "SearchStat") 
    content = getDirTxtList(webdataDir & "/SearchStat/") 
	content=contentNameSort(content,"")
    splStr = Split(content, vbCrLf) 
    Call hr() 
    For Each filePath In splStr
        fileName = getfilename(filePath) 
        If filePath <> "" And InStr("_#", Left(fileName, 1)) = False Then
            Call echo("搜索统计", filePath) 
            content = getftext(filePath) 
			'这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409 
			if instr(content,vbcrlf)=false then  
				content=replace(content,chr(10),vbcrlf)  
			end if
            splxx = Split(content, vbCrLf & "-------------------------------") 
            For Each s In splxx
                If InStr(s, "【title】") > 0 Then 
					s=s & vbcrlf 
                    title = newGetStrCut(s, "title") 
                    webtitle = newGetStrCut(s, "webtitle") 
                    webkeywords = newGetStrCut(s, "webkeywords") 
                    webdescription = newGetStrCut(s, "webdescription") 

                    customaurl = newGetStrCut(s, "customaurl") 
                    target = newGetStrCut(s, "target") 
                    isthrough = newGetStrCut(s, "isthrough")  
                    If isthrough = "0" Or LCase(isthrough) = "false" Then
                        isthrough = 0 
                    Else
                        isthrough = 1 
                    End If 
                    sortrank = newGetStrCut(s, "sortrank") 
                    If sortrank = "" Then sortrank = 0 
                    '是否启用生成html
                    isonhtml = newGetStrCut(s, "isonhtml") 
                    If isonhtml = "0" Or LCase(isonhtml) = "false" Then
                        isonhtml = 0 
                    Else
                        isonhtml = 1 
                    End If 
                    '是否为nofollow
                    nofollow = newGetStrCut(s, "nofollow") 
                    If nofollow = "1" Or LCase(nofollow) = "true" Then
                        nofollow = 1 
                    Else
                        nofollow = 0 
                    End If 
                    'call echo("title",title)
                    conn.Execute("insert into " & db_PREFIX & "SearchStat (title,webtitle,webkeywords,webdescription,customaurl,target,isthrough,sortrank,isonhtml,nofollow) values('" & title & "','" & webtitle & "','" & webkeywords & "','" & webdescription & "','" & customaurl & "','" & target & "'," & isthrough & "," & sortrank & "," & isonhtml & "," & nofollow & ")") 

                End If 
            Next 
        End If 
    Next 
	dim itemid,username,ip,reply,tablename			'评论
    '评论
    conn.Execute("delete from " & db_PREFIX & "TableComment")  
    content = getDirTxtList(webdataDir & "/TableComment/") 
	content=contentNameSort(content,"")
    splStr = Split(content, vbCrLf) 
    Call hr() 
    For Each filePath In splStr
        fileName = getfilename(filePath) 
        If filePath <> "" And InStr("_#", Left(fileName, 1)) = False Then
            Call echo("评论", filePath) 
            content = getftext(filePath) 
			'这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409 
			if instr(content,vbcrlf)=false then  
				content=replace(content,chr(10),vbcrlf)  
			end if
            splxx = Split(content, vbCrLf & "-------------------------------") 
            For Each s In splxx
                If InStr(s, "【title】") > 0 Then 
					s=s & vbcrlf 
					
                    tablename = newGetStrCut(s, "tablename") 
                    title = newGetStrCut(s, "title") 
                    itemid = getArticleId(newGetStrCut(s, "itemid"))
					if itemid="" then itemid=0
					'call echo("itemID",itemID)
                    adddatetime = newGetStrCut(s, "adddatetime") 
                    username = newGetStrCut(s, "username") 
                    ip = newGetStrCut(s, "ip") 
                    bodycontent = newGetStrCut(s, "bodycontent") 
                    reply = newGetStrCut(s, "reply") 
					 
					

                    isthrough = newGetStrCut(s, "isthrough")  
                    If isthrough = "0" Or LCase(isthrough) = "false" Then
                        isthrough = 0 
                    Else
                        isthrough = 1 
                    End If 
                   
				   
				   
                    'call echo("title",title)
                    conn.Execute("insert into " & db_PREFIX & "TableComment (tablename,title,itemid,adddatetime,username,ip,bodycontent,reply,isthrough) values('"& tablename &"','" & title & "',"& itemid &",'"& adddatetime &"','"& username &"','"& ip &"','"& bodycontent &"','"& reply &"',"& isthrough &")") 

                End If 
            Next 
        End If 
    Next 
	
    '友情链接
    conn.Execute("delete from " & db_PREFIX & "FriendLink")  
    content = getDirTxtList(webdataDir & "/FriendLink/") 
	content=contentNameSort(content,"")
    splStr = Split(content, vbCrLf) 
    Call hr() 
    For Each filePath In splStr
        fileName = getfilename(filePath) 
        If filePath <> "" And InStr("_#", Left(fileName, 1)) = False Then
            Call echo("评论", filePath) 
            content = getftext(filePath) 
			'这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409 
			if instr(content,vbcrlf)=false then  
				content=replace(content,chr(10),vbcrlf)  
			end if
            splxx = Split(content, vbCrLf & "-------------------------------") 
            For Each s In splxx
                If InStr(s, "【title】") > 0 Then 
					s=s & vbcrlf 
					
                    title = newGetStrCut(s, "title")  
                    httpurl = newGetStrCut(s, "httpurl") 
                    smallimage = newGetStrCut(s, "smallimage") 
                    flags = newGetStrCut(s, "flags") 
					target= newGetStrCut(s, "target") 
					

                    sortrank = newGetStrCut(s, "sortrank")  
                    If sortrank = "0" Or LCase(sortrank) = "false" Then
                        sortrank = 0 
                    Else
                        sortrank = 1 
                    End If
                    isthrough = newGetStrCut(s, "isthrough")  
                    If isthrough = "0" Or LCase(isthrough) = "false" Then
                        isthrough = 0 
                    Else
                        isthrough = 1 
                    End If  
                    'call echo("title",title)
                    conn.Execute("insert into " & db_PREFIX & "FriendLink (title,httpurl,smallimage,flags,sortrank,isthrough,target) values('"& title &"','" & httpurl & "','"& smallimage &"','"& flags &"',"& sortrank &","& isthrough &",'"& target &"')") 

                End If 
            Next 
        End If 
    Next 
	
    '留言
    conn.Execute("delete from " & db_PREFIX & "GuestBook")  
    content = getDirTxtList(webdataDir & "/GuestBook/") 
	content=contentNameSort(content,"")
    splStr = Split(content, vbCrLf) 
    Call hr() 
    For Each filePath In splStr
        fileName = getfilename(filePath) 
        If filePath <> "" And InStr("_#", Left(fileName, 1)) = False Then
            Call echo("留言", filePath) 
            content = getftext(filePath) 
			'这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409 
			if instr(content,vbcrlf)=false then  
				content=replace(content,chr(10),vbcrlf)  
			end if
            splxx = Split(content, vbCrLf & "-------------------------------") 
            For Each s In splxx
                If InStr(s, "【adddatetime】") > 0 Then 
					s=s & vbcrlf 
					
					adddatetime = newGetStrCut(s, "adddatetime") 					
                    bodycontent = newGetStrCut(s, "bodycontent") 
                    reply = newGetStrCut(s, "reply") 
                    isthrough = newGetStrCut(s, "isthrough")  
                    If isthrough = "0" Or LCase(isthrough) = "false" Then
                        isthrough = 0 
                    Else
                        isthrough = 1 
                    End If  
                    conn.Execute("insert into " & db_PREFIX & "GuestBook (adddatetime,bodycontent,reply,isthrough) values('"& adddatetime &"','" & bodycontent & "','"& reply &"',"& isthrough &")") 

                End If 
            Next 
        End If 
    Next 
	
	
    '采集网站
    conn.Execute("delete from " & db_PREFIX & "CaiWeb")  
    content = getDirTxtList(webdataDir & "/CaiWeb/") 
	content=contentNameSort(content,"")
    splStr = Split(content, vbCrLf) 
    Call hr() 
    For Each filePath In splStr
        fileName = getfilename(filePath) 
        If filePath <> "" And InStr("_#", Left(fileName, 1)) = False Then
            Call echo("采集网站", filePath) 
            content = getftext(filePath) 
			'这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409 
			if instr(content,vbcrlf)=false then  
				content=replace(content,chr(10),vbcrlf)  
			end if
            splxx = Split(content, vbCrLf & "-------------------------------") 
            For Each s In splxx
                If InStr(s, "【bigclassname】") > 0 Then 
					s=s & vbcrlf 
					
					
					bigclassname = newGetStrCut(s, "bigclassname") 
					httpurl = newGetStrCut(s, "httpurl") 
					morepageurl = newGetStrCut(s, "morepageurl") 					
					charset = newGetStrCut(s, "charset") 					
					
					
					adddatetime = newGetStrCut(s, "adddatetime") 					
                    bodycontent = newGetStrCut(s, "bodycontent") 
					
                    sortrank = newGetStrCut(s, "sortrank") 
                    If sortrank = "" Then sortrank = 0 
					
                    thispage = newGetStrCut(s, "thispage") 
                    If thispage = "" Then thispage = 0 
                    countpage = newGetStrCut(s, "countpage") 
                    If countpage = "" Then thispage = 0 
					
					
                    conn.Execute("insert into " & db_PREFIX & "CaiWeb (adddatetime,bodycontent,httpurl,morepageurl,charset,sortrank,thispage,countpage,bigclassname) values('"& adddatetime &"','" & bodycontent & "','"& httpurl &"','"& morepageurl &"','"& charset &"',"& sortrank &","& thispage &","& countpage &",'"& bigclassname &"')") 

                End If 
            Next 
        End If 
    Next 
	
	
	
	call writeSystemLog("","恢复默认数据" & db_PREFIX)			'系统日志

End Sub 

'内容转码
Function contentTranscoding(ByVal content)
    content = Replace(Replace(Replace(Replace(content, "<?", "&lt;?"), "?>", "?&gt;"), "<" & "%", "&lt;%"), "?>", "%&gt;") 


    Dim splStr, i, s, c, isTranscoding, isBR 
    isTranscoding = False 
    isBR = False 
    splStr = Split(content, vbCrLf) 
    For Each s In splStr
        If InStr(s, "[&html转码&]") > 0 Then
            isTranscoding = True 
        End If 
        If InStr(s, "[&html转码end&]") > 0 Then
            isTranscoding = False 
        End If 
        If InStr(s, "[&全部换行&]") > 0 Then
            isBR = True 
        End If 
        If InStr(s, "[&全部换行end&]") > 0 Then
            isBR = False 
        End If 

        If isTranscoding = True Then
            s = Replace(Replace(s, "[&html转码&]", ""), "<", "&lt;") 
        Else
            s = Replace(s, "[&html转码end&]", "") 
        End If 
        If isBR = True Then
            s = Replace(s, "[&全部换行&]", "") & "<br>" 
        Else
            s = Replace(s, "[&全部换行end&]", "") 
        End If 
        c = c & s & vbCrLf 
    Next
	c=replace(replace(c,"【b】","<b>"),"【/b】","</b>")
	c=replace(replace(c,"【strong】","<strong>"),"【/strong】","</strong>")
    contentTranscoding = c 
End Function 
%>   
