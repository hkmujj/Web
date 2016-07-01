<!--#Include virtual = "/Inc/Config.Asp"--> 
<% 
Dim httpurl : httpurl = "http://www.cncn.com/"                                  '网址
httpurl = "https://www.wpcom.cn/" 
httpurl = "http://www.baocms.com/" 
httpurl = "http://www.9466.com/" 
httpurl="https://blog.dandyweng.com/"
httpurl="http://www.eggboards.com/"
httpurl="http://www.leverandbloomcoffee.co.uk/"
httpurl="http://www.ufoer.com/"
httpurl="https://www.dandyweng.com/"
httpurl="http://www.camarts.cn/"
httpurl="http://www.moke8.com/wordpress/"
httpurl="http://www.jinhusns.com/"
httpurl="http://rednuht.org/genetic_cars_2/"
httpurl="http://www.uiessays.com/"					'NO
httpurl="http://www.5shiguang.net/"
httpurl="https://grails.org/"
httpurl="http://www.zsjujia.com/"				'模仿
Dim newWebDir : newWebDir = "newweb/"                                           '新网站路径
Dim Char_Set : Char_Set = "utf-8"                                               '编码
Dim isGetHttpUrl : isGetHttpUrl = False                                         'getHttpUrl方式获得内容
Dim isMakeWeb : isMakeWeb = False                                               '生成WEB文件夹
Dim isMakeTemplate : isMakeTemplate = True                                      '是否生成模板文件夹
Dim isPackWeb : isPackWeb = False                                               '是否打包WEB文件夹 xml
Dim isPackWebZip : isPackWebZip = False                                               '是否打包WEB文件夹 zip
Dim isWebToTemplateDir : isWebToTemplateDir = True                              '文件是否放到模板文件夹里
Dim templateDir : templateDir = "/templates/"                                   '模板存在目录
Dim isPackTemplate : isPackTemplate = False                                     '是否打包模板文件夹 xml
Dim isPackTemplateZip : isPackTemplateZip = False                                     '是否打包模板文件夹 zip
Dim templateName                                                                '模板文件名称
Dim isFormatting : isFormatting = True                                          '是否格式化HTML
dim isUniformCoding:isUniformCoding=true										'是否统一编码

Dim PubAHrefList, PubATitleList, DownFileToLocaList 

'httpUrl="http://127.0.0.1/1.html"
templateName = getWebSiteName(httpurl)                                          '模板名称从网址中获得

Select Case Request("act")
    Case "downweb" : downWeb()
    Case "download" : downloadFile()
	case "login" : login()
    Case Else : displayDefaultLayout()
End Select
'下载文件
Sub downloadFile()
    Dim downfilePath 
    downfilePath = xorDec(Trim(Request("downfile")), 31380) 
    'call eerr("downfilePath",downfilePath)
    If downfilePath = "" Then
        Call eerr("出错", "下载文件为空") 
    Else
        Call popupDownFile(downfilePath) 
    End If 
End Sub

'下载网站
Sub downweb()
    Dim verificationTime, nSecond ,s
    httpurl = Request("httpurl")                                                    '获得网址
    Char_Set = Request("Char_Set")                                                  '编码
    templateName = LCase(Request("templateName"))                                   '模板名称
    If ishttpurl(httpurl) = False Then
        Call eerr("请输入正确的网站", httpurl) 
    End If 
    isGetHttpUrl = IIF(Request("isGetHttpUrl") = "1", True, False)                  'getHttpUrl方式获得内容
    isMakeWeb = IIF(Request("isMakeWeb") = "1", True, False)                        '生成WEB文件夹
    isMakeTemplate = IIF(Request("isMakeTemplate") = "1", True, False)              '是否生成模板文件夹

    isPackWeb = IIF(Request("isPackWeb") = "1", True, False)                        '是否打包WEB文件夹
    isWebToTemplateDir = IIF(Request("isWebToTemplateDir") = "1", True, False)      '文件是否放到模板文件夹里

    templateDir = Request("templateDir")                                            '模板存在目录
    isWebToTemplateDir = Request("isWebToTemplateDir")                              '文件放到模板文件夹里
    isFormatting = IIF(Request("isFormatting") = "1", True, False)                  '是否格式化HTML
    
	
    verificationTime = Request("verificationTime") 
    If verificationTime = "" Then
        Call eerr("出错", "验证时间为空") 
    End If 

    verificationTime = xorDec(verificationTime, 31380) 
    nSecond = DateDiff("s", verificationTime, Now()) 
    If nSecond >= 180 And getIP <> "127.0.0.1" Then
        Call eerr("时间验证出错 <a href='?'>返回</a>", "nSecond(" & nSecond & ") verificationTime(" & verificationTime & ")") 
    End If 


    Call createfolder(newWebDir) 
    Call createfolder("./htmlweb") 

    httpurl = Request("httpurl") 
    Char_Set = Request("Char_Set") 
	isUniformCoding=request("isUniformCoding")
    newWebDir = newWebDir & setfilename(httpurl) & "/" 

    Call createDirFolder(newWebDir) 
    Call copyWeb() 

    'call rwend(handleConentUrl("/admin/js/", "<script src='aa/js.js' ><script src=""bb/js.js"" >","",""))                '待用

    '生成模板目录
    If Request("isMakeTemplate") = "1" Then
        Call webBeautify("/Templates/" & templateName & "/", "/Templates/" & templateName & "/", True) 
        If request("isWebToTemplateDir") = "1" Then
            Call deleteFolder("/Templates/" & templateName & "/") 
            Call copyFolder(newWebDir & "/Templates/" & templateName & "/", "/Templates/" & templateName & "/") 
			s=newWebDir & "/Templates/" & templateName & "/   ==>>  /Templates/" & templateName & "/"
            Call echoYellowB("放到模板文件夹里", s) 
        End If 
    End If 
    '生成WEB目录
    If Request("isMakeWeb") = "1" Then
        Call webBeautify("/web/", "", False) 
    End If 

    Call echo("成功", "复制网站完成") 
End Sub 
'美化网站 生成js,css分文件夹
Sub webBeautify(rootDir, resourcesDir, isTemplate)
    Dim content, splStr, s,c, fileName, fileExt, filePath, toFilePath, webBeautifyDir, imagesDir, cssDir, jsDir, indexFilePath, toIndexFilePath, indexFileContent 
    Dim cssFileList, splxx, url ,findStr
    webBeautifyDir = newWebDir & rootDir                                            '"/web/"
    imagesDir = webBeautifyDir & "/images/" 
    cssDir = webBeautifyDir & "/css/" 
    jsDir = webBeautifyDir & "/js/" 
    Call deleteFolder(webBeautifyDir) 
    Call createDirFolder(webBeautifyDir)                                            '批量创建文件夹
    Call createfolder(imagesDir) 
    Call createfolder(cssDir) 
    Call createfolder(jsDir) 
    Call echo("webBeautifyDir", webBeautifyDir) 
    Call echo("resourcesDir", resourcesDir) 
    Call echo("jsDir", jsDir) 

    indexFilePath = newWebDir & "/index.html" 
    '为模板
    If isTemplate = True Then
        toIndexFilePath = webBeautifyDir & "/Index_Model.html" 
    Else
        toIndexFilePath = webBeautifyDir & "/index.html" 
    End If 
    indexFileContent = readFile(indexFilePath, Char_Set)                            '读文件

    content = getFileFolderList(newWebDir, True, "全部", "名称", "", "", "") 
    splStr = Split(content, vbCrLf) 
    For Each fileName In splStr
        If fileName <> "" Then
            filePath = newWebDir & "/" & fileName 
            fileExt = LCase(getFileExt(fileName)) 
            If fileExt = "css" Then
                toFilePath = cssDir & fileName 
                'indexFileContent=replace(indexFileContent,fileName, resourcesDir & "css/" & fileName)
                cssFileList = cssFileList & toFilePath & vbCrLf 
            ElseIf fileExt = "js" Then
                toFilePath = jsDir & fileName 
            'indexFileContent=replace(indexFileContent,fileName,resourcesDir & "js/" & fileName)
            ElseIf fileExt = "txt" Or fileExt = "html" Then
                toFilePath = "" 
            ElseIf InStr(fileName, ".") > 0 Then
                toFilePath = imagesDir & fileName 
            'indexFileContent=replace(indexFileContent,fileName,resourcesDir & "images/" & fileName)
            Else
                toFilePath = "" 
            End If 
            'call echo(filePath,tofilepath)
            If toFilePath <> "" Then
				'判断文件大小不为零0    20160612
				if getFSize(filePath)<>0 then
	                Call copyfile(filePath, toFilePath) 
				end if
            End If 
            Call echo(fileName, resourcesDir & "images/" & fileName) 
        End If 
    Next 
    '为模板
    If isTemplate = True Then
        'link里有个favicon.ico  位置会有问题，不管了20160302
        indexFileContent = handleConentUrl("[$cfg_webcss$]/", indexFileContent, "|link|", "", "") 
        indexFileContent = handleConentUrl("[$cfg_webimages$]/", indexFileContent, "|img||embed|param|meta|", "", "") 
        indexFileContent = handleConentUrl("[$cfg_webjs$]/", indexFileContent, "|script|", "", "") 
    Else
        'link里有个favicon.ico  位置会有问题，不管了20160302
        indexFileContent = handleConentUrl("css/", indexFileContent, "|link|", "", "") 
        indexFileContent = handleConentUrl("images/", indexFileContent, "|img||embed|param|meta|", "", "") 
        indexFileContent = handleConentUrl("js/", indexFileContent, "|script|", "", "") 
    End If 
    '为模板，则替换标题与关键词与描述
    If isTemplate = True Then
        indexFileContent = replaceWebTitle(indexFileContent, "{$Web_Title$}") 
        indexFileContent = ReplaceWebMate(indexFileContent, "name", "keywords", " content=""{$Web_KeyWords$}"" ") 
        indexFileContent = ReplaceWebMate(indexFileContent, "name", "description", " content=""{$Web_Description$}"" ") 
        indexFileContent = regExp_Replace(indexFileContent, "utf-8", "gb2312") 
        Char_Set = "gb2312" 
    End If 
    '格式化HTML为真
    If isFormatting = True Then
        indexFileContent = formatting(indexFileContent, "") 
        indexFileContent = htmlFormatting(indexFileContent) 
    End If 
    Call WriteToFile(toIndexFilePath, phptrim(indexFileContent), Char_Set) 



    'call echo("cssFileList",cssFileList)
    'css列表不为空
    If cssFileList <> "" Then
        splxx = Split(cssFileList, vbCrLf) 
        For Each filePath In splxx
            If filePath <> "" Then
                content = readFile(filePath, "") 
			    content = Replace(Replace(content, """", ""), "'", "")                          '去掉点
                If isTemplate = True Then
                    content = regExp_Replace(content, "url\(", "url(" & Request("templateDir") & templateName & "/images/" & "") 
                    content = regExp_Replace(content, "url\(../images/data:image/", "url(data:image/") 
                Else
                    content = regExp_Replace(content, "url\(", "url(../images/") 
                    content = regExp_Replace(content, "url\(../images/data:image/", "url(data:image/") 
                End If 

                Call WriteToFile(filePath, phptrim(content), "") 
            End If 
        Next 
    End If 
    content = getFileFolderList(webBeautifyDir, True, "全部", "", "全部文件夹", "", "") 
    c = Replace(content, vbCrLf, "|") 
    'call rw(content)
    'call eerr(getThisUrlNoFileName(),getthisurl())
    c = c & "|||||" 
    url = getThisUrlNoFileName() & "/myZIP.php?webFolderName=web_&zipDir=htmlweb/&isPrintEchoMsg=0"
    'PHP版打包 网站
    If request("isPackWebZip")="1" or request("isWebToTemplateDir")="1" Then
		If request("isPackWebZip")="1" then
			s=XMLPost(url, "content=" & escape(c) & "&replaceZipFilePath=" & escape(handlePath(newWebDir)) & "\")
		else
			'PHP版打包 模板网站
			s=XMLPost(url, "content=" & escape(c) & "&replaceZipFilePath=" & escape(handlePath(isWebToTemplateDir)) & "\")
		end if
		findStr=strCut(s,"downfile.asp?downfile=","""",2)
		call echo("findStr",findStr)
		'用session方式传下载值
		session("downfilePath")="/" & findStr
		s=replace(s,"downfile.asp?downfile=" & findStr,"/tools/downfile.asp?downfile=session")
        Call echo("", s) 
    End If
	
    If(Request("isPackWeb") = "1" And isTemplate = False) Or(Request("isPackTemplate") = "1" And isTemplate = True) Then
        Dim xmlFileName,xmlSize
        xmlFileName = getIP() & "_update.xml" 

        Dim objXmlZIP : Set objXmlZIP = new xmlZIP
            Call objXmlZIP.run(handlePath(newWebDir), handlePath(newWebDir & rootDir), False, xmlFileName) 
			call echo("1111111111111","注意")
			call echo(handlePath(newWebDir), handlePath(newWebDir & rootDir))
        Set objXmlZIP = Nothing 
		doevents
		xmlSize=getFSize(xmlFileName)
		xmlSize=printSpaceValue(xmlSize)
        Call echo("下载xml打包文件", "<a href=?act=download&downfile=" & xorEnc(xmlFileName, 31380) & " title='点击下载'>点击下载" & xmlFileName & "("& xmlSize &")</a>") 
    End If 
End Sub 
'仿站
Sub copyWeb()
    Dim content, url, splStr, splxx, i, s, tempS, c, sFType, sFName, filePath 
    Dim YuanWebPath, SplTitle 
    If httpurl = "" Then
        Call echo("回显", "网址为空" & httpurl) 
        Exit Sub 
    End If 
    '创建文件夹
    Call createfolder(newWebDir) 
    '下载源码
    YuanWebPath = groupUrl(newWebDir, "/index_网页源码.html") 
    If checkFile(YuanWebPath) = False Then
        Call echo("回显", "正在下载网站源码" & httpurl) 
        If Request("isGetHttpUrl") = "1" Then
            content = getHttpUrl(httpurl, Char_Set) 
            Char_Set = "gb2312" 
            Call WriteToFile(YuanWebPath, content, Char_Set) 
        Else
            Call SaveRemoteFile(httpurl, YuanWebPath)
            content = readFile(YuanWebPath, Char_Set) 
        End If
		'对style中url()图片处理20160630
		if instr(lcase(content),"url(")>0 then
			content=handleCssStyleContent(httpurl, newWebDir,content)
            Call WriteToFile(YuanWebPath, content, Char_Set) 		
		end if

        content = regExp_Replace(content, "<head>", "<head>" & vbCrLf & "<base href=""" & getWebsite(httpurl) & """ />" & vbCrLf) 
        Call WriteToFile(groupUrl(newWebDir, "/index_网页源码_base.html"), content, Char_Set) 

    End If 
    content = readFile(YuanWebPath, Char_Set) 

    '让内容中网址完整
    content = handleConentUrl(httpurl, content, "|*|", PubAHrefList, PubATitleList) 
    '移动JS Css后面参数 如 /base.min.css?v=201510131920  为 /base.min.css
    Call remoteJsCssParam(content, PubAHrefList) 
	'call createfile("1.txt",PubAHrefList)
    '下面
    If PubAHrefList <> "" Then PubAHrefList = Left(PubAHrefList, Len(PubAHrefList) - 2) 
    If PubATitleList <> "" Then PubATitleList = Left(PubATitleList, Len(PubATitleList) - 2)

    Dim FileNameList, nCountArray 
    '下载图片等素材到本地
    If PubAHrefList <> "" Then
        splStr = Split(PubAHrefList, vbCrLf) 
        SplTitle = Split(PubATitleList, vbCrLf) 
        For i = 0 To UBound(splStr)
            s = splStr(i) 
            tempS = LCase(Trim(s)) 
            sFType = CStr(Right(tempS, Len(tempS) - InStrRev(tempS, ".") + 1)) 
            sFName = CStr(Right(tempS, Len(tempS) - InStrRev(tempS, "/"))) 
            If checkUrlFileNameParam(sFName, "js|css|jpg|png|bmp|swf|") = True Then
                sFName = Mid(sFName, 1, InStr(sFName, "?") - 1) 
                If InStr(sFType, "?") > 0 Then
                    sFType = Mid(sFType, 1, InStr(sFType, "?") - 1) 
                End If 
            End If 
            If InStr("|.jpg|.gif|.png|.swf|.js|.css|||", "|" & sFType & "|") > 0 Or InStr(sFName, "?") > 0 Then
                If InStr(sFName, "?") > 0 Then
                    sFName = setFileName(sFName) 
                End If 
                If InStr("|" & FileNameList & "|", "|" & LCase(sFName) & "|") > 0 Then

                    nCountArray = getArrayCount(FileNameList, LCase(sFName) & "|") + 1              '共多少条
                    FileNameList = FileNameList & LCase(sFName) & "|" 
                    'On Error Resume Next
                    If InStr(sFName, ".") > 0 Then
                        sFName = Mid(sFName, 1, InStrRev(sFName, ".") - 1) & "_" & nCountArray & Mid(sFName, InStrRev(sFName, ".")) 
                    End If 
                'MsgBox (nCountArray & vbCrLf & sFName & vbCrLf & fileNameList)
                Else
                    FileNameList = FileNameList & LCase(sFName) & "|" 
                End If
                filePath = newWebDir & "\" & sFName 
                If sFName = "pub.js" Then
                    msgBox(sFName & vbCrLf & s) 
                End If

                'MsgBox ("测试文件名称=" & sFName & vbCrLf & FilePath)
                content = regExp_Replace(content, s, sFName)                             '替换内容中网址
                content = Replace(content, s, sFName)                                    '替换内容中网址
                If checkFile(filePath) = False Then
                    Call echo("回显", "正在下载素材 " & s) 
					doevents
                    Call SaveRemoteFile(s, filePath) 
                    DownFileToLocaList = DownFileToLocaList & sFName & "|"                          '收集下载文件列表
                    If sFType = ".css" Then
                        Call HandleCssFile(s, newWebDir, sFName) 
                    End If 
					
	                If checkFile(filePath) = False Then
						call echoYellow("下载失败",s)
					end if
                End If 
            End If 
            doEvents 
        Next 
    End If 
    url = httpurl 
    url = Mid(url, 1, InStrRev(url, "/")) 
    content = Replace(Replace(content, url, ""), getWebSite(httpurl), "")           '去掉模板网址
    Call WriteToFile(newWebDir & "/index.html", phptrim(content), Char_Set)         '保存模板文件(去两边空格)
    Call createfile(newWebDir & "/图片网址列表.txt", PubAHrefList) 
End Sub 

'处理Css文件内容
Sub handleCssFile(httpurl, folderName, CssFileName)
    Dim content, TempContent, c, startStr, endStr, splStr, i, s, tempS, sFType, sFName, filePath, url, ToPath, Char_Set ,sFName2
    ToPath = folderName & "/" & CssFileName

	if request("isUniformCoding")="1" then
		Char_Set=Request("Char_Set")
	else	
	    Char_Set = checkCode(ToPath) 
	end if
    Call echo(ToPath, Char_Set) 
    content = readFile(ToPath, Char_Set): c = content : TempContent = content 
    content = LCase(content) 
    '重复保存css文件 让utf-8换成gb2312
    If LCase(Char_Set) = "utf-8" Or InStr(content, """utf-8""") > 0 Then
        If InStr(content, """utf-8""") > 0 Then TempContent = regExp_Replace(TempContent, """utf-8""", """gb2312""") 
        Call WriteToFile(ToPath, phptrim(TempContent), "gb2312") 
    End If 

	c=handleCssStyleContent(httpurl, folderName,c)

    Call WriteToFile(ToPath, phptrim(c), Char_Set) 
End Sub
'处理CSS样式内容
function handleCssStyleContent(httpurl, folderName,c)
    Dim content, TempContent, startStr, endStr, splStr, i, s, tempS, sFType, sFName, filePath, url, ToPath, Char_Set ,sFName2
	content=lcase(c)
    startStr = "url\(" 
    endStr = "\)" 
    content = getArray(content, startStr, endStr, False, False) 
    content = Replace(Replace(content, """", ""), "'", "")                          '去掉点
    If content <> "" Then
        splStr = Split(content, "$Array$") 
        For Each s In splStr
			
            tempS = LCase(Trim(s)) 
            sFType = Right(tempS, Len(tempS) - InStrRev(tempS, ".") + 1) 
            sFName = Right(tempS, Len(tempS) - InStrRev(tempS, "/")) 
			'sFName=replace(replace(sFName,"""",""),"'","")			'处理css路径20160612

            url = fullHttpUrl(httpurl, s) 
            If InStr("|.jpg|.gif|.png|.swf|.js|.css|||", "|" & sFType & "|") > 0 Then
				if instr(vbcrlf & PubAHrefList & vbcrlf, vbcrlf & url & vbcrlf)=false then
					PubAHrefList=PubAHrefList & url & vbcrlf
    				Call createfile(newWebDir & "/图片网址列表.txt", PubAHrefList) 		'时时更新
										
					filePath = folderName & "\" & sFName
					'对名称图片处理20160630
					If checkFile(filePath) = true Then
						for i=1 to 10
							sFName2=getFileAttr(sFName,3) & "_" & i  &"."  & getFileAttr(sFName,4)
							if checkFile(sFName2)=false then
								sFName=sFName2
								exit for
							end if
						next
						filePath = folderName & "\" & sFName
					end if
					c = caseInsensitiveReplace(c, s, sFName)                                 '不分大小写替换
					If checkFile(filePath) = False Then
						call echo("下载",url) 
						Call SaveRemoteFile(url, filePath) 
						DownFileToLocaList = DownFileToLocaList & sFName & "|" 
						
					End If 
				end if
            End If 
            doEvents 
        Next 
    End If 
	handleCssStyleContent=c
end function
'替换网站标题
Function replaceWebTitle(content, webTitle)
    Dim TempContent, startStr, endStr, nLen, bodyStart, bodyEnd 
    TempContent = LCase(content) 
    startStr = "<title>" 
    endStr = "</title>" 
    nLen = InStr(TempContent, startStr) 
    If nLen > 0 Then
        bodyStart = Mid(content, 1, nLen - 1)                                           '开始部分
        TempContent = Mid(TempContent, nLen) 
        content = Mid(content, nLen) 
    End If 
    nLen = InStr(TempContent, endStr) 
    If nLen > 0 Then
        bodyEnd = Mid(content, nLen + Len(endStr))                                      '结束部分
    End If 
    replaceWebTitle = bodyStart & startStr & webTitle & endStr & bodyEnd 
End Function 
'替换网站关键词
Function replaceWebMate(content, sArrt, sTypeName, valueStr)
    Dim TempContent, sourceContent, startStr, endStr, nLen, bodyStart, bodyEnd 
    sourceContent = content 
    TempContent = LCase(content) 
    startStr = " " & sArrt & "=""" & sTypeName & """" 
    endStr = "/>" 
    If InStr(TempContent, startStr) = False Then
        startStr = " " & sArrt & "='" & sTypeName & "'" 
    End If 
    nLen = InStr(TempContent, startStr) 
    If nLen > 0 Then
        bodyStart = Mid(content, 1, nLen - 1)                                           '开始部分
        TempContent = Mid(TempContent, nLen) 
        content = Mid(content, nLen) 
    End If 
    nLen = InStr(TempContent, endStr) 
    If nLen > 0 Then
        bodyEnd = Mid(content, nLen + Len(endStr))                                      '结束部分
    End If 
    If bodyStart = "" Or bodyEnd = "" Then
        replaceWebMate = sourceContent 
    Else
        replaceWebMate = bodyStart & startStr & valueStr & endStr & bodyEnd 
    End If 
End Function 
'登录
function login()
	dim password
	password=request("password")
	if password="2016" then
		session("fangzhang")="1"
	end if
	displayDefaultLayout()
end function
'显示默认布局
Sub displayDefaultLayout()
    If getIP() <> "127.0.0.1" Then
        If Session("adminusername") <> "ASPPHPCMS" and session("fangzhang")<>"1" Then
            Response.Write("不可用nolocal 登录后台")
        
%> 
        <form id="form1" name="form1" method="post" action="?act=login">
          密码：
          <input type="text" name="password" id="password" />
          <input type="submit" name="button" id="button" value="提交" />
        </form>
<%    	Response.End() 
    	    End If
	end if
%>
<form action="?act=downweb" method="post" name="form1" target="_blank" id="form1"> 
  网址： 
  <input name="httpurl" type="text" id="httpurl" value="<% =httpurl%>" size="90" /> 
编码： 
<select name="Char_Set" id="Char_Set"> 
  <option value="gb2312" selected="selected">gb2312</option> 
  <option value="utf-8" <% If Char_Set = "utf-8" Then rw("selected")%>>utf-8</option> 
</select> 
模板名称 
<input name="templateName" type="text" id="templateName" value="<% =templateName%>" size="12" /> 
<input type="submit" name="button" id="button" value="提交" /> 
<input name="verificationTime" type="hidden" id="verificationTime" value="<% =xorEnc(Now(), 31380)%>" /> 
<hr /> 
选择下载方式： 
<select name="isGetHttpUrl" id="isGetHttpUrl"> 
  <option value="0" selected="selected">直接下载</option> 
  <option value="1">getHttpUrl方式获得内容</option> 
</select> 
<hr /> 
<label for="isMakeWeb"><input name="isMakeWeb" type="checkbox" id="isMakeWeb" value="1" <% If isMakeWeb = True Then rw("checked")%>/> 
生成WEB文件夹 
<label for="isPackWeb"><input name="isPackWeb" type="checkbox" id="isPackWeb" value="1" <% If isPackWeb = True Then rw("checked")%>/> 
是否打包WEB文件夹(xml) 
</label> 
<label for="isPackWebZip"><input name="isPackWebZip" type="checkbox" id="isPackWebZip" value="1" <% If isPackWebZip = True Then rw("checked")%>/> 
是否打包WEB文件夹(Zip) 
</label> 
<br /><hr /> 
</label> 
<label for="isMakeTemplate"><input name="isMakeTemplate" type="checkbox" id="isMakeTemplate" value="1" <% If isMakeTemplate = True Then rw("checked")%> /> 
是否生成模板文件夹 
</label> 
<label for="isPackTemplate"><input name="isPackTemplate" type="checkbox" id="isPackTemplate" value="1" <% If isPackTemplate = True Then rw("checked")%> /> 
是否打包模板文件夹(xml)
<label for="isPackTemplateZip"><input name="isPackTemplateZip" type="checkbox" id="isPackTemplateZip" value="1" <% If isPackTemplateZip = True Then rw("checked")%> /> 
是否打包模板文件夹(Zip)
</label> 
<br /><hr /> 
模板存在目录 
<input name="templateDir" type="text" id="templateDir" value="<% =templateDir%>" /> 
<label for="isWebToTemplateDir"><input name="isWebToTemplateDir" type="checkbox" id="isWebToTemplateDir" value="1" <% If isWebToTemplateDir = True Then rw("checked")%> />文件放到模板文件夹里</label>  
<hr /> 
<label for="isFormatting"><input name="isFormatting" type="checkbox" id="isFormatting" value="1" <% If isFormatting = True Then rw("checked")%> />是否格式化HTML</label> 
<label for="isUniformCoding"><input name="isUniformCoding" type="checkbox" id="isUniformCoding" value="1" <% If isUniformCoding = True Then rw("checked")%> />是否统一编码</label> 
</form> 



<script type="text/javascript" src="/Jquery/Jquery.Min.js"></script> 
<script type="text/javascript"> 
$(function(){ 
    checkIsMakeWeb() 
    $("#isMakeWeb").click(function(){ 
        checkIsMakeWeb() 
        checkIsFormatting() 
    }) 
    checkIsMakeTemplate() 
    $("#isMakeTemplate").click(function(){ 
        checkIsMakeTemplate() 
        checkIsFormatting() 
    }) 
    //检测是否为操作网站 
    function checkIsMakeWeb(){ 
        if($("#isMakeWeb").is(':checked')) {             
            $("#isPackWeb").removeAttr("disabled");//要变成Enable，JQuery只能这么写   
            $("#isPackWebZip").removeAttr("disabled");//要变成Enable，JQuery只能这么写   
        }else{             
            $("#isPackWeb").attr("disabled","disabled");//再改成disabled  
            $("#isPackWebZip").attr("disabled","disabled");//再改成disabled  
        } 
    } 
    //检测是否为操作模板网站 
    function checkIsMakeTemplate(){ 
        if($("#isMakeTemplate").is(':checked')) {             
            $("#isPackTemplate").removeAttr("disabled");//要变成Enable，JQuery只能这么写   
            $("#isPackTemplateZip").removeAttr("disabled");//要变成Enable，JQuery只能这么写   
            $("#isWebToTemplateDir").removeAttr("disabled");//要变成Enable，JQuery只能这么写   
            $("#templateDir").removeAttr("disabled");//要变成Enable，JQuery只能这么写   
        }else{             
            $("#isPackTemplate").attr("disabled","disabled");//再改成disabled  
            $("#isPackTemplateZip").attr("disabled","disabled");//再改成disabled  
            $("#isWebToTemplateDir").attr("disabled","disabled");//再改成disabled  
            $("#templateDir").attr("disabled","disabled");//再改成disabled  
        } 
    } 
    //判断是否 是否格式化HTML 
    function checkIsFormatting(){ 
        if($("#isMakeWeb").is(':checked') || $("#isMakeTemplate").is(':checked')) { 
            $("#isFormatting").removeAttr("disabled");//要变成Enable，JQuery只能这么写               
        }else{ 
            $("#isFormatting").attr("disabled","disabled");//再改成disabled  
         
        } 
    } 
}) 
</script> 
<!-- 
http://www.thinkphp.cn/ 
--> 


<% 
End Sub

 


%> 


