<!--#Include virtual="/Inc/Config.Asp"-->
<%
dim httpUrl	: httpUrl="http://www.cncn.com/"					'网址
httpurl="https://www.wpcom.cn/"
httpurl="http://www.baocms.com/"
httpurl="http://www.9466.com/"
dim newWebDir  : newWebDir="newweb/"			'新网站路径
dim Char_Set	: Char_Set="utf-8"			'编码
dim isGetHttpUrl : isGetHttpUrl=false				'getHttpUrl方式获得内容
dim isMakeWeb : isMakeWeb=false				'生成WEB文件夹
dim isMakeTemplate : isMakeTemplate=true				'是否生成模板文件夹
dim isPackWeb : isPackWeb=false			'是否打包WEB文件夹
dim isWebToTemplateDir : isWebToTemplateDir=true	 			'文件是否放到模板文件夹里
dim templateDir : templateDir="/templates/"					'模板存在目录
dim isPackTemplate : isPackTemplate=false 	'是否打包模板文件夹
dim templateName							'模板文件名称
dim isFormatting : isFormatting=true							'是否格式化HTML
dim PubAHrefList, PubATitleList,DownFileToLocaList

'httpUrl="http://127.0.0.1/1.html"
templateName=getWebSiteName(httpurl)			'模板名称从网址中获得


select case request("act")
	case "downweb" : downWeb()
	case "download" : downloadFile()
	case else:displayDefaultLayout()
end select
'下载文件
sub downloadFile()
	dim downfilePath
	downfilePath=XorDec(trim(request("downfile")),31380)
	'call eerr("downfilePath",downfilePath)	
	if downfilePath="" then
		call eerr("出错","下载文件为空")
	else
		call popupDownFile(downfilePath)
	end if
end sub

'下载网站
sub downweb()
	dim verificationTime,nSecond
	
		
	httpUrl=request("httpurl")											'获得网址
	Char_Set=request("Char_Set")										'编码
	templateName=lcase(request("templateName"))							'模板名称
	if ishttpurl(httpUrl)=false then
		call eerr("请输入正确的网站", httpUrl)
	end if
	isGetHttpUrl =IIF(request("isGetHttpUrl")="1",true,false)			'getHttpUrl方式获得内容
	isMakeWeb =IIF(request("isMakeWeb")="1",true,false)			' 生成WEB文件夹
	isMakeTemplate =IIF(request("isMakeTemplate")="1",true,false)			'是否生成模板文件夹 
	
	isPackWeb =IIF(request("isPackWeb")="1",true,false)			'是否打包WEB文件夹
	isWebToTemplateDir =IIF(request("isWebToTemplateDir")="1",true,false)			'文件是否放到模板文件夹里
	
	templateDir=request("templateDir")					'模板存在目录
	isWebToTemplateDir=request("isWebToTemplateDir")			'文件放到模板文件夹里 
	isFormatting=IIF(request("isFormatting")="1",true,false)			'是否格式化HTML
	 	 			'
	
	'call die("")
	 
	
	
	verificationTime=request("verificationTime")
	if verificationTime="" then
		call eerr("出错","验证时间为空")
	end if 
	
	verificationTime=XorDec(verificationTime,31380)
	nSecond = DateDiff("s", verificationTime, Now()) 
	if nSecond>=180 and getIP<>"127.0.0.1" then
		call eerr("时间验证出错 <a href='?'>返回</a>", "nSecond("& nSecond &") verificationTime("& verificationTime &")")
	end if
	
	
	call createfolder(newWebDir)
	call createfolder("./htmlweb")
	
	httpurl=request("httpurl")
	Char_Set=request("Char_Set")
	newWebDir= newWebDir & setfilename(httpurl) & "/"

	call createDirFolder(newWebDir)
	call CopyWeb()
	
'call rwend(handleConentUrl("/admin/js/", "<script src='aa/js.js' ><script src=""bb/js.js"" >","",""))				'待用

	
	'生成模板目录
	if request("isMakeTemplate")="1" then
		call webBeautify("/Templates/"& templateName &"/", "/Templates/"& templateName &"/",true)
		if isWebToTemplateDir="1" then
			call deleteFolder("/Templates/"& templateName &"/")
			call copyFolder(newWebDir & "/Templates/"& templateName &"/","/Templates/"& templateName &"/")
			call echo("件放到模板文件夹里","/Templates/"& templateName &"/")
		end if
	end if
	'生成WEB目录
	if request("isMakeWeb")="1" then
		call webBeautify("/web/", "",false)
	end if
	
	call echo("成功","复制网站完成")
end sub
'美化网站 生成js,css分文件夹
sub webBeautify(rootDir,resourcesDir,isTemplate)
	dim content,splstr,i,s,c,fileName,fileExt,filePath,toFilePath,webBeautifyDir,imagesDir,cssDir,jsDir,indexFilePath,toIndexFilePath,indexFileContent
	dim cssFileList,splxx,url,templateContent
	webBeautifyDir=newWebDir & rootDir			'"/web/"
	imagesDir=webBeautifyDir & "/images/"
	cssDir=webBeautifyDir & "/css/"
	jsDir=webBeautifyDir & "/js/"
	call deleteFolder(webBeautifyDir)
	call createDirFolder(webBeautifyDir)			'批量创建文件夹
	call createfolder(imagesDir)
	call createfolder(cssDir)
	call createfolder(jsDir)
	call echo("webBeautifyDir",webBeautifyDir)
	call echo("resourcesDir",resourcesDir)
	call echo("jsDir",jsDir)
	
	indexFilePath=newWebDir & "/index.html"	
	'为模板
	if isTemplate=true then
		toIndexFilePath=webBeautifyDir & "/Index_Model.html" 
	else		
		toIndexFilePath=webBeautifyDir & "/index.html" 
	end if
    indexFileContent = ReadFile(indexFilePath, Char_Set)		'读文件
	
	content = getFileFolderList(newWebDir, True, "全部", "名称", "", "", "")
	splstr=split(content,vbcrlf)
	for each fileName in splstr
		if fileName<>"" then
			filePath=newWebDir & "/" & fileName
			fileExt=Lcase(getFileExt(fileName))
			if fileExt="css" then
				toFilePath=cssDir & fileName
				'indexFileContent=replace(indexFileContent,fileName, resourcesDir & "css/" & fileName)
				cssFileList=cssFileList & toFilePath & vbcrlf
			elseif fileExt="js" then
				toFilePath=jsDir & fileName
				'indexFileContent=replace(indexFileContent,fileName,resourcesDir & "js/" & fileName)
			elseif fileExt="txt" or fileExt="html" then
				tofilePath=""
			elseif instr(fileName,".")>0 then
				toFilePath=imagesDir & fileName
				'indexFileContent=replace(indexFileContent,fileName,resourcesDir & "images/" & fileName)
			else
				toFilePath=""
			end if
			'call echo(filePath,tofilepath)
			if toFilePath<>"" then
				call copyfile(filePath,tofilepath)
			end if
			call echo(fileName,resourcesDir & "images/" & fileName)
		end if
	next
	'为模板
	if isTemplate=true then 
		'link里有个favicon.ico  位置会有问题，不管了20160302
		indexFileContent=handleConentUrl("[$cfg_webcss$]/", indexFileContent,"|link|","","")		
		indexFileContent=handleConentUrl("[$cfg_webimages$]/", indexFileContent,"|img||embed|param|meta|","","")
		indexFileContent=handleConentUrl("[$cfg_webjs$]/", indexFileContent,"|script|","","")
	else
		'link里有个favicon.ico  位置会有问题，不管了20160302
		indexFileContent=handleConentUrl("css/", indexFileContent,"|link|","","")		
		indexFileContent=handleConentUrl("images/", indexFileContent,"|img||embed|param|meta|","","")
		indexFileContent=handleConentUrl("js/", indexFileContent,"|script|","","")
	end if
	'为模板，则替换标题与关键词与描述
	if isTemplate=true then 
		indexFileContent = replaceWebTitle(indexFileContent, "{$Web_Title$}") 
		indexFileContent = ReplaceWebMate(indexFileContent,"name","keywords", " content=""{$Web_KeyWords$}"" ") 
		indexFileContent = ReplaceWebMate(indexFileContent,"name", "description"," content=""{$Web_Description$}"" ")  
		indexFileContent=regExp_Replace(indexFileContent,"utf-8","gb2312")
		Char_Set="gb2312"
	end if
	'格式化HTML为真
	if isFormatting=true then
		indexFileContent=formatting(indexFileContent,"")
		indexFileContent=htmlFormatting(indexFileContent)
	end if
	Call WriteToFile(toIndexFilePath,phptrim(indexFileContent), Char_Set)	
	 
	
	
	'call echo("cssFileList",cssFileList)
	'css列表不为空
	if cssFileList<>"" then
		splxx=split(cssfilelist,vbcrlf)
		for each filePath in splxx
			if filePath <>""then
				content=readFile(filePath,"")
				if isTemplate=true then 
					content=regExp_Replace(content,"url\(","url("& request("templateDir") & templateName &"/images/"  &"")
					content=regExp_Replace(content,"url\(../images/data:image/","url(data:image/")
				else
					content=regExp_Replace(content,"url\(","url(../images/")
					content=regExp_Replace(content,"url\(../images/data:image/","url(data:image/")
				end if
				
				call writeToFile(filePath,phptrim(content),"")
			end if
		next
	end if
	content = getFileFolderList(webBeautifyDir, True, "全部", "", "全部文件夹", "", "")
	c=replace(content,vbcrlf,"|")
	'call rw(content)
' 	call eerr(getThisUrlNoFileName(),getthisurl())
	c=c & "|||||"
	url=getThisUrlNoFileName() & "/myZIP.php?webFolderName=web_&zipDir=htmlweb/&isPrintEchoMsg=0"
	'PHP版打包
	if 1=2 then
		Call Echo("",XMLPost(url, "content=" & escape(c) & "&replaceZipFilePath=" & escape(handlePath(newWebDir)) ))
	end if
	
	
	if (request("isPackWeb")="1" and isTemplate=false) or (request("isPackTemplate")="1" and isTemplate=true) then
	
		
		dim xmlFileName
		xmlFileName=getIP() & "_update.xml"
		call echo("下载xml打包文件","<a href=?act=download&downfile="& XorEnc(xmlFileName,31380) &" title='点击下载'>点击下载"& xmlFileName &"</a>")
		
 
		Dim objXmlZIP : Set objXmlZIP = new xmlZIP  
		call objXmlZIP.run(handlepath(newWebDir), handlePath(newWebDir & rootDir),false,xmlFileName) 
		Set objXmlZIP = Nothing
	end if
end sub
'复制网站
Sub CopyWeb()
    Dim content,  url, splStr, Splxx, SplThree, I, s, tempS, c, LalType, LalStr, nLen, sFType, sFName, FilePath, ImgList
    Dim ReplaceContent, startStr, endStr, ToPath, YuanWebPath, SplTitle
    Dim WebKeywords, WebTitle, WebDescription
    If HttpUrl = "" Then
        Call echo("回显","网址为空" & HttpUrl)
        Exit Sub
    End If
    '创建文件夹
    Call createFolder(newWebDir)
    '下载源码
    YuanWebPath = groupUrl(newWebDir, "/index_网页源码.html")
    If checkFile(YuanWebPath) = False Then
        Call echo("回显","正在下载网站源码" & HttpUrl) 
		if request("isGetHttpUrl")="1" then  
			content=getHttpUrl(HttpUrl,Char_Set)
			Char_Set="gb2312"
		    Call writeToFile(YuanWebPath, content,"gb2312")			
		else		
	        Call SaveRemoteFile(HttpUrl, YuanWebPath)
			content = ReadFile(YuanWebPath, Char_Set)
		end if
		content=regExp_Replace(content,"<head>","<head>" & vbcrlf & "<base href="""& getWebsite(httpUrl) &""" />" & vbcrlf)
		Call writeToFile( groupUrl(newWebDir, "/index_网页源码_base.html"), content,Char_Set)
			
    End If	
    content = ReadFile(YuanWebPath, Char_Set)
	 
   '让内容中网址完整
    content = handleConentUrl(httpurl, content, "|*|", PubAHrefList, PubATitleList)
     '移动JS Css后面参数 如 /base.min.css?v=201510131920  为 /base.min.css
    Call remoteJsCssParam(content, PubAHrefList)
    
    '下面
    If PubAHrefList <> "" Then PubAHrefList = Left(PubAHrefList, Len(PubAHrefList) - 2)
    If PubATitleList <> "" Then PubATitleList = Left(PubATitleList, Len(PubATitleList) - 2)
    
 
    Dim FileNameList, nCountArray
    '下载图片等素材到本地
    If PubAHrefList <> "" Then
        splStr = Split(PubAHrefList, vbCrLf)
        SplTitle = Split(PubATitleList, vbCrLf)
        For I = 0 To UBound(splStr)
            s = splStr(I)
            tempS = LCase(Trim(s))
            sFType = CStr(Right(tempS, Len(tempS) - InStrRev(tempS, ".") + 1))
            sFName = CStr(Right(tempS, Len(tempS) - InStrRev(tempS, "/")))
			if checkUrlFileNameParam(sFName,"js|css|jpg|png|bmp|swf|")=true then
                sFName = Mid(sFName, 1, InStr(sFName, "?") - 1)
                If InStr(sFType, "?") > 0 Then
                    sFType = Mid(sFType, 1, InStr(sFType, "?") - 1)
                End If
            End If
            If InStr("|.jpg|.gif|.png|.swf|.js|.css|||", "|" & sFType & "|") > 0 Or InStr(sFName, "?") > 0 Then
                If InStr(sFName, "?") > 0 Then
                    sFName = SetFileName(sFName)
                End If
                If InStr("|" & FileNameList & "|", "|" & LCase(sFName) & "|") > 0 Then
                    
                    nCountArray = getArrayCount(FileNameList, LCase(sFName) & "|") + 1        '共多少条
                    FileNameList = FileNameList & LCase(sFName) & "|"
                    'On Error Resume Next
                    If InStr(sFName, ".") > 0 Then
                        sFName = Mid(sFName, 1, InStrRev(sFName, ".") - 1) & "_" & nCountArray & Mid(sFName, InStrRev(sFName, "."))
                    End If
                    'MsgBox (nCountArray & vbCrLf & sFName & vbCrLf & fileNameList)
                Else
                    FileNameList = FileNameList & LCase(sFName) & "|"
                End If
                
                
                
                FilePath = newWebDir & "\" & sFName
                If sFName = "pub.js" Then
                    MsgBox (sFName & vbCrLf & s)
                End If
                
                
                'MsgBox ("测试文件名称=" & sFName & vbCrLf & FilePath)
                content = RegExp_Replace(content, s, sFName)        '替换内容中网址
                content = Replace(content, s, sFName)        '替换内容中网址
                If checkFile(FilePath) = False Then
                    Call echo("回显","正在下载素材 " & s)
                    Call SaveRemoteFile(s, FilePath)
                    DownFileToLocaList = DownFileToLocaList & sFName & "|"          '收集下载文件列表
                    If sFType = ".css" Then
                        Call HandleCssFile(s, newWebDir, sFName)
                    End If
                End If
            End If
            DoEvents
        Next
    End If 
    url = httpurl
    url = Mid(url, 1, InStrRev(url, "/"))
    content = Replace(Replace(content, url, ""), getWebSite(httpurl), "")      '去掉模板网址 
    Call WriteToFile(newWebDir & "/index.html", phptrim(content), Char_Set)        '保存模板文件(去两边空格)
	Call writeToFile(newWebDir & "/图片网址列表.txt", PubAHrefList,"gb2312") 
End Sub

'处理Css文件内容
Sub HandleCssFile(HttpUrl, folderName, CssFileName)
    Dim content, TempContent, c, startStr, endStr, splStr, I, s, tempS, sFType, sFName, FilePath, url, ToPath, Char_Set
    ToPath = folderName & "/" & CssFileName 
    Char_Set = CheckCode(ToPath)
	call echo(topath,char_set)
    content = ReadFile(ToPath, Char_Set): c = content: TempContent = content
    content = LCase(content)
    '重复保存css文件 让utf-8换成gb2312
    If LCase(Char_Set) = "utf-8" Or InStr(content, """utf-8""") > 0 Then
        If InStr(content, """utf-8""") > 0 Then TempContent = RegExp_Replace(TempContent, """utf-8""", """gb2312""")
        Call WriteToFile(ToPath, phptrim(TempContent), "gb2312")
    End If
    
    
    startStr = "url\("
    endStr = "\)"
    content = getArray(content, startStr, endStr, False, False)
    content = Replace(Replace(content, """", ""), "'", "")                         '去掉点
   
    If content <> "" Then
        splStr = Split(content, "$Array$")
        For Each s In splStr
		
            tempS = LCase(Trim(s))
            sFType = Right(tempS, Len(tempS) - InStrRev(tempS, ".") + 1)
            sFName = Right(tempS, Len(tempS) - InStrRev(tempS, "/"))
            url = fullHttpUrl(HttpUrl, s)
            If InStr("|.jpg|.gif|.png|.swf|.js|.css|||", "|" & sFType & "|") > 0 Then
                FilePath = folderName & "\" & sFName
                c = CaseInsensitiveReplace(c, s, sFName)            '不分大小写替换
                If checkFile(FilePath) = False Then
                    Call SaveRemoteFile(url, FilePath)
                    DownFileToLocaList = DownFileToLocaList & sFName & "|"
                End If
            End If
            DoEvents
        Next
    End If
    Call WriteToFile(ToPath, phptrim(c), Char_Set)
End Sub 
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
Function replaceWebMate(content,sArrt, sTypeName, valueStr) 
    Dim TempContent, sourceContent, startStr, endStr, nLen, bodyStart, bodyEnd 
    sourceContent = content 
    TempContent = LCase(content) 
    startStr = " "& sArrt &"=""" & sTypeName & """" 
    endStr = "/>" 
    If InStr(TempContent, startStr) = False Then
        startStr = " "& sArrt &"='" & sTypeName & "'" 
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
        replaceWebMate = bodyStart & startStr &  valueStr & endStr & bodyEnd 
    End If 
End Function 
'显示默认布局
sub displayDefaultLayout()
	if GetIP()<>"127.0.0.1"  then
		if Session("adminusername") <> "ASPPHPCMS" then
			response.Write("不可用nolocal")
			response.End()
		end if
	end if
%>
<form action="?act=downweb" method="post" name="form1" target="_blank" id="form1">
  网址：
  <input name="httpurl" type="text" id="httpurl" value="<%=httpurl%>" size="90" />
编码：
<select name="Char_Set" id="Char_Set">
  <option value="gb2312" selected="selected">gb2312</option>
  <option value="utf-8" <%if Char_Set="utf-8" then rw("selected")%>>utf-8</option>
</select>
模板名称
<input name="templateName" type="text" id="templateName" value="<%=templateName%>" size="12" />
<input type="submit" name="button" id="button" value="提交" />
<input name="verificationTime" type="hidden" id="verificationTime" value="<%=XorEnc(now(),31380)%>" />
<hr />
选择下载方式：
<select name="isGetHttpUrl" id="isGetHttpUrl">
  <option value="0" selected="selected">直接下载</option>
  <option value="1">getHttpUrl方式获得内容</option>
</select>
<hr />
<label for="isMakeWeb"><input name="isMakeWeb" type="checkbox" id="isMakeWeb" value="1" <%if isMakeWeb=true then rw("checked")%>/>
生成WEB文件夹
<label for="isPackWeb"><input name="isPackWeb" type="checkbox" id="isPackWeb" value="1" <%if isPackWeb=true then rw("checked")%>/>
是否打包WEB文件夹
</label>
<br /><hr />
</label>
<label for="isMakeTemplate"><input name="isMakeTemplate" type="checkbox" id="isMakeTemplate" value="1" <%if isMakeTemplate=true then rw("checked")%> />
是否生成模板文件夹
</label>
<label for="isPackTemplate"><input name="isPackTemplate" type="checkbox" id="isPackTemplate" value="1" <%if isPackTemplate=true then rw("checked")%> />
是否打包模板文件夹
</label>
<br /><hr />
模板存在目录
<input name="templateDir" type="text" id="templateDir" value="<%=templateDir%>" />
<label for="isWebToTemplateDir"><input name="isWebToTemplateDir" type="checkbox" id="isWebToTemplateDir" value="1" <%if isWebToTemplateDir=true then rw("checked")%> />文件放到模板文件夹里</label> 
<hr />
<label for="isFormatting"><input name="isFormatting" type="checkbox" id="isFormatting" value="1" <%if isFormatting=true then rw("checked")%> />是否格式化HTML</label>



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
		}else{			
			$("#isPackWeb").attr("disabled","disabled");//再改成disabled 
		}
	}
	//检测是否为操作模板网站
	function checkIsMakeTemplate(){
		if($("#isMakeTemplate").is(':checked')) {			
			$("#isPackTemplate").removeAttr("disabled");//要变成Enable，JQuery只能这么写  
			$("#isWebToTemplateDir").removeAttr("disabled");//要变成Enable，JQuery只能这么写  
			$("#templateDir").removeAttr("disabled");//要变成Enable，JQuery只能这么写  
		}else{			
			$("#isPackTemplate").attr("disabled","disabled");//再改成disabled 
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
end sub















Class xmlZIP
    Dim zipPathDir, zipPathFile, zipFileExt 
    Dim starTime, endTime 
    Dim pubXmlZipDir                                                             'xml打包目录
	dim pubIsEchoMsg									'是否回显打印信息
	dim webRootPathDir							'网站根目录路径

    Sub run(thisWebRootPathDir,xmlZipDir,isEchoMsg,xmlFileName)
        pubXmlZipDir = xmlZipDir & "/" 		'xml目录
		pubIsEchoMsg=isEchoMsg				'是否回显
        '此为默认当前文件夹
		webRootPathDir=	thisWebRootPathDir				'等于它
		'call echo("webRootPathDir",webRootPathDir)

        zipPathDir = Server.MapPath("./") & "\" 

        '在此更改要打包文件夹的路径
        'ZipPathDir="D:\MYWEB\WEBINFO"
        '生成的xml文件
		zipPathFile=xmlFileName
		if zipPathFile="" then
	        zipPathFile = "update.xml"
		end if
        '不进行打包的文件扩展名
        zipFileExt = "db;bak" 
        If Right(zipPathDir, 1) <> "\" Then zipPathDir = zipPathDir & "\" 

        Response.Write("保存路径：" & zipPathDir & zipPathFile & "<hr>") 
        '开始打包
        CreateXml(zipPathDir & zipPathFile) 
    End Sub 
    '遍历目录内的所有文件以及文件夹
    Sub loadData(dirPath)
        Dim xmlDoc 
        Dim fso                                                                         'fso对象
        Dim objFolder                                                                   '文件夹对象
        Dim objSubFolders                                                               '子文件夹集合
        Dim objSubFolder                                                                '子文件夹对象
        Dim objFiles                                                                    '文件集合
        Dim objFile                                                                     '文件对象
        Dim objStream 
        Dim pathname, xFolder, xFPath, xFile, xPath, xStream 
        Dim pathNameStr 
		
		if pubIsEchoMsg=true then
	        Response.Write("==========" & dirPath & "==========<br>") 
		end if
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
            Set objFolder = fso.GetFolder(dirPath)                                          '创建文件夹对象
				
				if pubIsEchoMsg=true then
					Response.Write dirPath 
					Response.Flush 
				end if
                Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
                    xmlDoc.load(Server.MapPath(zipPathFile)) 
                    xmlDoc.async = False 
                    '写入每个文件夹路径
                    Set xFolder = xmlDoc.SelectSingleNode("//root").AppendChild(xmlDoc.CreateElement("folder"))
                        Set xFPath = xFolder.AppendChild(xmlDoc.CreateElement("path"))
                            xFPath.text = Replace(dirPath, webRootPathDir, "") 	'zipPathDir  注意
                            Set objFiles = objFolder.files
                                For Each objFile In objFiles
                                    If LCase(dirPath & objFile.Name) <> LCase(Request.ServerVariables("PATH_TRANSLATED")) And LCase(dirPath & objFile.Name) <> LCase(dirPath & zipPathFile) Then
                                        If ext(objFile.Name) Then
                                            pathNameStr = dirPath & "" & objFile.Name  
											if pubIsEchoMsg=true then
                                            	Response.Write "---<br/>" 
												Response.Write pathNameStr & "" 
												Response.Flush 
											end if
                                            '================================================

                                            '写入文件的路径及文件内容
                                            Set xFile = xmlDoc.SelectSingleNode("//root").AppendChild(xmlDoc.CreateElement("file"))
                                                Set xPath = xFile.AppendChild(xmlDoc.CreateElement("path"))
													'call echo(zipPathDir,webRootPathDir) 
                                                    xPath.text = Replace(pathNameStr, webRootPathDir, "") 		'zipPathDir  注意
                                                    '创建文件流读入文件内容，并写入XML文件中
                                                    Set objStream = Server.CreateObject("ADODB.Stream")
                                                        objStream.Type = 1 
                                                        objStream.Open() 
                                                        objStream.loadFromFile(pathNameStr) 
                                                        objStream.position = 0 
                                                        Set xStream = xFile.AppendChild(xmlDoc.CreateElement("stream"))
                                                            xStream.SetAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes" 
                                                            '文件内容采用二制方式存放
                                                            xStream.dataType = "bin.base64" 
                                                            xStream.nodeTypedValue = objStream.read() 
                                                        Set objStream = Nothing 
                                                    Set xPath = Nothing 
                                                Set xStream = Nothing 
                                            Set xFile = Nothing 
                                            '================================================

                                        End If 
                                    End If 
                                Next 
								if pubIsEchoMsg=true then
                                	Response.Write "<p>" 
								end if
                                xmlDoc.Save(Server.MapPath(zipPathFile)) 
                            Set xFPath = Nothing 
                        Set xFolder = Nothing 
                    Set xmlDoc = Nothing 
                    '创建的子文件夹对象
                    Set objSubFolders = objFolder.subFolders
                        '调用递归遍历子文件夹
                        For Each objSubFolder In objSubFolders
                            pathname = dirPath & objSubFolder.Name & "\" 
                            loadData(pathname) 
                        Next 
                    Set objFolder = Nothing 
                Set objSubFolders = Nothing 
            Set fso = Nothing 
    End Sub
    '创建一个空的XML文件，为写入文件作准备
    Sub createXml(filePath)
        '程序开始执行时间
        starTime = Timer() 
        Dim xmlDoc, root 
        Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
            xmlDoc.async = False 
            Set root = xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
                xmlDoc.appendChild(root) 
                xmlDoc.appendChild(xmlDoc.CreateElement("root")) 
                If InStr(filePath, ":") = False Then filePath = Server.MapPath(filePath) 
                xmlDoc.Save(filePath) 
            Set root = Nothing 
        Set xmlDoc = Nothing 
        'call eerr(ZipPathDir & "inc/",pubXmlZipDir)
        'call echo(ZipPathDir & "newweb/http：揦揦www。thinkphp。cn揦/web/",pubXmlZipDir)
        loadData(pubXmlZipDir) 
        '程序结束时间
        endTime = Timer() 
        Response.Write("页面执行时间：" & FormatNumber((endTime - starTime), 3) & "秒") 
    End Sub 
    '判断文件类型是否合法
    Function ext(fileName)
        ext = True 
        Dim temp_ext, e 
        temp_ext = Split(zipFileExt, ";") 
        For e = 0 To UBound(temp_ext)
            If Mid(fileName, InStrRev(fileName, ".") + 1) = temp_ext(e) Then ext = False 
        Next 
    End Function 
End Class

 


%>


