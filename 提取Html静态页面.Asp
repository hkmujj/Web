<!--#Include File = "Inc/Config.Asp"-->
<%
'20151214  增加生成时把文件转成utf-8  和验证码转图片
dim loadingDefaultJsList			'加载默认Js列表
dim webFolderName					'网站目录名称
dim sourceWebFolder					'源网站文件夹路径
dim fileCodeSet						'文件编码
fileCodeSet="utf-8"						'为真，gb2312转utf-8

'把images文件里*.css文件也复制到css/文件夹里  等处理

select case request("act")
	case "copyWebHtml" : call copyWebHtml()
	case else : default()
end select

'loadingDefaultJsList="Jquery.Min.js|MSClass.js|"
'call copyHtmlFileToNewWeb("test3","","|Jquery.Min.js|")

'显示默认内容
sub default()
	dim content,splstr,folderName
	content=GetDirFolderNameList("Templates2015")
	splstr=split(content,vbcrlf)
	for each folderName in splstr
		call echo("文件夹【"& folderName &"】","<a href='?act=copyWebHtml&folderName="& folderName &"'>点击生成</a>")
	next
end sub
'获得网站文件夹名称
function getWebFolderName()
	call openconn()
	rs.Open"select * from [website]",conn,1,1
	if not rs.eof then
		getWebFolderName=rs("WebFolderName")
	end if:rs.close 
end function
'复制网站的HTML
sub copyWebHtml()
	dim folderName,c
	folderName=Rq("folderName"):webFolderName=folderName
	sourceWebFolder= "/htmlweb/" & folderName & "/"
	call CreateDirFolder(sourceWebFolder)
	
	'根据特殊文件夹 复制不一样的js文件列表
	if instr("|test1|","|"& folderName &"|") then
		loadingDefaultJsList="Jquery.Min.js|MSClass.js|"
	elseif instr("|test2|test3|网商时代|A站页面设计|B站页面设计|C站页面设计|天天有返利|","|"& folderName &"|") then
		loadingDefaultJsList="Jquery.Min.js|"
	elseif instr("|秒杀|","|"& folderName &"|") then
		loadingDefaultJsList="Jquery.Min.js|scrollPicLeft.js|" 
	end if
	
	call copyHtmlFileToNewWeb(folderName,"",loadingDefaultJsList)
	
	c="<a href='"& sourceWebFolder &"/Index.Html'>"& sourceWebFolder &"/Index.Html(首页)</a>"
	if checkfile(sourceWebFolder & "/1webmap.html") then
		c=c & " | <a href='"& sourceWebFolder &"/1webmap.html'>"& sourceWebFolder &"/1webmap.html(地图)</a>"
	end if
	call echo("成功","复制网站的HTML完成 " & c)
	
	
	dim backupWebPath
	call createfolder("/TestWeb")
	backupWebPath="/TestWeb/" & getWebFolderName() & Format_Time(now(),12)
	call copyfolder(sourceWebFolder,backupWebPath)
	 
	dim url	
	call makeHtmlWebToZip()		'压缩htmlweb文件夹，里面所有文件的最后修改时间都是对的20150909
	
	'这种打包会把修改时间都改成最新，这个不好20150909
	'用PHP打包程序
	url = "http://127.0.0.1/phpZip.php?toexlude=&dfile%5B%5D=htmlweb&todir=htmlweb&zipname=htmlweb"& Format_Time(now(),12) &".zip&pwd=jb51.net&myaction=dozip"
	'call getHttpUrl(url,"")
	'call echo("url",url)	
	
    C = "<script>" & vbCrlf
    C = C & "setTimeout(""zipWeb()"",2000);" & vbCrlf
    C = C & "function zipWeb(){" & vbCrlf
    C = C & "	document.write(""<iframe src=\"""& url &"\"" width=\""100%\"" height=\""200\""></iframe>"")" & vbCrlf
    C = C & "}" & vbCrlf
    C = C & "</script>" & vbCrlf
	'复制iis配置文件，方便调试
	'call copyFile("\VB工程\Template\htmlweb.config",sourceWebFolder & "\web.config")
end sub
'使htmlWeb文件夹用php压缩
function makeHtmlWebToZip()
	dim content,splstr,filePath,c,fileArray,fileName,fileType,isTrue
	dim cleanFileList			'干净文件列表 为了删除翻页文件
	content = GetFileFolderList(sourceWebFolder, True, "全部", "", "全部文件夹", "", "") 
	splstr=split(content,vbcrlf)
	for each filePath in splstr
		if checkfolder(filePath)=false then
			fileArray=HandleFilePathArray(filePath)
			fileName=lcase(fileArray(2))
			fileType=lcase(fileArray(4))
			fileName=remoteNumber(fileName)
			isTrue=true
			
			if instr("|" & cleanFileList & "|", "|" & fileName & "|")>0 and fileType="html" then
				isTrue=false
			end if
			if isTrue=true then
				'call echo(fileType,fileName)
				if c<>"" then c=c & "|"
				c=c & replace(filePath,HandlePath("/"),"")
				cleanFileList=cleanFileList & fileName & "|"
			end if
		end if
	next
	call rw(c)	
	c=c & "|||||"
	call createfile("htmlweb/1.txt",c)
	call echo("<hr>cccccccccccc",c)
	'Call Echo("",XMLPost("http://127.0.0.1/7.asp", "content=" & escape(c))) 
	Call Echo("",XMLPost("http://127.0.0.1/myZIP.php?webFolderName=" & webFolderName, "content=" & escape(c)))
	'call DeleteFile("htmlweb/1.txt")
End Function

'获得默认JS列表
function getDefaultJsList(jsList,JsDir)
	dim splstr,c,jsName,jsNameList,filePath,toFilePath
	splstr=split(jsList,"|")
	for each jsName in splstr
		jsName=PHPTrim(jsName)
		if jsName <>"" and instr( "|"& jsNameList &"|", "|" & LCase(jsName) & "|")=false then
			jsNameList=jsNameList & LCase(jsName) & "|"
			if c<>"" then c=c & vbcrlf
			c=c & "<script type=""text/javascript"" src=""Js/"& jsName &"""></script>"
 
			filePath="/jquery/" & jsName
			toFilePath=JsDir & "/" & jsName
			call echo("filePath",filePath & " | " & toFilePath)
			call copyfile(filePath,toFilePath)
		end if
	next	
	getDefaultJsList=c
end function
'复制网页
function copyHtmlFileToNewWeb(WebFolderName,sType,jsList)
	dim sourceFolder,toRootFolder,CssDir,ImagesDir,JqueryDir,JsDir,s
	dim splstr,content,filePath,toFilePath,fileName,newFileName
	dim jsListStr
	call echo("WebFolderName",WebFolderName)
	sourceFolder="\Templates2015\"& WebFolderName &"\"
	toRootFolder = sourceWebFolder		'"\htmlweb\"
	
	
	CssDir = toRootFolder & "Css"
	ImagesDir = toRootFolder & "Images"
	JqueryDir = toRootFolder & "Jquery"
	JsDir = toRootFolder & "Js"
	
	call deletefolder(toRootFolder)
	call createfolder(toRootFolder)
	call copyfolder(sourceFolder & "Css", CssDir)
	call copyfolder(sourceFolder & "Images", ImagesDir)
	'call copyfolder(sourceFolder & "Jquery", JqueryDir)
	call deletefolder(ImagesDir & "/文章信息展示样式")
	call deletefolder(ImagesDir & "/文章展示样式")
	call createfolder(JsDir)
		
	dim nErrLevel,editTime
	
	dim cssStr
	content=gethttpurl("http://127.0.0.1/webmap.asp","")
	cssstr="<style>"& getftext("/Images\map\slickmap.css") &"</style>"
	cssstr=CssCompression(cssstr, 0)
	content=replace(content,"<link rel=""stylesheet"" type=""text/css"" media=""screen, print"" href=""images/map/slickmap.css"" />", cssstr)
	
	'移除 它上面的注册登录
	s=GetStrCut(content,"<ul id=""utilityNav"">","</ul>",1)
	if s<>"" then
		content=replace(content,s,"")
	end if
	'这里面设置那些网站可以网站地图(20150923)
	call echo("webFolderName", webFolderName)
	if instr("|B站页面设计|A站页面设计|C站页面设计|", "|" & webFolderName & "|")>0 then
		call deletefile("\wwwroot\"& WebFolderName & "\1webmap.html")		
		content=regexp_replace(content,"href=""http://127.0.0.1/wwwroot/"& WebFolderName &"/""","href=""index.html""") 
		'call rwend(content)
		call createfile("\wwwroot\"& WebFolderName & "\1webmap.html",content)
	end if
	
	jsListStr=getDefaultJsList(jsList,JsDir)
	'处理html
	content=GetDirHtmlList("\wwwroot\"& WebFolderName &"") & vbcrlf & GetDirHtmlList("\html") 
	splstr=split(content,vbcrlf)
	for each filePath in splstr
		'路径不为空则创建
		if filePath<>"" then
			editTime=GetFileEditDate(filePath)
			call echo("filePath",filePath)
			content = getftext(filePath)
			 
			nErrLevel=0
			call HandleHtmlFormatting(content,false,nErrLevel,"|删除空行|")			'|删除空行|
			
			nErrLevel=checkHtmlFormatting(content)
			if checkHtmlFormatting(content)=false then
				call eerr(filepath & "(格式化错误)",nErrLevel)
			end if
			
			'call echo("WebFolderName",WebFolderName)
			content=replace(content,"http://127.0.0.1/wwwroot/"& WebFolderName &"/","")			'追加20150821	
			content=replace(content,"http://127.0.0.1/wwwroot/"& LCase(WebFolderName) &"/","")			'追加20150821			
			content=replace(content,"http://127.0.0.1/Templates/"& WebFolderName &"/","")
			content=replace(content,"http://127.0.0.1/Templates2015/"& WebFolderName &"/","")    '2015
			
			'content=replace(content,"src=""/Jquery/Common.js""","src=""Js/Jquery.Min.js""")
			
			's="<script type=""text/javascript"" src=""Js/Jquery.Min.js"">< /script>" & vbcrlf & "<script type=""text/javascript"" src=""Js/MSClass.js""></ script>"
			content=replace(content,"<script type=""text/javascript"" src=""/Jquery/Common.js""></script>",jsListStr)			'用已经处理好的JS
			 
			content=replace(content,"""/Templates/"& WebFolderName &"/","""")
			content=replace(content,"/Templates2015/"& WebFolderName &"/","")				'2015
			content=replace(content,"http://127.0.0.1/html/","")
			content=replace(content,"http://127.0.0.1/wwwroot/"& WebFolderName &"/","")
			
			content=replace(content,"<script src=""Images/","<script src=""Js/")
			content=replace(content," src=""/Jquery/"," src=""Js/")
			
			'处理验证码
			if instr(content,"Code_7.asp")>0 then
				call copyFile("\VB工程\Images\code.jpg",ImagesDir & "\verification.jpg")
				content=replace(content,"\Inc/Code_7.asp","Images/verification.jpg")
				content=replace(content,"/Inc/Code_7.asp","Images/verification.jpg")
				content=replace(content,"\Inc\Code_7.asp","Images/verification.jpg")
			end if
			
			
			newFileName = toRootFolder & HandleFilePathArray(filePath)(2)
			'把.Html后缀改成.html
			'call echo("newFileName",newFileName)
			if right(lcase(newFileName),5)=".html" then
				newFileName=mid(newFileName,1,len(newFileName)-5) & ".html"
			end if
			
			
			
			'images文件夹css文件放到css文件里
			content = regExpReplaceCssDir(content, "css/")
			content = replaceContentJsDir(content,"Js/","","")
			
			if fileCodeSet="utf-8" then
				content = replace(content,"; charset=gb2312"" />","; charset=utf-8"" />")
			end if
			
			newFileName=lcase(newFileName)		'改小写
			call writeToFile(newFileName,content,fileCodeSet)
			Call ModifyFileModifydate(newFileName, editTime) 
		end if
	next
	'处理CSS目录
	content=GetDirFileList(CssDir,"")
	splstr=split(content,vbcrlf)
	for each filePath in splstr
		call echo("filePath",filePath)		
		content = getftext(filePath)		
		content=CssCompression(content,0)			'删除CSS里注释(20150827)
		
		content=replace(content,"url(/Skins/"& WebFolderName &"/","url(../")
		content=replace(content,"url(/Templates/"& WebFolderName &"/","url(../") 
		content=replace(content,"url(/Templates2015/"& WebFolderName &"/","url(../") 			'2015
	
		if fileCodeSet="utf-8" then
			content = replace(content,"@charset ""gb2312""","@charset ""utf-8""")
		end if
		
		content=cssCompression(content,2)			'压缩CSS
		call writeToFile(filePath,content,fileCodeSet)
	next
	
	'最后处理Images里的Js文件列表，把js文件移到Js文件夹里
	content = GetFileFolderList(ImagesDir, True, "js", "名称", "", "", "") 
	splstr=split(content,vbcrlf)
	for each fileName in splstr
		if fileName<>"" then
			filePath = ImagesDir & "/" & fileName
			toFilePath = JsDir & "/" & fileName
			call moveFile(filePath,toFilePath)
			
			if fileCodeSet="utf-8" then
				if instr("|jquery.min.js|","|"& fileName &"|")=false then
					'不为utf-8 则转为utf-8
					if getFileCharSet(toFilePath)<>"utf-8" then
						content=getftext(toFilePath)					
						call writeToFile(toFilePath,content,fileCodeSet)
					end if
				end if
			end if
			
		end if
	next	
	'最后处理Images里的css文件列表，把css文件移到css文件夹里   上面已经处理了，这里无效20151222
	content = GetFileFolderList(ImagesDir, True, "css", "名称", "", "", "") 
	splstr=split(content,vbcrlf)
	for each fileName in splstr
		if fileName<>"" then
			filePath = ImagesDir & "/" & fileName
			toFilePath = cssDir & "/" & fileName
			'call echo(filePath,"filePath")
			'call echo(toFilePath,"toFilePath")
			call moveFile(filePath,toFilePath)
		end if
	next
	Call ModifyFileModifydate(sourceWebFolder & "\Js\Jquery.Min.js","2009/11/8 10:20:10")
end function 
%>