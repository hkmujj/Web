<!--#Include virtual="/Inc/Config.Asp"-->
<%
dim httpUrl	: httpUrl="http://www.cncn.com/"					'��ַ
httpurl="https://www.wpcom.cn/"
httpurl="http://www.baocms.com/"
httpurl="http://www.9466.com/"
dim newWebDir  : newWebDir="newweb/"			'����վ·��
dim Char_Set	: Char_Set="utf-8"			'����
dim isGetHttpUrl : isGetHttpUrl=false				'getHttpUrl��ʽ�������
dim isMakeWeb : isMakeWeb=false				'����WEB�ļ���
dim isMakeTemplate : isMakeTemplate=true				'�Ƿ�����ģ���ļ���
dim isPackWeb : isPackWeb=false			'�Ƿ���WEB�ļ���
dim isWebToTemplateDir : isWebToTemplateDir=true	 			'�ļ��Ƿ�ŵ�ģ���ļ�����
dim templateDir : templateDir="/templates/"					'ģ�����Ŀ¼
dim isPackTemplate : isPackTemplate=false 	'�Ƿ���ģ���ļ���
dim templateName							'ģ���ļ�����
dim isFormatting : isFormatting=true							'�Ƿ��ʽ��HTML
dim PubAHrefList, PubATitleList,DownFileToLocaList

'httpUrl="http://127.0.0.1/1.html"
templateName=getWebSiteName(httpurl)			'ģ�����ƴ���ַ�л��


select case request("act")
	case "downweb" : downWeb()
	case "download" : downloadFile()
	case else:displayDefaultLayout()
end select
'�����ļ�
sub downloadFile()
	dim downfilePath
	downfilePath=XorDec(trim(request("downfile")),31380)
	'call eerr("downfilePath",downfilePath)	
	if downfilePath="" then
		call eerr("����","�����ļ�Ϊ��")
	else
		call popupDownFile(downfilePath)
	end if
end sub

'������վ
sub downweb()
	dim verificationTime,nSecond
	
		
	httpUrl=request("httpurl")											'�����ַ
	Char_Set=request("Char_Set")										'����
	templateName=lcase(request("templateName"))							'ģ������
	if ishttpurl(httpUrl)=false then
		call eerr("��������ȷ����վ", httpUrl)
	end if
	isGetHttpUrl =IIF(request("isGetHttpUrl")="1",true,false)			'getHttpUrl��ʽ�������
	isMakeWeb =IIF(request("isMakeWeb")="1",true,false)			' ����WEB�ļ���
	isMakeTemplate =IIF(request("isMakeTemplate")="1",true,false)			'�Ƿ�����ģ���ļ��� 
	
	isPackWeb =IIF(request("isPackWeb")="1",true,false)			'�Ƿ���WEB�ļ���
	isWebToTemplateDir =IIF(request("isWebToTemplateDir")="1",true,false)			'�ļ��Ƿ�ŵ�ģ���ļ�����
	
	templateDir=request("templateDir")					'ģ�����Ŀ¼
	isWebToTemplateDir=request("isWebToTemplateDir")			'�ļ��ŵ�ģ���ļ����� 
	isFormatting=IIF(request("isFormatting")="1",true,false)			'�Ƿ��ʽ��HTML
	 	 			'
	
	'call die("")
	 
	
	
	verificationTime=request("verificationTime")
	if verificationTime="" then
		call eerr("����","��֤ʱ��Ϊ��")
	end if 
	
	verificationTime=XorDec(verificationTime,31380)
	nSecond = DateDiff("s", verificationTime, Now()) 
	if nSecond>=180 and getIP<>"127.0.0.1" then
		call eerr("ʱ����֤���� <a href='?'>����</a>", "nSecond("& nSecond &") verificationTime("& verificationTime &")")
	end if
	
	
	call createfolder(newWebDir)
	call createfolder("./htmlweb")
	
	httpurl=request("httpurl")
	Char_Set=request("Char_Set")
	newWebDir= newWebDir & setfilename(httpurl) & "/"

	call createDirFolder(newWebDir)
	call CopyWeb()
	
'call rwend(handleConentUrl("/admin/js/", "<script src='aa/js.js' ><script src=""bb/js.js"" >","",""))				'����

	
	'����ģ��Ŀ¼
	if request("isMakeTemplate")="1" then
		call webBeautify("/Templates/"& templateName &"/", "/Templates/"& templateName &"/",true)
		if isWebToTemplateDir="1" then
			call deleteFolder("/Templates/"& templateName &"/")
			call copyFolder(newWebDir & "/Templates/"& templateName &"/","/Templates/"& templateName &"/")
			call echo("���ŵ�ģ���ļ�����","/Templates/"& templateName &"/")
		end if
	end if
	'����WEBĿ¼
	if request("isMakeWeb")="1" then
		call webBeautify("/web/", "",false)
	end if
	
	call echo("�ɹ�","������վ���")
end sub
'������վ ����js,css���ļ���
sub webBeautify(rootDir,resourcesDir,isTemplate)
	dim content,splstr,i,s,c,fileName,fileExt,filePath,toFilePath,webBeautifyDir,imagesDir,cssDir,jsDir,indexFilePath,toIndexFilePath,indexFileContent
	dim cssFileList,splxx,url,templateContent
	webBeautifyDir=newWebDir & rootDir			'"/web/"
	imagesDir=webBeautifyDir & "/images/"
	cssDir=webBeautifyDir & "/css/"
	jsDir=webBeautifyDir & "/js/"
	call deleteFolder(webBeautifyDir)
	call createDirFolder(webBeautifyDir)			'���������ļ���
	call createfolder(imagesDir)
	call createfolder(cssDir)
	call createfolder(jsDir)
	call echo("webBeautifyDir",webBeautifyDir)
	call echo("resourcesDir",resourcesDir)
	call echo("jsDir",jsDir)
	
	indexFilePath=newWebDir & "/index.html"	
	'Ϊģ��
	if isTemplate=true then
		toIndexFilePath=webBeautifyDir & "/Index_Model.html" 
	else		
		toIndexFilePath=webBeautifyDir & "/index.html" 
	end if
    indexFileContent = ReadFile(indexFilePath, Char_Set)		'���ļ�
	
	content = getFileFolderList(newWebDir, True, "ȫ��", "����", "", "", "")
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
	'Ϊģ��
	if isTemplate=true then 
		'link���и�favicon.ico  λ�û������⣬������20160302
		indexFileContent=handleConentUrl("[$cfg_webcss$]/", indexFileContent,"|link|","","")		
		indexFileContent=handleConentUrl("[$cfg_webimages$]/", indexFileContent,"|img||embed|param|meta|","","")
		indexFileContent=handleConentUrl("[$cfg_webjs$]/", indexFileContent,"|script|","","")
	else
		'link���и�favicon.ico  λ�û������⣬������20160302
		indexFileContent=handleConentUrl("css/", indexFileContent,"|link|","","")		
		indexFileContent=handleConentUrl("images/", indexFileContent,"|img||embed|param|meta|","","")
		indexFileContent=handleConentUrl("js/", indexFileContent,"|script|","","")
	end if
	'Ϊģ�壬���滻������ؼ���������
	if isTemplate=true then 
		indexFileContent = replaceWebTitle(indexFileContent, "{$Web_Title$}") 
		indexFileContent = ReplaceWebMate(indexFileContent,"name","keywords", " content=""{$Web_KeyWords$}"" ") 
		indexFileContent = ReplaceWebMate(indexFileContent,"name", "description"," content=""{$Web_Description$}"" ")  
		indexFileContent=regExp_Replace(indexFileContent,"utf-8","gb2312")
		Char_Set="gb2312"
	end if
	'��ʽ��HTMLΪ��
	if isFormatting=true then
		indexFileContent=formatting(indexFileContent,"")
		indexFileContent=htmlFormatting(indexFileContent)
	end if
	Call WriteToFile(toIndexFilePath,phptrim(indexFileContent), Char_Set)	
	 
	
	
	'call echo("cssFileList",cssFileList)
	'css�б�Ϊ��
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
	content = getFileFolderList(webBeautifyDir, True, "ȫ��", "", "ȫ���ļ���", "", "")
	c=replace(content,vbcrlf,"|")
	'call rw(content)
' 	call eerr(getThisUrlNoFileName(),getthisurl())
	c=c & "|||||"
	url=getThisUrlNoFileName() & "/myZIP.php?webFolderName=web_&zipDir=htmlweb/&isPrintEchoMsg=0"
	'PHP����
	if 1=2 then
		Call Echo("",XMLPost(url, "content=" & escape(c) & "&replaceZipFilePath=" & escape(handlePath(newWebDir)) ))
	end if
	
	
	if (request("isPackWeb")="1" and isTemplate=false) or (request("isPackTemplate")="1" and isTemplate=true) then
	
		
		dim xmlFileName
		xmlFileName=getIP() & "_update.xml"
		call echo("����xml����ļ�","<a href=?act=download&downfile="& XorEnc(xmlFileName,31380) &" title='�������'>�������"& xmlFileName &"</a>")
		
 
		Dim objXmlZIP : Set objXmlZIP = new xmlZIP  
		call objXmlZIP.run(handlepath(newWebDir), handlePath(newWebDir & rootDir),false,xmlFileName) 
		Set objXmlZIP = Nothing
	end if
end sub
'������վ
Sub CopyWeb()
    Dim content,  url, splStr, Splxx, SplThree, I, s, tempS, c, LalType, LalStr, nLen, sFType, sFName, FilePath, ImgList
    Dim ReplaceContent, startStr, endStr, ToPath, YuanWebPath, SplTitle
    Dim WebKeywords, WebTitle, WebDescription
    If HttpUrl = "" Then
        Call echo("����","��ַΪ��" & HttpUrl)
        Exit Sub
    End If
    '�����ļ���
    Call createFolder(newWebDir)
    '����Դ��
    YuanWebPath = groupUrl(newWebDir, "/index_��ҳԴ��.html")
    If checkFile(YuanWebPath) = False Then
        Call echo("����","����������վԴ��" & HttpUrl) 
		if request("isGetHttpUrl")="1" then  
			content=getHttpUrl(HttpUrl,Char_Set)
			Char_Set="gb2312"
		    Call writeToFile(YuanWebPath, content,"gb2312")			
		else		
	        Call SaveRemoteFile(HttpUrl, YuanWebPath)
			content = ReadFile(YuanWebPath, Char_Set)
		end if
		content=regExp_Replace(content,"<head>","<head>" & vbcrlf & "<base href="""& getWebsite(httpUrl) &""" />" & vbcrlf)
		Call writeToFile( groupUrl(newWebDir, "/index_��ҳԴ��_base.html"), content,Char_Set)
			
    End If	
    content = ReadFile(YuanWebPath, Char_Set)
	 
   '����������ַ����
    content = handleConentUrl(httpurl, content, "|*|", PubAHrefList, PubATitleList)
     '�ƶ�JS Css������� �� /base.min.css?v=201510131920  Ϊ /base.min.css
    Call remoteJsCssParam(content, PubAHrefList)
    
    '����
    If PubAHrefList <> "" Then PubAHrefList = Left(PubAHrefList, Len(PubAHrefList) - 2)
    If PubATitleList <> "" Then PubATitleList = Left(PubATitleList, Len(PubATitleList) - 2)
    
 
    Dim FileNameList, nCountArray
    '����ͼƬ���زĵ�����
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
                    
                    nCountArray = getArrayCount(FileNameList, LCase(sFName) & "|") + 1        '��������
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
                
                
                'MsgBox ("�����ļ�����=" & sFName & vbCrLf & FilePath)
                content = RegExp_Replace(content, s, sFName)        '�滻��������ַ
                content = Replace(content, s, sFName)        '�滻��������ַ
                If checkFile(FilePath) = False Then
                    Call echo("����","���������ز� " & s)
                    Call SaveRemoteFile(s, FilePath)
                    DownFileToLocaList = DownFileToLocaList & sFName & "|"          '�ռ������ļ��б�
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
    content = Replace(Replace(content, url, ""), getWebSite(httpurl), "")      'ȥ��ģ����ַ 
    Call WriteToFile(newWebDir & "/index.html", phptrim(content), Char_Set)        '����ģ���ļ�(ȥ���߿ո�)
	Call writeToFile(newWebDir & "/ͼƬ��ַ�б�.txt", PubAHrefList,"gb2312") 
End Sub

'����Css�ļ�����
Sub HandleCssFile(HttpUrl, folderName, CssFileName)
    Dim content, TempContent, c, startStr, endStr, splStr, I, s, tempS, sFType, sFName, FilePath, url, ToPath, Char_Set
    ToPath = folderName & "/" & CssFileName 
    Char_Set = CheckCode(ToPath)
	call echo(topath,char_set)
    content = ReadFile(ToPath, Char_Set): c = content: TempContent = content
    content = LCase(content)
    '�ظ�����css�ļ� ��utf-8����gb2312
    If LCase(Char_Set) = "utf-8" Or InStr(content, """utf-8""") > 0 Then
        If InStr(content, """utf-8""") > 0 Then TempContent = RegExp_Replace(TempContent, """utf-8""", """gb2312""")
        Call WriteToFile(ToPath, phptrim(TempContent), "gb2312")
    End If
    
    
    startStr = "url\("
    endStr = "\)"
    content = getArray(content, startStr, endStr, False, False)
    content = Replace(Replace(content, """", ""), "'", "")                         'ȥ����
   
    If content <> "" Then
        splStr = Split(content, "$Array$")
        For Each s In splStr
		
            tempS = LCase(Trim(s))
            sFType = Right(tempS, Len(tempS) - InStrRev(tempS, ".") + 1)
            sFName = Right(tempS, Len(tempS) - InStrRev(tempS, "/"))
            url = fullHttpUrl(HttpUrl, s)
            If InStr("|.jpg|.gif|.png|.swf|.js|.css|||", "|" & sFType & "|") > 0 Then
                FilePath = folderName & "\" & sFName
                c = CaseInsensitiveReplace(c, s, sFName)            '���ִ�Сд�滻
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
'�滻��վ����
Function replaceWebTitle(content, webTitle)
    Dim TempContent, startStr, endStr, nLen, bodyStart, bodyEnd 
    TempContent = LCase(content) 
    startStr = "<title>" 
    endStr = "</title>" 
    nLen = InStr(TempContent, startStr) 
    If nLen > 0 Then
        bodyStart = Mid(content, 1, nLen - 1)                                           '��ʼ����
        TempContent = Mid(TempContent, nLen) 
        content = Mid(content, nLen) 
    End If 
    nLen = InStr(TempContent, endStr) 
    If nLen > 0 Then
        bodyEnd = Mid(content, nLen + Len(endStr))                                      '��������
    End If 
    replaceWebTitle = bodyStart & startStr & webTitle & endStr & bodyEnd 
End Function 
'�滻��վ�ؼ���
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
        bodyStart = Mid(content, 1, nLen - 1)                                           '��ʼ����
        TempContent = Mid(TempContent, nLen) 
        content = Mid(content, nLen) 
    End If 
    nLen = InStr(TempContent, endStr) 
    If nLen > 0 Then
        bodyEnd = Mid(content, nLen + Len(endStr))                                      '��������
    End If 
    If bodyStart = "" Or bodyEnd = "" Then
        replaceWebMate = sourceContent 
    Else
        replaceWebMate = bodyStart & startStr &  valueStr & endStr & bodyEnd 
    End If 
End Function 
'��ʾĬ�ϲ���
sub displayDefaultLayout()
	if GetIP()<>"127.0.0.1"  then
		if Session("adminusername") <> "ASPPHPCMS" then
			response.Write("������nolocal")
			response.End()
		end if
	end if
%>
<form action="?act=downweb" method="post" name="form1" target="_blank" id="form1">
  ��ַ��
  <input name="httpurl" type="text" id="httpurl" value="<%=httpurl%>" size="90" />
���룺
<select name="Char_Set" id="Char_Set">
  <option value="gb2312" selected="selected">gb2312</option>
  <option value="utf-8" <%if Char_Set="utf-8" then rw("selected")%>>utf-8</option>
</select>
ģ������
<input name="templateName" type="text" id="templateName" value="<%=templateName%>" size="12" />
<input type="submit" name="button" id="button" value="�ύ" />
<input name="verificationTime" type="hidden" id="verificationTime" value="<%=XorEnc(now(),31380)%>" />
<hr />
ѡ�����ط�ʽ��
<select name="isGetHttpUrl" id="isGetHttpUrl">
  <option value="0" selected="selected">ֱ������</option>
  <option value="1">getHttpUrl��ʽ�������</option>
</select>
<hr />
<label for="isMakeWeb"><input name="isMakeWeb" type="checkbox" id="isMakeWeb" value="1" <%if isMakeWeb=true then rw("checked")%>/>
����WEB�ļ���
<label for="isPackWeb"><input name="isPackWeb" type="checkbox" id="isPackWeb" value="1" <%if isPackWeb=true then rw("checked")%>/>
�Ƿ���WEB�ļ���
</label>
<br /><hr />
</label>
<label for="isMakeTemplate"><input name="isMakeTemplate" type="checkbox" id="isMakeTemplate" value="1" <%if isMakeTemplate=true then rw("checked")%> />
�Ƿ�����ģ���ļ���
</label>
<label for="isPackTemplate"><input name="isPackTemplate" type="checkbox" id="isPackTemplate" value="1" <%if isPackTemplate=true then rw("checked")%> />
�Ƿ���ģ���ļ���
</label>
<br /><hr />
ģ�����Ŀ¼
<input name="templateDir" type="text" id="templateDir" value="<%=templateDir%>" />
<label for="isWebToTemplateDir"><input name="isWebToTemplateDir" type="checkbox" id="isWebToTemplateDir" value="1" <%if isWebToTemplateDir=true then rw("checked")%> />�ļ��ŵ�ģ���ļ�����</label> 
<hr />
<label for="isFormatting"><input name="isFormatting" type="checkbox" id="isFormatting" value="1" <%if isFormatting=true then rw("checked")%> />�Ƿ��ʽ��HTML</label>



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
	//����Ƿ�Ϊ������վ
	function checkIsMakeWeb(){
		if($("#isMakeWeb").is(':checked')) {			
			$("#isPackWeb").removeAttr("disabled");//Ҫ���Enable��JQueryֻ����ôд  
		}else{			
			$("#isPackWeb").attr("disabled","disabled");//�ٸĳ�disabled 
		}
	}
	//����Ƿ�Ϊ����ģ����վ
	function checkIsMakeTemplate(){
		if($("#isMakeTemplate").is(':checked')) {			
			$("#isPackTemplate").removeAttr("disabled");//Ҫ���Enable��JQueryֻ����ôд  
			$("#isWebToTemplateDir").removeAttr("disabled");//Ҫ���Enable��JQueryֻ����ôд  
			$("#templateDir").removeAttr("disabled");//Ҫ���Enable��JQueryֻ����ôд  
		}else{			
			$("#isPackTemplate").attr("disabled","disabled");//�ٸĳ�disabled 
			$("#isWebToTemplateDir").attr("disabled","disabled");//�ٸĳ�disabled 
			$("#templateDir").attr("disabled","disabled");//�ٸĳ�disabled 
		}
	}
	//�ж��Ƿ� �Ƿ��ʽ��HTML
	function checkIsFormatting(){
		if($("#isMakeWeb").is(':checked') || $("#isMakeTemplate").is(':checked')) {
			$("#isFormatting").removeAttr("disabled");//Ҫ���Enable��JQueryֻ����ôд  			
		}else{
			$("#isFormatting").attr("disabled","disabled");//�ٸĳ�disabled 
		
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
    Dim pubXmlZipDir                                                             'xml���Ŀ¼
	dim pubIsEchoMsg									'�Ƿ���Դ�ӡ��Ϣ
	dim webRootPathDir							'��վ��Ŀ¼·��

    Sub run(thisWebRootPathDir,xmlZipDir,isEchoMsg,xmlFileName)
        pubXmlZipDir = xmlZipDir & "/" 		'xmlĿ¼
		pubIsEchoMsg=isEchoMsg				'�Ƿ����
        '��ΪĬ�ϵ�ǰ�ļ���
		webRootPathDir=	thisWebRootPathDir				'������
		'call echo("webRootPathDir",webRootPathDir)

        zipPathDir = Server.MapPath("./") & "\" 

        '�ڴ˸���Ҫ����ļ��е�·��
        'ZipPathDir="D:\MYWEB\WEBINFO"
        '���ɵ�xml�ļ�
		zipPathFile=xmlFileName
		if zipPathFile="" then
	        zipPathFile = "update.xml"
		end if
        '�����д�����ļ���չ��
        zipFileExt = "db;bak" 
        If Right(zipPathDir, 1) <> "\" Then zipPathDir = zipPathDir & "\" 

        Response.Write("����·����" & zipPathDir & zipPathFile & "<hr>") 
        '��ʼ���
        CreateXml(zipPathDir & zipPathFile) 
    End Sub 
    '����Ŀ¼�ڵ������ļ��Լ��ļ���
    Sub loadData(dirPath)
        Dim xmlDoc 
        Dim fso                                                                         'fso����
        Dim objFolder                                                                   '�ļ��ж���
        Dim objSubFolders                                                               '���ļ��м���
        Dim objSubFolder                                                                '���ļ��ж���
        Dim objFiles                                                                    '�ļ�����
        Dim objFile                                                                     '�ļ�����
        Dim objStream 
        Dim pathname, xFolder, xFPath, xFile, xPath, xStream 
        Dim pathNameStr 
		
		if pubIsEchoMsg=true then
	        Response.Write("==========" & dirPath & "==========<br>") 
		end if
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
            Set objFolder = fso.GetFolder(dirPath)                                          '�����ļ��ж���
				
				if pubIsEchoMsg=true then
					Response.Write dirPath 
					Response.Flush 
				end if
                Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
                    xmlDoc.load(Server.MapPath(zipPathFile)) 
                    xmlDoc.async = False 
                    'д��ÿ���ļ���·��
                    Set xFolder = xmlDoc.SelectSingleNode("//root").AppendChild(xmlDoc.CreateElement("folder"))
                        Set xFPath = xFolder.AppendChild(xmlDoc.CreateElement("path"))
                            xFPath.text = Replace(dirPath, webRootPathDir, "") 	'zipPathDir  ע��
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

                                            'д���ļ���·�����ļ�����
                                            Set xFile = xmlDoc.SelectSingleNode("//root").AppendChild(xmlDoc.CreateElement("file"))
                                                Set xPath = xFile.AppendChild(xmlDoc.CreateElement("path"))
													'call echo(zipPathDir,webRootPathDir) 
                                                    xPath.text = Replace(pathNameStr, webRootPathDir, "") 		'zipPathDir  ע��
                                                    '�����ļ��������ļ����ݣ���д��XML�ļ���
                                                    Set objStream = Server.CreateObject("ADODB.Stream")
                                                        objStream.Type = 1 
                                                        objStream.Open() 
                                                        objStream.loadFromFile(pathNameStr) 
                                                        objStream.position = 0 
                                                        Set xStream = xFile.AppendChild(xmlDoc.CreateElement("stream"))
                                                            xStream.SetAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes" 
                                                            '�ļ����ݲ��ö��Ʒ�ʽ���
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
                    '���������ļ��ж���
                    Set objSubFolders = objFolder.subFolders
                        '���õݹ�������ļ���
                        For Each objSubFolder In objSubFolders
                            pathname = dirPath & objSubFolder.Name & "\" 
                            loadData(pathname) 
                        Next 
                    Set objFolder = Nothing 
                Set objSubFolders = Nothing 
            Set fso = Nothing 
    End Sub
    '����һ���յ�XML�ļ���Ϊд���ļ���׼��
    Sub createXml(filePath)
        '����ʼִ��ʱ��
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
        'call echo(ZipPathDir & "newweb/http���X�Xwww��thinkphp��cn�X/web/",pubXmlZipDir)
        loadData(pubXmlZipDir) 
        '�������ʱ��
        endTime = Timer() 
        Response.Write("ҳ��ִ��ʱ�䣺" & FormatNumber((endTime - starTime), 3) & "��") 
    End Sub 
    '�ж��ļ������Ƿ�Ϸ�
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


