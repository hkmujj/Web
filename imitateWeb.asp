<!--#Include virtual = "/Inc/Config.Asp"--> 
<% 
Dim httpurl : httpurl = "http://www.cncn.com/"                                  '��ַ
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
httpurl="http://www.zsjujia.com/"				'ģ��
Dim newWebDir : newWebDir = "newweb/"                                           '����վ·��
Dim Char_Set : Char_Set = "utf-8"                                               '����
Dim isGetHttpUrl : isGetHttpUrl = False                                         'getHttpUrl��ʽ�������
Dim isMakeWeb : isMakeWeb = False                                               '����WEB�ļ���
Dim isMakeTemplate : isMakeTemplate = True                                      '�Ƿ�����ģ���ļ���
Dim isPackWeb : isPackWeb = False                                               '�Ƿ���WEB�ļ��� xml
Dim isPackWebZip : isPackWebZip = False                                               '�Ƿ���WEB�ļ��� zip
Dim isWebToTemplateDir : isWebToTemplateDir = True                              '�ļ��Ƿ�ŵ�ģ���ļ�����
Dim templateDir : templateDir = "/templates/"                                   'ģ�����Ŀ¼
Dim isPackTemplate : isPackTemplate = False                                     '�Ƿ���ģ���ļ��� xml
Dim isPackTemplateZip : isPackTemplateZip = False                                     '�Ƿ���ģ���ļ��� zip
Dim templateName                                                                'ģ���ļ�����
Dim isFormatting : isFormatting = True                                          '�Ƿ��ʽ��HTML
dim isUniformCoding:isUniformCoding=true										'�Ƿ�ͳһ����

Dim PubAHrefList, PubATitleList, DownFileToLocaList 

'httpUrl="http://127.0.0.1/1.html"
templateName = getWebSiteName(httpurl)                                          'ģ�����ƴ���ַ�л��

Select Case Request("act")
    Case "downweb" : downWeb()
    Case "download" : downloadFile()
	case "login" : login()
    Case Else : displayDefaultLayout()
End Select
'�����ļ�
Sub downloadFile()
    Dim downfilePath 
    downfilePath = xorDec(Trim(Request("downfile")), 31380) 
    'call eerr("downfilePath",downfilePath)
    If downfilePath = "" Then
        Call eerr("����", "�����ļ�Ϊ��") 
    Else
        Call popupDownFile(downfilePath) 
    End If 
End Sub

'������վ
Sub downweb()
    Dim verificationTime, nSecond ,s
    httpurl = Request("httpurl")                                                    '�����ַ
    Char_Set = Request("Char_Set")                                                  '����
    templateName = LCase(Request("templateName"))                                   'ģ������
    If ishttpurl(httpurl) = False Then
        Call eerr("��������ȷ����վ", httpurl) 
    End If 
    isGetHttpUrl = IIF(Request("isGetHttpUrl") = "1", True, False)                  'getHttpUrl��ʽ�������
    isMakeWeb = IIF(Request("isMakeWeb") = "1", True, False)                        '����WEB�ļ���
    isMakeTemplate = IIF(Request("isMakeTemplate") = "1", True, False)              '�Ƿ�����ģ���ļ���

    isPackWeb = IIF(Request("isPackWeb") = "1", True, False)                        '�Ƿ���WEB�ļ���
    isWebToTemplateDir = IIF(Request("isWebToTemplateDir") = "1", True, False)      '�ļ��Ƿ�ŵ�ģ���ļ�����

    templateDir = Request("templateDir")                                            'ģ�����Ŀ¼
    isWebToTemplateDir = Request("isWebToTemplateDir")                              '�ļ��ŵ�ģ���ļ�����
    isFormatting = IIF(Request("isFormatting") = "1", True, False)                  '�Ƿ��ʽ��HTML
    
	
    verificationTime = Request("verificationTime") 
    If verificationTime = "" Then
        Call eerr("����", "��֤ʱ��Ϊ��") 
    End If 

    verificationTime = xorDec(verificationTime, 31380) 
    nSecond = DateDiff("s", verificationTime, Now()) 
    If nSecond >= 180 And getIP <> "127.0.0.1" Then
        Call eerr("ʱ����֤���� <a href='?'>����</a>", "nSecond(" & nSecond & ") verificationTime(" & verificationTime & ")") 
    End If 


    Call createfolder(newWebDir) 
    Call createfolder("./htmlweb") 

    httpurl = Request("httpurl") 
    Char_Set = Request("Char_Set") 
	isUniformCoding=request("isUniformCoding")
    newWebDir = newWebDir & setfilename(httpurl) & "/" 

    Call createDirFolder(newWebDir) 
    Call copyWeb() 

    'call rwend(handleConentUrl("/admin/js/", "<script src='aa/js.js' ><script src=""bb/js.js"" >","",""))                '����

    '����ģ��Ŀ¼
    If Request("isMakeTemplate") = "1" Then
        Call webBeautify("/Templates/" & templateName & "/", "/Templates/" & templateName & "/", True) 
        If request("isWebToTemplateDir") = "1" Then
            Call deleteFolder("/Templates/" & templateName & "/") 
            Call copyFolder(newWebDir & "/Templates/" & templateName & "/", "/Templates/" & templateName & "/") 
			s=newWebDir & "/Templates/" & templateName & "/   ==>>  /Templates/" & templateName & "/"
            Call echoYellowB("�ŵ�ģ���ļ�����", s) 
        End If 
    End If 
    '����WEBĿ¼
    If Request("isMakeWeb") = "1" Then
        Call webBeautify("/web/", "", False) 
    End If 

    Call echo("�ɹ�", "������վ���") 
End Sub 
'������վ ����js,css���ļ���
Sub webBeautify(rootDir, resourcesDir, isTemplate)
    Dim content, splStr, s,c, fileName, fileExt, filePath, toFilePath, webBeautifyDir, imagesDir, cssDir, jsDir, indexFilePath, toIndexFilePath, indexFileContent 
    Dim cssFileList, splxx, url ,findStr
    webBeautifyDir = newWebDir & rootDir                                            '"/web/"
    imagesDir = webBeautifyDir & "/images/" 
    cssDir = webBeautifyDir & "/css/" 
    jsDir = webBeautifyDir & "/js/" 
    Call deleteFolder(webBeautifyDir) 
    Call createDirFolder(webBeautifyDir)                                            '���������ļ���
    Call createfolder(imagesDir) 
    Call createfolder(cssDir) 
    Call createfolder(jsDir) 
    Call echo("webBeautifyDir", webBeautifyDir) 
    Call echo("resourcesDir", resourcesDir) 
    Call echo("jsDir", jsDir) 

    indexFilePath = newWebDir & "/index.html" 
    'Ϊģ��
    If isTemplate = True Then
        toIndexFilePath = webBeautifyDir & "/Index_Model.html" 
    Else
        toIndexFilePath = webBeautifyDir & "/index.html" 
    End If 
    indexFileContent = readFile(indexFilePath, Char_Set)                            '���ļ�

    content = getFileFolderList(newWebDir, True, "ȫ��", "����", "", "", "") 
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
				'�ж��ļ���С��Ϊ��0    20160612
				if getFSize(filePath)<>0 then
	                Call copyfile(filePath, toFilePath) 
				end if
            End If 
            Call echo(fileName, resourcesDir & "images/" & fileName) 
        End If 
    Next 
    'Ϊģ��
    If isTemplate = True Then
        'link���и�favicon.ico  λ�û������⣬������20160302
        indexFileContent = handleConentUrl("[$cfg_webcss$]/", indexFileContent, "|link|", "", "") 
        indexFileContent = handleConentUrl("[$cfg_webimages$]/", indexFileContent, "|img||embed|param|meta|", "", "") 
        indexFileContent = handleConentUrl("[$cfg_webjs$]/", indexFileContent, "|script|", "", "") 
    Else
        'link���и�favicon.ico  λ�û������⣬������20160302
        indexFileContent = handleConentUrl("css/", indexFileContent, "|link|", "", "") 
        indexFileContent = handleConentUrl("images/", indexFileContent, "|img||embed|param|meta|", "", "") 
        indexFileContent = handleConentUrl("js/", indexFileContent, "|script|", "", "") 
    End If 
    'Ϊģ�壬���滻������ؼ���������
    If isTemplate = True Then
        indexFileContent = replaceWebTitle(indexFileContent, "{$Web_Title$}") 
        indexFileContent = ReplaceWebMate(indexFileContent, "name", "keywords", " content=""{$Web_KeyWords$}"" ") 
        indexFileContent = ReplaceWebMate(indexFileContent, "name", "description", " content=""{$Web_Description$}"" ") 
        indexFileContent = regExp_Replace(indexFileContent, "utf-8", "gb2312") 
        Char_Set = "gb2312" 
    End If 
    '��ʽ��HTMLΪ��
    If isFormatting = True Then
        indexFileContent = formatting(indexFileContent, "") 
        indexFileContent = htmlFormatting(indexFileContent) 
    End If 
    Call WriteToFile(toIndexFilePath, phptrim(indexFileContent), Char_Set) 



    'call echo("cssFileList",cssFileList)
    'css�б�Ϊ��
    If cssFileList <> "" Then
        splxx = Split(cssFileList, vbCrLf) 
        For Each filePath In splxx
            If filePath <> "" Then
                content = readFile(filePath, "") 
			    content = Replace(Replace(content, """", ""), "'", "")                          'ȥ����
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
    content = getFileFolderList(webBeautifyDir, True, "ȫ��", "", "ȫ���ļ���", "", "") 
    c = Replace(content, vbCrLf, "|") 
    'call rw(content)
    'call eerr(getThisUrlNoFileName(),getthisurl())
    c = c & "|||||" 
    url = getThisUrlNoFileName() & "/myZIP.php?webFolderName=web_&zipDir=htmlweb/&isPrintEchoMsg=0"
    'PHP���� ��վ
    If request("isPackWebZip")="1" or request("isWebToTemplateDir")="1" Then
		If request("isPackWebZip")="1" then
			s=XMLPost(url, "content=" & escape(c) & "&replaceZipFilePath=" & escape(handlePath(newWebDir)) & "\")
		else
			'PHP���� ģ����վ
			s=XMLPost(url, "content=" & escape(c) & "&replaceZipFilePath=" & escape(handlePath(isWebToTemplateDir)) & "\")
		end if
		findStr=strCut(s,"downfile.asp?downfile=","""",2)
		call echo("findStr",findStr)
		'��session��ʽ������ֵ
		session("downfilePath")="/" & findStr
		s=replace(s,"downfile.asp?downfile=" & findStr,"/tools/downfile.asp?downfile=session")
        Call echo("", s) 
    End If
	
    If(Request("isPackWeb") = "1" And isTemplate = False) Or(Request("isPackTemplate") = "1" And isTemplate = True) Then
        Dim xmlFileName,xmlSize
        xmlFileName = getIP() & "_update.xml" 

        Dim objXmlZIP : Set objXmlZIP = new xmlZIP
            Call objXmlZIP.run(handlePath(newWebDir), handlePath(newWebDir & rootDir), False, xmlFileName) 
			call echo("1111111111111","ע��")
			call echo(handlePath(newWebDir), handlePath(newWebDir & rootDir))
        Set objXmlZIP = Nothing 
		doevents
		xmlSize=getFSize(xmlFileName)
		xmlSize=printSpaceValue(xmlSize)
        Call echo("����xml����ļ�", "<a href=?act=download&downfile=" & xorEnc(xmlFileName, 31380) & " title='�������'>�������" & xmlFileName & "("& xmlSize &")</a>") 
    End If 
End Sub 
'��վ
Sub copyWeb()
    Dim content, url, splStr, splxx, i, s, tempS, c, sFType, sFName, filePath 
    Dim YuanWebPath, SplTitle 
    If httpurl = "" Then
        Call echo("����", "��ַΪ��" & httpurl) 
        Exit Sub 
    End If 
    '�����ļ���
    Call createfolder(newWebDir) 
    '����Դ��
    YuanWebPath = groupUrl(newWebDir, "/index_��ҳԴ��.html") 
    If checkFile(YuanWebPath) = False Then
        Call echo("����", "����������վԴ��" & httpurl) 
        If Request("isGetHttpUrl") = "1" Then
            content = getHttpUrl(httpurl, Char_Set) 
            Char_Set = "gb2312" 
            Call WriteToFile(YuanWebPath, content, Char_Set) 
        Else
            Call SaveRemoteFile(httpurl, YuanWebPath)
            content = readFile(YuanWebPath, Char_Set) 
        End If
		'��style��url()ͼƬ����20160630
		if instr(lcase(content),"url(")>0 then
			content=handleCssStyleContent(httpurl, newWebDir,content)
            Call WriteToFile(YuanWebPath, content, Char_Set) 		
		end if

        content = regExp_Replace(content, "<head>", "<head>" & vbCrLf & "<base href=""" & getWebsite(httpurl) & """ />" & vbCrLf) 
        Call WriteToFile(groupUrl(newWebDir, "/index_��ҳԴ��_base.html"), content, Char_Set) 

    End If 
    content = readFile(YuanWebPath, Char_Set) 

    '����������ַ����
    content = handleConentUrl(httpurl, content, "|*|", PubAHrefList, PubATitleList) 
    '�ƶ�JS Css������� �� /base.min.css?v=201510131920  Ϊ /base.min.css
    Call remoteJsCssParam(content, PubAHrefList) 
	'call createfile("1.txt",PubAHrefList)
    '����
    If PubAHrefList <> "" Then PubAHrefList = Left(PubAHrefList, Len(PubAHrefList) - 2) 
    If PubATitleList <> "" Then PubATitleList = Left(PubATitleList, Len(PubATitleList) - 2)

    Dim FileNameList, nCountArray 
    '����ͼƬ���زĵ�����
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

                    nCountArray = getArrayCount(FileNameList, LCase(sFName) & "|") + 1              '��������
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

                'MsgBox ("�����ļ�����=" & sFName & vbCrLf & FilePath)
                content = regExp_Replace(content, s, sFName)                             '�滻��������ַ
                content = Replace(content, s, sFName)                                    '�滻��������ַ
                If checkFile(filePath) = False Then
                    Call echo("����", "���������ز� " & s) 
					doevents
                    Call SaveRemoteFile(s, filePath) 
                    DownFileToLocaList = DownFileToLocaList & sFName & "|"                          '�ռ������ļ��б�
                    If sFType = ".css" Then
                        Call HandleCssFile(s, newWebDir, sFName) 
                    End If 
					
	                If checkFile(filePath) = False Then
						call echoYellow("����ʧ��",s)
					end if
                End If 
            End If 
            doEvents 
        Next 
    End If 
    url = httpurl 
    url = Mid(url, 1, InStrRev(url, "/")) 
    content = Replace(Replace(content, url, ""), getWebSite(httpurl), "")           'ȥ��ģ����ַ
    Call WriteToFile(newWebDir & "/index.html", phptrim(content), Char_Set)         '����ģ���ļ�(ȥ���߿ո�)
    Call createfile(newWebDir & "/ͼƬ��ַ�б�.txt", PubAHrefList) 
End Sub 

'����Css�ļ�����
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
    '�ظ�����css�ļ� ��utf-8����gb2312
    If LCase(Char_Set) = "utf-8" Or InStr(content, """utf-8""") > 0 Then
        If InStr(content, """utf-8""") > 0 Then TempContent = regExp_Replace(TempContent, """utf-8""", """gb2312""") 
        Call WriteToFile(ToPath, phptrim(TempContent), "gb2312") 
    End If 

	c=handleCssStyleContent(httpurl, folderName,c)

    Call WriteToFile(ToPath, phptrim(c), Char_Set) 
End Sub
'����CSS��ʽ����
function handleCssStyleContent(httpurl, folderName,c)
    Dim content, TempContent, startStr, endStr, splStr, i, s, tempS, sFType, sFName, filePath, url, ToPath, Char_Set ,sFName2
	content=lcase(c)
    startStr = "url\(" 
    endStr = "\)" 
    content = getArray(content, startStr, endStr, False, False) 
    content = Replace(Replace(content, """", ""), "'", "")                          'ȥ����
    If content <> "" Then
        splStr = Split(content, "$Array$") 
        For Each s In splStr
			
            tempS = LCase(Trim(s)) 
            sFType = Right(tempS, Len(tempS) - InStrRev(tempS, ".") + 1) 
            sFName = Right(tempS, Len(tempS) - InStrRev(tempS, "/")) 
			'sFName=replace(replace(sFName,"""",""),"'","")			'����css·��20160612

            url = fullHttpUrl(httpurl, s) 
            If InStr("|.jpg|.gif|.png|.swf|.js|.css|||", "|" & sFType & "|") > 0 Then
				if instr(vbcrlf & PubAHrefList & vbcrlf, vbcrlf & url & vbcrlf)=false then
					PubAHrefList=PubAHrefList & url & vbcrlf
    				Call createfile(newWebDir & "/ͼƬ��ַ�б�.txt", PubAHrefList) 		'ʱʱ����
										
					filePath = folderName & "\" & sFName
					'������ͼƬ����20160630
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
					c = caseInsensitiveReplace(c, s, sFName)                                 '���ִ�Сд�滻
					If checkFile(filePath) = False Then
						call echo("����",url) 
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
        replaceWebMate = bodyStart & startStr & valueStr & endStr & bodyEnd 
    End If 
End Function 
'��¼
function login()
	dim password
	password=request("password")
	if password="2016" then
		session("fangzhang")="1"
	end if
	displayDefaultLayout()
end function
'��ʾĬ�ϲ���
Sub displayDefaultLayout()
    If getIP() <> "127.0.0.1" Then
        If Session("adminusername") <> "ASPPHPCMS" and session("fangzhang")<>"1" Then
            Response.Write("������nolocal ��¼��̨")
        
%> 
        <form id="form1" name="form1" method="post" action="?act=login">
          ���룺
          <input type="text" name="password" id="password" />
          <input type="submit" name="button" id="button" value="�ύ" />
        </form>
<%    	Response.End() 
    	    End If
	end if
%>
<form action="?act=downweb" method="post" name="form1" target="_blank" id="form1"> 
  ��ַ�� 
  <input name="httpurl" type="text" id="httpurl" value="<% =httpurl%>" size="90" /> 
���룺 
<select name="Char_Set" id="Char_Set"> 
  <option value="gb2312" selected="selected">gb2312</option> 
  <option value="utf-8" <% If Char_Set = "utf-8" Then rw("selected")%>>utf-8</option> 
</select> 
ģ������ 
<input name="templateName" type="text" id="templateName" value="<% =templateName%>" size="12" /> 
<input type="submit" name="button" id="button" value="�ύ" /> 
<input name="verificationTime" type="hidden" id="verificationTime" value="<% =xorEnc(Now(), 31380)%>" /> 
<hr /> 
ѡ�����ط�ʽ�� 
<select name="isGetHttpUrl" id="isGetHttpUrl"> 
  <option value="0" selected="selected">ֱ������</option> 
  <option value="1">getHttpUrl��ʽ�������</option> 
</select> 
<hr /> 
<label for="isMakeWeb"><input name="isMakeWeb" type="checkbox" id="isMakeWeb" value="1" <% If isMakeWeb = True Then rw("checked")%>/> 
����WEB�ļ��� 
<label for="isPackWeb"><input name="isPackWeb" type="checkbox" id="isPackWeb" value="1" <% If isPackWeb = True Then rw("checked")%>/> 
�Ƿ���WEB�ļ���(xml) 
</label> 
<label for="isPackWebZip"><input name="isPackWebZip" type="checkbox" id="isPackWebZip" value="1" <% If isPackWebZip = True Then rw("checked")%>/> 
�Ƿ���WEB�ļ���(Zip) 
</label> 
<br /><hr /> 
</label> 
<label for="isMakeTemplate"><input name="isMakeTemplate" type="checkbox" id="isMakeTemplate" value="1" <% If isMakeTemplate = True Then rw("checked")%> /> 
�Ƿ�����ģ���ļ��� 
</label> 
<label for="isPackTemplate"><input name="isPackTemplate" type="checkbox" id="isPackTemplate" value="1" <% If isPackTemplate = True Then rw("checked")%> /> 
�Ƿ���ģ���ļ���(xml)
<label for="isPackTemplateZip"><input name="isPackTemplateZip" type="checkbox" id="isPackTemplateZip" value="1" <% If isPackTemplateZip = True Then rw("checked")%> /> 
�Ƿ���ģ���ļ���(Zip)
</label> 
<br /><hr /> 
ģ�����Ŀ¼ 
<input name="templateDir" type="text" id="templateDir" value="<% =templateDir%>" /> 
<label for="isWebToTemplateDir"><input name="isWebToTemplateDir" type="checkbox" id="isWebToTemplateDir" value="1" <% If isWebToTemplateDir = True Then rw("checked")%> />�ļ��ŵ�ģ���ļ�����</label>  
<hr /> 
<label for="isFormatting"><input name="isFormatting" type="checkbox" id="isFormatting" value="1" <% If isFormatting = True Then rw("checked")%> />�Ƿ��ʽ��HTML</label> 
<label for="isUniformCoding"><input name="isUniformCoding" type="checkbox" id="isUniformCoding" value="1" <% If isUniformCoding = True Then rw("checked")%> />�Ƿ�ͳһ����</label> 
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
            $("#isPackWebZip").removeAttr("disabled");//Ҫ���Enable��JQueryֻ����ôд   
        }else{             
            $("#isPackWeb").attr("disabled","disabled");//�ٸĳ�disabled  
            $("#isPackWebZip").attr("disabled","disabled");//�ٸĳ�disabled  
        } 
    } 
    //����Ƿ�Ϊ����ģ����վ 
    function checkIsMakeTemplate(){ 
        if($("#isMakeTemplate").is(':checked')) {             
            $("#isPackTemplate").removeAttr("disabled");//Ҫ���Enable��JQueryֻ����ôд   
            $("#isPackTemplateZip").removeAttr("disabled");//Ҫ���Enable��JQueryֻ����ôд   
            $("#isWebToTemplateDir").removeAttr("disabled");//Ҫ���Enable��JQueryֻ����ôд   
            $("#templateDir").removeAttr("disabled");//Ҫ���Enable��JQueryֻ����ôд   
        }else{             
            $("#isPackTemplate").attr("disabled","disabled");//�ٸĳ�disabled  
            $("#isPackTemplateZip").attr("disabled","disabled");//�ٸĳ�disabled  
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
End Sub

 


%> 


