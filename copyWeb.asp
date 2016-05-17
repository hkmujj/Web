<!--#Include File = "Inc/Config.Asp"--> 
<% 


select case request("act")
	case "copyweb" : call copyweb()
	case "copyASPPHPCMS" : call copyASPPHPCMS()
	case else displayDefault()
end select

'�������ܽ���
'delASP=true   Ϊɾ��ASP
'delPHP=true   Ϊɾ��PHP
'templates2015=web_default
'jqueryCommon1=true
sub displayDefault()
	call echo("������վ", "<a href='?act=copyweb' target='_blank'>���������վ</a>")
	
	call echo("����ASPPHPCMS��ASP��", "<a href='?act=copyASPPHPCMS' target='_blank'>�������ASPPHPCMS ��ģ��</a>")
	
	call echo("����ASPPHPCMS��ASP��", "<a href='?act=copyASPPHPCMS&templates2015=sharembweb|web_default&templates=ufoer&jqueryCommon1=true' target='_blank'>�������ASPPHPCMS[sharembweb|web_default|ufoer]</a>")
	
	
	call echo("����ASPPHPCMS��ASP��", "<a href='?act=copyASPPHPCMS&delPHP=true&jqueryCommon1=true&templates2015=web_default' target='_blank'>ASP��[web_default]</a>") 
	
	call echo("����ASPPHPCMS��ASP��", "<a href='?act=copyASPPHPCMS&isnote=true&isaspcodehx=true&isdelcodenode=true&templates=ufoer|legendstu&templates2015=sharembweb' target='_blank'>�������ASPPHPCMS��isnote|isaspcodehx|isdelcodenode��[sharembweb|ufoer|legendstu]</a>")
	
	call echo("����ASPPHPCMS��ASP��(anyi)", "<a href='?act=copyASPPHPCMS&templates=&templates2015=anyi&delASP=true' target='_blank'>�������ASPPHPCMS [anyi]</a>")
	 
	call echo("����ASPPHPCMS��ASP+PHP��(��Դ��)", "<a href='?act=copyASPPHPCMS&isnote=true&delsavefile=true&templates=ufoer|legendstu&templates2015=sharembweb' target='_blank'>�������ASPPHPCMS��isnote|delsavefile��[sharembweb|ufoer|legendstu]</a>")
	
	
	'delAsp=true
end sub

'�������ASPPHPCMS��ASP��
sub copyASPPHPCMS()
	dim rootDir,incDir,adminDir,jqueryDir,splstr,content,fileName,filePath,toFilePath,fileSuffix,templatesDir,templates2015Dir,indexFile,s,i
	dim UploadFilesDir,dataDir,folderPath,folderName
	
	rootDir="/../ASPPHPCMS ga1/"
	'call deletefolder(rootdir) 
	if checkfolder(rootDir)=true then
		splstr=split(getDirFolderList(rootDir),vbcrlf)
		for each folderPath in splstr
			call deleteFolder(folderPath)
			call echo("del folderPath",folderPath)
		next
		splstr=split(getDirFileList(rootDir,""),vbcrlf)
		for each filePath in splstr
			call deleteFile(filePath)
			call echo("del filePath",filePath)
		next 
	end if
	
	incDir=rootDir & "/Inc/"
	adminDir=rootDir & "/Admin/"
	jqueryDir=rootDir & "/Jquery/"
	templatesDir=rootDir & "/Templates/"
	templates2015Dir=rootDir & "/Templates2015/"
	indexFile=rootDir & "/index.asp"
	UploadFilesDir=rootDir & "/UploadFiles/"
	dataDir=rootDir & "/Data/"
	
	call createDirFolder(rootDir)
	call copyFolder("/inc/",incDir)
	call copyFolder("/web/",adminDir)
	'ɾ��admin/data
	call deleteFolder(adminDir & "/Data")	
	call createFolder(adminDir & "/Data")
	call createFolder(adminDir & "/Data/BackUpDateBases")
	call createFolder(adminDir & "/Data/Stat")
	call createFolder(adminDir & "/Data/SystemLog")
	call createfile(adminDir & "/Data/BackUpDateBases/index.html","")		'�����յ�index.htmlҳ��
	call createfile(adminDir & "/Data/Stat/index.html","")		'�����յ�index.htmlҳ��
	call createfile(adminDir & "/Data/SystemLog/index.html","")		'�����յ�index.htmlҳ��
	'call copyFolder("/Templates/",templatesDir)
	call createFolder(templatesDir)
	call createFolder(templates2015Dir)
	'template    templates=ufoer
	splstr=split(request("templates"),"|")
	for each s in splstr
		s=trim(s)
		if s <>"" then
			call copyFolder("/Templates/"& s &"/",templatesDir & "/"& s &"/") 
		end if
	next
	'template2015      templates2015=sharembweb|anyi
	splstr=split(request("templates2015"),"|")
	for each s in splstr
		s=trim(s)
		if s <>"" then
			call copyFolder("/Templates2015/"& s &"/",templates2015Dir & "/"& s &"/") 
		end if
	next
	
	
	'����Ŀ¼
	call createFolder(dataDir)
	if rq("templates2015")="sharembweb" then
		call copyFolder("/templates2015/sharembweb/WebData/", dataDir & "/WebData/")
	elseif rq("templates2015")="anyi" then		
		call copyFolder("/templates2015/anyi/WebData/", dataDir & "/WebData/")
	else
		call copyFolder("/data/WebData/", dataDir & "/WebData/")
	end if
	call copyfile("/Data/Data.mdb",dataDir & "/Data.mdb")						'�������ݿ� 
	Call Echo("ѹ�����ݿ�", CompactDB(dataDir & "/Data.mdb", False))		'ѹ��MDB���ݿ�
	'UploadFilesDir
	call createFolder(UploadFilesDir)
	call createFolder(UploadFilesDir & "/image/")
	call copyfile("/UploadFiles/Down.rar", UploadFilesDir & "/Down.rar")
	call copyfile("/UploadFiles/NoImg.jpg", UploadFilesDir & "/NoImg.jpg")
	call copyfile("/UploadFiles/baidu.gif", UploadFilesDir & "/baidu.gif")
	call copyfile("/UploadFiles/google.gif", UploadFilesDir & "/google.gif")
	call copyfile("/UploadFiles/yahoo.gif", UploadFilesDir & "/yahoo.gif")
	call copyfile("/UploadFiles/NoImg.gif", UploadFilesDir & "/NoImg.gif")
	
	'jquery
	call createFolder(jqueryDir)
	call copyFolder("/jquery/skins",jqueryDir & "/skins")
	call copyFolder("/jquery/syntaxhighlighter",jqueryDir & "/syntaxhighlighter")
	call copyfile("/jquery/Callcontext_menu.js",jqueryDir & "/Callcontext_menu.js") 
	call copyfile("/jquery/contextmenu.css",jqueryDir & "/contextmenu.css") 
	call copyfile("/jquery/jquery.Min.js",jqueryDir & "/jquery.Min.js") 
	call copyfile("/jquery/lhgdialog.min.js",jqueryDir & "/lhgdialog.min.js") 
	call copyfile("/jquery/mmenubg.gif",jqueryDir & "/mmenubg.gif") 
	
	'admin �ļ����ƴ���
	content=getDirFolderNameList(adminDir)
	splstr=split(content,vbcrlf)
	for each folderName in splstr
		if folderName<>"" then
			if instr("#_",left(folderName,1))>0 or folderName="����" then
				call deleteFolder(adminDir & "/" & folderName)
				call echo("del folderName",folderName)
			end if
		end if
	next
	
	'admin �ļ���������
	content=getDirFileNameList(adminDir,"")
	
	splstr=split(content,vbcrlf)
	for each fileName in splstr
		if fileName<>"" then
			fileSuffix=Lcase(getFileExt(fileName))
			filePath=adminDir & fileName
			if fileSuffix="txt" or lcase(left(fileName,6))="zcase_" or filename="1.html" then
				call deletefile(filePath)
				call echo("filePath",filePath)
			elseif fileSuffix="asp" or fileSuffix="html" then
				content=getftext(filePath)
				'call echo("aspfile",filepath)
				if instr(content,"aspweb.asp")>0 then
					if fileSuffix="asp" then
						content=replace(content,"aspweb.asp","index.asp")
					else
						content=replace(content,"aspweb.asp","index.{$EDITORTYPE$}")
					end if
					call writeToFile(filePath,content,"gb2312")
					call echo("aspweb.asp",filepath)
				end if
				'ɾ������Ҫ�����ݲ���
				if instr(content,"<!--#del# start-->")>0 or instr(content,"<!--#del# end-->")>0 then
					for i = 1 to 10
						s=getStrCut(content,"<!--#del# start-->","<!--#del# end-->",1)
						if s <>"" then
							content=replace(content,s,"")
						else
							exit for
						end if
					next
				end if
				
				'asp���滻
				content=replace(content, "../"" & EDITORTYPE & ""web."" & EDITORTYPE & """, "../index.asp")
				content=replace(content, """../"" & EDITORTYPE & ""web."" & EDITORTYPE", """../index.asp""")	
				content=replace(content, """function.Asp""", """../Inc/admin_function.asp""")
				content=replace(content, """setAccess.Asp""", """../Inc/admin_setAccess.asp""") 
				content=replace(content, "template.asp?", "../inc/admin_template.asp?")
				
				
				Content=codeSafe(Content,fileSuffix)			'���봦��
				call writeToFile(filePath,content,"gb2312")
				
			end if
		end if
	next 
	'ɾ��inc/���ļ��б�
	dim delFileList : delFileList="|Install.Asp|Create_Html.Asp|Create_VB.Asp|Create_VBNet.Asp|ZClassAspCode.Asp|ZClass_Maiside.Asp|ZClass_YouKou.Asp|ztemporary_Cache.Asp|ztemporary_DirManage.Asp|2014_NetWorkClass.asp|2014_Action.Asp|2014_HelpStudy.Asp|2014_Links.Asp|2014_Links.Asp|2014_MainInfo.Asp|2014_Nav.Asp|2014_News.Asp|2014_Search.Asp|2014_SiteMap.Asp|2015_TempSub.Asp|2015_ToPhpCms.Asp|2014_Class.Asp|AutoAdd.Asp|AjAx.Asp|App.Asp|CallWeb.Asp|Fun_Table.Asp|Function_SEO.Asp|Function2014.Asp|ProductGl.Asp|Web2014.Asp|2014_Banner.Asp|"
	splstr=split(delFileList,"|")
	for each fileName in splstr
		if fileName<>"" then 
			filePath=incDir & fileName
			call echo("ɾ���ļ�",filePath)
			call deleteFile(filePath)
		end if
	next
	
	'inc �ļ���������
	content=getDirFileNameList(incDir,"")
	splstr=split(content,vbcrlf)
	for each fileName in splstr
		if fileName<>"" then
			fileSuffix=Lcase(getFileExt(fileName))
			filePath=incDir & fileName
			if fileSuffix="php" or fileSuffix="txt" or lcase(left(fileName,7))="ZClass_" or lcase(left(fileName,4))="dll_" then
				call deletefile(filePath)
				call echo("inc/filePath",filePath)				
			elseif fileSuffix="asp" and (instr(lcase(delFileList),"|"& lcase(fileName) &"|"))>0 and request("delsavefile")="true" then
				call echo("delsavefile",filePath)
				call deletefile(filePath)
			elseif fileSuffix="asp" or fileSuffix="html" then
				content=getftext(filePath)
				content=replace(content, "/web/1.asp", "/admin/index.asp")
				content=replace(content, """/web/""", """/admin/""")
				content=replace(content, "aspweb.asp", "index.asp")
				
				
				Content=codeSafe(Content,fileSuffix)			'���봦��
				call writeToFile(filePath,content,"gb2312")
			end if
		end if
	next	
	call deleteFile(incDir & "/config.asp")
	call moveFile(incDir & "/_config.asp",incDir & "/config.asp")
	'����  inc/config.asp
	filePath=incDir & "/config.asp"
	content=getftext(filePath)
	splstr=split(delFileList,"|")
	for each fileName in splstr
		if fileName<>"" then
			content=regExp_Replace(content,"<!--#" & "Include File = """& fileName &"""-->","")
		end if	
	next
	call writeToFile(filePath,content,"gb2312")
	
	
	'�ƶ��ļ�
	call moveFile(adminDir & "/function.asp", incDir & "/admin_function.asp")
	call moveFile(adminDir & "/setAccess.asp", incDir & "/admin_setAccess.asp")
	call moveFile(adminDir & "/template.asp", incDir & "/admin_template.asp")
 
 
	'admin
	
	content=getftext(adminDir & "/1.asp")
	content=replace(content,"_Config.Asp","Config.Asp")
	call writeToFile(adminDir & "/index.asp",content,"gb2312") 
	
	call deleteFile(adminDir & "/1.asp")
	
	'root  index.asp
	content=getftext("/aspweb.asp")
	content=replace(content,"web/","admin/")
	content=replace(content, """admin/function.Asp""", """inc/admin_function.Asp""")
	content=replace(content,"_Config.Asp","Config.Asp")  

	Content=codeSafe(Content,fileSuffix)			'���봦��
	call writeToFile(indexFile,content,"gb2312")
	'root url.asp
	content=getftext("/url.asp")
	Content=codeSafe(Content,fileSuffix)			'���봦��
	call writeToFile(rootDir & "/url.asp",content,"gb2312")
	 
	
	call writeToFile(rootDir & "web.config",getWebConfig(),"gb2312")
	
	
	
	
	'����PHP����
	incDir=rootDir & "/PhpInc/"
	call createFolder(incDir)
	
	splstr=split("ASP.php|sys_FSO.php|Conn.php|MySqlClass.php|sys_Cai.php|sys_System.php|startInstall.php|Install.php|sys_Url.php|sys_System.php|","|")
	for each fileName in splstr
		filePath="/PHP2/ImageWaterMark/Include/" & fileName
		toFilePath=incDir & fileName
		if checkFile(filePath)=true then
			content=getftext(filePath) 
			if lcase(fileName)="startinstall.php" then
				content=replace(content,"/phpweb.php","../index.php")
				content=replace(content,"/web/1.php","../admin/index.php")
				 
				
			elseif lcase(fileName)="install.php" then
				content=replace(content,"./../../Web/Inc/","./")
				content=replace(content,"./../../../web/","./../admin/")
				content=replace(content,"/web/1.php","../admin/index.php")
				
				'call rwend(content)
				
				
				
				
			end if
			content=codeSafe(Content,"php")			'���봦��
			call writeToFile(toFilePath,content,"gb2312")
			
		end if
	next
	 
	
	'phpinc conn.php
	filePath=incDir & "/Conn.php"
	content=getftext(filePath) 
	content=replace(content, "PHP2/ImageWaterMark/Include/", "phpinc/") 
	content=replace(content, "$dbhost='localhost';", "$dbhost='localhostNO';") 
	
	
	
	Content=codeSafe(Content,"php")			'���봦��
	call writeToFile(filePath,content,"gb2312") 
	
	'phpinc conn.php
	filePath=incDir & "/Install.php"
	content=getftext(filePath) 
	content=replace(content, "/web/1.php", "/admin/index.php") 
	content=replace(content, "?act=resetAccessData", "?act=setAccess") 
	Content=codeSafe(Content,"php")			'���봦��
	call writeToFile(filePath,content,"gb2312") 
	
	


	'�ƶ��ļ�
	call moveFile(adminDir & "/function.php", incDir & "/admin_function.php")
	call moveFile(adminDir & "/setAccess.php", incDir & "/admin_setAccess.php")
	call moveFile(adminDir & "/template.php", incDir & "/admin_template.php")
	call moveFile(adminDir & "/config.php", incDir & "/config.php")
	

	filePath=incDir & "/admin_function.php"
	content=getftext(filePath) 
	content=replace(content, "aspweb.asp", "../index.php") 
	
	Content=codeSafe(Content,"php")			'���봦��
	call writeToFile(filePath,content,"gb2312")
	
	filePath=incDir & "/admin_setAccess.php"
	content=getftext(filePath) 
	content=replace(content, "/' . EDITORTYPE . 'web.' . EDITORTYPE . '", "../index.php")
	content=replace(content, "\'..../index.php\'","\'../index.php\'")		'�Ľ�����  setaction.php����ļ�20160409
	content=replace(content, "1.' . EDITORTYPE . '", "index.php") 
	Content=codeSafe(Content,"php")			'���봦��
	call writeToFile(filePath,content,"gb2312") 
	'config.php
	filePath=incDir & "/config.php"
	content=getftext(filePath)  
	content=replace(content, "'/web/'", "'/admin/'") 
	content=replace(content, "'../phpweb.php'", "'../index.php'") 
	content=replace(content, "'/web/1.php'", "'/admin/index.php'")
	call writeToFile(filePath,content,"gb2312") 
	
	
	 
	
	'inc �ļ���������
	content=getFileFolderList("/PHP2/Web/Inc/", True, "php", "����", "", "", "") 
	splstr=split(content,vbcrlf)
	for each fileName in splstr
		if fileName<>"" then
			call echo("php",fileName)
			content=getftext("/PHP2/Web/Inc/" & fileName)
			content=replace(content, "/web/1.asp", "/admin/index.php")
			Content=codeSafe(Content,"php")			'���봦��
			call writeToFile(incDir & fileName,content,"gb2312")
		end if
	next 
	
	'admin
 
 
	'php��ҳ
	indexFile=rootDir & "/index.php"
	content=getftext("/phpweb.php")
	content=replace(content, "./Web/setAccess.php", "./phpInc/admin_setAccess.php") 
	content=replace(content, "./Web/function.php", "./phpInc/admin_function.php") 
	content=replace(content, "./Web/config.php", "./phpInc/config.php")
	content=replace(content, "/web/1.asp", "/admin/index.asp")
	
	content=replace(content, "PHP2/ImageWaterMark/Include/", "phpInc/")
	content=replace(content, "PHP2/Web/Inc/", "phpInc/")
	content=replace(content, "Web/", "phpInc/")
	 
	 
	
	Content=codeSafe(Content,"php")			'���봦��
	call writeToFile(indexFile,content,"gb2312")
	
	'php��̨��ҳ
	call moveFile(adminDir & "/1.php",adminDir & "/index.php") 
	filePath=adminDir & "/index.php"
	content=getftext(filePath) 
	content=replace(content, "PHP2/ImageWaterMark/Include/", "phpInc/")
	content=replace(content, "PHP2/Web/Inc/", "phpInc/")
	content=replace(content, "Web/", "phpInc/")
	content=replace(content, "setAccess.php", "../phpInc/admin_setAccess.php")
	content=replace(content, "function.php", "../phpInc/admin_function.php")
	content=replace(content, "config.php", "../phpInc/config.php")
	content=replace(content, "'../phpweb.php'", "'../index.php'")
	Content=codeSafe(Content,"php")			'���봦��
	
	
	call writeToFile(filePath,content,"gb2312")
	s="<a href='http://apcms.n/index.asp'>ASP��Ԥ��</a>  |  <a href='http://apcms.n/admin/index.asp'>ASP���̨</a>"
	s=s & " &nbsp; &nbsp; &nbsp; <a href='http://apcms.n/index.php'>PHP��Ԥ��</a>  |  <a href='http://apcms.n/admin/index.php'>PHP���̨</a>"
	'ɾ��ASP�汾
	if request("delASP")="true" then
		call echo("ɾ��ASP�汾","true")
		call deleteFile(rootDir & "/url.asp")
		call deleteFile(rootDir & "/index.asp")
		call deleteFile(rootDir & "/admin/index.asp")
		call deleteFile(rootDir & "/admin\Data\Data.mdb")
		call deleteFolder(rootDir & "/inc")
	end if
	'ɾ��PHP�汾
	if request("delPHP")="true" then
		call deleteFile(rootDir & "/index.php")
		call deleteFile(rootDir & "/admin/index.php")
		call deleteFolder(rootDir & "/phpInc")		
	end if
	'׷��jquery Common1
	if request("jqueryCommon1")="true" then
		jqueryDir=rootDir & "/Jquery/"
		call copyFile("/jquery/Jquery.Min.Js", jqueryDir & "Jquery.Min.Js")
		call copyFile("/jquery/ShowTime.Js", jqueryDir & "ShowTime.Js")
		call copyFile("/jquery/MSClass.Js", jqueryDir & "MSClass.Js")
		call copyFile("/jquery/ScrollPicLeft.Js", jqueryDir & "ScrollPicLeft.Js")
		call copyFile("/jquery/swfobject.Js", jqueryDir & "swfobject.Js")
		call copyFile("/jquery/Function.Js", jqueryDir & "Function.Js")
		call copyFile("/jquery/Common1.Js", jqueryDir & "Common1.Js")
	end if
	'�����԰��ô���
	if request("templates2015")="anyi" then
		call deleteFolder(rootDir & "/Templates")
		call deleteFolder(rootDir & "/data/WebData")
		call deleteFile(rootDir & "/data/data.mdb")
		call copyFolder(rootDir & "/\Templates2015\anyi\WebData",rootDir & "/data/WebData")		
	end if
	
	call echo("�������","����ASPPHPCMS�ĳɹ�  " & s)
end sub
'���봦��
function codeSafe(Content,fileSuffix)	
	dim addtoAspHeadStr : addtoAspHeadStr="<%" & vbcrlf & authorInfo("") & "%" & ">" & vbcrlf			'Asp�ļ�ͷ׷������
	dim addtoPhpHeadStr : addtoPhpHeadStr="<?php " & vbcrlf & handleAuthorInfo("","php") & "?>" & vbcrlf			'Asp�ļ�ͷ׷������
	
	if fileSuffix="asp" then
		'ɾ��ASP��ע��
		if request("isdelcodenode")="true" then
			Content = DelAspNote(Content)
		end if 
		'Asp�����������
		if request("isaspcodehx")="true" then
			Content = AspCodeConfusion(Content)
		end if
		'���MYע��
		if request("isnote")="true" then
			Content = addtoAspHeadStr & Content
		end if
		
		'Content = OptimizeAspCode(Content)		'�Ż�ASP����	 
		'Content = ASPVariableGarbled(Content )	'Asp�������� ֻ�Դ�Сд��
	elseif fileSuffix="php" then
		'���MYע��
		if request("isnote")="true" then
			Content = addtoPhpHeadStr & Content
		end if
	end if
	
	codeSafe=Content
end function

'��������webConfig����
function getWebConfig()
   Dim c 
    c = c & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf 
    c = c & "<configuration>" & vbCrLf 
    c = c & "    <system.webServer>" & vbCrLf 
'    c = c & "        <directoryBrowse enabled=""true"" />" & vbCrLf 
    c = c & "        <defaultDocument>" & vbCrLf 
    c = c & "            <files>" & vbCrLf 
    c = c & "            	<clear />" & vbCrLf 
    c = c & "                <add value=""index.html"" />" & vbCrLf 
    c = c & "                <add value=""index.asp"" />" & vbCrLf 
    c = c & "                <add value=""index.php"" />" & vbCrLf 
    c = c & "                <add value=""index.htm"" />" & vbCrLf 
    c = c & "                <add value=""Default.htm"" />" & vbCrLf 
    c = c & "                <add value=""Default.asp"" />" & vbCrLf 
    c = c & "                <add value=""iisstart.htm"" />" & vbCrLf 
    c = c & "                <add value=""default.aspx"" />" & vbCrLf 
    c = c & "            </files>" & vbCrLf 
    c = c & "        </defaultDocument>" & vbCrLf 
    c = c & "    </system.webServer>" & vbCrLf 
    c = c & "</configuration>" & vbCrLf 
	getWebConfig=c
end function
 

'������վ
sub copyweb()
	dim splstr,s
	'Admin,Inc,Images,Jquery,Templates2015,����,UploadFiles,Web,PHP,PHP2,VB����
	if 1=1 then
		Call makeHanleleFile(handlePath("/Admin/"), handlePath("E:/Web/Admin/"), "ȫ���ļ���") 
		Call makeHanleleFile(handlePath("/Inc/"), handlePath("E:/Web/Inc/"), "ȫ���ļ���") 
		Call makeHanleleFile(handlePath("/Images/"), handlePath("E:/Web/Images/"), "ȫ���ļ���") 
		Call makeHanleleFile(handlePath("/Jquery/"), handlePath("E:/Web/Jquery/"), "ȫ���ļ���") 
		Call makeHanleleFile(handlePath("/Templates/"), handlePath("E:/Web/Templates/"), "ȫ���ļ���") 
		Call makeHanleleFile(handlePath("/Templates2015/"), handlePath("E:/Web/Templates2015/"), "ȫ���ļ���") 
		Call makeHanleleFile(handlePath("/UploadFiles/"), handlePath("E:/Web/UploadFiles/"), "ȫ���ļ���") 
		Call makeHanleleFile(handlePath("/Web/"), handlePath("E:/Web/Web/"), "ȫ���ļ���") 
		Call makeHanleleFile(handlePath("/����/"), handlePath("E:/Web/FormattingTools/"), "ȫ���ļ���") 
		Call makeHanleleFile(handlePath("/PHP2/"), handlePath("E:/Web/PHP2/"), "ȫ���ļ���") 
		Call makeHanleleFile(handlePath("/Web_Soft/"), handlePath("E:/Web/Web_Soft/"), "ȫ���ļ���") 
		Call makeHanleleFile(handlePath("/Data/"), handlePath("E:/Web/Data/"), "ȫ���ļ���")
		
		Call makeHanleleFile(handlePath("/VB����/"), handlePath("E:/Web/VB����/"), "ȫ���ļ���")
	end if
	Call makeHanleleFile(handlePath("/"), handlePath("E:/Web/"), "") 
	Call vBEchoTimer()
end sub

'�����ļ�
Function makeHanleleFile(dirPath, toDirPath, sFFType)
    Dim content, splStr, fileName, fileExt, filePath, toFilePath, isHandle, nLen,sMsg
	dim templateDefaultStartStr,defaultStartStr
	templateDefaultStartStr="<%" & vbcrlf & authorInfo("") & "%" & ">" & vbcrlf			'׷�Ӹ�����Ϣ

    Call createfolder(toDirPath) 
    nLen = Len(dirPath) + 1 
    content = getFileFolderList(dirPath, True, "ȫ��", "", sFFType, "", "") 

    splStr = Split(content, vbCrLf) 
    For Each filePath In splStr
        toFilePath = toDirPath & Mid(filePath, nLen) 
        '�ļ���
        If checkfolder(filePath) = True Then
            Call createfolder(toFilePath) 
        Else
            isHandle = True 
            fileName = getFileName(filePath) 
            fileExt = LCase(getFileExt(fileName)) 

            If checkFile(toFilePath) = True Then
                If getFileEditTime(filePath) = getFileEditTime(toFilePath) Then
                    isHandle = False 
                End If 
            End If 
            If isHandle = True Then  
                If instr("|asp|bas|cls|frm#|","|" & fileExt & "|")>0 and instr(filePath,"\KEditor\")=false and instr("|data.gif.asp|API.bas|","|"& filename &"|")=false Then
                    content = getftext(filePath)
					defaultStartStr=""
					if instr("|asp|bas|cls|frm|","|" & fileExt & "|")>0 and content<>"" then
						content="<%"& vbcrlf & content &"%" & vbcrlf & ">"
						defaultStartStr=templateDefaultStartStr
					end if 
					'�����ļ�
					if filename="1AspToPhp.Asp" then
						content=replace(content,"""atemp.asp""", """sharembweb.com.asp""")
						content=replace(content,"""atemp.php""", """sharembweb.com.php""")
					end if 
                    content =aspCodeSafe(content)			'asp���밲ȫ
					 
					
					if instr("|asp|bas|cls|","|" & fileExt & "|")>0 and content<>"" then
						content=phptrim(content)
						content= defaultStartStr & mid(content,5,len(content)-8)
					end if
					sMsg="�������滻�ļ�"
                    Call writeToFile(toFilePath, content,"gb2312") 
                Else
                    Call deleteFile(toFilePath) 
                    Call copyfile(filePath, toFilePath) 
					sMsg="�����ļ�"
                End If 
                Call editFileEditDate(toFilePath, getFileEditTime(filePath)) 
                Call echo(sMsg, toFilePath) 
            End If 
        End If 
    Next 
End Function 

%> 

