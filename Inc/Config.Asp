<%Option Explicit%>
<!--#Include File = "Md5.Asp"--> 
<!--#Include File = "Common.Asp"--> 
<!--#Include File = "Function.Asp"--> 
<!--#Include File = "FSO.Asp"--> 
<!--#Include File = "AdodbStream.Asp"--> 
<!--#Include File = "XML.Asp"--> 
<!--#Include File = "Time.Asp"--> 
<!--#Include File = "RegExp.Asp"--> 
<!--#Include File = "Cai.Asp"--> 
<!--#Include File = "StringNumber.Asp"--> 
<!--#Include File = "EncDec.Asp"--> 
<!--#Include File = "Dictionary.Asp"--> 
<!--#Include File = "Access.Asp"--> 
<!--#Include File = "Incpage.Asp"--> 
<!--#Include File = "Check.Asp"--> 
<!--#Include File = "GetWord.Asp"--> 
<!--#Include File = "Url.Asp"--> 
<!--#Include File = "Conn.Asp"--> 
<!--#Include File = "Fun_Table.Asp"--> 
<!--#Include File = "FunHTML.Asp"--> 
<!--#Include File = "ProductGl.Asp"--> 
<!--#Include File = "Html.Asp"--> 
<!--#Include File = "AutoAdd.Asp"-->
<!--#Include File = "Web2014.Asp"-->   
<!--#Include File = "Print.Asp"--> 
<!--#Include File = "Img.Asp"--> 
<!--#Include File = "PinYin.Asp"-->
<!--#Include File = "Function2014.Asp"-->
<!--#Include File = "IE.Asp"-->  
<!--#Include File = "ASPPHPAccess.Asp"-->   
<!--#Include File = "2014_Links.Asp"--> 
<!--#Include File = "2014_MainInfo.Asp"-->  
<!--#Include File = "2014_News.Asp"--> 
<!--#Include File = "2014_Banner.Asp"--> 
<!--#Include File = "2014_Array.Asp"--> 
<!--#Include File = "2014_Js.Asp"--> 
<!--#Include File = "2014_Css.Asp"--> 
<!--#Include File = "2014_Template.Asp"-->  
<!--#Include File = "2014_Action.Asp"-->  
<!--#Include File = "2014_Nav.Asp"-->  
<!--#Include File = "2014_Class.Asp"-->  
<!--#Include File = "2014_Search.Asp"-->  
<!--#Include File = "2014_GBUTF.Asp"-->   
<!--#Include File = "2014_Form.Asp"-->  
<!--#Include File = "2014_SiteMap.Asp"-->   
<!--#Include File = "2014_NetWorkClass.Asp"-->
<!--#Include File = "2014_Doc.Asp"-->
<!--#Include File = "2014_Xls.Asp"-->
<!--#Include File = "2014_Cache.Asp"-->
<!--#Include File = "2014_Cookies.Asp"-->
<!--#Include File = "2014_Author.Asp"-->
<!--#Include File = "2015_Special.Asp"--> 
<!--#Include File = "2015_Formatting.Asp"--> 
<!--#Include File = "2015_Color.Asp"--> 
<!--#Include File = "2015_Param.Asp"-->
<!--#Include File = "2015_ASP.Asp"-->
<!--#Include File = "2015_BieRen.Asp"--> 
<!--#Include File = "2015_SqlIn.Asp"-->  
<!--#Include File = "2015_Json.Asp"-->
<!--#Include File = "2015_PHP.Asp"-->  
<!--#Include File = "2015_WebSql.Asp"-->
<!--#Include File = "2015_hexSwitch.Asp"-->  
<!--#Include File = "2015_APGeneral.Asp"-->  
<!--#Include File = "2015_ToMyPHP.Asp"-->   
<!--#Include File = "2015_Editor.Asp"-->   
<!--#Include File = "2015_Editor.Asp"-->    
<!--#Include File = "2015_NewWebFunction.Asp"-->   
<!--#Include File = "2016_WebControl.Asp"-->    
<!--#Include File = "2016_SaveData.Asp"-->     

<!--#Include File = "ZClass_YouKou.Asp"--> 
<!--#Include File = "ZClass_Maiside.Asp"--> 
<!--#Include File = "SystemInfo.Asp"--> 
<!--#Include File = "2016_PackXML.Asp"--> 
 
<% 

'ExecuteGlobal("S=""bb""")				'可动态加载文件不过不能有<%出现

Server.ScriptTimeOut = 9999	'超时
Public db, Conn, Rs, Rsx, Rss, Rst, Rsd, TempRs, TempRs2, RsTemp
Public nCount, nPageSize, MaxPage, Page 

Public WebTitle, WebUrl, WebKeywords, WebDescription, TemplatePath, TemplateContent, CreateHtmlPath, ClassStr, CreatePage
Dim MainNav,MainTitle,MainContent
Public SysStyle(9)
Public PubTimer : PubTimer = now()		'Timer()不要用这个(20150728)
'主程序全局变量设置
Public PubNavDid,PubNavSid,PubProDid,PubProSid,PubProTid,PubActionType,PubID,PubClassName,PubWD
Dim PubNavType				'网站导航类型
Public ReadBlockList			'读出块列表 配置FindModuleStr函数使用
Public ExecuteResut				'执行返回结果

dim makeHtmlFileToLCase	 : makeHtmlFileToLCase=true		'生成HTML文件转小写
dim isWebLabelClose : isWebLabelClose=true					'闭合标签(20150831)

dim HandleisCache : HandleisCache=false						'缓冲是否处理了
dim adminDir :adminDir="/admin/" 
MDBPath = adminDir &"/Data/OLDData.mdb"

SysStyle(0) = "#999999"
Response.Addheader "Content-Type","text/html; charset=gb2312"
'Response.Addheader "Content-Type","application/json; charset=gb2312"
'Response.Addheader "Content-Type","text/xml; charset=gb2312"

	'禁止缓冲
	Response.Buffer = True
	Response.Expires = 0
	Response.ExpiresAbsolute = Now() - 1
	Response.CacheControl = "no-cache"
	Response.AddHeader "Expires",Date()
	Response.AddHeader "Pragma","no-cache"
	Response.AddHeader "Cache-Control","private, no-cache, must-revalidate" 
Select Case Request("A")
	Case "HandWebImgPath" : 	'处理网站图片路径
	Call OpenConn()
	Rs.Open"Select * From [Product]",Conn,1,3
	Call Echo("表[Product]", "共有数据["& Rs.RecordCount &"]条")
	doEvents
	While Not Rs.Eof
		Rs("Banner") = HandleHttpUrl(Rs("Banner"))
		Rs("BigFiles") = HandleHttpUrl(Rs("BigFiles"))
		Rs("SmallFiles") = HandleHttpUrl(Rs("SmallFiles"))
		Rs.Update
		doEvents
	Rs.MoveNext:Wend:Rs.Close
	Call Echo("把数据库中\改成/","完成")
	Response.End()
End Select



%> 

