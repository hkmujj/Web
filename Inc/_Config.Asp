<%Option Explicit%>
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
<!--#Include File = "Print.Asp"--> 
<!--#Include File = "Img.Asp"--> 
<!--#Include File = "PinYin.Asp"-->
<!--#Include File = "Md5.Asp"-->  
<!--#Include File = "IE.Asp"-->  
<!--#Include File = "ASPPHPAccess.Asp"-->    
<!--#Include File = "2014_Array.Asp"--> 
<!--#Include File = "2014_Js.Asp"--> 
<!--#Include File = "2014_Css.Asp"--> 
<!--#Include File = "2014_Template.Asp"-->   
<!--#Include File = "2014_Nav.Asp"-->   
<!--#Include File = "2014_GBUTF.Asp"-->   
<!--#Include File = "2014_Form.Asp"-->    
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
<!--#Include File = "2016_Log.Asp"-->      
<!--#Include File = "SystemInfo.Asp"-->    
<!--#Include File = "Json.Asp"-->
<!--#Include File = "2016_PackXML.Asp"--> 
<% 

 
Server.ScriptTimeOut = 9999	'超时
Response.Addheader "Content-Type","text/html; charset=gb2312"
	'禁止缓冲
	Response.Buffer = True
	Response.Expires = 0
	Response.ExpiresAbsolute = Now() - 1
	Response.CacheControl = "no-cache"
	Response.AddHeader "Expires",Date()
	Response.AddHeader "Pragma","no-cache"
	Response.AddHeader "Cache-Control","private, no-cache, must-revalidate" 
dim Conn, Rs, Rsx, Rss, Rst, Rsd, TempRs, TempRs2, RsTemp
dim PubTimer : PubTimer = now()		'Timer()不要用这个(20150728)

'Note:开发程序原则，能简单，绝对不复杂，能修改html的，绝对不修改asp，能傻瓜式，绝对不要用脑(20160121)
Dim WEBCOLUMNTYPE : WEBCOLUMNTYPE = "首页|文本|产品|新闻|视频|下载|案例|留言|反馈|招聘|订单" 
Dim WEB_VIEWURL : WEB_VIEWURL = "/aspweb.asp"                '前台显示
dim WEB_ADMINURL : WEB_ADMINURL ="/web/1.asp"				'后台显示
Dim EDITORTYPE : EDITORTYPE = "asp"                                             '程序类型，是ASP,或PHP,或jSP,或.NET
dim webDir:webDir=""
MDBPath = webDir & "/Data/Data.mdb"
If checkFile(MDBPath) = False Then
    webDir = "/" & getUrlDir(remoteUrlFileName(getThisUrl())) 
	webDir=mid(webDir,1,len(webDir)-1)
    webDir = Replace(LCase(webDir), "admin/", "") 
    MDBPath = webDir & "/Data/Data.mdb" 
'call echo(webDir,mdbpath)
End If
dim onCacheHtml:onCacheHtml=true
'=========


dim ReadBlockList

dim SysStyle(9)
SysStyle(0) = "#999999"
dim makeHtmlFileToLCase	 : makeHtmlFileToLCase=true		'生成HTML文件转小写
dim isWebLabelClose : isWebLabelClose=true					'闭合标签(20150831)

dim HandleisCache : HandleisCache=false						'缓冲是否处理了
Dim db_PREFIX : db_PREFIX = "xy_"    		                                         '表前缀
dim adminDir :adminDir="/web/"							'后台目录


Dim openErrorLog : openErrorLog = true                                             '开启错误日志
Dim openWriteSystemLog : openWriteSystemLog = "|txt|database|"                                             '开启写系统日志 txt写入文本 database写入数据库
dim isTestEcho: isTestEcho=true											'开关测试回显
Dim webVersion : webVersion="ASPPHPCMS v1.5"												'网站版本

 
dim WEB_CACHEFile : WEB_CACHEFile=webDir & "/web/"& EDITORTYPE &"cachedata.txt"								'缓冲文件
dim WEB_CACHEContent : WEB_CACHEContent=""								'缓冲文件内容
dim isCacheTip : isCacheTip=false			'是否开启缓冲提示

dim language : language="#en-us"			'en-us  | zh-cn | zh-tw

'批量替换语言
function batchReplaceL(content,str)
	dim splstr,s,i
	splstr=split(str,"|*|")
	for i = 0 to ubound(splstr)
		s=splstr(i)
		if s <>"" then
			content=replaceL(content,s)
		end if
	next
	batchReplaceL=content
end function
'替换语言
function replaceL(content,str)  
	replaceL=replace(content,str, setL(str))	
end function
'语言
function setL(str)
	dim c
	c=str
	if language="en-us" then
		c=languageEN(c)
	end if
	setL=c
end function  
'处理显语言  c=handleDisplayLanguage(c,"loginok")
function handleDisplayLanguage(c,sType)
	'繁体就直接转换了，不要再一个一个转了，
	if language="zh-tw" then
		handleDisplayLanguage=simplifiedTransfer(c)
		exit function
	end if
	if sType="login" then
	
		c = batchReplaceL(c, "请不要输入特殊字符|*|输入正确|*|用户名可以用字母|*|用户名可以用字母|*|您的用户名为空|*|密码可以用字母|*|您的密码为空")
		c = batchReplaceL(c, "登录后台|*|管理员登录|*|如果您不是管理员|*|请立即停止您的登陆行为|*|用户名|*|版")
		c = batchReplaceL(c, "密&nbsp;码|*|密码|*|请输入|*|登 录|*|登录|*|重 置|*|重置")
		
		
	elseif sType="loginok" then
		c = batchReplaceL(c, "后台地图|*|清除缓冲|*|超级管理员|*|当前位置|*|管理员信息|*|修改密码|*|最新访客信息|*|开发团队|*|版权所有|*|开发与支持团队")
		c = batchReplaceL(c, "进入在线修改模式|*|系统信息|*|免费开源版|*|授权信息|*|服务器名称|*|服务器版本|*|交流群|*|相关链接|*|登录后台")
		c = batchReplaceL(c, "用户名|*|表前缀|*|帮助|*|退出|*|您好|*|首页|*|权限|*|端口|*|邮箱|*|官网|*|版|*|云端")	
		c = batchReplaceL(c, "系统管理|*|我的面板|*|栏目管理|*|模板管理|*|会员管理|*|生成管理|*|更多设置")
		
		c = batchReplaceL(c, "站点配置|*|网站统计|*|生成|*|后台操作日志|*|后台管理员|*|网站栏目|*|分类信息|*|评论|*|搜索统计|*|单页管理|*|友情链接|*|招聘管理")
		c = batchReplaceL(c, "反馈管理|*|留言管理|*|会员配置|*|竞价词|*|网站布局|*|网站模块|*|后台菜单|*|执行|*|仿站")
		
		
	end if
	handleDisplayLanguage=c
end function

'为英文
function languageEN(str) 
	dim c
	if str="登录成功，正在进入后台..." then
		c="Login successfully, is entering the background..."
	elseif str="账号密码错误<br>登录次数为 " then
		c="Account password error <br> login "
	elseif str="登录成功，正在进入后台..." then
		c="Login successfully, is entering the background..."
	elseif str="退出成功" then
		c="Exit success"
	elseif str="退出成功，正在进入登录界面..." then
		c="Quit successfully, is entering the login screen..."
	elseif str="清除缓冲完成，正在进入后台界面..." then
		c="Clear buffer finish, is entering the background interface..."
	elseif str="提示信息" then
		c="Prompt info"
	elseif str="如果您的浏览器没有自动跳转，请点击这里" then
		c="If your browser does not automatically jump, please click here"
	elseif str="倒计时"	 then
		c="Countdown "
	elseif str="后台地图"	 then
		c="Admin map"
	elseif str="清除缓冲"	 then
		c="Clear buffer"
	elseif str="超级管理员"	 then
		c="Super administrator"
	elseif str="当前位置"	 then
		c="current location"
	elseif str="管理员信息"	 then
		c="Admin info"
	elseif str="修改密码"	 then
		c="Modify password"
	elseif str="用户名"	 then
		c="username"
	elseif str="表前缀"	 then
		c="Table Prefix"
	elseif str="进入在线修改模式"	 then
		c="online modification"
	elseif str="系统信息"	 then
		c="system info"
	elseif str="授权信息"	 then
		c="Authorization information"
	elseif str="免费开源版"	 then
		c="Free open source"
	elseif str="服务器名称"	 then
		c="Server name"
	elseif str="服务器版本"	 then
		c="Server version"		
	elseif str="最新访客信息"	 then
		c="visitor info"
	elseif str="开发团队"	 then
		c="team info"		
	elseif str="版权所有"	 then
		c="copyright"
	elseif str="开发与支持团队"	 then
		c="Develop and support team"
	elseif str="交流群"	 then
		c="Exchange group"
	elseif str="相关链接"	 then
		c="Related links"
	elseif str="系统管理"	 then
		c="System"
	elseif str="我的面板"	 then
		c="My panel"
	elseif str="栏目管理"	 then
		c="Column"
	elseif str="模板管理"	 then
		c="Template"
	elseif str="会员管理"	 then
		c="Member"
	elseif str="生成管理"	 then
		c="Generation"
	elseif str="更多设置"	 then
		c="More settings"
		 
		
	elseif str="登录后台"	 then
		c="Login management background"
	elseif str="管理员登录"	 then
		c="Administrator login "
	elseif str="如果您不是管理员"	 then
		c="If you are not an administrator"
	elseif str="请立即停止您的登陆行为"	 then
		c="Please stop your login immediately"
	elseif str="密&nbsp;码"	or str="密码" then
		c="password"
	elseif str="请输入"	 then
		c="Please input"		
	elseif str="登 录"	or str="登录" then
		c="login"
	elseif str="重 置" or str="重置" then
		c="reset"
	elseif str="请不要输入特殊字符"	 then
		c="Please do not enter special characters"		
	elseif str="输入正确"	 then
		c="Input correct"		
	elseif str="用户名可以用字母"	 then
		c="Username use "		
	elseif str="您的用户名为空"	 then
		c="Your username is empty"
	elseif str="密码可以用字母"	 then
		c="Passwords use "	
	elseif str="您的密码为空"	 then
		c="Your password is empty"		 
	elseif str="站点配置"	 then
		c="Site configuration"		
	elseif str="网站统计"	 then
		c="Website statistics"		
	elseif str="后台操作日志"	 then
		c="Admin log "		
	elseif str="后台管理员"	 then
		c="Background manager"		
	elseif str="网站栏目"	 then
		c="Website column"		
	elseif str="分类信息"	 then
		c="Classified information"		
	elseif str="搜索统计"	 then
		c="Search statistics"		
	elseif str="单页管理"	 then
		c="Single page"		
	elseif str="友情链接"	 then
		c="Friendship link"		
	elseif str="招聘管理"	 then
		c="Recruitment management"		
	elseif str="反馈管理"	 then
		c="Feedback management"		
	elseif str="留言管理"	 then
		c="message management"		
	elseif str="会员配置"	 then
		c="Member allocation"		
	elseif str="竞价词"	 then
		c="Bidding words"	
	elseif str="网站布局"	 then
		c="Website layout"		
	elseif str="网站模块"	 then
		c="Website module"		
	elseif str="后台菜单"	 then
		c="Background menu"		
	elseif str="仿站"	 then
		c="Template website "	 
		
	elseif str="11111"	 then
		c="1111111"		
		
	 
		
	elseif str="执行"	 then
		c="implement "		
	elseif str="评论"	 then
		c="comment "		
	elseif str="生成"	 then
		c="generate "		
	elseif str="权限"	 then
		c="jurisdiction"
	elseif str="帮助"	 then
		c="Help"
	elseif str="退出"	 then
		c="sign out"
	elseif str="您好"	 then
		c="hello"
	elseif str="首页"	 then
		c="home"
	elseif str="端口"	 then
		c="port"
	elseif str="官网"	 then
		c="official website"
	elseif str="邮箱"	 then
		c="Emai"
	elseif str="云端"	 then
		c="Cloud"
		 
	elseif str="版"	 then
		c=" edition"
		
		 
		
		
	else
		c=str
	end if 
	languageEN=c
end function
%> 

