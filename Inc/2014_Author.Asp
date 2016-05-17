<% 
'关于作者信息
Sub aboutAuthor()
    Dim c 
    c = c & "<pre>" & vbCrLf 
    c = c & "作者：小孙" & vbCrLf 
    c = c & "联系方式" & vbCrLf 
    c = c & "QQ：313801120" & vbCrLf 
    c = c & "邮箱：313801120@qq.com" & vbCrLf 
    c = c & "微信：mq313801120" & vbCrLf 
    c = c & "交流群35915100(群里已有几百人)" & vbCrLf 
    c = c & "" & vbCrLf 
    c = c & "业务及特长" & vbCrLf 
    c = c & "精通ASP,VB程序开发，独立开发出一套ASP网站后台和VB辅助软件。" & vbCrLf 
    c = c & "熟练掌握HTML、DIV、CSS、JS" & vbCrLf 
    c = c & "熟练使用Dreamweaver、Fireworks、 Flash、Photoshop等软件" & vbCrLf 
    c = c & "自觉PHP、Android等编程语言" & vbCrLf 
    c = c & "" & vbCrLf 
    c = c & "" & vbCrLf 
    c = c & "自我评价" & vbCrLf 
    c = c & "自学能力强、新知识接受快、勇于面对困难，敢于挑战。" & vbCrLf 
    c = c & "" & vbCrLf 
    c = c & "" & vbCrLf 
    c = c & "工作经历" & vbCrLf 
    c = c & "2007年1月 至 2012年1月 上海子映网络" & vbCrLf 
    c = c & "工作内容：网站开发" & vbCrLf 
    c = c & "" & vbCrLf 
    c = c & "2013---2014，南京麦思德餐饮有限公司" & vbCrLf 
    c = c & "工作内容：网站开发" & vbCrLf 
    c = c & "" & vbCrLf 
    c = c & "2014---至今，微战略网络有限公司" & vbCrLf 
    c = c & "工作内容：网站整站开发，自己用VB开发出一款网站制作辅助软件。" & vbCrLf 
    c = c & "" & vbCrLf 
    c = c & "" & vbCrLf 
    c = c & "网站案例：" & vbCrLf 
    c = c & "http://www.863health.com/" & vbCrLf 
    c = c & "http://www.wzl99.com/" & vbCrLf 
    c = c & "http://www.jfh6666.com/" & vbCrLf 

    c = c & "</pre>" & vbCrLf 
    Response.Write(c) : Response.End() 
End Sub 

'作者信息
Function authorInfo(FileInfo)
    authorInfo = handleAuthorInfo(fileInfo,"asp")
End Function
'处理作者信息
function handleAuthorInfo(fileInfo,sType)	
    Dim c,phpS,aspS
	if sType="php" then
		phpS="/"
	else
		aspS="'"
	end if
    c = aspS & phpS & "************************************************************" & vbCrLf 
    If FileInfo <> "" Then c = c & aspS & "  文件：" & FileInfo & vbCrLf 
    c = c & aspS  &"作者：云端 (精通ASP/VB/PHP/JS/Flash，交流合作可联系本人)" & vbCrlf
    c = c & aspS &"版权：源代码公开，各种用途均可免费使用。 " & vbCrlf
    c = c & aspS &"创建：" & Format_Time(Now(), 2) & vbCrLf 
    c = c & aspS &"联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com" & vbCrlf
    c = c & aspS &"更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得" & vbCrlf
    c = c & aspS &"*                                    Powered by ASPPHPCMS " & vbCrlf
    c = c & aspS &"************************************************************" & phpS & vbCrlf
    handleAuthorInfo = c 
end function
 

Function authorInfo2()
    Dim c 
    c = "                '''" & vbCrLf 
    c = c & "               (0 0)" & vbCrLf 
    c = c & "   +-----oOO----(_)------------+" & vbCrLf 
    c = c & "   |                           |" & vbCrLf 
    c = c & "   |    让我们一起来体验       |" & vbCrLf 
    c = c & "   |    QQ:313801120           |" & vbCrLf 
    c = c & "   |    sharembweb.com         |" & vbCrLf 
    c = c & "   |                           |" & vbCrLf 
    c = c & "   +------------------oOO------+" & vbCrLf 
    c = c & "              |__|__|" & vbCrLf 
    c = c & "               || ||" & vbCrLf 
    c = c & "              ooO Ooo" & vbCrLf 

    authorInfo2 = c 
End Function 
%> 

