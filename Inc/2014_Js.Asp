<% 
'Js

'删除JS注释 20160602
function delJsNote(content)
	dim splstr,s,c,isMultiLineNote,s2
	isMultiLineNote=false			'多行注释默认为假
	splstr=split(content,vbcrlf)
	for each s in splstr
		s2=phptrim(s)
		if isMultiLineNote=true then
			if len(s2)>=2 then
				if right(s2,2)="*/" then
					isMultiLineNote=false
				end if
			end if				
			s=""
		else	
			if left(s2,2)="/*" then
				if right(s2,2)<>"*/" then
					isMultiLineNote=true
				end if
				s=""
			elseif left(s2,2)="//" then
				s=""
			end if
		end if
		c=c & s & vbcrlf
	next
	delJsNote=c
end function

'JS转换，引用别人
Function JsEncode__(ByVal s)

	If isNul(s) Then JsEncode__ = "" : Exit Function
	Dim arr1, arr2, i, j, c, p, t
	arr1 = Array(chr(34),chr(92),chr(47),chr(8),chr(12),chr(10),chr(13),chr(9))		 '34|",92|\,47|/,8|,12|,10| ,13| ,9|	,
	arr2 = Array(chr(34),chr(92),chr(47),chr(98),chr(102),chr(110),chr(114))		 '34|",92|\,47|/,98|b,102|f,110|n,114|r,1865|,
	For i = 1 To Len(s)
		p = True
		c = Mid(s, i, 1)
		For j = 0 To Ubound(arr1)
			If c = arr1(j) Then
				t = t & "\" & arr2(j)
				p = False
				Exit For
			End If
		Next
		If p Then t = t & c
	Next
	JsEncode__ = t
End Function


'远程网站会员统计2010330
'<script>document.writeln("<script src=\'http://127.0.0.1/web_soft/R.Asp?act=Stat&GoToUrl="+escape(document.referrer)+"&ThisUrl="+escape(window.location.href)+"&screen="+escape(window.screen.width+"x"+window.screen.height)+"&co="+escape(document.cookie)+" \'><\/script>");<'/script>
Function showStatJSCode(url)
    showStatJSCode = "<script>document.writeln(""<script src=\'" & url & "act=Stat&GoToUrl=""+escape(document.referrer)+""&ThisUrl=""+escape(window.location.href)+""&screen=""+escape(window.screen.width+""x""+window.screen.height)+""&co=""+escape(document.cookie)+"" \'><\/script>"");</script>" 
End Function 


'Js定时跳转 Timing = 定时 时间测定 例：Call Rw("账号或密码错误，" & JsTiming("返回", 5))
Function jsTiming(url, seconds)
    Dim c 
    c = c & "<span id=mytimeidboyd>倒计时</span>" & vbCrLf 
    c = c & "<script type=""text/javascript"">" & vbCrLf 
    c = c & "//配置Config" & vbCrLf 
    c = c & "var coutnumb" & vbCrLf 
    c = c & "coutnumb=" & seconds & "" & vbCrLf 
    c = c & "" & vbCrLf 
    c = c & "//定时跳转" & vbCrLf 
    c = c & "function Countdown(){" & vbCrLf                                        'Countdown=倒数计秒
    c = c & "    coutnumb-=1" & vbCrLf 
    c = c & "    mytimeidboyd.innerHTML=""倒计时<font color=#000000>""+coutnumb+""</font>""" & vbCrLf 
    c = c & "    if(coutnumb<1){    " & vbCrLf 

    If url = "back" Or url = "返回" Then                                            '当Action为back是返回上页
        c = c & "        history.back();" & vbCrLf 
    Else
        c = c & "        location.href='" & url & "';" & vbCrLf 
    End If 


    c = c & "    }else{" & vbCrLf 
    c = c & "        setTimeout(""Countdown()"",1000);" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "}setTimeout(""Countdown()"",1)" & vbCrLf 
    c = c & "</script>" & vbCrLf 
    jsTiming = c 
End Function 
'JS弹窗 Call Javascript("返回", "操作成功", "")
Function javascript(action, msg, url)
    If msg <> "" Then msg = "alert('" & msg & "');"                                 '当Msg不为空则弹出信息
    If action = "back" Or action = "返回" Then                                      '当Action为back是返回上页
        Response.Write("<script>" & msg & "history.back();</script>") 
    ElseIf url <> "" Then                                                           '当Url不为空
        Response.Write("<script>" & msg & "location.href='" & url & "';</script>") '跳转Url页
    Else
        Response.Write("<script>" & msg & "</script>") 
    End If 
    Response.End() 
End Function 
'创建Ajax对象实例
Function createAjax()
    Dim c 
    c = "<script language=""javascript"">" & vbCrLf 
    c = c & "//AjAX XMLHTTP对象实例" & vbCrLf 
    c = c & "function createAjax() { " & vbCrLf 
    c = c & "    var _xmlhttp;" & vbCrLf 
    c = c & "    try {    " & vbCrLf 
    c = c & "        _xmlhttp=new ActiveXObject(""Microsoft.XMLHTTP"");    //IE的创建方式" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "    catch (e) {" & vbCrLf 
    c = c & "        try {" & vbCrLf 
    c = c & "            _xmlhttp=new XMLHttpRequest();    //FF等浏览器的创建方式" & vbCrLf 
    c = c & "        }" & vbCrLf 
    c = c & "        catch (e) {" & vbCrLf 
    c = c & "            _xmlhttp=false;        //如果创建失败，将返回false" & vbCrLf 
    c = c & "        }" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "    return _xmlhttp;    //返回xmlhttp对象实例" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "//Ajax" & vbCrLf 
    c = c & "function Ajax(URL,ShowID) {  " & vbCrLf 
    c = c & "    var xmlhttp=createAjax();" & vbCrLf 
    c = c & "    if (xmlhttp) {" & vbCrLf 
    c = c & "        URL+= ""&n=""+Math.random() " & vbCrLf 
    c = c & "        xmlhttp.open('post', URL, true);//基本方法" & vbCrLf 
    c = c & "        xmlhttp.setRequestHeader(""cache-control"",""no-cache""); " & vbCrLf 
    c = c & "        xmlhttp.setRequestHeader(""Content-Type"", ""application/x-www-form-urlencoded"");         " & vbCrLf 
    c = c & "        xmlhttp.onreadystatechange=function() {        " & vbCrLf 
    c = c & "            if (xmlhttp.readyState==4 && xmlhttp.status==200) {     " & vbCrLf 
    c = c & "                document.getElementById(ShowID).innerHTML = ""操作完成""// unescape(xmlhttp.responseText); " & vbCrLf 
    c = c & "            }" & vbCrLf 
    c = c & "            else {                " & vbCrLf 
    c = c & "                document.getElementById(ShowID).innerHTML = ""正在加载中...""" & vbCrLf 
    c = c & "            }" & vbCrLf 
    c = c & "        }" & vbCrLf 
    'c=c & "alert(document.all.TEXTContent.value)" & vbcrlf
    c = c & "        xmlhttp.send(""Content=""+escape(document.all.TEXTContent.value)+"""");    " & vbCrLf 
    c = c & "        //alert(""网络错误"");" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "function GetIDHTML(Root){" & vbCrLf 
    c = c & "    alert(document.all[Root].innerHTML)" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "</script>" & vbCrLf 
    createAjax = c 
End Function 
'JS在线编辑
Function onLineEditJS()
    Dim c 
    c = c & "<script language=""javascript"">" & vbCrLf 
    c = c & "//显示与编辑内容，但不修改，ASP代码分析器用 创作于2013,10,5" & vbCrLf 
    c = c & "function TestInput(Root){ " & vbCrLf 
    c = c & "    var TempContent" & vbCrLf 
    c = c & "    TempContent="""" " & vbCrLf 
    c = c & "    " & vbCrLf 
    c = c & "    document.all[Root].title=""""" & vbCrLf 
    c = c & "    if(document.all[Root].innerHTML.indexOf(""<TEXTAREA"")==-1){" & vbCrLf 
    c = c & "            TempContent=document.all[Root].innerHTML" & vbCrLf 
    c = c & "            TempContent=TempContent.replace(/<BR><BR>/g,""<BR>"");     " & vbCrLf 
    c = c & "            TempContent=TempContent.replace(/<BR>/g,""\n"");     " & vbCrLf 
    c = c & "            if(TempContent==""&nbsp;""){TempContent=""""}" & vbCrLf 
    c = c & "            document.all[Root].innerHTML=""<textarea name=TEXT""+Root+"" style='width:50%;height:50%' onblur=if(this.value!=''){document.all.""+Root+"".title='点击可编辑';document.all.""+Root+"".innerHTML=ReplaceNToBR(this.value)}else{document.all.""+Root+"".innerHTML='&nbsp;'};>"" + TempContent + ""</textarea>"";" & vbCrLf 
    c = c & "            document.all[""TEXT""+Root].focus();" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "function ReplaceNToBR(Content){" & vbCrLf 
    c = c & "    return Content.replace(/\n/g,""<BR>"")" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "</script>" & vbCrLf 
    onLineEditJS = c 
End Function 
'在线编辑
Function editTXT(content, jsId)
    content = IIF(content = "", "&nbsp;", content) 
    editTXT = "<span id='" & jsId & "' onClick=""TestInput('" & jsId & "');"" title='点击可编辑'>" & content & "</span>" 
End Function 
'在线编辑  (辅助)
Function onLineEdit(content, jsId)
    onLineEdit = editTXT(content, jsId) 
End Function 
'****************************************************
'函数名：JSGoTo
'作  用：显示文本
'时  间：2013年12月14日
'参  数：Url
'*       SetTime
'返回值：字符串
'调  试：Call Echo("测试函数 JSGoTo", JSGoTo("", "",""))
'****************************************************
Function jsGoTo(title, url, setTime)
    Dim c 
    If title = "" Then title = "添加成功" 
    If setTime = "" Then setTime = 4                                                '默认为4秒
    c = c & "<script>" & vbCrLf 
    c = c & "//通用定时器 如：MyTimer('Show', 'alert(1+1)', 5)" & vbCrLf 
    c = c & "var StopTimer = """"" & vbCrLf 
    c = c & "function MyTimer(ID, ActionStr,TimeNumb){" & vbCrLf 
    c = c & "    if(StopTimer == ""停止"" || StopTimer == ""停止定时器""){" & vbCrLf 
    c = c & "        StopTimer = """"" & vbCrLf 
    c = c & "        return false" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "    TimeNumb--" & vbCrLf 
    c = c & "    document.all[ID].innerHTML = ""倒计时："" + TimeNumb" & vbCrLf 
    c = c & "    if(TimeNumb<1){" & vbCrLf 
    c = c & "        setTimeout(ActionStr,100);" & vbCrLf 
    c = c & "    }else{" & vbCrLf 
    c = c & "        setTimeout(""MyTimer('""+ID+""', '""+ActionStr+""',""+TimeNumb+"")"",1000);" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "function GotoURL(){" & vbCrLf 
    c = c & "    location.href='" & url & "'" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "</script>" & vbCrLf 
    c = c & "<div id=""Show"">Loading...</div><script>MyTimer('Show', 'GotoURL()', " & setTime & ")</script>" & vbCrLf 
    jsGoTo = c 
End Function 

'JS图片滚动
Function jsPhotoScroll(id, width, height)
    Dim c 
    c = c & "<script type=""text/javascript"">" & vbCrLf 
    c = c & "    var marqueeB = new Marquee(""" & id & """)    " & vbCrLf 
    c = c & "    marqueeB.Direction =2;" & vbCrLf 
    c = c & "    marqueeB.Step = 1;" & vbCrLf 
    c = c & "    marqueeB.Width = " & width & ";" & vbCrLf 
    c = c & "    marqueeB.Height = " & height & ";" & vbCrLf 
    c = c & "    marqueeB.Timer = 1;" & vbCrLf 
    c = c & "    marqueeB.DelayTime = 0;" & vbCrLf 
    c = c & "    marqueeB.WaitTime = 0;" & vbCrLf 
    c = c & "    marqueeB.ScrollStep = 20;" & vbCrLf 
    c = c & "    marqueeB.Start();    " & vbCrLf 
    c = c & "</script>" & vbCrLf 
    jsPhotoScroll = c 
End Function 
'图片向左滚动（暂不用）
Function photoLeftScroll(demo, demo1, demo2)
    Dim c 
    c = c & "<!--图片向左轮番滚动-->" & vbCrLf 
    c = c & "<script language=""javascript"">" & vbCrLf 
    c = c & "var speed=30" & vbCrLf 
    c = c & "" & demo2 & ".innerHTML=" & demo1 & ".innerHTML" & vbCrLf 
    c = c & "function Marquee(){" & vbCrLf 
    c = c & "    if(" & demo2 & ".offsetWidth-" & demo & ".scrollLeft<=0)" & vbCrLf 
    c = c & "        " & demo & ".scrollLeft-=" & demo1 & ".offsetWidth" & vbCrLf 
    c = c & "    else{" & vbCrLf 
    c = c & "        " & demo & ".scrollLeft++" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "var MyMar=setInterval(Marquee,speed)" & vbCrLf 
    c = c & "" & demo & ".onmouseover=function() {clearInterval(MyMar)}" & vbCrLf 
    c = c & "" & demo & ".onmouseout=function() {MyMar=setInterval(Marquee,speed)}" & vbCrLf 
    c = c & "</script> " & vbCrLf 
    photoLeftScroll = c 
End Function 
%> 

