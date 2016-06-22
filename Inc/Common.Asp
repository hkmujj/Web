<% 
'常用函数大全 (2013,9,27Option Explicit)

'成功永远都是缓慢的 2013,10,4,悟

'显示弹窗对话框20150312
Function msgBox(ByVal content)
    content = Replace(Replace(content, Chr(10), "\n"), Chr(13), "\n") 
    Response.Write("<script>alert('" & content & "');</script>") 
End Function 
'提示20150729
Function MBInfo(title, content)
    MBInfo = msgBox(title) 
End Function 
'给Queststring赋值
Sub addRq(GookeName, valueStr)
    Response.QueryString(GookeName) = valueStr 
End Sub 
'获得Cookies值
Function rc(GookeName)
    rc = Request.Cookies(GookeName) 
End Function 
'给Cookies赋值
Sub addRc(GookeName, valueStr, DateStr)
    Response.Cookies(GookeName) = valueStr 
    If DateStr <> "" Then Response.Cookies(GookeName).Expires = DateStr 
End Sub 
'ASP自带的跳转
Sub rr(url)
    Response.Redirect(url) 
End Sub 
'替换Request.Form对象
Function rf(str)
    rf = Request.Form(str) 
End Function 
Function ADSqlRf(inputName)
    ADSqlRf = Replace(Request.Form(inputName), "'", "''") 
End Function 
'获得传值
Function rq(str)
    rq = Request.QueryString(str) 
End Function 
'获得传值
Function rfq(str)
    rfq = Request(str) 
End Function 
'替换Response.Write对象
Function rw(str)
    Response.Write(str) 
End Function 
'输出内容加换行
Function rwBr(str)
    Response.Write(str & vbCrLf) 
End Function 
'替换Response.Write对象 + Response.End()
Function rwEnd(str)
    Response.Write(str) 
    Response.End() 
End Function 
'HTML结束
Function htmEnd(str)
    Call RwEnd(str) 
End Function 
'替换Response.Write对象+Response.End()
Function die(str)
    Response.Write(str) 
    Response.End() 
End Function 
'替换Response.Write对象
Function debug(str)
    Response.Write("<div  style=""border:solid 1px #000000;margin-bottom:2px;"">调试" & str & "</div>" & vbCrLf) 
End Function 
'跟踪
Function trace(str)
    trace = debug(str) 
End Function 
'测试显示信息
Function echo(Word, str)
    Call echoPrompt(Word, str) 
End Function 
'测试显示信息+红色
Function echoRed(Word, str)
    Response.Write("<font color=red>" & Word & "</font>：" & str & "<br>") 
End Function 
'测试显示信息+红色+粗
Function echoRedB(Word, str)
    Response.Write("<b><font color=red>" & Word & "</font>：" & str & "</b><br>") 
End Function 
'测试显示信息+黄色
Function echoYellow(Word, str)
    Response.Write("<font color=#B38704>" & Word & "</font>：" & str & "<br>") 
End Function 
'测试显示信息+黄色+粗
Function echoYellowB(Word, str)
    Response.Write("<b><font color=#B38704>" & Word & "</font>：" & str & "</b><br>") 
End Function 
'回显内容
Function echoPrompt(Word, str)
    Response.Write("<font color=Green>" & Word & "</font>：" & str & "<br>") 
End Function 
'回显内容
Function echoStr(Word, str)
    Call echoPrompt(Word, str) 
End Function 
'打印数组 打印PHP里用到
Function print_R(content)
    Dim i, c 
    If TypeName(content) = "Variant()" Then
        For i = 0 To UBound(content)
            c = c & "[" & i & "] => " & content(i) & vbCrLf 
        Next 
    Else
        c = content 
    End If 
    Response.Write(c) 
End Function 
'测试显示信息 颜色
Function setColorEcho(color, Word, str)
    Response.Write("<font color=" & color & ">" & Word & "</font>：" & str & "<br>") 
End Function 
'测试显示信息暂停
Function eerr(Word, str)
    'Response.Write(TypeName(Word) & "-" & TypeName(Str)):Response.End()
    Response.Write("<font color=red>" & Word & "</font>：" & str) 
    Response.End() 
End Function 
'立即回显内容
Sub doEvents()
    Response.flush() 
End Sub 
'功能:ASP里的IIF 如：IIf(1 = 2, "a", "b")
Function IIf(bExp, sVal1, sVal2)
    If(bExp) Then IIf = sVal1 Else IIf = sVal2 
End Function 
'Hr
Function HR()
    Response.Write("<hr size='1' color='#666666'> ") 
End Function 
'BR 20160517
Function BR()
    Response.Write("<br/>") 
End Function 

'输出字符串 引用别人20141217
'Public Sub Echo(ByVal s) : Response.Write(s) : End Sub
'输出字符串和一个换行符
Public Sub print(ByVal s)
    Response.Write(s & vbCrLf) : Response.flush() 
End Sub 
'输出字符串和一个html换行符
Public Sub println(ByVal s)
    Response.Write(s & "<br />" & vbCrLf) : Response.flush() 
End Sub 
'输出字符串并将HTML标签转为普通字符
Public Sub printHtml(ByVal s)
    Response.Write(displayHtml(s) & vbCrLf) 
End Sub 
Public Sub printlnHtml(ByVal s)
    Response.Write(displayHtml(s) & "<br />") & vbCrLf 
End Sub 
'将任意变量直接输出为字符串(Json格式)
'Public Sub PrintString(ByVal s) : Response.Write(Str.ToString(s) & VbCrLf) : End Sub
'Public Sub PrintlnString(ByVal s) : Response.Write(Str.ToString(s)) & "<br />" & VbCrLf : End Sub
'输出经过格式化的字符串
'Public Sub PrintFormat(ByVal s, ByVal f) : Response.Write(Str.Format(s, f)) & VbCrLf : End Sub
'Public Sub PrintlnFormat(ByVal s, ByVal f) : Response.Write(Str.Format(s, f)) & "<br />" & VbCrLf : End Sub
'输出字符串并终止程序运行
Public Sub printEnd(ByVal s)
    Response.Write(s) : Response.End() 
End Sub 




'判断是否一样，一样返回checked,否者返回空值
Function isChecked(ByVal str, ByVal str2)
    If str = str2 Then isChecked = "checked='checked'" Else isChecked = "" 
End Function 
'判断是否一样，一样返回selected,否者返回空值
Function isSelected(ByVal str, ByVal str2)
    If str = str2 Then isSelected = "selected='selected'" Else isSelected = "" 
End Function 

'处理错误信息 If Err Then doError Err.Description, "Msg"
Sub doError(s, msg)
    'On Error Resume Next
    Dim nRnd, c 
    Randomize 
    nRnd = CLng(Rnd() * 29252888) 
    c = "<br />" 
    c = c & "<div style=""width:100%; font-size:12px;;line-height:150%"">" 
    c = c & "  <label onClick=""ERRORDIV" & nRnd & ".style.display=(ERRORDIV" & nRnd & ".style.display=='none'?'':'none')"">" 
    c = c & "  <span style=""background-color:#820222;color:#FFFFFF;height:23px;font-size:14px;cursor:pointer"">〖 出错 提示信息 ERROR 〗</span><br />" 
    c = c & "  </label>" 
    c = c & "  <div id=""ERRORDIV" & nRnd & """ style=""width:100%;border:1px solid #820222;padding:5px;overflow:hidden;"">" 
    c = c & " <span style=""color:#FF0000;"">出错描述</span> " & s & "<br />" 
    c = c & " <span style=""color:#FF0000;"">回显信息</span> " & msg & "<br />" 
    c = c & "  </div>" 
    c = c & "</div>" 
    c = c & "<br />" 
    Response.Write(c) 
    Response.End()                                                                  '终止，程序停止
End Sub 

'执行ASP脚本
Function exec(ByVal tStr)
    If tStr = "" Then Exit Function : End If 
    On Error Resume Next 
    Execute(tStr) 
    On Error Goto 0 
End Function 

%>  

