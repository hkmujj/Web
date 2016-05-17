<% 
'输出内容(2014,03,28)


'空网页内容 20160108
Function getBlankHtmlBody(webTitle, webBody)
    Dim c 
    c = "<!DOCTYPE html PUBLIC>" & vbCrLf 
    c = c & "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbCrLf 
    c = c & "<head>" & vbCrLf 
    c = c & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />" & vbCrLf 
    c = c & "<title>" & webTitle & "</title>" & vbCrLf 
    c = c & "</head>" & vbCrLf 
    c = c & "<body>" & vbCrLf 
    c = c & webBody & vbCrLf 
    c = c & "</body>" & vbCrLf 
    c = c & "</html>" & vbCrLf 
    getBlankHtmlBody = c 
End Function 



'错误提示 之前网站用到
Function errorText(Refresh, str, url)
    If Refresh<>"" Then 
		rw("<meta http-equiv=""refresh"" content=""" & Refresh & ";URL=" & url & """"""">") & vbCrLf 
    end if
	rw("<fieldset>") & vbCrLf 
    'rw("<legend></legend>") & vbCrLf 
    rw("<div style=""padding-left:20px;padding-top:10px;color:red;font-weight:bold;text-align:center;"">" & str & "</div>") & vbCrLf 
    rw("<div style=""height:200p;text-align:center;""><P>") & vbCrLf 
    rw("<a href=""" & url & """>如果您的游览器没有自动跳转,请点这里>></a><P>") & vbCrLf 
    rw("</div></fieldset>") : Response.End() 
End Function
 

'愿形显示内容
Function printPre(content)
    content = Replace(content, "<", "&lt;") 
    printPre = "<pre>" & content & "</pre>" 
End Function 


'折叠菜单
Function foldingMenu(id, s, msg)
    'On Error Resume Next
    Dim nRnd, c 
    If id = "" Then
        Randomize 
        nRnd = CLng(Rnd() * 29252888) 
    Else
        nRnd = id 
    End If 
    c = "<div style=""width:100%; font-size:12px;;line-height:150%;margin-bottom:4px;"">" 
    c = c & "  <label onClick=""ERRORDIV" & nRnd & ".style.display=(ERRORDIV" & nRnd & ".style.display=='none'?'':'none')"">" 
    c = c & "  <span style=""background-color:#666;color:#FFFFFF;height:23px;font-size:14px;cursor:pointer"">〖 " & s & " 〗</span><br />" 
    c = c & "  </label>" 
    c = c & "  <div id=""ERRORDIV" & nRnd & """ style=""width:100%;border:1px solid #820222;padding:5px;overflow:hidden;display:none;"">" 
    c = c & msg 
    c = c & "  </div>" 
    c = c & "</div>" 
    foldingMenu = c 
End Function 
'返回标题红色
Function returnRen(Lable, msg)
    returnRen = "<font color=red>" & Lable & "</font>：" & msg 
End Function 
'返回Hr
Function returnHr()
    returnHr = "<hr size='1' color='#666666'> " 
End Function 
'返回Hr
Function returnRenHr()
    returnRenHr = "<hr size='1' color='red'> " 
End Function 
'处理错误信息 第二种
Sub showErr(ByVal ErrCode, ByVal ErrDesc)
    Dim c 
    c = "<style>.ab-showerr{width:400px;font-size:12px;font-family:Consolas;margin:10px auto;padding:0;background-color:#FFF;}" 
    c = c & ".ab-showerr h3,.ab-showerr h4{font-size:12px;margin:0;line-height:24px;text-align:center;background-color:#999;border:1px solid #555;color:#FFF;border-bottom:none;}.ab-showerr h4{padding:5px;line-height:1.5em;text-align:left;background-color:#FFC;color:#000; font-weight:normal;}" 
    c = c & ".ab-showerr h4 strong{color:red;}.ab-showerr table{width:100%;margin:0;padding:0;border-collapse:collapse;border:1px solid #555;border-bottom:none;}.ab-showerr th{background-color:#EEE;white-space:nowrap;}.ab-showerr thead th{background-color:#CCC;}.ab-showerr th,.ab-showerr td{font-size:12px;border:1px solid #999;padding:6px;line-height:20px;word-break:break-all;}.ab-showerr span.info{color:#F30;}</style>" 
    c = c & "<div class=""ab-showerr""><h3>Microsoft VBScript 编译器错误</h3><h4>程序出错了，错误代码： <strong>" & ErrCode & "</strong> ，以下是错误描述：</h4><table><tr><td>" & ErrDesc & "</td></tr></table></div>" 
    Response.Write(c) : Response.End() 
End Sub 

'打印Form表单转送的值，方便写程序  因为它不能与VB软件共存  从Print.Asp引用过来
Sub printFormInfo()
    Dim content, NameList, DimVariablesList, GetFormList, AspCodeList, ShowContent, PrintFieldList 
    content = Request.Form 
    Dim v, a, b                                                                     '所有表单值
    Dim t                                                                           '所有表单数量
    ReDim a(Request.Form.Count)                                                     '表单名称数组
    ReDim b(Request.Form.Count)                                                     '对应表单值数组
    t = Request.Form.Count 
    v = Request.Form 
    DimVariablesList = "Dim " 
    For i = 0 To t - 1
        a(i) = Split(Split(v, "&")(i), "=")(0) 
        b(i) = Split(Split(v, "&")(i), "=")(1) 
        NameList = NameList & a(i) & "|" 
        DimVariablesList = DimVariablesList & a(i) & ", " 
        ShowContent = ShowContent & a(i) & "=" & b(i) & "<br>" 
        GetFormList = GetFormList & CopyStr("&nbsp;", 4) & a(i) & " = Trim(Rf(""" & a(i) & """))<br>" 
        AspCodeList = AspCodeList & CopyStr("&nbsp;", 8) & "Rs(""" & a(i) & """) = " & a(i) & "<br>" & vbCrLf 
        PrintFieldList = PrintFieldList & CopyStr("&nbsp;", 8) & a(i) & " = Rs(""" & a(i) & """) " & "<br>" & vbCrLf 
    Next 
    If Right(DimVariablesList, 2) = ", " Then DimVariablesList = Mid(DimVariablesList, 1, Len(DimVariablesList) - 2) 
    Call rw("名称列表 ========<br>" & NameList & "<br>") 
    Call rw("定义变量 ========<br>" & DimVariablesList & "<br>") 
    Call rw("获得Form ========<br>" & GetFormList & "<br>") 
    Call rw("ASP代码 ========<br>" & AspCodeList & "<br>") 
    Call rw("打印ASP ========<br>" & PrintFieldList & "<br>") 
    Call rw("显示内容 ========<br>" & ShowContent & "<br>") 
End Sub 


'DedeCms回显样式 20150113
Function dedeCMSMsg()
    Dim c 
    c = "<style> " & vbCrLf 
    c = c & ".msgbox {" & vbCrLf 
    c = c & "    width: 450px;" & vbCrLf 
    c = c & "    border: 1px solid #DADADA;" & vbCrLf 
    c = c & "    margin:0 auto;" & vbCrLf 
    c = c & "    margin-top:20px;" & vbCrLf 
    c = c & "    line-height:20px;" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & ".msgbox .ptitle{" & vbCrLf 
    c = c & "    padding: 6px;" & vbCrLf 
    c = c & "    font-size: 12px;" & vbCrLf 
    c = c & "    border-bottom: 1px solid #DADADA;" & vbCrLf 
    c = c & "    background: #DBEEBD url(/plus/img/wbg.gif);" & vbCrLf 
    c = c & "    font-weight:bold;" & vbCrLf 
    c = c & "    text-align:center;" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & ".msgbox .pcontent{" & vbCrLf 
    c = c & "    height: 100px;" & vbCrLf 
    c = c & "    font-size: 10pt;" & vbCrLf 
    c = c & "    background: #ffffff;" & vbCrLf 
    c = c & "    text-align:center;" & vbCrLf 
    c = c & "    padding-top:30px;" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "</style> " & vbCrLf 
    c = c & "<div class=""msgbox"">" & vbCrLf 
    c = c & "    <div class=""ptitle"">提示信息！</div>" & vbCrLf 
    c = c & "    <div class=""pcontent"">" & vbCrLf 
    c = c & "        成功登录，正在转向管理管理主页！<br>" & vbCrLf 
    c = c & "        <a href=""#"">如果你的浏览器没反应，请点击这里...</a>" & vbCrLf 
    c = c & "" & vbCrLf 
    c = c & "    </div>" & vbCrLf 
    c = c & "</div>" & vbCrLf 
    dedeCMSMsg = c 
End Function 
%>  

