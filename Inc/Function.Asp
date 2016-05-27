<% 
'发送邮箱 call Send_Email("313801120@qq.com", "标题", "111", "内容") 
Function Send_Email(ToMail, ToTitle, MyEmalName, ToContent)
    Dim JMail, isgo, MailBody 
    Response.AddHeader "Content-Type", "text/html; charset=gb2312" 
    Set JMail = CreateObject("JMail.Message")
        'JMail.ISOEncodeHeaders = False ' 是否进行ISO编码，默认为True
        JMail.ContentTransferEncoding = "base64" 
        JMail.Encoding = "base64" 
        JMail.ContentType = "text/html"    '正常显示内容的代码 
        JMail.silent = True 
        JMail.Logging = True 
        JMail.CharSet = "gb2312" 
        JMail.MailServerUserName = "m18251922007"    '此为您邮箱的登录帐号，使用时请更改为自己的邮箱登录帐号 
        JMail.MailServerPassword = "mydd3a"    '此为您邮箱的登录密码，使用时请更改为自己的邮箱登录密码 
        JMail.From = "m18251922007@163.com"    '"m18251922007@163.com" '发件人Email 
        JMail.FromName = MyEmalName    '发件人姓名 
        JMail.AddRecipient ToMail    '收件人Email 
        JMail.Subject = ToTitle    '邮件主题 
        '邮件主体（HTML(注意信件内链接附件的方式)）
        MailBody = MailBody & "<html><head><META content=zh-cn http-equiv=Content-Language><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312""><style type=text/css>BODY {FONT-SIZE: 9pt}</style></head><body>" 
        MailBody = MailBody & ToContent 
        MailBody = MailBody & "</body></html>" 
        JMail.Body = MailBody    '邮件正文 
        Send_Email = JMail.Send("smtp.163.com")    'SMTP服务器地址         //返回发送是否成功
        JMail.Close 
    Set JMail = Nothing 
End Function 
'邮箱发送 例：Response.Write( ServerSend_Email("m18251922007","mydd3a","313801120@qq.com", "标题", "11@aa.com", "内容"  ) )
'邮箱发送 
Function ServerSend_Email(ServerUserName, ServerPassword, ToMail, ToTitle, MyEmalName, ToContent)
    Dim JMail, isgo, MailBody 
    Response.AddHeader "Content-Type", "text/html; charset=gb2312" 
    Set JMail = CreateObject("JMail.Message")
        'JMail.ISOEncodeHeaders = False ' 是否进行ISO编码，默认为True
        JMail.ContentTransferEncoding = "base64" 
        JMail.Encoding = "base64" 
        JMail.ContentType = "text/html"    '正常显示内容的代码 
        JMail.silent = True 
        JMail.Logging = True 
        JMail.CharSet = "gb2312"
        JMail.MailServerUserName = ServerUserName    '此为您邮箱的登录帐号，使用时请更改为自己的邮箱登录帐号 
        JMail.MailServerPassword = ServerPassword    '此为您邮箱的登录密码，使用时请更改为自己的邮箱登录密码 
        JMail.From = "m18251922007@163.com"    '"m18251922007@163.com" '发件人Email 
        JMail.FromName = MyEmalName    '发件人姓名 
        JMail.AddRecipient ToMail    '收件人Email 
        JMail.Subject = ToTitle    '邮件主题 
        '邮件主体（HTML(注意信件内链接附件的方式)）
        MailBody = MailBody & "<html><head><META content=zh-cn http-equiv=Content-Language><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312""><style type=text/css>BODY {FONT-SIZE: 9pt}</style></head><body>" 
        MailBody = MailBody & ToContent 
        MailBody = MailBody & "</body></html>" 
        JMail.Body = MailBody    '邮件正文 
        ServerSend_Email = JMail.Send("smtp.163.com")    'SMTP服务器地址         //返回发送是否成功 
        JMail.Close 
    Set JMail = Nothing 
End Function 
'找ASP函数列表如：Sub a(b,c)
Function FindAspCode(ByVal Content) 
	Content = GetAspCode(Content)
    Content = GetASPFunctionList(Content)    '获得函数列表 
    FindAspCode =Content
End Function
'扫描ASP中函数名称列表
Function FindAspFunList(ByVal Content)
	Dim SplStr,nLen,SubStr,S,TempS,C
	Content = FindAspCode(Content)				'获得ASP代码
	SplStr=Split(Content,vbCrlf)
	For Each S In SplStr
		TempS = S : S = LCase(S)
		If S <>"" Then
			nLen=0
			If InStr(S,"function ")>0 Then
				nLen=InStr(S,"function ")+8
			ElseIf InStr(S,"sub ")>0 Then
				nLen=InStr(S,"sub ")+3
			End If
			If nLen > 0 Then
				TempS = Mid(TempS,nLen)
				If InStr(TempS,"(")>0 Then
					TempS=Mid(TempS,1,InStr(TempS,"(")-1)
				Else
					TempS=Mid(TempS,1,InStr(TempS & " "," ")-1)
				End If				
                C = C & Trim(TempS) & vbCrLf            '去除函数名称两边空格
			End If
		End If	
        DoEvents
	Next
	FindAspFunList=C 
End Function  
'找到指定Asp函数全部代码 感觉不精准
Function FindAspFunctionCode(ByVal Content, ByVal FunctionName)
    Dim FunctionType, SplStr, S, C 
    If CheckStr(FunctionName, "sub ") Then
        FunctionType = "end sub"
    Else
        FunctionType = "end function" 
    End If 
    Content = Right(Content, Len(Content) - InStr(LCase(Content), LCase(FunctionName)) + 1) 
    SplStr = Split(Content, vbCrlf) 
    For Each S In SplStr
        C = C & S & vbCrlf 
        If CheckStr(S, FunctionType) = True Then Exit For 
    Next 
    FindAspFunctionCode = C 
End Function 
'处理替换列表 (ReplaceList改成handleReplaceList)为了兼容php程序
Function handleReplaceList(ByVal Content, YStr, TStr)
    Dim YSplStr, TSplStr, I, S 
    YSplStr = Split(YStr, "|") 
    TSplStr = Split(TStr & "||||||||||||||||||||||||||||||||", "|") 
    For I = 0 To UBound(YSplStr)
        S = YSplStr(I) 
        If S <> "" Then
            Content = Replace(Content, S, TSplStr(I)) 
        End If 
    Next 
    handleReplaceList = Content 
End Function
'获得函数名称与变量 ，加强版,添加于2013,10,5 
Function GetFunNameDim(ByVal C, FunName, FunDim)
    Dim FunType, S1, S2, FunValue 
    C = Trim(C) 
    If InStr(LCase(C), "sub ") > 0 Then
        FunType = "sub" 
    ElseIf InStr(LCase(C), "function ") > 0 Then
        FunType = "function" 
    Else
        Exit Function 
    End If 
    C = Replace(C, ":", ")") & ")" 
    S1 = InStr(C, FunType & " ") + Len(FunType & " ") 
    S2 = InStr(S1, C, ")") 
    C = Mid(C, S1, S2 - S1) 
    FunValue = FunType & " " & C 
    If InStr(C, "(") > 0 Then
        FunName = Left(C, InStr(C, "(") - 1) 
        If InStr(FunName, " ") Then
            FunName = Mid(FunName, InStr(FunName, " "))    '可把Priver与Public去除掉 
        End If 
        FunValue = FunValue & ")" 
        S1 = InStr(FunValue, "(") + 1 
        S2 = InStr(S1, FunValue, ")") 
        FunDim = Mid(FunValue, S1, S2 - S1) 
        FunDim = RegExp_Replace(FunDim, "ByVal ", "")    'ByVal重新定义变更 
        FunDim = RegExp_Replace(FunDim, "ByRef ", "")    'ByRef为不重新定义变量,默认为它，也就量用上一级函数变量 
        FunDim = Replace(FunDim, " ", "") 
    Else
        FunName = C 
    End If 
    'Response.Write("FunType=" & FunType & "|C="&C&"|FunValue="&FunValue&"|FunName="&FunName&"|FunDim="&FunDim&"<hr>")
    GetFunNameDim = FunValue 
End Function 
'从右边开始截取 作用：删除左边空格
Function LeftSpace(ByVal Content, ByVal AddNumb)
    Dim I, nCount 
    nCount = 0 
    For I = Len(Content) To 1 Step - 1
        If Mid(Content, I, 1) <> " " Then Exit For
        nCount = nCount + 1 
    Next 
    If nCount > AddNumb And Replace(Content, " ", "") <> "" Then
        Content = Left(Content, Len(Content) - nCount + AddNumb) 
    ElseIf nCount < AddNumb And 1 = 1 Then
        For I = nCount To AddNumb - 1
            Content = Content & " " 
        Next 
    End If 
    LeftSpace = Content 
End Function 
'从右边开始截取 作用：删除右边空格
Function RightSpace(ByVal Content, ByVal AddNumb)
    Dim I, nCount 
    nCount = 0 
    For I = 1 To Len(Content)
        If Mid(Content, I, 1) <> " " Then Exit For 
        nCount = nCount + 1 
    Next 
    If nCount > AddNumb Then
        Content = Right(Content, Len(Content) - nCount + AddNumb) 
    ElseIf nCount < AddNumb Then
        For I = nCount To AddNumb - 1
            Content = " " & Content 
        Next 
    End If 
    RightSpace = Content 
End Function 
'在函数里加乱的Dim，这个Dim出现次数等于函数参数值，这个Dim是没有任何意思的 
Function InFunAddDim(ByVal Content)
    Dim SplStr, Splxx, SplTemp, S, I, VbCrlfNumb, FunDimNumb, ExceptNumb, C, RootNumb 
    Splxx = Split("A_p,B_p,C_4,D_2,E_4,F_9,G_1,H_2,I_2,J_3,K_1,L_o,M_m,N_u,O_o,P_n,Q_o,R_l,S_i,T_2,U_2,V_2,W_2,X_2,Y_2,Z_2", ",") 
    SplStr = Split(Content, vbCrlf) 
    VbCrlfNumb = UBound(SplStr) - 1    '换行数 
    S = StrCut(Content, "(", ")", 2) 
    SplTemp = Split(S, ",") 
    FunDimNumb = UBound(SplTemp) + 1    '函数 
    '为0表示没有则退出，
    If FunDimNumb = 0 Then
        InFunAddDim = Content 
        Exit Function 
    End If 
    ExceptNumb = CInt(VbCrlfNumb / FunDimNumb)    '除数 
    RootNumb = 0 
    For I = 0 To UBound(SplStr)
        S = SplStr(I) 
        If Trim(S) <> "" Then
            C = C & SplStr(I) & vbCrlf 
        End If 
        If ExceptNumb <> 0 Then
            If I Mod ExceptNumb = 0 And UBound(SplTemp) >= RootNumb And I <> UBound(SplStr) Then
                C = C & "dIm " & Splxx(RootNumb) & vbCrlf 
                RootNumb = RootNumb + 1 
            End If 
        End If 
    Next 
    If C <> "" Then C = Left(C, Len(C) - 2) 
    InFunAddDim = C 
End Function 
'=============== 下面为2013,11,1添加 
'获得QQ 
Function GetQQ(ByVal Content, nOK)
    Dim SplStr, I, J, S, C, QQ, nErr, QQYes 
    Content = Replace(Replace(LCase(Content), "&nbsp;", ""), " ", "") 
    nOK = 0 
    If InStr(Content, "qq") > 0 Then
        SplStr = Split(Content, "qq") 
        For I = 1 To UBound(SplStr)
            S = SplStr(I) 
            QQ = "" : nErr = 14 : QQYes = False 
            For J = 1 To Len(S)
                If InStr("0123456789", Mid(S, J, 1)) = False Then
                    If QQYes = True Then Exit For    'qq开始累加时则退出 
                    If nErr = 0 Then Exit For    '错误大于N退出 
                    nErr = nErr - 1 
                    If Mid(S, J, 1) = "群" Then Exit For    '为QQ群退出 
                Else
                    QQYes = True 
                    QQ = QQ & Mid(S, J, 1) 
                End If 
                If J > 30 Then Exit For    'j大于20退出 
            Next 
            If Len(QQ) >= 6 And Len(QQ) <= 10 And InStr(vbCrlf & C, vbCrlf & QQ & vbCrlf) = False Then
                If InStr(vbCrlf & C, vbCrlf & QQ & vbCrlf) = False Then
                    C = C & QQ & vbCrlf 
                    nOK = nOK + 1 
                End If 
            End If 
        Next 
    End If 
    GetQQ = C 
End Function 
'获得手机 
Function GetTel(ByVal Content, nOK)
    Dim SplStr, I, J, S, C, Tel, nErr, TelYes 
    Content = Replace(Replace(LCase(Content), "&nbsp;", ""), "手机", "tel") 
    nOK = 0 
    If InStr(Content, "tel") > 0 Then
        SplStr = Split(Content, "tel") 
        For I = 1 To UBound(SplStr)
            S = SplStr(I) 
            Tel = "" : nErr = 14 : TelYes = False 
            For J = 1 To Len(S)
                If InStr("0123456789", Mid(S, J, 1)) = False Then
                    If TelYes = True Then Exit For    'Tel开始累加时则退出 
                    If nErr = 0 Then Exit For    '错误大于N退出 
                    nErr = nErr - 1 
                Else
                    TelYes = True 
                    Tel = Tel & Mid(S, J, 1) 
                End If 
                If J > 30 Then Exit For    'j大于20退出 
            Next 
            If Len(Tel) = 11 Then
                If InStr(vbCrlf & C, vbCrlf & Tel & vbCrlf) = False Then
                    C = C & Tel & vbCrlf 
                    nOK = nOK + 1 
                End If 
            End If 
        Next 
    End If 
    GetTel = C 
End Function 
'获得邮箱 
Function GetMail(ByVal Content, nOK)
    Dim SplStr, I, J, S, C, Mail, nErr, MailYes 
    Content = Replace(Replace(LCase(Content), "&nbsp;", ""), "邮箱", "mail") 
    nOK = 0 
    If InStr(Content, "mail") > 0 Then
        SplStr = Split(Content, "mail") 
        For I = 1 To UBound(SplStr)
            S = SplStr(I) 
            Mail = "" : nErr = 14 : MailYes = False 
            For J = 1 To Len(S)
                If InStr("0123456789abcdefghijklmnopqrstuvwxyz.@", Mid(S, J, 1)) = False Then
                    If MailYes = True Then Exit For    'Mail开始累加时则退出 
                    If nErr = 0 Then Exit For    '错误大于N退出 
                    nErr = nErr - 1 
                Else
                    MailYes = True 
                    Mail = Mail & Mid(S, J, 1) 
                End If 
                If J > 30 Then Exit For    'j大于20退出 
            Next 
            If InStr(Mail, ".") > 0 And InStr(Mail, "@") > 0 Then
                If InStr(vbCrlf & C, vbCrlf & Mail & vbCrlf) = False Then
                    C = C & Mail & vbCrlf 
                    nOK = nOK + 1 
                End If 
            End If 
        Next 
    End If 
    GetMail = C 
End Function 
'获得图片列表 
Function GetImgStr(HttpUrl, Content, nOK)
    Dim SplStr, I, C, Url, UrlList 
    Content = GetIMG(Content) 
    SplStr = Split(Content, vbCrlf) 
    nOK = 0 
    For Each Url In SplStr
        If InStr("|" & UrlList & "|", "|" & Url & "|") = False Then
            If Left(Url, 1) = "/" Or(InStr(Url, "http://") = False And InStr(Url, "www.") = False) Then
                Url = FullHttpUrl(HttpUrl, Url) 
            End If 
            nOK = nOK + 1 
            C = C & Url & "<br>" 
        End If 
    Next 
    GetImgStr = C 
End Function 
'获得Css列表 
Function GetCssStr(HttpUrl, Content, nOK)
    Dim SplStr, I, C, Url, UrlList 
    Content = GetCssHref(Content) 
    SplStr = Split(Content, vbCrlf) 
    nOK = 0 
    For Each Url In SplStr
        If InStr("|" & UrlList & "|", "|" & Url & "|") = False Then
            If Left(Url, 1) = "/" Or(InStr(Url, "http://") = False And InStr(Url, "www.") = False) Then
                Url = FullHttpUrl(HttpUrl, Url) 
            End If 
            nOK = nOK + 1 
            C = C & Url & "<br>" 
        End If 
    Next 
    GetCssStr = C 
End Function 
'获得Js列表 
Function GetJsStr(HttpUrl, Content, nOK)
    Dim SplStr, I, C, Url, UrlList 
    Content = GetJsSrc(Content) 
    SplStr = Split(Content, vbCrlf) 
    nOK = 0 
    For Each Url In SplStr
        If InStr("|" & UrlList & "|", "|" & Url & "|") = False Then
            If Left(Url, 1) = "/" Or(InStr(Url, "http://") = False And InStr(Url, "www.") = False) Then
                Url = FullHttpUrl(HttpUrl, Url) 
            End If 
            nOK = nOK + 1 
            C = C & Url & "<br>" 
        End If 
    Next 
    GetJsStr = C 
End Function 

'处理成AspExecute
Function HandleAspExecute(ByVal Content)
	Dim SplStr,S,C 
	If Left(Content,2)="<" & "%" Then
		Content = Right(Content,Len(Content)-2) 
	End If	
	If Right(Content,2)="%" & ">" Then
		Content = Left(Content, Len(Content)-2) 
	End If 
	'Content=Replace(Content,"""","""""")				'这个不需要，要不会出错
	SplStr=Split(Content,vbCrlf)
	For Each S In SplStr
		S = TrimVbCrlf(S)
		If S <> "" Then
			C=C & S & ":"
		End If
	Next
	If C <>"" Then C=Left(C, Len(C)-1)
	If Left(C,1)="=" Then C = "Response.Write " & Right(C,Len(C)-1)
	HandleAspExecute = C
End Function
'获得网址列表 Rw(GetUrlStr("", "链接", Content, "", 0))
Function GetUrlStr(HttpUrl, SType, Content, SearchValue, nOK)
    Dim HrefList, SplStr, Title, Url, C, S, LcaseUrl, UrlList, TitleList, SVal, SVal2 
    If SType = "链接" Then
        HrefList = GetAURL(Content) 
    ElseIf SType = "链接标题" Then
        HrefList = GetATitle(Content) 
    Else
        HrefList = GetAURLTitle(Content) 
    End If 
    SplStr = Split(HrefList, vbCrlf) 
    nOK = 0 
    For Each S In SplStr
        If S <> "" Then
            If SType = "链接" Then
                Url = Trim(S) 
                If InStr(LCase(Url), "javascript:") = False And Url <> "#" And Url <> "" Then
                    If InStr("|" & UrlList & "|", "|" & Url & "|") = False Then
                        If Left(Url, 1) = "/" Or(InStr(Url, "http://") = False And InStr(Url, "www.") = False) Then
                            Url = FullHttpUrl(HttpUrl, Url) 
                        End If 
                        nOK = nOK + 1 
                        C = C & Url & vbCrlf 
                    End If 
                End If 
            ElseIf SType = "链接标题" Then
                If InStr("|" & TitleList & "|", "|" & S & "|") = False Then
                    TitleList = TitleList & S & "|" 
                    nOK = nOK + 1 
                    C = C & S & vbCrlf 
                End If 
            Else
                If InStr(S, "【_|-】") > 0 Then
                    Url = Mid(S, 1, InStr(S, "【_|-】") - 1) : LcaseUrl = Url 
                    Title = Mid(S, Len(Url) + 6) 
                    If InStr(LCase(LcaseUrl), "javascript:") = False And Url <> "#" And Url <> "" Then
                        If InStr("|" & UrlList & "|", "|" & Url & "|") = False Then
                            If Left(Url, 1) = "/" Or(InStr(LcaseUrl, "http://") = False And InStr(LcaseUrl, "www.") = False) Then
                                Url = FullHttpUrl(HttpUrl, Url) 
                            End If 
                            nOK = nOK + 1 
                            C = C & Url & "   &nbsp;  " & Title & vbCrlf 
                        End If 
                    End If 
                End If 
            End If 
        End If 
    Next 
    If SearchValue <> "" Then
        SearchValue = Replace(SearchValue, " and ", " And ") 
        SplStr = Split(SearchValue, " And ") 
        SVal = SplStr(0) 
        If UBound(SplStr) > 0 Then SVal2 = SplStr(1) 
        SplStr = Split(C, vbCrlf) : nOK = 0 : C = "" 
        For Each S In SplStr
            If SVal2 <> "" Then
                If InStr(S, SVal) > 0 And InStr(S, SVal2) > 0 Then
                    If InStr(vbCrlf & C & vbCrlf, vbCrlf & S & vbCrlf) = False Then
                        nOK = nOK + 1 
                        C = C & S & "<br>" 
                    End If 
                End If 
            Else
                If InStr(S, SVal) > 0 Then
                    If InStr(vbCrlf & C & vbCrlf, vbCrlf & S & vbCrlf) = False Then
                        nOK = nOK + 1 
                        C = C & S & "<br>" 
                    End If 
                End If 
            End If 
        Next 
    End If 
    GetUrlStr = C 
End Function 
'获得Html中翻页配置 (采集用到)
Function GetPageConfig(HttpUrl, Content, SType)
    Dim SplStr, I, S, S2, S3, C, Url, TempUrl, ArrUrl(99), UrlList, PageUrl, nLen, Splxx 
    Content = GetAURL(Content) 
    SplStr = Split(Content, vbCrlf) 
    For Each Url In SplStr
        If Url <> "" Then
            TempUrl = Url 
            If InStr(LCase(Url), "javascript:") = False And Url <> "#" Then
                Url = Replace(Replace(Replace(Replace(Replace(Url, "0", ""), "1", ""), "2", ""), "3", ""), "4", "") 
                Url = Replace(Replace(Replace(Replace(Replace(Url, "5", ""), "6", ""), "7", ""), "8", ""), "9", "") 
                If SType = "注入" And InStr(Url, "?") > 0 Then Url = HandlSqlInUrl(Url) 
                C = C & Url & vbCrlf 
                nLen = InStr(vbCrlf & UrlList, vbCrlf & Url & "【】") 
                If nLen > 0 Then
                    S = Mid(UrlList, nLen) 
                    S = Mid(S, 1, InStr(S, vbCrlf) - 1) 
                    Splxx = Split(S, "【】") 
                    S2 = Splxx(0) 
                    S3 = Splxx(1) + 1 
                    UrlList = Replace(UrlList, S, S2 & "【】" & S3 & "【】" & Splxx(2)) 
                    PageUrl = Url 
                Else
                    UrlList = UrlList & Url & "【】0【】" & FullHttpUrl(HttpUrl, TempUrl) & vbCrlf 
                End If 
            End If 
        End If 
    Next 
    SplStr = Split(UrlList, vbCrlf) : C = "" 
    If SType = "注入" Then
        For Each S In SplStr
            If InStr(S, "【】") > 0 Then
                Splxx = Split(S, "【】") 
                Url = Trim(Splxx(2)) 
                If InStr(Url, HttpUrl) > 0 Then
                    If InStr(vbCrlf & C, vbCrlf & Url & vbCrlf) = False Then
                        C = C & Url & vbCrlf 
                    End If 
                End If 
            End If 
        Next 
        GetPageConfig = C : Exit Function 
    End If 
    '处理出页配置
    For Each S In SplStr
        If InStr(S, "【】") > 0 Then
            Splxx = Split(S, "【】") 
            If Splxx(1) > 0 Then ArrUrl(Splxx(1)) = Splxx(0) & "  &nbsp; | &nbsp; " & Splxx(2) 
        End If 
    Next 
    For I = 99 To 0 Step - 1
        If ArrUrl(I) <> "" Then
            If InStr(ArrUrl(I), HttpUrl) > 0 Then
                C = C & ArrUrl(I) & "，   出现[" & I & "]次<br>" 
            End If 
        End If 
    Next 
    GetPageConfig = C 
End Function 
'获得网页中注入网址 
Function GetSqlInUrl(HttpUrl, Content, SType)
    Dim SplStr, I, S, S2, S3, C, Url, TempUrl, ArrUrl(99), UrlList, PageUrl, nLen, Splxx 
    Content = GetAURL(Content) 
    SplStr = Split(Content, vbCrlf) 
    For Each Url In SplStr
        If Url <> "" Then
            TempUrl = Url 
            If InStr(Url, "?") > 0 Then
                C = C & Url & vbCrlf 
                Url = HandlSqlInUrl(Url) 
                nLen = InStr(vbCrlf & UrlList, vbCrlf & Url & "【】") 
                If nLen > 0 Then
                    S = Mid(UrlList, nLen) 
                    S = Mid(S, 1, InStr(S, vbCrlf) - 1) 
                    Splxx = Split(S, "【】") 
                    S2 = Splxx(0) 
                    S3 = Splxx(1) + 1 
                    UrlList = Replace(UrlList, S, S2 & "【】" & S3 & "【】" & Splxx(2)) 
                    PageUrl = Url 
                Else
                    UrlList = UrlList & Url & "【】0【】" & FullHttpUrl(HttpUrl, TempUrl) & vbCrlf 
                End If 
            End If 
        End If 
    Next 
    SplStr = Split(UrlList, vbCrlf) 
    For Each S In SplStr
        If InStr(S, "【】") > 0 Then
            Splxx = Split(S, "【】") 
            If S3 > 0 Then
                If SType = "注入" Then
                    ArrUrl(Splxx(1)) = Splxx(2) 
                Else
                    ArrUrl(Splxx(1)) = Splxx(0) & "  &nbsp; | &nbsp; " & Splxx(2) 
                End If 
            End If 
        End If 
    Next 
    C = "" 
    For I = 99 To 0 Step - 1
        If ArrUrl(I) <> "" Then
            If SType = "注入" Then
                C = C & ArrUrl(I) & vbCrlf 
            Else
                C = C & ArrUrl(I) & "，   出现[" & I & "]次<br>" 
            End If 
        End If 
    Next 
    GetSqlInUrl = C 
End Function 
'处理注入网址，配置获得网站注入网址 
Function HandlSqlInUrl(HttpUrl)
    Dim Url, SplStr, I, S 
    SplStr = Split(HttpUrl, "=") 
    For I = 0 To UBound(SplStr)
        If I = UBound(SplStr) Then Url = Url & "=" : Exit For 
        S = SplStr(I) 
        If I Mod 2 = 0 Then
            Url = Url & SplStr(I) 
        Else
            If InStr(S, "&") > 0 Then
                Url = Url & "=" & Mid(S, InStr(S, "&")) 
            End If 
        End If 
    Next 
    HandlSqlInUrl = Url 
End Function

'获得扫描后函数名称列表
Function GetScanFunctionNameList(Content)
	On Error Resume Next
    Dim SplStr, I, YesASP, YesWord, Sx, S, Wc, Zc, S1, AspCode, SYHCount, UpWord,FunList,S2
    Dim UpWordn ,TempS,DimList
	Dim YesFunction								'函数是否为真 
    YesASP = False    '是ASP 默认为假 
    YesFunction = False    '是函数 默认为假 
    YesWord = False    '是单词 默认为假 
    SYHCount = 0    '双引号默认为0 
    SplStr = Split(Content, vbCrlf)    '分割行 
    '循环分行
    For Each S In SplStr
        '循环每个字符
        For I = 1 To Len(S)
            Sx = Mid(S, I, 1) 
            'Asp开始
            If Sx = "<" And Wc = "" Then    '输出文本必需为空 Wc为输出内容 如"<%" 排除 修改于20140412
                If Mid(S, I + 1, 1) = "%" Then
                     YesASP = True    'ASP为真 
                     I = I + 1    '加1而不能加2，要不然<%function Test() 就截取不到 
                End If 
            'ASP结束
            ElseIf Sx = "%" And Mid(S, I + 1, 1) = ">" And Wc = "" Then				'Wc为输出内容
                YesASP = False    'ASP为假 
                I = I + 1    '不能加2，只能加1，因为这里定义ASP为假，它会在下一次显示上面的 'ASP运行为假 
            End If 
            If YesASP = True Then
                '输入文本
                If(Sx = """" Or Wc <> "") Then
                    '双引号累加
                    If Sx = """" Then SYHCount = SYHCount + 1
                    '判断是否"在最后
                    If SYHCount Mod 2 = 0 Then
                        If Mid(S, I + 1, 1) <> """" Then
                            Wc = Wc & Sx 
                            S1 = Right(Replace(Mid(S, 1, I - Len(Wc)), " ", ""), 1)    '必需放在这里，要不会出错 
                            If YesFunction = True Then AspCode = AspCode & Wc 			'函数代码累加
                            SYHCount = 0 : Wc = ""    '清除 
                        Else
                            Wc = Wc & Sx 
                        End If 
                    Else
                        Wc = Wc & Sx
                    End If
                ElseIf Sx = "'" Then    '注释则退出
                    If YesFunction = True Then AspCode = AspCode & Mid(S, I)
                    Exit For
                '字母 
				ElseIf CheckABC(Sx) = True Or (Sx = "_" And Zc<>"") Or Zc <> "" Then
                    Zc = Zc & Sx 
                    S1 = LCase(Mid(S & " ", I + 1, 1)) 
                    If InStr("abcdefghijklmnopqrstuvwxyz0123456789", S1) = 0 And Not(S1 = "_" And Zc <> "") Then    '最简单判断	
						TempS = Mid(S,I+1)
                        If InStr("|function|sub|", "|" & LCase(Zc) & "|") Then							
							'函数开始
                            If YesFunction = False And LCase(UpWord) <> "end" Then
                                YesFunction = True  
								S2 = Mid(S,I+2)
								S2 = Mid(S2,1,InStr(S2,"(")-1)
								FunList = FunList & S2 & vbCrlf
                            ElseIf YesFunction = True And LCase(UpWord) = "end" Then    '获得上一个单词
                                AspCode = AspCode & Zc & vbCrlf 
								
                                YesFunction = False 
                            End If
						 
                        End If 
                        UpWord = Zc    '记住当前单词
                        If YesFunction = True Then AspCode = AspCode & Zc
                        Zc = ""  
                    End If 
                
                End If 
            End If	
        	DoEvents 
        Next		
        
    Next  
	GetScanFunctionNameList = FunList
End Function 
'获得函数名称20150402
Function GetFunName(ByVal C)
	If InStr(C,"(")>0 Then
		C=Mid(C,1,InStr(C,"(")-1)
		C=PHPTrim(C)
	End If
	GetFunName = C
End Function 
'获得函数名称中变量
Function GetFunDimName(ByVal C)
	Dim StartStr,EndStr,S
	C = LCase(C)
	StartStr = "("
	EndStr = ")"
 	If InStr(C, StartStr)>0 And InStr(C, EndStr)>0 Then
		C = StrCut(C,StartStr,EndStr,2)
	End If
	If C <> "" Then
		C=Replace(Replace(C,"byref ",""),"byref,","")
		C=Replace(Replace(C,"byval ",""),"byval,","")
		C=Replace(C," ","")
	End If
	GetFunDimName = C
End Function 
'获得变量定义的名称
Function GetDimName(ByVal C)
	Dim StartStr,EndStr,S
	C = LCase(C)
	StartStr = ":"
 	If InStr(C, StartStr)>0 Then
		C = Mid(C,1,InStr(C,":")-1)
	End If
	If C <> "" Then
		C=Replace(Replace(C,"byref ",""),"byref,","")
		C=Replace(Replace(C,"byval ",""),"byval,","")
		C=Replace(C," ","")
	End If
	If InStr(C,"'")>0 Then
		C= PHPTrim(Mid(C,1,InStr(C,"'")-1))				'刪除里面的'注释部分	20150330
	End If
	GetDimName = C
End Function
'获得Set变量定义的名称 暂时用不上
Function GetSetName(ByVal C)
	C = PHPTrim(LCase(C))
 	C=Mid(C, 1,InStr(C,"=")-1)
	GetSetName=Trim(C)
End Function
'替换变量
Function ReplaceDim(ByVal Content)
	Dim SplStr, S, TempS, C, LCaseS, DimInTHNumb 
	SplStr = Split(Content, ",") 
	For Each S In SplStr
		S = Trim(S) 
		LCaseS = LCase(S) 
		If S <> "" Then
			'对变量中()处理，
			If InStr(S, "(") > 0 Then
				S = Mid(S, 1, InStr(S, "(") - 1) 
			End If 
			DimInTHNumb = InStr(LCase(ReplaceDimList), "," & LCaseS & "=")    '替换变量 
			If DimInTHNumb > 0 Then    '替换变量
				TempS = Mid(ReplaceDimList, DimInTHNumb + 1) 
				TempS = Mid(TempS, 1, InStr(TempS, ",") - 1) 
				TempS = Mid(TempS, InStr(TempS, "=") + 1) 
				If InStr(FunDim & RootFunDim, "," & TempS & ",") = False Then
					S = TempS 
				End If 
			End If 
			C = C & S & "," 
		End If 
	Next 
	If C <> "" Then C = Left(C, Len(C) - 1) 
	ReplaceDim = C 
End Function 
'替换变量名称(最优化处理) 加强于20141017
Function ReplaceDim2(ByVal DimList,DimName)
	Dim ZD,SplStr,I,S,nMod,nInt,C
	ReplaceDim2  = DimName
	ZD="abcdefghijklmnopqrstuvwxyz"
	DimName=LCase(DimName)
	SplStr=Split(DimList,",")
	For I=0 To Ubound(SplStr)
		S = SplStr(I)
		If S = DimName Then
			nMod = (I) Mod Len(ZD)+1
			nInt = Fix((I)/Len(ZD))
			
			If nMod<>0 Then
				C=C & Mid(ZD,nMod,1)
			End If
			If nInt<>0 Then
				C=C & CopyStr(Mid(ZD,nInt,1),nInt)
			End If
			'Call Echo(I,Len(ZD))
			'Call Echo("nMod",nMod)
			'Call Echo("nInt",nInt)
			'Call Echo("C",C)
			ReplaceDim2 = C
			Exit Function
		End If
	Next 
End Function
'找当前文件夹重复函数主程序
Function FindFolderRepeatFunction(FolderPath)
    Dim FilePath, S, C, Content, Funs, FunList, AllFunList, nOK, nErr, SplStr, Splxx, nAllOK, nAllErr, nI 
	Dim ErrFunList,AllErrFunList
    Call HandlePath(FolderPath)                                                     '获得完整路径
    C = "操作文件夹" & FolderPath & vbCrLf 
    Content = GetDirFileList(FolderPath,"")
    SplStr = Split(Content, vbCrLf) : nI = 0 
    For Each FilePath In SplStr
        nI = nI + 1 
        S = nI & "、" & FilePath 
    	Content = GetFText(FilePath)
	    Content = GetScanFunctionNameList(Content)				'获得ASP函数名称列表
        nOK = 0 : nErr = 0 : nAllOK = 0 : nAllErr = 0 : FunList = "" : ErrFunList = "" : AllErrFunList=""
        Splxx = Split(Content, vbCrLf) 
        For Each Funs In Splxx
            If Funs <> "" Then
                If InStr("|" & FunList & "|", "|" & Funs & "|") = 0 Then
                    FunList = FunList & Funs & "|" 
                    nOK = nOK + 1 
                Else
					ErrFunList=ErrFunList & Funs & "|" 
                    nErr = nErr + 1 
                End If 
                If InStr("|" & AllFunList & "|", "|" & Funs & "|") = 0 Then
                    AllFunList = AllFunList & Funs & "|" '全部函数
                    nAllOK = nAllOK + 1 
                Else
					AllErrFunList=AllErrFunList & Funs & "|" 
                    nAllErr = nAllErr + 1 
                End If 
            End If 
            DoEvents 
        Next 
		'Call CreateFile("allfun.txt", AllFunList)
        C = C & S & "，函数（" & UBound(SplStr) + 1 & "），重复(" & nErr & "["& ErrFunList &"])全部函数重复(" & nAllErr & "["& AllErrFunList &"])" & vbCrLf
        DoEvents 
    Next 
    FindFolderRepeatFunction = C 
End Function 
'找当前文件重复函数 (辅助)
Function FindFileRepeatFunction(FilePath)
    FindFileRepeatFunction = HandleContentRepeatFunction(GetFText("FilePath"), 2)
End Function
'找当前内容重复函数
Function FindContentRepeatFunction(Content)
    FindContentRepeatFunction = HandleContentRepeatFunction(Contnet, 2)
End Function
'处理内容重复函数列表 sType为0为不显示，1为显示函数列表，2为显示重复函数列表，3为显示函数与重复函数列表
Function HandleContentRepeatFunction(Content, sType)
    Dim C, Funs, FunList, nOK, nErr, Splxx, ErrFunList 
    sType = CStr(sType)           
    Content = GetScanFunctionNameList(Content)                                      '获得ASP函数名称列表
    nOK = 0 : nErr = 0 
    Splxx = Split(Content, vbCrLf) 
    For Each Funs In Splxx
        If Funs <> "" Then
            If InStr("|" & FunList & "|", "|" & Funs & "|") = 0 Then
                FunList = FunList & Funs & "|" 
                nOK = nOK + 1 
            Else
                ErrFunList = ErrFunList & Funs & vbCrLf 
                nErr = nErr + 1 
            End If 
        End If 
        DoEvents 
    Next 
    C = "找到函数共（" & UBound(Splxx) + 1 & "），重复(" & nErr & ")" & vbCrLf 
	'函数列表
	If sType = "1" Or SType="3" Then
		 C = C & vbCrLf & "函数列表" & vbCrLf & FunList 
	End If
	'重复函数列表
    If ErrFunList <> "" And (sType = "1" Or SType="3") Then
        C = C & vbCrLf & "重复函数列表" & vbCrLf & ErrFunList 
    End If 
    HandleContentRepeatFunction = C 
End Function 
'替换字符内容2 自己写得一个 区分大小写
Function Replace2(Content,SearchStr,ReplaceStr)
	Dim LeftStr,RightStr
	If InStr(Content,SearchStr) > 0 Then
		LeftStr = Mid(Content, 1, InStr(Content,SearchStr)-1)
		RightStr = Mid(Content, Len(LeftStr)+Len(SearchStr)+1)
		Content = LeftStr & ReplaceStr & RightStr
	End If
	Replace2 = Content
End Function 
'替换全部字符内容2 自己写得一个 区分大小写
Function AllReplace(Content,SearchStr,ReplaceStr)
	Dim LeftStr,RightStr,I
	For I = 1 To 99
		If InStr(Content,SearchStr) > 0 Then
			LeftStr = Mid(Content, 1, InStr(Content,SearchStr)-1)
			RightStr = Mid(Content, Len(LeftStr)+Len(SearchStr)+1)
			Content = LeftStr & ReplaceStr & RightStr
		Else
			Exit For
		End If
	Next
	AllReplace = Content
End Function 
'替换一次，不区分大小写
Function ReplaceOneNOLU(Content,SearchStr,ReplaceStr)
	Dim LeftStr,RightStr,LCaseContent
	SearchStr=LCase(SearchStr)
	LCaseContent=LCase(Content)
	If InStr(LCaseContent,SearchStr) > 0 Then
		LeftStr = Mid(Content, 1, InStr(LCaseContent,SearchStr)-1)
		RightStr = Mid(Content, Len(LeftStr)+Len(SearchStr)+1)
		Content = LeftStr & ReplaceStr & RightStr
	End If
	ReplaceOneNOLU = Content
End Function 
'优化ASP代码 删除左右空格
Function OptimizeAspCode(Content)
    Dim SplStr,S,C,I
	SplStr=Split(Content,vbCrlf)
	For Each S In SplStr
		S = TrimVbTab(S)
		If S <> "" Then C=C & TrimVbTab(S) & vbCrlf
	Next    
    OptimizeAspCode = C
End Function
'Asp代码混淆处理
Function AspCodeConfusion(Content)
    Dim SplStr, I, YesASP, YesWord, Sx, S, Wc, Zc, S1, AspCode, SYHCount, UpWord
    Dim UpWordn ,TempS,DimList
	Dim YesFunction								'函数是否为真
	Dim StartFunction							'开始函数 目的是为了让function default 处理函数后面没有()   20150218
	StartFunction = False						'默认开始函数为假
	'If nType="" Then  nType  = 0
    YesASP = False    '是ASP 默认为假 
    YesFunction = False    '是函数 默认为假 
    YesWord = False    '是单词 默认为假 
    SYHCount = 0    '双引号默认为0 
    SplStr = Split(Content, vbCrlf)    '分割行 
    '循环分行
    For Each S In SplStr
        '循环每个字符
        For I = 1 To Len(S)
            Sx = Mid(S, I, 1) 
            'Asp开始
            If Sx = "<" And Wc = "" Then    '输出文本必需为空 Wc为输出内容 如"<%" 排除 修改于20140412
                If Mid(S, I + 1, 1) = "%" Then
                     YesASP = True    'ASP为真 
                     I = I + 1    '加1而不能加2，要不然<%function Test() 就截取不到 
					 Sx = Mid(S, I, 1) 
					 AspCode = AspCode & "<"
                End If 
            'ASP结束
            ElseIf Sx = "%" And Mid(S, I + 1, 1) = ">" And Wc = "" Then				'Wc为输出内容
                YesASP = False    'ASP为假 
                I = I + 1    '不能加2，只能加1，因为这里定义ASP为假，它会在下一次显示上面的 'ASP运行为假 
				Sx = Mid(S, I, 1) 
				AspCode = AspCode & "%" 
            End If 
            If YesASP = True Then
                '输入文本
                If(Sx = """" Or Wc <> "") Then
                    '双引号累加
                    If Sx = """" Then SYHCount = SYHCount + 1
                    '判断是否"在最后
                    If SYHCount Mod 2 = 0 Then
                        If Mid(S, I + 1, 1) <> """" Then
                            Wc = Wc & Sx 
                            S1 = Right(Replace(Mid(S, 1, I - Len(Wc)), " ", ""), 1)    '必需放在这里，要不会出错 
                            AspCode = AspCode & Wc 			'函数代码累加
                            SYHCount = 0 : Wc = ""    '清除 
                        Else
                            Wc = Wc & Sx 
                        End If 
                    Else
                        Wc = Wc & Sx
                    End If
                ElseIf Sx = "'" Then    '注释则退出
                    AspCode = AspCode & Mid(S, I)
                    Exit For 
                '字母 
				ElseIf CheckABC(Sx) = True Or (Sx = "_" And Zc<>"") Or Zc <> "" Then
                    Zc = Zc & Sx 
                    S1 = LCase(Mid(S & " ", I + 1, 1)) 
                    If InStr("abcdefghijklmnopqrstuvwxyz0123456789", S1) = 0 And Not(S1 = "_" And Zc <> "") Then    '最简单判断	
						TempS = Mid(S,I+1)
                        If InStr("|function|sub|", "|" & LCase(Zc) & "|") Then					
							'函数开始
                            If YesFunction = False And LCase(UpWord) <> "end" Then
                                YesFunction = True
								DimList = GetFunDimName(TempS)
								StartFunction=True
                            ElseIf YesFunction = True And LCase(UpWord) = "end" Then    '获得上一个单词                                
                                YesFunction = False
                            End If
						ElseIf YesFunction = True And LCase(Zc) = "dim" Then
							DimList = DimList & "," & GetDimName(TempS)
						ElseIf YesFunction = True  Then
							'排除函数后面每一个名称
							If StartFunction=False Then 
								Zc = ReplaceDim2(DimList, Zc)
							End If
							StartFunction=False
                        End If
                        UpWord = Zc    '记住当前单词
                        AspCode = AspCode & Zc
                        Zc = ""  
                    End If 
				Else
					AspCode = AspCode & Sx 
                End If 
			Else
				AspCode = AspCode & Sx 
            End If
        	DoEvents 
        Next		
        AspCode = RTrim(AspCode)            '去除右边空格
        AspCode = AspCode & vbCrlf    'Asp换行  	
        DoEvents
    Next 
    AspCodeConfusion = AspCode
End Function
'Asp变量乱码 只对大小写乱
Function ASPVariableGarbled(Content)
    Dim SplStr, I, YesASP, YesWord, Sx, S, Wc, Zc, S1, AspCode, SYHCount, UpWord
    Dim UpWordn ,TempS,DimList
	Dim YesFunction								'函数是否为真
	'If nType="" Then  nType  = 0
    YesASP = False    '是ASP 默认为假 
    YesFunction = False    '是函数 默认为假 
    YesWord = False    '是单词 默认为假 
    SYHCount = 0    '双引号默认为0 
    SplStr = Split(Content, vbCrlf)    '分割行 
    '循环分行
    For Each S In SplStr
        '循环每个字符
        For I = 1 To Len(S)
            Sx = Mid(S, I, 1) 
            'Asp开始
            If Sx = "<" And Wc = "" Then    '输出文本必需为空 Wc为输出内容 如"<%" 排除 修改于20140412
                If Mid(S, I + 1, 1) = "%" Then
                     YesASP = True    'ASP为真 
                     I = I + 1    '加1而不能加2，要不然<%function Test() 就截取不到 
					 Sx = Mid(S, I, 1) 
					 AspCode = AspCode & "<"
                End If 
            'ASP结束
            ElseIf Sx = "%" And Mid(S, I + 1, 1) = ">" And Wc = "" Then				'Wc为输出内容
                YesASP = False    'ASP为假 
                I = I + 1    '不能加2，只能加1，因为这里定义ASP为假，它会在下一次显示上面的 'ASP运行为假 
				Sx = Mid(S, I, 1) 
				AspCode = AspCode & "%" 
            End If 
            If YesASP = True Then
                '输入文本
                If(Sx = """" Or Wc <> "") Then
                    '双引号累加
                    If Sx = """" Then SYHCount = SYHCount + 1
                    '判断是否"在最后
                    If SYHCount Mod 2 = 0 Then
                        If Mid(S, I + 1, 1) <> """" Then
                            Wc = Wc & Sx 
                            S1 = Right(Replace(Mid(S, 1, I - Len(Wc)), " ", ""), 1)    '必需放在这里，要不会出错 
                            AspCode = AspCode & Wc 			'函数代码累加
                            SYHCount = 0 : Wc = ""    '清除 
                        Else
                            Wc = Wc & Sx 
                        End If 
                    Else
                        Wc = Wc & Sx
                    End If
                ElseIf Sx = "'" Then    '注释则退出
                    AspCode = AspCode & Mid(S, I)
                    Exit For 
                '字母 
				ElseIf CheckABC(Sx) = True Or (Sx = "_" And Zc<>"") Or Zc <> "" Then
                    Zc = Zc & Sx 
                    S1 = LCase(Mid(S & " ", I + 1, 1)) 
                    If InStr("abcdefghijklmnopqrstuvwxyz0123456789", S1) = 0 And Not(S1 = "_" And Zc <> "") Then    '最简单判断	
						TempS = Mid(S,I+1)
                        
                        AspCode = AspCode & LUCaseGarbled(Zc)
                        Zc = ""  
                    End If 
				Else
					AspCode = AspCode & Sx 
                End If 
			Else
				AspCode = AspCode & Sx 
            End If
        	DoEvents 
        Next		
        AspCode = RTrim(AspCode)            '去除右边空格
        AspCode = AspCode & vbCrlf    'Asp换行  	
        DoEvents
    Next 
    ASPVariableGarbled = AspCode
End Function
' 让大小写乱掉
Function LUCaseGarbled(Content)
    Dim LenStr, SplStr, I, S, C, Numb
    Numb = 0
    For I = 1 To Len(Content)
        S = Mid(Content, I, 1) 
        Numb = Numb + 1
        If Numb = 1 Then
            S = LCase(S)
        ElseIf Numb = 2 Or Numb = 3 Then
            S = UCase(S)
        ElseIf Numb = 4 Then
            S = LCase(S): Numb = 1
        End If
        C = C & S
    Next
    LUCaseGarbled = C
End Function

'Asp函数名称混淆处理20160120   call rw(AspFunctionConfusion(getftext("5.asp")))
Function AspFunctionConfusion(Content)
    Dim SplStr, splxx,I,j, YesASP, YesWord, Sx, S, Wc, Zc, S1, AspCode, SYHCount, UpWord,nNumb
    Dim UpWordn ,TempS,funList
	Dim YesFunction								'函数是否为真
	Dim StartFunction							'开始函数 目的是为了让function default 处理函数后面没有()   20150218
	StartFunction = False						'默认开始函数为假
	dim nFunIndex								'函数索引
	nFunIndex=0
	dim splNewFun
	'注意as是不能用的
	splNewFun=split("aa|ab|ac|ad|ae|af|ag|ah|ai|aj|ak|al|am|an|ao|ap|aq|ar|at|au|av|aw|ax|ay|az|ba|bb|bc|bd|be|bf|bg|bh|bi|bj|bk|bl|bm|bn|bo|bp|bq|br|bs|bt|bu|bv|bw|bx|by|bz|ca|cb|cc|cd|ce|cf|cg|ch|ci|cj|ck|cl|cm|cn|co|cp|cq|cr|cs|ct|cu|cv","|")
	dim splYunFun,yunFunList
	yunFunList= "|"& lcase(getAspFunNameList(Content)) &"|"
	splYunFun=split(yunFunList,"|")
 
	'If nType="" Then  nType  = 0
    YesASP = False    '是ASP 默认为假 
    YesFunction = False    '是函数 默认为假 
    YesWord = False    '是单词 默认为假 
    SYHCount = 0    '双引号默认为0 
	funlist="|"			'函数名称列表
    SplStr = Split(Content, vbCrlf)    '分割行
    '循环分行
    For Each S In SplStr
        '循环每个字符
        For I = 1 To Len(S)
            Sx = Mid(S, I, 1) 
            'Asp开始
            If Sx = "<" And Wc = "" Then    '输出文本必需为空 Wc为输出内容 如"<%" 排除 修改于20140412
                If Mid(S, I + 1, 1) = "%" Then
                     YesASP = True    'ASP为真 
                     I = I + 1    '加1而不能加2，要不然<%function Test() 就截取不到 
					 Sx = Mid(S, I, 1) 
					 AspCode = AspCode & "<"
                End If 
            'ASP结束
            ElseIf Sx = "%" And Mid(S, I + 1, 1) = ">" And Wc = "" Then				'Wc为输出内容
                YesASP = False    'ASP为假 
                I = I + 1    '不能加2，只能加1，因为这里定义ASP为假，它会在下一次显示上面的 'ASP运行为假 
				Sx = Mid(S, I, 1) 
				AspCode = AspCode & "%" 
            End If 
            If YesASP = True Then
                '输入文本
                If(Sx = """" Or Wc <> "") Then
                    '双引号累加
                    If Sx = """" Then SYHCount = SYHCount + 1
                    '判断是否"在最后
                    If SYHCount Mod 2 = 0 Then
                        If Mid(S, I + 1, 1) <> """" Then
                            Wc = Wc & Sx 
                            S1 = Right(Replace(Mid(S, 1, I - Len(Wc)), " ", ""), 1)    '必需放在这里，要不会出错 
                            AspCode = AspCode & Wc 			'函数代码累加
                            SYHCount = 0 : Wc = ""    '清除 
                        Else
                            Wc = Wc & Sx 
                        End If 
                    Else
                        Wc = Wc & Sx
                    End If
                ElseIf Sx = "'" Then    '注释则退出
                    AspCode = AspCode & Mid(S, I)
                    Exit For 
                '字母 
				ElseIf CheckABC(Sx) = True Or (Sx = "_" And Zc<>"") Or Zc <> "" Then
                    Zc = Zc & Sx 
                    S1 = LCase(Mid(S & " ", I + 1, 1)) 
                    If InStr("abcdefghijklmnopqrstuvwxyz0123456789", S1) = 0 And Not(S1 = "_" And Zc <> "") Then    '最简单判断	
						TempS = Mid(S,I+1)
                        If InStr("|function|sub|", "|" & LCase(Zc) & "|") Then					
							'函数开始
                            If YesFunction = False And LCase(UpWord) <> "end" and LCase(UpWord) <> "exit" Then
                               funList=funList & lcase(GetFunName(TempS)) & "," & splNewFun(nFunIndex)  & "|"
							   nFunIndex=nFunIndex+1
							   'call echo("funList",funList)
								StartFunction=True
                            ElseIf YesFunction = True And LCase(UpWord) = "end" Then    '获得上一个单词    
								StartFunction=False
                            End If
						ElseIf YesFunction = True And LCase(Zc) = "dim" Then
							 
					 
                        End If
						
						nNumb=instr(yunFunList ,"|" & lcase(Zc) & "|")
						if nNumb>0 then
							for j=1 to ubound(splYunFun)
								if splYunFun(j)=lcase(zc) then
									'call echo(zc,splNewFun(j))
									zc=splNewFun(j)
									exit for
								end if
							next
						end if
						
                        UpWord = Zc    '记住当前单词
                        AspCode = AspCode & Zc
                        Zc = ""  
                    End If 
				Else
					AspCode = AspCode & Sx 
                End If 
			Else
				AspCode = AspCode & Sx 
            End If
        	DoEvents 
        Next		
        AspCode = RTrim(AspCode)            '去除右边空格
        AspCode = AspCode & vbCrlf    'Asp换行  	
        DoEvents
    Next 
    AspFunctionConfusion = AspCode
End Function

'获得ASP函数名称列表 20160120 如  aa|bb|cc
Function getAspFunNameList(Content)
    Dim SplStr, splxx,I, YesASP, YesWord, Sx, S, Wc, Zc, S1, AspCode, SYHCount, UpWord,nNumb
    Dim UpWordn ,TempS,funList
	Dim YesFunction								'函数是否为真
	Dim StartFunction							'开始函数 目的是为了让function default 处理函数后面没有()   20150218
	StartFunction = False						'默认开始函数为假
	'If nType="" Then  nType  = 0
    YesASP = False    '是ASP 默认为假 
    YesFunction = False    '是函数 默认为假 
    YesWord = False    '是单词 默认为假 
    SYHCount = 0    '双引号默认为0 
	 
    SplStr = Split(Content, vbCrlf)    '分割行 
    '循环分行
    For Each S In SplStr
        '循环每个字符
        For I = 1 To Len(S)
            Sx = Mid(S, I, 1) 
            'Asp开始
            If Sx = "<" And Wc = "" Then    '输出文本必需为空 Wc为输出内容 如"<%" 排除 修改于20140412
                If Mid(S, I + 1, 1) = "%" Then
                     YesASP = True    'ASP为真 
                     I = I + 1    '加1而不能加2，要不然<%function Test() 就截取不到 
					 Sx = Mid(S, I, 1) 
					 AspCode = AspCode & "<"
                End If 
            'ASP结束
            ElseIf Sx = "%" And Mid(S, I + 1, 1) = ">" And Wc = "" Then				'Wc为输出内容
                YesASP = False    'ASP为假 
                I = I + 1    '不能加2，只能加1，因为这里定义ASP为假，它会在下一次显示上面的 'ASP运行为假 
				Sx = Mid(S, I, 1) 
				AspCode = AspCode & "%" 
            End If 
            If YesASP = True Then
                '输入文本
                If(Sx = """" Or Wc <> "") Then
                    '双引号累加
                    If Sx = """" Then SYHCount = SYHCount + 1
                    '判断是否"在最后
                    If SYHCount Mod 2 = 0 Then
                        If Mid(S, I + 1, 1) <> """" Then
                            Wc = Wc & Sx 
                            S1 = Right(Replace(Mid(S, 1, I - Len(Wc)), " ", ""), 1)    '必需放在这里，要不会出错 
                            AspCode = AspCode & Wc 			'函数代码累加
                            SYHCount = 0 : Wc = ""    '清除 
                        Else
                            Wc = Wc & Sx 
                        End If 
                    Else
                        Wc = Wc & Sx
                    End If
                ElseIf Sx = "'" Then    '注释则退出
                    AspCode = AspCode & Mid(S, I)
                    Exit For 
                '字母 
				ElseIf CheckABC(Sx) = True Or (Sx = "_" And Zc<>"") Or Zc <> "" Then
                    Zc = Zc & Sx 
                    S1 = LCase(Mid(S & " ", I + 1, 1)) 
                    If InStr("abcdefghijklmnopqrstuvwxyz0123456789", S1) = 0 And Not(S1 = "_" And Zc <> "") Then    '最简单判断	
						TempS = Mid(S,I+1)
                        If InStr("|function|sub|", "|" & LCase(Zc) & "|") Then					
							'函数开始
                            If YesFunction = False And LCase(UpWord) <> "end" and LCase(UpWord) <> "exit" Then
                               funList=funList & GetFunName(TempS) & "|"
							    
                            End If					 
                        End If 
                        UpWord = Zc    '记住当前单词 
						zc="" 
                    End If  
                End If  
            End If
        	DoEvents 
        Next		 
    Next 
    getAspFunNameList = funList
End Function
'ASP安全处理
Function AspCodeSafe(Content) 
	'这是处理后会出错，因为它调用了多个文件了
	if instr(content,"'【Asp函数" & "名称混淆】")>0 then			'不分开的话会处理此页
		Content = AspFunctionConfusion(Content )	'Asp函数名称混淆处理 20160120 
	end if
	Content = DelAspNote(Content)			'删除ASP里注释
	Content = OptimizeAspCode(Content)		'优化ASP代码	
	Content = AspCodeConfusion(Content )	'Asp代码混淆处理
	Content = ASPVariableGarbled(Content )	'Asp变量乱码 只对大小写乱
	AspCodeSafe = Content
End Function 

'截取函数里大括号
function cutDaKuoHao(content)
	cutDaKuoHao=cutFunctionvValue(content,"{","}")
end function
'截取函数里小括号
function cutXianKuoHao(content)
	cutXianKuoHao=cutFunctionvValue(content,"(",")")
end function

'截取函数值20150716
function cutFunctionvValue(content,startStr,endStr)
	dim ku1,ku2,i,s,c
	ku1=1:ku2=0
	for i=1 to len(content)
		s=mid(content,i,1)
		c=c & s
		if s=startStr then
			ku1=ku1+1
		elseif s=endStr then
			ku2=ku2+1
			if ku1=ku2 then
				exit for
			end if
		end if 
	next
	cutFunctionvValue = c
end function
'获得字符在内容里出现次数20150721
function getStrIntContentNumb(content,findStr)
	dim splstr
	if instr(content,findStr)>0 then
		splstr=split(content,findStr)
		getStrIntContentNumb = Ubound(splstr)
	else
		getStrIntContentNumb=0
	end if
end function
%>





