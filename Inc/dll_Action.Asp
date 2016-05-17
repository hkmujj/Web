<%
Option Explicit
'HTML标签嵌套框架 作者：小云   联系QQ313801120   时间20150730
dim c
'c=getftext("/1.html")
c="{$getLableValue title='标题 now()'content=""len(now())""$}"
call rw(HandleAction(c))



'处理动作   ReplaceValueParam为控制字符显示方式
Function HandleAction(Content)
    Dim StartStr, EndStr, ActionList, SplStr, Action, S, HandYes
    StartStr = "{\$" : EndStr = "\$}"
    ActionList = GetArray(Content, StartStr, EndStr, True, True)
    SplStr = Split(ActionList, "$Array$")
    For Each S In SplStr
        Action = Trim(S)
        If Action <> "" Then
          Action = Trim(Mid(Action,3,Len(Action)-4)) & " "
            HandYes = True          '处理为真
            '{VB #} 这种是放在图片路径里，目的是为了在VB里不处理这个路径
            If CheckFunValue(Action, "# ") = True Then
                Action = ""
            '读模板样式并设置标题与内容   软件里有个栏目Style进行设置
            ElseIf CheckFunValue(Action, "getLableValue ") = True Then
                Action = XY_getLableValue(Action)
            Else
                HandYes = False         '处理为假
            End If
            '注意这样，有的则不显示 晕 And IsNul(Action)=False
            If isNul(Action) = True Then Action = ""
            If HandYes = True Then
                Content = Replace(Content, S, Action)
            End If
        End If
    Next
    HandleAction = Content
End Function
'获得标签内容
Function XY_getLableValue(Action)
    Dim title,content,c
    title = RParam(Action, "title")
    content = RParam(Action, "content")
	c=c & "title=" & GetContentRunStr(title) & "<hr>"
	c=c & "content=" & GetContentRunStr(content) & "<hr>"
    XY_getLableValue = c
End Function

'替换参数值 2014  12 01
Function ReplaceValueParam(Content,ParamName,ReplaceStr)
	Dim StartStr,EndStr,LabelStr,nLen,nTimeFormat,DelHtmlYes,TrimYes
	'ReplaceStr = ReplaceStr & "这里面放上内容在这时碳呀。"
'	ReplaceStr = CStr(ReplaceStr)			'转成字符类型
	If IsNul(ReplaceStr)=True Then ReplaceStr=""
	
	StartStr="[$"& ParamName : EndStr = "$]"
	If InStr(Content,StartStr)>0 And InStr(Content,EndStr)>0 Then
		LabelStr = StrCut(Content,StartStr,EndStr,1)
		'删除Html
		DelHtmlYes = RParam(LabelStr,"DelHtml")			'是否删除Html
		If DelHtmlYes="true" Then ReplaceStr = Replace(DelHtml(ReplaceStr),"<","&lt;")			'HTML处理
		'删除两边空格
		TrimYes = RParam(LabelStr,"Trim")			'是否删除两边空格
		If TrimYes="true" Then ReplaceStr = TrimVbCrlf(ReplaceStr)
	
		'截取字符处理
		nLen = RParam(LabelStr,"Len")			'字符长度值
		nLen = HandleNumber(nLen)
		'If nLen<>"" Then ReplaceStr = CutStr(ReplaceStr,nLen,"null")' Left(ReplaceStr,nLen)
		If nLen<>"" Then ReplaceStr = CutStr(ReplaceStr,nLen,"...")' Left(ReplaceStr,nLen)
		
		'时间处理
		nTimeFormat  = RParam(LabelStr,"Format_Time")			'时间处理值
		If nTimeFormat<>"" Then 
			ReplaceStr = Format_Time(ReplaceStr, nTimeFormat)
		End If		
		Content = Replace(Content,LabelStr,ReplaceStr)
		
	End If
	ReplaceValueParam = Content
End Function
Function DelHtml(ByVal strHtml)
    On Error Resume Next
    Dim Rep, strOutput
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.IgnoreCase = True
        Rep.Global = True
        Rep.Pattern = "(<[a-zA-Z].*?>)|(<[\/][a-zA-Z].*?>)"
        strOutput = Rep.Replace(strHtml, "")
        DelHtml = strOutput
    Set Rep = Nothing
    If Err Then doError Err.Description, "DelHtml 删除HTML标签 函数出错，strHtml=" & strHtml
End Function
Function TrimVbCrlf(Str)
    TrimVbCrlf = PHPRTrim(PHPLTrim(Str))
End Function

Function HandleNumber(ByVal Content)
    Dim I, S, C
    C = ""
    For I = 1 To Len(Content)
        S = Mid(Content, I, 1)
        If InStr("0123456789", S) > 0 Then
            C = C & S
        End If
    Next
    HandleNumber = C
End Function

Function CutStr(ByVal Content, ByVal CutNumb, ByVal MoreStr)
    Dim I, S, N
    N = 0
    CutNumb = Cint(CutNumb)'转换成数字类型    追加于20141107
    If MoreStr = "" Then MoreStr = "..."
    If LCase(MoreStr) = "none" Or LCase(MoreStr) = "null" Then MoreStr = ""
    CutStr = Content
    For I = 1 To Len(Content)
        S = Asc(Mid(Cstr(Content), I, 1))
        If S < 0 Then S = S + 65536
        If S < 255 Then N = N + 1
        If S > 255 Then N = N + 2
        If N >= CutNumb Then CutStr = Left(Content, I) & MoreStr : Exit Function
    Next
End Function 
    Function Format_Time(s_Time, n_Flag)
        Dim Y, M, D, H, Mi, S
        Format_Time = ""
        If IsDate(s_Time) = False Then Exit Function
        Y = Cstr(Year(s_Time))
        M = Cstr(Month(s_Time))
        If Len(M) = 1 Then M = "0" & M
        D = Cstr(Day(s_Time))
        If Len(D) = 1 Then D = "0" & D
        H = Cstr(Hour(s_Time))
        If Len(H) = 1 Then H = "0" & H
        Mi = Cstr(Minute(s_Time))
        If Len(Mi) = 1 Then Mi = "0" & Mi
        S = Cstr(Second(s_Time))
        If Len(S) = 1 Then S = "0" & S
        Select Case n_Flag
            Case 1
                                    'yyyy-mm-dd hh:mm:ss
                Format_Time = Y & "-" & M & "-" & D & " " & H & ":" & Mi & ":" & S
            Case 2
                                    'yyyy-mm-dd
                Format_Time = Y & "-" & M & "-" & D
            Case 3
                                    'hh:mm:ss
                Format_Time = H & ":" & Mi & ":" & S
            Case 4
                                    'yyyy年mm月dd日
                Format_Time = Y & "年" & M & "月" & D & "日"
            Case 5
                                    'yyyymmdd
                Format_Time = Y & M & D
            Case 6
                                    'yyyymmddhhmmss
                Format_Time = Y & M & D & H & Mi & S
            Case 7
                                    'mm-dd
                Format_Time = M & "-" & D
            Case 8
                                    'yyyy年mm月dd日
                Format_Time = Y & "年" & M & "月" & D & "日" & " " & H & ":" & Mi & ":" & S
            Case 9
                                    'yyyy年mm月dd日H时mi分S秒 早上
                Format_Time = Y & "年" & M & "月" & D & "日" & " " & H & "时" & Mi & "分" & S & "秒，" & GetDayStatus(H, 1)
            Case 10
                                    'yyyy年mm月dd日H时
                Format_Time = Y & "年" & M & "月" & D & "日" & H & "时"
            Case 11
                                    'yyyy年mm月dd日H时mi分S秒
                Format_Time = Y & "年" & M & "月" & D & "日" & " " & H & "时" & Mi & "分" & S & "秒"
            Case 12
                                    'yyyy年mm月dd日H时mi分
                Format_Time = Y & "年" & M & "月" & D & "日" & " " & H & "时" & Mi & "分"
            Case 13
                                    'yyyy年mm月dd日H时mi分 早上
                Format_Time = M & "月" & D & "日" & " " & H & ":" & Mi & " " & GetDayStatus(H, 0)
            Case 14
                                    'yyyy年mm月dd日
                Format_Time = Y & "/" & M & "/" & D
            Case 15
                                    'yyyy年mm月 第1周
                Format_Time = Y & "年" & M & "月 第" & GetCountPage(D, 7) & "周"
        End Select
End Function

Function PHPRTrim(Str)
    Dim Pos, isBlankChar
    Pos = Len(Str)
    isBlankChar = True
    While isBlankChar and Pos >= 2
        If Mid(Str, Pos, 1) = " " or Mid(Str, Pos, 1) = vbTab Then'把vbTab也去掉
            Pos = Pos - 1
        ElseIf Mid(Str, Pos - 1, 2) = vbCrLf Then
            Pos = Pos - 2
        Else
            isBlankChar = False
        End If
    Wend
    PHPRTrim = RTrim(Left(Str, Pos))
End Function

Function PHPLTrim(Str)
    Dim Pos, isBlankChar
    Pos = 1
    isBlankChar = True
    While isBlankChar
        If Mid(Str, Pos, 1) = " " Or Mid(Str, Pos, 1) = vbTab Then'把vbTab也去掉
            Pos = Pos + 1
        ElseIf Mid(Str, Pos, 2) = vbCrLf Then
            Pos = Pos + 2
        Else
            isBlankChar = False
        End If
    Wend
    PHPLTrim = Right(Str, Len(Str) - Pos + 1)
End Function

    Function GetDayStatus(H, SType)
        Dim C
        If Left(H, 1) = "0" Then
            H = Right(H, 1)
        End If
        H = CInt(H)
        If H >= 0 And H <= 5 Then
            C = "凌晨"
        ElseIf H >= 6 And H <= 8 Then
            C = "早上"
        ElseIf H >= 9 And H <= 12 Then
            C = "上午"
        ElseIf H >= 13 And H <= 18 Then
            C = "下午"
        ElseIf H >= 19 And H <= 24 Then
            C = "晚上"
        Else
            C = "深夜"
        End If
        If SType = 1 Then C = "<b>" & C & "</b>"
        GetDayStatus = C
End Function

    Function GetCountPage(nCount, nPageSize)
            '把负数转成正确进行计算20150502
        if nCount < 0 then
            nCount = nCount * - 1
        end if
        if nPageSize < 0 then
            nPageSize = nPageSize * - 1
        end if
        GetCountPage = Fix(nCount / nPageSize)
        If InStr(nCount / nPageSize, ".") > 0 Then GetCountPage = GetCountPage + 1
End Function



Function echoPrompt(Word, Str)
    Response.Write("<font color=Green>" & Word & "</font>：" & Str & "<br>")
End Function
Function Echo(Word, Str)
    Call echoPrompt(Word, Str)
End Function


Function GetArray(Content, StartStr, EndStr, StartType, EndType)
    GetArray = GetArrayList(Content, StartStr, EndStr, StartType, EndType, "")
End Function

Function CheckFunValue(Action, FunName)
    CheckFunValue =(Left(Action, Len(FunName)) = FunName)
End Function

Function isNul(ByVal s)
    On Error Resume Next : IF Err.Number <> 0 Then Err.Clear
    isNul = False
    Select Case VarType(s)
        Case vbEmpty, vbNull
            isNul = True : Exit Function
        Case vbString
            If s = "" Then isNul = True : Exit Function
        Case vbObject
            Select Case TypeName(s)
                Case "Nothing", "Empty"
                    isNul = True : Exit Function
                Case "Recordset"
                    If s.State = 0 Then isNul = True : Exit Function
                    If s.Bof And s.Eof Then isNul = True : Exit Function
                Case "Dictionary"
                    If s.Count = 0 Then isNul = True : Exit Function
            End Select
    Case vbArray, 8194, 8204, 8209
        If Ubound(s) = -1 Then isNul = True : Exit Function
    End Select
    On Error Goto 0
End Function


Function GetArrayList(ByVal Content, ByVal StartStr, ByVal EndStr, ByVal StartType, ByVal EndType, ByVal SType)
    'on error resume next
    If IsNull(Content) = True Or IsNull(StartStr) = True Or IsNull(EndStr) = True Then
        GetArrayList = ""
        Exit Function
    End If
    Dim C, Rep, Matches, TheMatch
    Content = Replace(Replace(Content, Chr(10), "【换行】"), Chr(13), "【换行】")'替换掉换行
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.IgnoreCase = True
        Rep.Global = True
        Rep.Pattern = "(" & StartStr & ").+?(" & EndStr & ")"
        Set Matches = Rep.Execute(Content)
            For Each TheMatch In Matches
                C = C & "$Array$" & TheMatch.Value
            Next
        Set Matches = Nothing
        If C = "" Then
            GetArrayList = ""
            Exit Function
        End If
        C = Right(C, Len(C) - 7)
        If StartType = False Then
            Rep.Pattern = StartStr
            C = Rep.Replace(C, "")
        End If
        If EndType = False Then
            Rep.Pattern = EndStr
            C = Rep.Replace(C, "")
        End If
    Set Rep = Nothing
    If SType = "1" Or SType = "去掉字符" Or SType = True Then
        C = Replace(C, """", "")
        C = Replace(C, "'", "")
        C = Replace(C, " ", "")
        C = Replace(C, "(", "")
        C = Replace(C, ")", "")
    End If
    GetArrayList = C
End Function


Function RParam(ByVal Action, ByVal LableStr)
    dim s
	'原始 单引号
    s = handleRParam(Action, LableStr, "'")
	'原始 双引号
    if s = "" then
        s = handleRParam(Action, LableStr, """")
    end if
	'原始 空
    if s = "" then
        s = handleRParam(Action, LableStr, "")
    end if
	
	'小写 单引号
	if s = "" then
	    s = handleRParam(Action, LCase(LableStr), "'")
	end if
	'小写 双引号
    if s = "" then
        s = handleRParam(Action, LCase(LableStr), """")
    end if
	'小写 空
    if s = "" then
        s = handleRParam(Action, LCase(LableStr), "")
    end if
	
	'大写 单引号
	if s = "" then
	    s = handleRParam(Action, UCase(LableStr), "'")
	end if
	'大写 双引号
    if s = "" then
        s = handleRParam(Action, UCase(LableStr), """")
    end if
	'大写 空
    if s = "" then
        s = handleRParam(Action, UCase(LableStr), "")
    end if	
    RParam = s
End Function
function handleRParam(ByVal Action, ByVal LableStr, ByVal typeStr)
    Dim LalbeName,endTypeStr 
	endTypeStr = IIF(typeStr<>"",typeStr," ")
    Action = vbCrLf & " " & Action'给它也加个空格，要不然在没有函数，前面就没有空格
	'默认前面加空格
    LalbeName = " " & LableStr'加个空格是为了精准
    '不存在  前面加点
    If InStr(Action, LalbeName & "=" & typeStr) = False Then
        LalbeName = "'" & LableStr
    End If
    '不存在 前面加双引号
    If InStr(Action, LalbeName & "=" & typeStr) = False Then
        LalbeName = """" & LableStr
    End If
    '不存在	前面加TAB
    If InStr(Action, LalbeName & "=" & typeStr) = False Then
        LalbeName = vbTab & LableStr
    End If
    '不存在	前面加换行
    If InStr(Action, LalbeName & "=" & typeStr) = False Then
        LalbeName = vbCrLf & LableStr
    End If
    If InStr(Action, LalbeName & "=" & typeStr) > 0 And InStr(Action, endTypeStr) > 0 Then
        handleRParam = StrCut(Action, LalbeName & "=" & typeStr, endTypeStr, 2)
        handleRParam = HandleInModule(handleRParam, "end")'处理里面参数 追加于20141031
            '判断是否对参数进行动作制作
        If InStr(handleRParam, "{$") > 0 And InStr(handleRParam, "$}") > 0 Then
            handleRParam = HandleAction(handleRParam)'给AddSqL处理一下动作
        End If
    End If
End Function
Function IIf(bExp, sVal1, sVal2)
    If(bExp) Then IIf = sVal1 Else IIf = sVal2
End Function
Function StrCut(ByVal Content, ByVal StartStr, ByVal EndStr, ByVal CutType)
    On Error Resume Next
    Dim S1, S2, S3
    Select Case CutType
        Case 1
            S1 = InStr(Content, StartStr)
            S2 = InStr(S1, Content, EndStr) + Len(EndStr)
        Case 2, 0, ""
            S1 = InStr(Content, StartStr) + Len(StartStr)
            S2 = InStr(S1, Content, EndStr)
    End Select
    S3 = S2 - S1
    If S3 >= 0 Then
        StrCut = Mid(Content, S1, S3)
    Else
        StrCut = ""
    End If
    If Err <> 0 Then Call Eerr(StartStr, Content)
    'doError Err.Description, "StrCut 截取字符串 函数出错，StartStr=" & EchoHTML(StartStr) & "<hr>EndStr=" & EchoHTML(EndStr)
End Function

Function HandleInModule(Content, SType)
    SType = LCase(CStr(SType))
    If SType = "1" Or SType = "start" Then
        Content = Replace(Content, "\'", "\|*|\")
        Content = Replace(Content, "\=", "\|&|\")'后加20141024
    ElseIf SType = "2" Or SType = "end" Then
        Content = Replace(Content, "\|*|\", "'")
        Content = Replace(Content, "\$", "$")
        Content = Replace(Content, "\}", "}")

        Content = Replace(Content, "\|&|\", "=")'后加20141024
    End If
    HandleInModule = Content
End Function

Function Eerr(Word, Str)
'Response.Write(TypeName(Word) & "-" & TypeName(Str)):Response.End()
    Response.Write("<font color=red>" & Word & "</font>：" & Str)
    Response.End()
End Function

Function Rw(Str)
    Response.Write(Str)
End Function

Function GetFText(FileName)
    GetFText = GetFileText(FileName)
End Function
Function GetFileText(ByVal FileName)
    On Error Resume Next
    Dim Fso, FText, OpenFile
    'GetFileText = ""   '它默认返回的就是空， 这个是多此一举 (2013,9,30)
    Call HandlePath(FileName)'获得完整路径
    Set Fso = CreateObject("Scripting.FileSystemObject")
        If Fso.FileExists(FileName) = True Then
            Set FText = Fso.OpenTextFile(FileName, 1)
                            '加强 读空文件出错
                Set OpenFile = Fso.GetFile(FileName)
                    If OpenFile.Size = 0 Then Exit Function'文件为空则退出
                Set OpenFile = Nothing
                GetFileText = FText.ReadAll
            Set FText = Nothing
        End If
    Set Fso = Nothing
    If Err Then doError Err.Description, "GetFileText 读取文件内容 函数出错，FileName=" & FileName
End Function


Function HandlePath(Path)'Path前面不加ByVal 重定义，这样是为了让前面函数里可以使用这个路径完整调用
        Path = Replace(Path, "/", "\")
        Path = Replace(Path, "\\", "\")
        Path = Replace(Path, "\\", "\")
        dim isDir'为目录
        isDir = false
        if right(Path, 1) = "\" then
            isDir = true
        end if
        If InStr(Path, ":") = 0 Then
            If Left(Path, 1) = "\" Then
                Path = Server.MapPath("\") & "\" & Path
            Else
                Path = Server.MapPath(".\") & "\" & Path
            End If
        End If
        Path = Replace(Path, "/", "\")
        Path = Replace(Path, "\\", "\")
        Path = Replace(Path, "\\", "\")
        path = FullPath(Path)
        if isDir = true then
            Path = Path & "\"
        end if
        HandlePath = Path
End Function

Sub DoError(S, Msg)
    'On Error Resume Next
    Dim nRnd, C
    Randomize
    nRnd = CLng(Rnd() * 29252888)
    C = "<br />"
    C = C & "<div style=""width:100%; font-size:12px;;line-height:150%"">"
    C = C & "  <label onClick=""ERRORDIV" & nRnd & ".style.display=(ERRORDIV" & nRnd & ".style.display=='none'?'':'none')"">"
    C = C & "  <span style=""background-color:#820222;color:#FFFFFF;height:23px;font-size:14px;cursor:pointer"">〖 出错 提示信息 ERROR 〗</span><br />"
    C = C & "  </label>"
    C = C & "  <div id=""ERRORDIV" & nRnd & """ style=""width:100%;border:1px solid #820222;padding:5px;overflow:hidden;"">"
    C = C & " <span style=""color:#FF0000;"">出错描述</span> " & S & "<br />"
    C = C & " <span style=""color:#FF0000;"">回显信息</span> " & Msg & "<br />"
    C = C & "  </div>"
    C = C & "</div>"
    C = C & "<br />"
    Response.Write(C)
    Response.End()'终止，程序停止
End Sub
 Function FullPath(ByVal Path)
        Dim SplStr, S, C
        Path = Replace(Path, "/", "\")
        SplStr = Split(Path, "\")
        For Each S In SplStr
            S = Trim(S)
            If S <> "" And S <> "." Then
                If InStr(C, "\") > 0 And S = ".." Then
                    C = Mid(C, 1, InStrRev(C, "\") - 1)
                Else
                    If C <> "" And Right(C, 1) <> "\" Then C = C & "\"
                    C = C & S
                End If
            End If
        Next
        FullPath = C
End Function
Function GetContentRunStr(ByVal Content)
    GetContentRunStr = HandleContentCode(Content, "")
End Function
Function HandleContentCode(ByVal Content, SType)
    Dim StartStr, EndStr, S, YunStr, SplStr, Splxx, SplContent
    Dim LabelName'标签名称
    Dim nStartStr'开始字符

    Dim ContentLCase'定义内容小写
    ContentLCase = LCase(Content)'内容小写
	
	'call echo("Content",Content)
	
    If SType = "check" Or SType = "检测" Then
        HandleContentCode = False
    End If
    SplContent = "|printpre|authorinfo|foldingmenu|returnren|returnhr|returnrenhr|showerr|printforminfo|dedecmsmsg|getdidsidtidsql|checkbigclass|checksmallclass|checkthreeclass|checkproduct|autoaddmaininfo|handleautoaddmaininfo|"
    SplContent = SplContent & "|autoaddbigclass|autoaddsmallclass|autoaddthreeclass|autoaddarticle|autoaddarticleaction|gethttppage|geturlcontent|gethttpurl|bytestobstr|strcut|getstrcut|cutstr|cutstrnolu|setcuttdstr|specialhtmluploadencryptiondecrypt|"
    SplContent = SplContent & "|encaspcontent|uppercaseorlowercase|enccode|deccode|tounicode|japan|japan26|unjapan26|send_email|serversend_email|findaspcode|findaspfunlist|findaspfunctioncode|handleReplaceList|getfunnamedim|getkeyword|gethandlekeyword|"
    SplContent = SplContent & "|getchinastr|checkstr|regexp_checkstr|existsstr|regexp_replace|replacecolor|delhtml|regexp_delhtml|getaurltitlelist|phprand|phprnd|deleterepeatstr|replacen|fillzero|caseinsensitivereplace|saveremotefile|"
    SplContent = SplContent & "|downfile|xmlgetstatus|xmlpost|utfxmlpost|handlexmlpost|format_time|getthisurlnoparam|getgotourl|getgotourlnoparam|getgotourlnofilename|getip2|getip|getfiletext|getftext|getfiletype|checkfile|existsfile|"
    SplContent = SplContent & "|createfile|checkurl|is_url|checkerr|strtotruefalse|strtruefalsetoint|checkvbcrlf|checkbr|isparity|checkeval|htmlcode|showhtml|unhtmlcode|echohtml|errortext|testhtml|checkcode|seniorcheckcode|readfile|testgetpy|"
    SplContent = SplContent & "|pinyin|deleterepeatarray|getarraycount|randomshow|arrayrandomshow|csscompression|deletecssnote|uncsscompression|trimvbcrlf|phptrim|PHPLTrim|PHPRTrim|utf2gb|checkutftogb2312|convchinese|getwebtemplate|"
    SplContent = SplContent & "|getwebskins|detailpageaddhtmlprefix|gethandlefoldername|getmaininfourl|htmlformatting|getcontentrunstr|checkcontentrunstr|handledoublequotation|strdq|handlecontentcode|getaspcode|handlegetaspcode|getrandcolor|"

SplContent = SplContent & "|now|date|len|left|right|mid|lcase|ucase|"

    SplStr = Split(SplContent, "|")
    For Each LabelName In SplStr
        If LabelName <> "" Then
            StartStr = LCase(LabelName & "(") : EndStr = ")"
            nStartStr = InStr(" " & ContentLCase, " " & StartStr)
            If nStartStr > 0 And InStr(ContentLCase, EndStr) > 0 Then
                If SType = "check" Or SType = "检测" Then
                    HandleContentCode = True'为检测 有函数  退出
                    Exit Function
                End If
                Exit For
            End If
        End If
    Next
    If LabelName = "" Then
        If SType <> "check" And SType <> "检测" Then
            HandleContentCode = Content
        End If
        Exit Function
    End If

    LabelName = LCase(LabelName)'小写

    StartStr = LabelName & "(" : EndStr = ")"
    nStartStr = InStr(ContentLCase, StartStr)
    '记录查的原值，不分大小写
    YunStr = Mid(Content, InStr(ContentLCase, StartStr))
    YunStr = Mid(YunStr, 1, InStrRev(YunStr, EndStr))

    S = Mid(Content, InStr(ContentLCase, StartStr) + Len(StartStr))

    S = Mid(S, 1, InStrRev(S, EndStr) - 1)
    '多次处理 注意和下面的判断
    If CheckContentRunStr(Content) = True Then
        S = GetContentRunStr(S)
    End If
    '处理内容里有""，则给它删除掉20150329
    S = HandleDoubleQuotation(S)

    Splxx = Split(S & ",1", ",")

    '-----------------------------------------------------------

    If LabelName = "printpre" Then
        S = PrintPre(strDQ(S)) 
  
	'----------------------- ASP自身内部函数处理 -----------------------------
	'系统函数 Now
    ElseIf LabelName = "now" Then
        S = Now()
    '系统函数 Date
    ElseIf LabelName = "date" Then
        S = Date
    '系统函数 Len
    ElseIf LabelName = "len" Then
        S = Len(S)
    '系统函数 LCase
    ElseIf LabelName = "lcase" Then
        S = LCase(S)
    '系统函数 UCase
    ElseIf LabelName = "ucase" Then
        S = UCase(S)
    '系统函数 Left
    ElseIf LabelName = "left" Then
        Splxx = Split(S & ",1", ",")
        S = Left(Trim(Splxx(0)), Trim(Splxx(1)))
    '系统函数 Right
    ElseIf LabelName = "right" Then
        S = Right(Trim(Splxx(0)), Trim(Splxx(1)))
    '系统函数 Mid
    ElseIf LabelName = "mid" Then
        Splxx = Split(S & ",99999", ",")
        If UBound(Splxx) >= 2 Then
            S = Mid(Trim(Splxx(0)), Trim(Splxx(1)), Trim(Splxx(2)))
        Else
            S = Mid(Trim(Splxx(0)), Trim(Splxx(1)))
        End If
    End If
    Content = Replace(Content, YunStr, S)
    HandleContentCode = Content
End Function


Function CheckContentRunStr(ByVal Content)
    CheckContentRunStr = HandleContentCode(Content, "check")
End Function

Function HandleDoubleQuotation(Byval S)
    Dim NewS
    NewS = PHPTrim(S)
    If Left(NewS, 1) = """" And Right(NewS, 1) = """" Then
        S = Mid(NewS, 2, Len(NewS) - 2)
    End If
    HandleDoubleQuotation = S
End Function

Function PrintPre(Content)
    Content = Replace(Content, "<", "&lt;")
    PrintPre = "<pre>" & Content & "</pre>"
End Function

Function strDQ(ByVal S)
    strDQ = HandleDoubleQuotation(S)
End Function
Function PHPTrim(Str)
    PHPTrim = TrimVbCrlf(Str)
End Function


%>