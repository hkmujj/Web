<% 
'格式化asp代码 不是最新20151223
'处理案例网站操作 
'Call AnQuanWeb("E:\E盘\WEB网站\至前网站\11\") 
'处理安全网站 
Sub AnQuanWeb(FolderPath)
    Dim IncFolder, AdminFolder 
    If FolderPath <> "" Then
    ElseIf Request("Folder") <> "" Then
        FolderPath = Request("Folder") 
    ElseIf Request("FolderPath") <> "" Then
        FolderPath = Request("FolderPath") 
    End If 
    Call Echo("文件夹路径", FolderPath) 
    IncFolder = FolderPath & "\Inc\" 
    Call Echo("操作文件夹", IncFolder) 
    Call DelSelFile(IncFolder, "Install.Asp") 
    Call HandleAnQuan(IncFolder) 
    AdminFolder = FolderPath & "\Admin\" 
    Call Echo("操作文件夹", AdminFolder) 
    Call DelSelFile(AdminFolder, "test.html|复件 MemberManage.Asp") 
    Call HandleAnQuan(AdminFolder) 
    Response.End() 
End Sub 
'获得全部文件（夹）列表 
Function HandleAnQuan(FolderPath)
    Dim Fso, F, Fc, MyFile, F1, FileType, C, Content, Char_Set, nLen, LeftStr, RightStr, ContentStr 
    ContentStr = vbCrlf & "'******** 压缩处理 ********" & vbCrlf 
    Call HandlePath(FolderPath)    '获得完整路径 
    Set Fso = CreateObject("Scripting.FileSystemObject")
        If Not Fso.FolderExists(FolderPath) Then Exit Function 
        Set F = Fso.GetFolder(FolderPath)
            For Each MyFile In F.Files
                FileType = LCase(Right(MyFile.Name, Len(MyFile.Name) - InStrRev(MyFile.Name, "."))) 
                If FileType = "asp" Then
                    Content = GetFText(MyFile) 
                    If InStr(Content, "<" & "%") > 0 Then
                        Char_Set = CheckCode(MyFile) 
                        Content = ReadFile(MyFile, Char_Set) 
                        If InStr(Content, ContentStr) > 0 Then
                            Call Echo(MyFile, "已经压缩[" & Char_Set & "]") 
                        ElseIf Content <> "" Then
                            '读出文件安全内容
                            Content = GetAnquanContent(MyFile, Char_Set) 
                            nLen = InStr(Content, "<" & "%" & vbCrlf) 
                            If nLen > 0 Then
                                LeftStr = Mid(Content, 1, nLen + 2) 
                                RightStr = Mid(Content, nLen + 2) 
                                Content = LeftStr & ContentStr & RightStr 
                            Else
                                Content = Content & vbCrlf & "<" & "%" & ContentStr & "%" & ">" 
                            End If 
                            Call WriteToFile(MyFile, Content, Char_Set) 
                            Call Echo(MyFile, "修改成功[" & Char_Set & "]") 
                        End If 
                        Doevents 
                    End If 
                End If 
            Next 
            Call Echo(FolderPath, "共有文件数【" & F.Files.Count & "】<hr>") 
            Set Fc = F.SubFolders
                For Each F1 In Fc
                'HandleAnQuan(F1)
                Next 
            Set Fc = Nothing 
        Set F = Nothing 
    Set Fso = Nothing 
    If Err Then doError Err.Description, "GetFileFolderList 获得全部文件（夹）列表 函数出错，FolderName=" & FolderName 
End Function 
'获得安全代码 
Function GetAnquanContent(FileName, Char_Set)
    '想要函数变量简洁处理 成立  clickdim=false  为开启
    Dim ObjCode, C, Content, TableHtml, IDList, I, Ac, Char, Path, DirPath, CountRow 
    Set ObjCode = New ASPCode
        CountRow = 19999 
        ObjCode.SetCountRow = CountRow 
        ObjCode.SetAnQuanType = "|行不退格|NO处理成一行|NO函数变量简洁处理|函数全部小写|" 
        Action = "|note=false|color=false|clickdim=false|html=false|code=true|anquan=true|" 
        ObjCode.SetAspSetType = Action    'ASPCode动作列表 
        Content = ObjCode.HandleAspCode(ReadFile(FileName, Char_Set)) 
        Content = Replace(Replace(Replace(Content, "&lt;", "<"), "&#x5B;", "["), "&amp;", "&") 
        GetAnquanContent = Content 
End Function
'显示ASP代码 
Function ShowAspCodeColor(Content)
    Dim ObjCode, C, TableHtml, IDList, I, Ac, Char, Path, DirPath, CountRow 
    Set ObjCode = New ASPCode
        CountRow = 19999 
        ObjCode.SetCountRow = CountRow 
        Action = "|note=false|color=true|clickdim=false|html=false|code=true|anquan=true|" 
        ObjCode.SetAspSetType = Action    'ASPCode动作列表 
        Content = ObjCode.HandleAspCode(Content) 
        Content = "<pre>" & Content & "</pre>" 
        ShowAspCodeColor = Content 
End Function
'外部调用 
Function ASPShowCode(Content)
    Dim ObjCode, C, TableHtml, IDList, I, Ac, Char, Path, DirPath, CountRow 
    Set ObjCode = New ASPCode
        CountRow = 19999 
        ObjCode.SetCountRow = CountRow 
        Action = "|note=true|color=true|clickdim=false|html=false|code=true|anquan=true|" 
        ObjCode.SetAspSetType = Action    'ASPCode动作列表 
        Content = ObjCode.HandleAspCode(Content) 
        Content = "<pre>" & Content & "</pre>" 
        ASPShowCode = Content 
End Function

'优化ASP代码 20150127
Function OptimizationAspCode(Content)
    Dim ObjCode, CountRow, Action
    Set ObjCode = New AspCode
        CountRow = 19999 
        ObjCode.SetCountRow = CountRow 
        Action = "|empty=true|fuhao=false|note=true|color=false|clickdim=true|html=false|code=true|anquan=false|" 
        ObjCode.SetAspSetType = Action                                                  'ASPCode动作列表
 		
        Content = ObjCode.HandleAspCode(Content) 
		
		Content=Replace(Content,"&lt;","<")
        OptimizationAspCode = Content 
End Function


Private Char_Set, FileName, Action, Note, Color, ClickDim, Html, Code
'ASP代码分析器(2013,10,7)  遇到问题解决问题 待有时间把多余的代码去掉 
Class ASPCode
'VB开始
    Private AspD, AspHelpTitle, AspHelpContent, CountRow, RowMsg 
    Private ArrayColor(199), BlueStr, RedStr, MyDimList, WriteList 
    Private FunValue, FunName, FunDim, FunDimStr, FunDimNumb, OKFunDim, NOFunDim, NoDimNumb 
    Private RootFunDim, RootFunDimStr, RootFunDimNumb, OKRootFunDim, NORootFunDim, AspSetType, CaseNumb 
    Private RSpaceStr, RSpaceNumb, ARF    'ARF为ASP中收集蓝色变量列表(=AspRowFunDimList) 
    Private ArrStr(20), ArrNumb, Css, nRow, NoteC, TempNote, nLen, nLen1, nLen2, nLen3 
    Private EvalList, BadDimList, ReplaceDimList, AnQuanDim, AnQuanType, BR 
    '让Asp内置函数突出显示
    Function HandleAspCode(ByVal Content)
        'On Error Resume Next
        Dim S, YuanS, S1, S2, LCaseS, nLen, RowList, YesOK, Color, ErrList, Title 
        Dim DotS, DotLeftC, DotRightC, StartStr, EndStr, StartYes, EndYes, DotYes, LeftC, RightC 
        Dim NoteNumb, SpaceNumb, TempStr, TempC, RowIFList, TheIFNumb, RI 
        Dim SplStr, Splxx, C, I, J, N, Sx, Sc, Nc, LCaseCode, Wc, Zc, SYHCount    '双引号总数 
        Dim ASPYes, IncludeYes, HeadYes, ASPCode, Tc    '为暂存内容 
        Dim CodeS, MHList, MHNumb, TheMH    'MaoHaoNumb冒号数 
        'AspSetType = "|note=false|color=false |clickdim=false|html=false|code=false|"         
        'Call Echo("AspSetType",AspSetType)
        Content = Replace(Content, vbTab, "    ")    '替换Tab为四个空格 
        SplStr = Split(Content, vbCrlf)    '让分行变为数组 
        ASPYes = True    '默认ASP为运行状态 
        nRow = 0 : CaseNumb = 0    '当前行数 
        RSpaceStr = "    " : RSpaceNumb = 0    '行空格值与数 
        IncludeYes = False : HeadYes = False    'ASP的Include与头像配置默认为假 
        NoDimNumb = 0    '函数中无意义的变量数 
        OKFunDim = "," : OKRootFunDim = "," 
        If UBound(SplStr) > CountRow Then
            RowMsg = "显示[" & CountRow & "/" & UBound(SplStr) & "]" 
        Else
            CountRow = UBound(SplStr) 
        End If 
        For RI = 0 To CountRow
            S = SplStr(RI) : YuanS = S 
            '当原代码设置为假时,清空S两边空格
            If InStr(AspSetType, "|code=false|") = False Then
                S = Trim(S)    '去除两边空格 
            End If 
            CodeS = Trim(S)    '获得一行干净的ASP代码 
            ARF = ","    'Asp行函数列表 
            SYHCount = 0    '双引号数 
            MHList = "" : MHNumb = 0 : TheMH = 0    '冒号列表与冒号数 
            RowIFList = "" : TheIFNumb = 1    '行IF分割列表 
            LCaseS = LCase(S)    '将S转成小写 
            'ASP为真 ASP代码不为空 则ASP代码累加
            If ASPYes = True And ASPCode <> "" Then
                ASPCode = ASPCode & vbCrlf 
            End If 
            Sc = "" : Wc = "" : Nc = "" : Zc = ""    'Sc为S循环累加值,Wc为"号里面内容 Nc为数字 Zc单词 
            '逐个循环字符
            For I = 1 To Len(S)
                Sx = Mid(S, I, 1) 
                '如果ASP运行 则收集代码
                If ASPYes = True Then
                    ASPCode = ASPCode & Sx 
                End If 
                '收集当前字符左边值与右边，去空格后
                LeftC = Replace(Mid(S, 1, I - 1), " ", "") 
                RightC = Replace(Mid(S, I + 1), " ", "") 
                '对 < 单独处理
                If Sx = "<" Then
                    'ASP头部
                    If Mid(S, I + 1, 2) = "%@" Then    '待应用，可用变量改成深红色 #990033
                        ASPYes = True    'ASP为真 
                        HeadYes = True    'ASP头部为真 
                        Sc = Sc & "&lt;%@"    '等于< 
                        ASPCode = ASPCode & "&lt;%@"    'ASPCode收集累加 
                        I = I + 3 
                        Sx = Mid(S, I, 1) 
                    '判断ASP开始
                    ElseIf Mid(S, I + 1, 1) = "%" Then
                        S1 = Left(Replace(Mid(S, I), " ", ""), 4) 
                        If S1 = "<" & "%%" & ">" Then    '判断这是不是一个空ASP定义
                            I = I + InStr(Mid(S, I), "%" & ">") + 1 
                        Else
                            ASPYes = True    'ASP为真 
                            Sc = Sc & "&lt;%"    '等于< 
                            ASPCode = ASPCode & "&lt;%"    'ASPCode收集累加 
                            I = I + 2 
                        End If 
                        Sx = Mid(S, I, 1)    '获得字符，好让下面程序继续运行 
                    'Include开始
                    ElseIf LCase(Mid(S, I, 4)) = "<!--" And InStr(LCase(Mid(S, I)), "#include ") > 0 And Wc = "" Then    'Wc为输入值
                        Tc = Mid(S, I)    '获得I后面全部代码 
                        nLen = InStr(LCase(Tc), "#include ") + 8 
                        'Tc = Replace(Mid(S,I, nLen), "<", "&lt;")                         
                        'Tc = Replace(Tc," ","") & " "                '让Include加载整齐
                        Tc = "&lt;!--" & "#Include " 
                        Sc = Sc & Tc 
                        ASPYes = True    'ASP为真 
                        IncludeYes = True    'ASP的Include为真 
                        I = I + nLen 
                        Sx = Mid(S, I, 1)    '获得字符，好让下面程序继续运行 
                    ElseIf Sc = "" Then    '当代码的第一个字符是<则为Html代码
                        ASPYes = False 
                    End If 
                End If 
                'ASP运行为假
                If ASPYes = False Then
                    If Sx = "<" Then
                        Sx = "&lt;" 
                    ElseIf Sx = "&" Then
                        Sx = "&amp;" 
                    End If 
                    Sc = Sc & Sx 
                '判断ASP结束
                ElseIf Sx = "%" And Mid(S, I + 1, 1) = ">" Then
                    ASPYes = False : IncludeYes = False : HeadYes = False 
                    Sx = Sx & ">" 
                    '判断是否为原来代码
                    If InStr(AspSetType, "|code=false|") = False Then
                        'Call Echo(Mid(S & "<"&"%",I+2,2)="<"&"%" ,InStr(AspCode,vbCrlf))
                        If InStr(ASPCode, vbCrlf) > 0 Then    '判断ASP结束后面还有代码 则换行
                            If InStr(AspSetType, "|html=false|") = False Then
                                Sx = Sx & "<br>" 
                            End If 
                            nRow = nRow + 1 
                        End If 
                    'Mid(S & "<"&"%",I+2,2)="<"&"%" And
                    End If 
                    Sc = Sc & Sx 
                    I = I + 1    '不能加2，只能加1，因为这里定义ASP为假，它会在下一次显示上面的 'ASP运行为假 
                    Sx = Mid(S, I, 1) 
                    ASPCode = "" 
                'Include结束
                ElseIf Sx = ">" And IncludeYes = True Then
                    ASPYes = False : IncludeYes = False    'ASP的Include与ASP为假 
                    S1 = LCase(Trim(Mid(S & " ", I + 1))) 
                    If S1 <> "" Then    '判断Include结束后面还有代码 则换行
                        If Left(S1, 5) = "<!--#" Then    '后面跟回函文件则换行
                            Sx = Sx & vbCrlf 
                        Else
                            Sx = Sx & "<br>" 
                        End If 
                        nRow = nRow + 1    '行数加一  【<br>】在复制时不换行，这样好，查看方便，复制代码也完整棒棒棒 
                    End If 
                    Sc = Sc & Sx 
                '输入文本
                ElseIf(Sx = """" Or Wc <> "") And ASPYes = True Then
                    '双引号累加
                    If Sx = """" Then SYHCount = SYHCount + 1 
                    '判断是否"在最后
                    If SYHCount Mod 2 = 0 Then
                        If Mid(S, I + 1, 1) <> """" Then
                            Wc = Wc & Sx 
                            SpaceNumb = 1 
                            S1 = Right(Replace(Mid(S, 1, I - Len(Wc)), " ", ""), 1)    '必需放在这里，要不会出错 
                            Wc = Replace(Wc, "&", "&amp;") 
                            Wc = Replace(Wc, "<", "&lt;") 
                            Wc = Replace(Wc, "[", "&#x5B;")    '为了对ClassASPCode自身文件处理时出错用 
                            '判断是否为组件定义
                            If InStr(Mid(S, 1, I), "CreateObject") > 0 Then
                                Wc = HandleWrite(Wc)    '处理输入的值 
                            End If 
                            '判断是否突出显示输入文本颜色
                            If InStr(AspSetType, "|color=false|") = False Then
                                Wc = "<font color=#008000>" & Wc & "</font>" 
                            End If 
                            If S1 = "(" Then
                                SpaceNumb = 0 
                            End If 
                            Sc = LeftSpace(Sc, SpaceNumb)    '让左边回头空格 
                            If IncludeYes = True Then
                                Wc = Replace(Wc, ".asp", ".Asp") 
                            End If 
                            Sc = Sc & Wc 
                            SYHCount = 0 : Wc = ""    '清除 
                        Else
                            Wc = Wc & Sx 
                        End If 
                    Else
                        Wc = Wc & Sx 
                    End If 
                    '最后代表 输入的有错误
                    If I = Len(S) Then
                        '判断是否突出显示输入文本颜色
                        If InStr(AspSetType, "|color=false|") = False Then
                            Wc = Replace(Wc, "&", "&amp;") 
                            Wc = Replace(Wc, "<", "&lt;") 
                            Sc = Sc & "<span style=""background:#FFB3B0;"" title=""错误：双引号有问题"">" & Wc & "</span>" 
                        End If 
                    End If 
                '注释
                ElseIf Sx = "'" Then    '没想到注释这么简单就搞定了，爽
                    S1 = Trim(Mid(S, I + 1))    '获得注释内容且清空两边空格 
                    CodeS = Mid(S, 1, I - 1)    '获得干净ASP代码  给找IF用的，不可删除 注意不可删除 
                    Sx = Sx & S1 
                    TempStr = Right(Trim(S1), 1) 
                    If TempStr = "=" Or TempStr = "/" Or((TempStr = ")" Or TempStr = "）") And InStr(S1, "(201") > 0) Then    '追加于20140311
                        Sx = Sx & vbCrlf 
                    End If 
                    Sx = Replace(Sx, "&", "&amp;") 
                    Sx = Replace(Sx, "<", "&lt;") 
                    Sx = Replace(Sx, "[", "&#x5B;")    '为了对ClassASPCode自身文件处理时出错用 
                    '判断是否突出显示输入文本颜色
                    If InStr(AspSetType, "|color=false|") = False Then
                        NoteNumb = NoteNumb + 1 
                        Sx = "<span id=""Note_" & NoteNumb & """ style=""color:#999999"">" & Sx & "</span>" 
                    End If 
                    '判断是否为原来代码为假
                    If InStr(AspSetType, "|code=false|") = False Then
                        '判断是否显示注释代码
                        If InStr(AspSetType, "|note=false|") = False Then
                            If I <> 1 Then Sc = LeftSpace(Sc, 4)    '让左边 追加空格 
                            Sc = Sc & Sx 
                        End If 
                    Else
                        Sc = Sc & Sx 
                    End If 
                    Exit For 
                '+-*\=&  加蓝色
                ElseIf InStr("+-*/\=&<>()^,%:", Sx) > 0 And ASPYes = True And IncludeYes = False Or(IncludeYes = True And Sx = "=") Then    '因为Include 中有=号
                    S1 = Right(LeftC, 1) : YesOK = True 
                    Color = "" : SpaceNumb = 1 
                    If(Sx = ">" And S1 = "<") Or Sx = "(" Or Sx = ")" Then
                        SpaceNumb = 0 
                        If Sx = "(" And I > 1 Then
                            S2 = Mid(S, 1, I - 1) 
                            S2 = Right(Trim(S2), 1) 
                            If InStr("+-*\/", S2) > 0 Then    '（前面为+-*\/算术符号时加空格2013,11,16
                                Sx = " " & Sx    '给(前面加个空格 添加2013,11,1 
                            End If 
                        End If 
                    ElseIf(Sx = "=" And S1 = "<") Or(Sx = "=" And S1 = ">") Then
                        SpaceNumb = 0 
                    ElseIf Sx = "," Then
                        SpaceNumb = 0 
                    End If 
                    If Sx = ":" Then
                        If S1 = ":" Or Replace(RightC, ":", "") = "" Then
                            YesOK = False 
                        Else    ':冒号处理待完善
                            MHNumb = MHNumb + 1    '冒号数累加 
                            TempStr = Mid(S, TheMH + 1, I - TheMH - 1)    '获得当前:值 
                            TempStr = Replace(TempStr, "&", "&amp;") 
                            TempStr = Replace(TempStr, "<", "&lt;") 
                            TempStr = Replace(TempStr, "[", "&#x5B;")    '为了对ClassASPCode自身文件处理时出错用 
                            TempStr = Replace(TempStr, "%", "&#x25;")    '为了对ClassASPCode自身文件处理时出错用 
                            TheMH = I    '让当前冒号等于I 
                            MHList = MHList & TempStr & "[云端]"    '累加冒号列表 
                        End If 
                    ElseIf Sx = "-" Then
                        TempStr = LCase(Replace(Mid(S, 1, I), " ", "")) 
                        If TempStr = "chr(-" Then SpaceNumb = 0    '解决Chr(-中空格  追加于20140312 
                    End If 
                    '判断是否突出显示输入文本颜色
                    If InStr(AspSetType, "|color=false|") = False Then
                        If Sx = "," Or Sx = ":" Then
                            Color = "#000000" 
                        ElseIf Sx = "(" Or Sx = ")" Then    '深蓝色
                            Color = "#000099" 
                        Else
                            Color = "#0000FF"    '蓝色 
                        End If 
                        If Sx <> "" Then Sx = "<font color=" & Color & ">" & Sx & "</font>"    '加个判断，要不于<%会有值 2013,10,25 
                    End If 
                    If YesOK = True Then    '当:过多时就不显示
                        Sc = LeftSpace(Sc, SpaceNumb)    '处理左边空格 
                        Sc = Sc & Sx 
                    End If 
                '字母
                ElseIf CheckABC(Sx) = True Or Sx = "_" Or Zc <> "" Then
                    Zc = Zc & Sx 
                    YesOK = True 
                    S1 = Mid(S & " ", I + 1, 1) 
                    S2 = Mid(Zc, 1, 1) 
                    If CheckABC(S1) <> True And S1 <> "_" Then
                        YesOK = False 
                    End If 
                    '允许变量后面是数字
                    If CheckNumber(S1) = True And CheckABC(S2) = True Then
                        YesOK = True 
                    End If 
                    If YesOK = False Then
                        '对IF单独处理
                        If MHNumb = 0 And(LCase(Zc) = "then" Or LCase(Zc) = "else" Or LCase(Zc) = "end") Then
                            If TheIFNumb = 1 Then
                                S1 = Mid(S, TheIFNumb, I - Len(Zc) - TheIFNumb + 1) 
                                RowIFList = Mid(S, 1, I) & "[云端]" 
                            Else
                                S1 = Mid(S, TheIFNumb + 1, I - Len(Zc) - TheIFNumb) 
                                RowIFList = RowIFList & Trim(S1) & "[云端]" & Zc & "[云端]" 
                            End If 
                            TheIFNumb = I 
                        End If 
                        Tc = Zc    'Tc暂存Zc内容 
                        '判断前面为空还是有个空格
                        Color = "" : SpaceNumb = 1 
                        S1 = Replace(Mid(S, 1, I - Len(Zc)), " ", "") 
                        S2 = Right(S1, 1) 
                        If S1 = "" Or S2 = "." Then    '.前面为空
                            SpaceNumb = 0 
                        ElseIf UCase(Zc) = "LANGUAGE" And S2 = "@" Then    'ASP@头前面为空
                            SpaceNumb = 0 
                            Color = "#990033" 
                        ElseIf UCase(Zc) = "CODEPAGE" And InStr(S, "@") > 0 Then    'ASP@头前面变颜色
                            SpaceNumb = 1 
                            Color = "#990033" 
                        ElseIf S2 = "(" Or Right(S1, 3) = "<" & "%=" Then    '前面是(无空格
                            SpaceNumb = 0 
                        ElseIf UCase(Left(Zc, 1)) = "H" And Mid(S, I - Len(Zc), 1) = "&" And Zc <> "Hour" Then    '&HF0000 排除 Hour
                            SpaceNumb = 0 
                        End If 
                        S1 = Zc 
                        YesOK =(LCase(Zc) = "eval" Or LCase(Zc) = "execute") And Mid(" " & S, I - Len(Zc) + 1, 1) = " " 
                        Sc = LeftSpace(Sc, SpaceNumb)    '处理左边空格 
                        Zc = ASPColor(S, I, Zc, AspSetType, Color)    'Code代表自动清空 
                        If InStr(AspSetType, "|html=false|") = False Then
                            If YesOK = True Then
                                Zc = "<font style=""background:#FFB3B0;"" title=""注意：" & S1 & "后门程序"">" & Zc & "</font>" 
                                Call HandleEvalList(YuanS, "<font color=red>" & S1 & "()函数可以执行任意ASP代码</font>", RI) 
                            ElseIf LCase(Zc) = "createtextfile" Then
                                Call HandleEvalList(YuanS, "使用了FSO的CreateTextFile写文件", RI) 
                            ElseIf LCase(Zc) = "savetofile" And Mid(" " & S, I - Len(S1) + 1, 1) = "." Then
                                Call HandleEvalList(YuanS, "使用了Stream的SaveToFile函数写文件", RI) 
                            ElseIf LCase(Zc) = "save" And Mid(" " & S, I - Len(S1) + 1, 1) = "." Then
                                Call HandleEvalList(YuanS, "使用了XMLHTTP的Save函数写文件", RI) 
                            End If 
                        End If 
                        'Rem注释
                        If LCase(Tc) = "rem" Then
                            Sx = Mid(S, I + 1) : I = I + Len(Sx) 
                            '注释改进(20131013)

                            NoteNumb = NoteNumb + 1 
                            '判断是否显示注释代码
                            If InStr(AspSetType, "|note=false|") = False Then
                                Sx = "<span id=""Note_" & NoteNumb & """ style=""color:#999999"">" & Sx & "</span>" 
                                Sc = Sc & Zc & Sx 
                            Else
                                Zc = ""    '清空，不需要Rem 
                            End If 
                        Else
                            Sc = Sc & Zc : Zc = "" 
                        End If 
                    End If 
                '为数字
                ElseIf CheckNumber(Sx) = True Or Nc <> "" Then
                    Nc = Nc & Sx 
                    S1 = Right(Replace(Mid(S, 1, I - Len(Nc)), " ", ""), 2) 
                    S2 = Mid(S & " ", I + 1, 1) 
                    Color = "" : SpaceNumb = 1 
                    If CheckNumber(S2) = False And S2 <> "." Then
                        '判断是否突出显示输入文本颜色
                        If InStr(AspSetType, "|color=false|") = False Then
                            Nc = "<font color=#FF0000>" & Nc & "</font>" 
                        End If 
                        If Right(S1, 1) = "(" Then
                            SpaceNumb = 0 
                        ElseIf Right(S1, 2) = "(-" Then    '解决Chr(-中空格  追加于20140312
                            SpaceNumb = 0 
                        ElseIf S1 = "=-" Then
                            SpaceNumb = 0 
                        End If 
                        Sc = LeftSpace(Sc, SpaceNumb)    '处理左边空格 
                        Sc = Sc & Nc 
                        Nc = "" 
                    End If 
                '最后输出
                Else
                    If Sx = "<" Then Sx = "&lt;" 
                    Sc = Sc & Sx 
                End If 
            Next 
            'Sc不为空
            If Sc <> "" Then    '不用ARF<>""是因为有的是其它ASP代码，但它要退格的，1010测试
                YesOK = True 
                '冒号
                If MHNumb > 1 And 1 = 2 Then    '不对:处理
                    Splxx = Split(MHList, "[云端]") 
                    TempC = "" 
                    For J = 0 To UBound(Splxx)
                        If J = 1 Or J = 3 Then
                            TempC = TempC & CopyStrNumb(RSpaceStr, RSpaceNumb + 1) & Splxx(J) & vbCrlf 
                        Else
                            TempC = TempC & CopyStrNumb(RSpaceStr, RSpaceNumb) & Splxx(J) & vbCrlf 
                        End If 
                    Next 
                    Sc = TempC 
                '函数结束
                ElseIf InStr(ARF, ",end,function,") > 0 Or InStr(ARF, ",end,sub,") > 0 Then
                    S1 = "Function" : S2 = "End Function" 
                    If InStr(ARF, ",end,sub,") > 0 Then
                        S1 = "Sub" : S2 = "End Sub" 
                    End If 
                    Call RowLocation(Sc, S1, S2, YesOK, True) 
                    nRow = nRow + 1    '函数结束要让行数累加一 
                    '判断是否显示HTML代码
                    If InStr(AspSetType, "|html=false|") = False Then
                        Sc = Sc & "</div>" 
                    End If 
                    If NOFunDim <> "," Then
                        Splxx = Split(NOFunDim, ",") 
                        For Each S1 In Splxx
                            If S1 <> "" Then
                                Css = Css & "." & FunName & S1 & "{background:#333333;color:#666666}" 
                                NoDimNumb = NoDimNumb + 1    '函数中无意义变量数 
                            End If 
                        Next 
                    End If 
                    FunName = ""    '清空函数名 
                    'Sc = Sc & vbCrlf    '<br>复制不换行 vbCrlf 复制有换行(怪)                     
                '函数的开始
                ElseIf(InStr(ARF, ",function,") > 0 Or InStr(ARF, ",sub,") > 0) And InStr(ARF, ",exit,") = False Then
                    Sc = GetNoteC(NoteC, "") & CopyStrNumb(RSpaceStr, RSpaceNumb) & Sc 
                    '清空数组内容 20个退格足够
                    For J = RSpaceNumb To 20
                        ArrStr(J) = "" 
                    Next 
                    S1 = "Function" 
                    If InStr(ARF, ",sub,") > 0 Then S1 = "Sub" 
                    ArrStr(RSpaceNumb) = S1 
                    RSpaceNumb = RSpaceNumb + 1 
                    YesOK = False 
                    '判断是否显示HTML代码
                    If InStr(AspSetType, "|html=false|") = False Then
                        S1 = "style='cursor:hand' title='点击收缩'" 
                        Sc = Replace(Sc, "Function</", "<font " & S1 & " onclick=""OnOffDiv(this,'ID" & FunName & "')"">Function</font></") 
                        Sc = Replace(Sc, "Sub</", "<font " & S1 & " onclick=""OnOffDiv(this,'ID" & FunName & "')"">Sub</font></") 
                        Sc = "<div id=""ID" & FunName & """>" & Sc 
                    End If 
                    'nRow = nRow + 1            '行数累加，给上面的函数用 这个不能用，要不然就会多行                     
                'Class结束
                ElseIf InStr(ARF, ",end,class,") > 0 Then
                    Call RowLocation(Sc, "Class", "End Class", YesOK, True) 
                'Class开始
                ElseIf InStr(ARF, ",class,") > 0 Then
                    ArrStr(RSpaceNumb) = "Class"    '把信息写入数组 
                    Sc = GetNoteC(NoteC, "") & CopyStrNumb(RSpaceStr, RSpaceNumb) & Sc 
                    RSpaceNumb = RSpaceNumb + 1 : YesOK = False 
                'Property开始
                ElseIf InStr(ARF, ",get,") > 0 Or InStr(ARF, ",let,") > 0 Then
                    ArrStr(RSpaceNumb) = "GetLet"    '把信息写入数组 
                    Sc = GetNoteC(NoteC, "") & CopyStrNumb(RSpaceStr, RSpaceNumb) & Sc 
                    RSpaceNumb = RSpaceNumb + 1 : YesOK = False 
                'Property结束
                ElseIf InStr(ARF, ",end,") > 0 And InStr(LCase(CodeS), "end property") > 0 Then
                    Call RowLocation(Sc, "GetLet", "End Property", YesOK, True) 
                '一个一行就结束的IF判断
                ElseIf InStr(ARF, ",if,") > 0 And InStr(ARF, ",then,") > 0 And(InStr(ARF, ",end,if,") > 0 Or InStr(ARF, ",else,") > 0) Then    '将If分割
                    If MHNumb = 0 Then
                        'Call Echo(RowIFList,"")
                        TempStr = Replace(RowIFList, "[云端] [云端]", " ") 
                        Splxx = Split(TempStr, "[云端]") 
                        TempC = "" 
                        For J = 0 To UBound(Splxx) - 1
                            If J = 1 Or J = 3 Then
                                TempC = TempC & CopyStrNumb(RSpaceStr, RSpaceNumb + 1) & Splxx(J) & vbCrlf 
                            Else
                                TempC = TempC & CopyStrNumb(RSpaceStr, RSpaceNumb) & Splxx(J) & vbCrlf 
                            End If 
                        Next 
                        TempC = Left(TempC, Len(TempC) - 2) & " If" 
                        'Call Echo(RowIFList,TempC)                         
                    'Sc = TempC             '出错了，以后再改
                    End If 
                'Else中间
                ElseIf(InStr(ARF, ",else,") > 0 Or InStr(ARF, ",elseif,") > 0) And InStr(ARF, ",case,else,") = False Then
                    Call RowLocation(Sc, "IF", "ElseIf", YesOK, "else") 
                'IF开始
                ElseIf InStr(ARF, ",if,") > 0 And InStr(ARF, ",then,") > 0 And Right(LCase(Replace(Replace(Trim(CodeS), "%" & ">", ""), " ", "")), 4) = "then" Then    '" then"前面有空格会出错
                    ArrStr(RSpaceNumb) = "IF"    '把信息写入数组 
                    Sc = GetNoteC(NoteC, "") & CopyStrNumb(RSpaceStr, RSpaceNumb) & Sc 
                    RSpaceNumb = RSpaceNumb + 1 : YesOK = False 
                'IF结束
                ElseIf InStr(ARF, ",end,if,") > 0 Then    '获得IF最后那个数
                    Call RowLocation(Sc, "IF", "End IF", YesOK, True) 
                'For 开始
                ElseIf InStr(ARF, ",for,") > 0 And InStr(ARF, ",exit,for,") = False Then
                    ArrStr(RSpaceNumb) = "For"    '把For信息写入数组 
                    Sc = GetNoteC(NoteC, "") & CopyStrNumb(RSpaceStr, RSpaceNumb) & Sc 
                    RSpaceNumb = RSpaceNumb + 1 : YesOK = False 
                'For结束
                ElseIf InStr(ARF, ",next,") > 0 And InStr(ARF, ",on,resume,next,") = False And InStr(ARF, ",for,next,") = False Then
                    Call RowLocation(Sc, "For", "Next", YesOK, True) 
                'Do开始
                ElseIf InStr(ARF, ",do,while,") > 0 Or InStr(ARF, ",do,until,") > 0 Or ARF = ",do," Then
                    ArrStr(RSpaceNumb) = "Do"    '把For信息写入数组 
                    Sc = GetNoteC(NoteC, "") & CopyStrNumb(RSpaceStr, RSpaceNumb) & Sc 
                    RSpaceNumb = RSpaceNumb + 1 : YesOK = False 
                'Do 结束
                ElseIf InStr(ARF, ",loop,") > 0 Then
                    Call RowLocation(Sc, "Do", "Loop", YesOK, True) 
                'While开始
                ElseIf InStr(ARF, ",while,") > 0 Then
                    ArrStr(RSpaceNumb) = "While"    '把For信息写入数组 
                    Sc = GetNoteC(NoteC, "") & CopyStrNumb(RSpaceStr, RSpaceNumb) & Sc 
                    RSpaceNumb = RSpaceNumb + 1 : YesOK = False 
                'While 结束
                ElseIf InStr(ARF, ",wend,") > 0 Then
                    Call RowLocation(Sc, "While", "Wend", YesOK, True) 
                'Set结束
                ElseIf InStr(ARF, ",set,nothing,") > 0 Then
                    Call RowLocation(Sc, "Set", "Set Nothing", YesOK, True) 
                'Set开始
                ElseIf InStr(ARF, "set,") > 0 Then    '在没有:符号 不可对连接代码起作用
                    ArrStr(RSpaceNumb) = "Set"    '把Set信息写入数组 
                    Sc = GetNoteC(NoteC, "") & CopyStrNumb(RSpaceStr, RSpaceNumb) & Sc 
                    RSpaceNumb = RSpaceNumb + 1 : YesOK = False 
                'With结束
                ElseIf InStr(ARF, ",end,with,") > 0 Then
                    Call RowLocation(Sc, "With", "End Width", YesOK, True) 
                'With开始
                ElseIf InStr(ARF, "with,") > 0 Then    '在没有:符号 不可对连接代码起作用
                    ArrStr(RSpaceNumb) = "With"    '把Set信息写入数组 
                    Sc = GetNoteC(NoteC, "") & CopyStrNumb(RSpaceStr, RSpaceNumb) & Sc 
                    RSpaceNumb = RSpaceNumb + 1 : YesOK = False 
                'Case 开始
                ElseIf InStr(ARF, ",select,case,") > 0 Then
                    ArrStr(RSpaceNumb) = "Case"    '把Set信息写入数组 
                    Sc = GetNoteC(NoteC, "") & CopyStrNumb(RSpaceStr, RSpaceNumb) & Sc 
                    RSpaceNumb = RSpaceNumb + 2 : CaseNumb = 1 : YesOK = False 
                'Case 中间
                ElseIf InStr(ARF, ",case,") > 0 Then
                    Call RowLocation(Sc, "Case", "Case", YesOK, "else") 
                'Case 结束
                ElseIf InStr(ARF, ",end,select,") > 0 Then
                    CaseNumb = 0 
                    Call RowLocation(Sc, "Case", "End Select", YesOK, True) 
                '最后为ASP结尾不用退格
                ElseIf Sc = "%" & ">" Then
                    RSpaceNumb = 0 
                End If 
                '最后给此行加空格，为了完美
                If YesOK = True Then
                    TempNote = NoteC : NoteC = "" 
                    If Left(S, 1) = "'" And RSpaceNumb > 0 Then
                        NoteC = Sc : Sc = "" 
                    End If 
                    If TempNote <> "" Then
                        TempStr = Sc 
                        TempStr = GetNoteC(TempNote, "") 
                        If Sc = "" Then    '当Sc为空时去掉TempStr尾部的vbCrlf
                            TempStr = Left(TempStr, Len(TempStr) - 2) 
                        End If 
                        Sc = TempStr & CopyStrNumb(RSpaceStr, RSpaceNumb) & Sc    '拷贝多次空格 
                    Else
                        If Sc <> "" Then Sc = CopyStrNumb(RSpaceStr, RSpaceNumb) & Sc    '拷贝多次空格 
                    End If 
                    If Trim(Sc) = "&lt;%" Then Sc = Trim(Sc)    '当一行中只有&lt;%时 去掉两边空格% 
                    If Sc <> "" And InStr(AspSetType, "|html=false|") = False Then Sc = Sc & "<span id=""Id" & nRow & """ onClick=""TestInput('Id" & nRow & "');"" style=""color:red;"" title=点击写笔记>&nbsp;</span>" 
                End If 
            End If 
            If InStr(AnQuanType, "|行不退格|") > 0 Then
                Sc = Trim(Sc) 
            End If 
            BR = vbCrlf 
            If InStr(AnQuanType, "|处理成一行|") > 0 Then
                BR = ": " 
            End If 
            '判断是否为原来代码为假
            If InStr(AspSetType, "|code=false|") = False Then
                If Sc <> "" Then
                    If ASPYes = False And Right(Trim(YuanS), 2) <> "%" & ">" Then    '处理   <td>< %=now()% > 要前面的空格
                        S1 = Ltrim(YuanS) 
                        S2 = Left(YuanS, Len(YuanS) - Len(S1)) 
                        Sc = S2 & Trim(Sc) 
                    ElseIf ASPYes = False And Right(Trim(YuanS), 2) = "%" & ">" Then    '处理 当一行中只有 %"> 不要空格
                        Sc = Trim(Sc) 
                    End If 
                    C = C & Sc & BR : nRow = nRow + 1    '行数累加 
                End If 
            Else
                C = C & Sc & BR : nRow = nRow + 1    '行数累加 
            End If 
        Next 
        '最后判断主程序中无意思变量
        If NORootFunDim <> "," Then
            Splxx = Split(NORootFunDim, ",") 
            For Each S1 In Splxx
                If S1 <> "" Then
                    Css = Css & ".call" & S1 & "{background:#333333;color:#666666}" 
                    NoDimNumb = NoDimNumb + 1    '函数中无意义变量数 
                End If 
            Next 
        End If 
        '自定变量里没有的变量列表
        If BadDimList <> "" Then
            Call CreateAddFile("自定变量里没有的变量列表.ini", BadDimList) 
        End If 
        HandleAspCode = C 
    End Function 
    '获得暂存注释代码
    Function GetNoteC(TempNote, SType)
        Dim J, N, ArrNumb 
        If TempNote = "" Then Exit Function 
        ArrNumb = RSpaceNumb    '默认 
        For J = 20 To 0 Step - 1
            If ArrStr(J) <> "" Then
                ArrNumb = J 
                Exit For 
            End If 
        Next 
        N = RSpaceNumb 
        If SType = "else" Then N = N - 1 
        GetNoteC = CopyStrNumb(RSpaceStr, N) & TempNote & vbCrlf 
        TempNote = "" 
    End Function 
    '替换ASP标签颜色
    Function ASPColor(ByVal S, ByVal I, Code, AspSetType, Color)
        Dim C, SplStr, J, Splxx, LeftC, RightC, StartStr, EndStr, StartYes, EndYes, LCaseCode, Sel, DimStr 
        Dim ColorStr, BlueInStrNumb, RedInStrNumb, HelpTip, TempCode, BlueRedType, ValueStr, SplDim 
        TempCode = Code 
        Code = ""    '让上一函数里对应Code值清空 
        LCaseCode = LCase(TempCode) 
        BlueRedType = ""    '为蓝色函数或红色函数，默认为空 
        '为蓝色查找
        BlueInStrNumb = InStr(LCase(BlueStr), "," & LCaseCode & ",") 
        '红色查找
        RedInStrNumb = InStr(LCase(RedStr), "," & LCaseCode & ",") 
        'ASP中回车     '淡蓝色 特殊
        If LCaseCode = "vbcrlf" Then
            ASPColor = "vbCrlf" 
            '判断是否显示颜色
            If InStr(AspSetType, "|color=false|") = False Then
                ASPColor = "<font color=#009999>vbCrlf</font>" 
            End If 
            Exit Function 
        '蓝色
        ElseIf BlueInStrNumb > 0 Then
            If LCaseCode = "dim" Then    '函数里定义的变量
                C = Replace(Mid(S, I + 1), " ", "") & ","    '获得Dim后面的变量 
                If InStr(C, "'") > 0 Then    '注意Dim中有注释，给去掉 添加于2013,10,5
                    C = Left(C, InStr(C, "'") - 1) & "," 
                End If 
                If InStr(C, ":") > 0 Then    '注意Dim中有:分割
                    C = Left(C, InStr(C, ":") - 1) & "," 
                End If 
                If FunName <> "" Then C = ReplaceDim(C) & ","    '替换变量 
                C = MyFunDim(C) & ","    '处理成自己偏爱变量 
                '为最头部
                If FunName = "" Then
                    If RootFunDimStr = "" Then C = "," & C 
                    RootFunDimStr = RootFunDimStr & C    '加入到函数定义列表中 
                    RootFunDim = RootFunDim & C    '加入定义函数中 
                    NORootFunDim = NORootFunDim & C    '为了判断定义的变量是否使用 
                '为函数里面
                Else
                    If FunDimStr = "" Then C = "," & C 
                    FunDimStr = FunDimStr & C    '加入到函数定义列表中 
                    FunDim = FunDim & C    '加入定义函数中 
                    NOFunDim = NOFunDim & C    '为了判断定义的变量是否使用 
                    AnQuanDim = AnQuanDim & C & "," 
                End If 
            ElseIf LCaseCode = "set" Then    '函数里定义的变量
                C = Replace(Mid(S, I + 1), " ", "") & ","    '获得Dim后面的变量 
                '添加Set变量
                If InStr(C, "=") Then
                    C = Left(C, InStr(C, "=") - 1) 
                    'C = ReplaceDim(C) & "," '替换变量
                    C = MyFunDim(C) & ","    '处理成自己偏爱变量 
                End If 
                '为最头部
                If FunName = "" Or InStr(FunDim, "," & C) = False Then    '当函数里没有Set定义的 就让外面的外面定义set去掉 修改于2013,11,17
                    If RootFunDimStr = "" Then C = "," & C 
                    RootFunDimStr = RootFunDimStr & C    '加入到函数定义列表中 
                    RootFunDim = RootFunDim & C    '加入定义函数中 
                    NORootFunDim = Replace(NORootFunDim, "," & MyFunDim(C) & ",", ",")    '去掉NORootFunDim 
                '为函数里面
                Else
                    If FunDimStr = "" Then C = "," & C 
                    FunDimStr = FunDimStr & C    '加入到函数定义列表中 
                    'FunDim = FunDim & C         '加入定义函数中
                    NOFunDim = Replace(NOFunDim, "," & MyFunDim(C) & ",", ",")    '去掉NOFunDim 
                End If 
            ElseIf LCaseCode = "const" Then    '函数里定义的变量
                C = Replace(Mid(S, I + 1), " ", "") & ","    '获得Dim后面的变量 
                If InStr(C, "=") > 0 Then
                    C = Left(C, InStr(C, "=") - 1) 
                    C = Replace(C, " ", "") 
                    'C = ReplaceDim(C) & "," '替换变量
                    C = MyFunDim(C) & ","    '处理成自己偏爱变量 
                    '为最头部
                    If FunName = "" Then
                        If RootFunDimStr = "" Then C = "," & C 
                        RootFunDimStr = RootFunDimStr & C    '加入到函数定义列表中 
                        RootFunDim = RootFunDim & C    '加入定义函数中 
                        NORootFunDim = NORootFunDim & C    '为了判断定义的变量是否使用 
                    '为函数里面
                    Else
                        If FunDimStr = "" Then C = "," & C 
                        FunDimStr = FunDimStr & C    '加入到函数定义列表中 
                        FunDim = FunDim & C    '加入定义函数中 
                        RootFunDim = RootFunDim & C    '为了判断定义的变量是否使用 
                    End If 
                End If 
            ElseIf LCaseCode = "call" Then    '函数里定义的变量
                C = Replace(Mid(S, I + 1), " ", "") & ","    '获得Dim后面的变量 
                If InStr(C, "(") > 0 Then
                    C = Left(C, InStr(C, "(") - 1) 
                    C = Replace(C, " ", "") 
                    'C = ReplaceDim(C) & "," '替换变量
                    C = MyFunDim(C) & ","    '处理成自己偏爱变量 
                    '为头部
                    If RootFunDimStr = "" Then C = "," & C 
                    RootFunDimStr = RootFunDimStr & C    '加入到函数定义列表中 
                    RootFunDim = RootFunDim & C    '加入定义函数中 
                End If 
            '获得定义函数信息
            ElseIf(LCaseCode = "sub" Or LCaseCode = "function") And FunName = "" Then    '函数
                If GetFunNameDim(S, FunName, FunDim) <> "" Then
                    FunName = Replace(Trim(FunName), vbCrlf, "") 
                    If InStr(FunName, " ") Then    '添加于2013,10,23 对函数前private publick处理
                        FunName = Mid(FunName, InStrRev(FunName, " ") + 1) 
                    End If 
                    FunDim = Replace(FunDim, vbCrlf, "") 
                    AnQuanDim = FunDim & ","    '初始化安全变量列表 
                    FunDim = "," & ReplaceDim(FunName & "," & FunDim) & ","    '替换变量 
                    FunDim = "," & MyFunDim(FunDim) & ","    '处理成自己偏爱变量 
                    NOFunDim = ","    '函数中无效变量列表 
                    FunDimStr = FunDim 
                    FunDimNumb = 0 
                End If 
            End If 
            ARF = ARF & LCaseCode & ","    '蓝色ASP行函数累加 
            '判断是否为原来代码
            If InStr(AspSetType, "|code=false|") = False Then
                ASPColor = Mid(BlueStr, BlueInStrNumb + 1, Len(LCaseCode)) 
                TempCode = ASPColor    '让暂存Code为大小写整齐 
            Else
                ASPColor = TempCode 
            End If 
            '判断是否显示颜色
            If InStr(AspSetType, "|color=false|") = False Then
                If Color = "" Then Color = "#0000FF" 
                ASPColor = "<font color=" & Color & ">" & ASPColor & "</font>" 
            End If 
            BlueRedType = "错误 ASP内置函数不能用 " 
        '为红颜色
        ElseIf RedInStrNumb > 0 Then
            HelpTip = ASPHelpTip(TempCode, Mid(S, I + 1))    '找到函数帮助说明 
            If HelpTip <> "" Then HelpTip = "title=""" & HelpTip & """" 
            '判断是否为原来代码
            If InStr(AspSetType, "|code=false|") = False Then
                ASPColor = Mid(RedStr, RedInStrNumb + 1, Len(LCaseCode)) 
                TempCode = ASPColor    '让暂存Code为大小写整齐 
            Else
                ASPColor = TempCode 
            End If 
            '判断是否显示颜色
            If InStr(AspSetType, "|color=false|") = False Then
                If Color = "" Then Color = "#990099" 
                'Call Echo(TempCode,Color)
                ASPColor = "<font color=" & Color & " " & HelpTip & ">" & ASPColor & "</font>" 
            End If 
            BlueRedType = "注意 ASP内置函数 " 
        End If 
        'LCaseCode = LCase(MyFunDim(LCaseCode))    '让函数处理一下        '这里不能处理，因为它会把Set If这种系统变量给处理了 除非在系统判断后用才可以         
        '给变量加JS
        If InStr(AspSetType, "|clickdim=false|") = False Then
            '先从函数自身找起
            If FunDim <> "" Then
                Splxx = Split(FunDim, ",") 
                For J = 0 To UBound(Splxx)
                    DimStr = Splxx(J) 
                    If LCaseCode = LCase(DimStr) Then
                        FunDimNumb = FunDimNumb + 1 
                        Sel = "" 
                        '为第一个变量则可点击
                        If InStr(FunDimStr, "," & DimStr & ",") > 0 Then
                            '背景颜色要小写，否则JS运行出错
                            Sel = " title='" & BlueRedType & "点击让变量突出显示" & J & "-" & FunDimNumb & "_有[]处' onclick=""AddColor('" & FunName & DimStr & "_', '" & LCase(ArrayColor(J)) & "');""" 
                            FunDimStr = Replace(FunDimStr, "," & DimStr & ",", ",")    '去掉，只有第一个可用 
                            If InStr(OKFunDim, "," & DimStr & ",") = False Then
                                OKFunDim = OKFunDim & DimStr & "," 
                            Else
                                NOFunDim = Replace(NOFunDim, "," & DimStr & ",", ",")    '去掉NOFunDim 
                            End If 
                        Else    '处理未被使用变量
                            NOFunDim = Replace(NOFunDim, "," & DimStr & ",", ",")    '去掉NOFunDim 
                        End If 
                        If BlueRedType <> "" Then
                            ValueStr = ASPColor 
                        Else
                            ValueStr = MyFunDim(DimStr) 
                        End If 
                        HelpTip = ASPHelpTip(ValueStr, "")    '找到函数帮助说明 
                        If HelpTip <> "" Then HelpTip = "title=""" & HelpTip & """" 
                        If InStr(AnQuanType, "|函数变量简洁处理|") > 0 Then
                            '安全处理函数
                            If InStr("," & AnQuanDim, "," & ValueStr & ",") > 0 Then
                                SplDim = Split(AnQuanDim, ",") 
                                For I = 0 To UBound(SplDim)
                                    If SplDim(I) = ValueStr Then
                                        If I + 1 > 26 Then    '超出26字母
                                            ValueStr = Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ",(I + 1) / 26, 1) 
                                            ValueStr = ValueStr & Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ",(I + 1) Mod 26, 1) 
                                        Else
                                            ValueStr = Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ", I + 1, 1) 
                                        End If 
                                        Exit For 
                                    End If 
                                Next 
                            End If 
                        End If 
                        '加_是为了区分末尾是数字变量防 aa,aa1,aa2乱掉，因为aa变量会显示出aa1变量
                        ASPColor = "<span name='" & DimStr & "' id='" & FunName & DimStr & "_" & FunDimNumb & "' " & HelpTip & " class='" & FunName & DimStr & "' " & Sel & ">" & ValueStr & "</span>" 
                        Exit Function 
                    End If 
                Next 
            End If 
            '再从函数外变量找
            If RootFunDim <> "" And InStr(AspSetType, "|clickdim=false|") = False Then
                Splxx = Split(RootFunDim, ",") 
                For J = 0 To UBound(Splxx)
                    DimStr = Splxx(J) 
                    If LCaseCode = LCase(DimStr) Then
                        RootFunDimNumb = RootFunDimNumb + 1 
                        Sel = "" 
                        '第一个变量可点击
                        If InStr(RootFunDimStr, "," & DimStr & ",") > 0 Then
                            '背景颜色要小写，否则JS运行出错
                            Sel = " title='" & BlueRedType & "点击让变量突出显示" & J & "-" & RootFunDimNumb & "_有[]处' onclick=""AddColor('" & DimStr & "_', '" & LCase(ArrayColor(J)) & "');""" 
                            RootFunDimStr = Replace(RootFunDimStr, "," & DimStr & ",", ",")    '去掉，只有第一个可用 
                            If InStr(OKRootFunDim, "," & DimStr & ",") = False Then    '主OK变量累加
                                OKRootFunDim = OKRootFunDim & DimStr & "," 
                            Else
                                'OKRootFunDim = Replace(OKRootFunDim, ","& DimStr &",", ",")    '去掉OKFunDim
                                NORootFunDim = Replace("," & NORootFunDim, "," & DimStr & ",", ",")    '去掉NORootFunDim 
                            End If 
                        Else    '处理未被使用变量
                            NORootFunDim = Replace("," & NORootFunDim, "," & DimStr & ",", ",")    '去掉NORootFunDim 
                        End If 
                        If BlueRedType <> "" Then
                            ValueStr = ASPColor 
                        Else
                            ValueStr = MyFunDim(DimStr) 
                        End If 
                        HelpTip = ASPHelpTip(ValueStr, "")    '找到函数帮助说明 
                        If HelpTip <> "" Then HelpTip = "title=""" & HelpTip & """" 
                        '加_是为了区分末尾是数字变量防 aa,aa1,aa2乱掉，因为aa变量会显示出aa1变量
                        ASPColor = "<span name='call" & DimStr & "' id='" & DimStr & "_" & RootFunDimNumb & "' " & HelpTip & " class='call" & DimStr & "' " & Sel & ">" & ValueStr & "</span>" 
                        Exit Function 
                    End If 
                Next 
            End If 
        End If 
        '返回为空 偏好变量处理
        If ASPColor = "" Then
            ASPColor = TempCode 
            '判断是否为原来代码
            If InStr(AspSetType, "|code=false|") = False Then
                ASPColor = MyFunDim(ASPColor) 
            End If 
            ASPColor = ASPColor 
        End If 
        If InStr(AnQuanType, "|函数全部小写|") > 0 Then
            ASPColor = LCase(ASPColor) 
        End If 
    End Function 
    'ASP结尾错位判断
    Sub RowLocation(Sc, ByVal SType, ErrStr, YesOK, ClearYes)
        Dim J, N, S1, ErrList, Title, StrList,Color
        StrList = "|" 
        ArrNumb = -1    '默认 
        For J = 20 To 0 Step - 1
            If ArrStr(J) = SType Then
                ArrNumb = J 
                Exit For 
            ElseIf ArrStr(J) <> "" Then
                StrList = StrList & ArrStr(J) & "|" 
            End If 
        Next 
        N = RSpaceNumb - 1 - CaseNumb 
        If N < 0 Then N = -9 
        If ArrNumb = N Then
            If ClearYes = "else" Then
                Sc = GetNoteC(NoteC, "else") & CopyStrNumb(RSpaceStr, RSpaceNumb - 1) & Sc    '自身往前进一格 
                YesOK = False 
            Else
                If ClearYes = True Then ArrStr(ArrNumb) = ""    '结尾 清空数组 
                RSpaceNumb = RSpaceNumb - 1    '后面的行跟它一样齐 
            End If 
        Else
            S1 = False 
            If ArrNumb >= 0 Then
                '判断前面有IF，就不操作，因为Set可以 StrList为上面累加寻找的定义
                If(InStr(StrList, "|IF|") > 0 Or InStr(StrList, "|Function|") > 0) And ArrStr(ArrNumb) = "Set" Then
                    S1 = "不操作" 
                Else
                    RSpaceNumb = ArrNumb    '退到这个位置 
                    S1 = GetNoteC(NoteC, "") & CopyStrNumb(RSpaceStr, ArrNumb)    '拷贝多次空格 
                    RSpaceNumb = ArrNumb    '让下一行退一格 
                    If ClearYes = True Then ArrStr(ArrNumb) = ""    '结尾 清空数组 
                    YesOK = False    '必需为假，因为S1已经有退格 下面有用到 
                    If ClearYes = "else" Then
                        RSpaceNumb = RSpaceNumb + 1    '让下一行退一格 
                    End If 
                End If 
            End If 
            If S1 <> "不操作" Then
                ErrList = "" 
                If ArrNumb >= 0 Then
                    For J = 10 To ArrNumb + 1 Step - 1
                        If ArrStr(J) <> "" Then
                            ErrList = ErrList & ArrStr(J) & "|" 
                        End If 
                    Next 
                Else    '在没有函数的情况下，让判断正确，添加于2013,10,14
                    S1 = False 
                    StrList = "" 
                End If 
                If S1 <> False Then
                    Color = "#CCCCCC" : Title = "注意：" & ErrStr & "前面有【" & ErrList & "】未关闭" 
                Else
                    Color = "#FFB3B0" : Title = "错误：" & ErrStr & "，未定义" & SType 
                    S1 = ""    '让S1为"" 
                End If 
                If StrList <> "|" And InStr(AspSetType, "|html=false|") = False Then
                    Sc = S1 & "<font style=""border:1px #000000 dotted;background:" & Color & ";"" title=""" & Title & """>" & Sc & "</font>" 
                Else
                    Sc = S1 & Sc 
                End If 
            End If 
        End If 
    End Sub 
    '初始化
    Private Sub Class_Initialize()
        Dim C, SplStr, I 
        Set AspD = CreateObject("Scripting.Dictionary")
            AspD.Add "Abs", "返回数字的绝对值" 
            AspD.Add "Sqr", "返回数值表达式的平方根" 
            AspD.Add "Sgn", "返回表示数字符号的整数" 
            AspD.Add "Rnd", "返回一个随机生成的数字" 
            AspD.Add "Log", "返回指定数值的自然对数" 
            AspD.Add "Exp", "返回e的指定次幂" 
            AspD.Add "Sin", "返回指定弧度的正弦值" 
            AspD.Add "Cos", "返回指定弧度的余弦值" 
            AspD.Add "Tan", "返回指定弧度的正切值" 
            AspD.Add "Atn", "返回指定数值的反正切值" 
            AspD.Add "Int", "返回指定数字的整数部分" 
            AspD.Add "Round", "对参数按指定的位数进行四舍五入" 
            AspD.Add "Fix", "返回指定数值的整数部分" 
            AspD.Add "Now", "返回当前的日期时间" 
            AspD.Add "Date", "返回当前的系统日期" 
            AspD.Add "Time", "返回当前的系统时间" 
            AspD.Add "Year", "从指定的表达式中提取年份" 
            AspD.Add "Month", "从指定的表达式中提取月份" 
            AspD.Add "Day", "从指定表达式中提取日期" 
            AspD.Add "Hour", "从指定的表达式中提取整点值" 
            AspD.Add "Minute", "从指定的表达式中提取分钟数" 
            AspD.Add "Second", "从指定的表达式中提取秒数" 
            AspD.Add "DateAdd", "将指定日期添加上指定时间间隔并返回新值" 
            AspD.Add "DateDiff", "计算两个日期之间的时间间隔" 
            AspD.Add "DatePart", "返回给定日期中的某一部分" 
            AspD.Add "DateSerial", "根据提供的年月日的数值返回一个Date子类型的数据" 
            AspD.Add "DateValue", "将指定的字符串转化为Date子类型的数据" 
            AspD.Add "MonthName", "根据指定的数字返回代表该月份的字符串" 
            AspD.Add "Weekday", "返回代表一周中某一天的整数" 
            AspD.Add "WeekdayName", "根据周次返回代表该周次的字符串" 
            AspD.Add "TimeSerial", "根据提供的时分秒的数值返回一个Date子类型的数据" 
            AspD.Add "TimeValue", "将指定的字符串转化为Date子类型的数据" 
            AspD.Add "CBool", "将表达式转化为Boolean子类型" 
            AspD.Add "CByte", "将表达式转化为Byte子类型" 
            AspD.Add "CDbl", "将表达式转化为Double子类型" 
            AspD.Add "Chr", "将ANSI字符代码（～）转化为对应的字符" 
            AspD.Add "CInt", "将表达式转化为Integer数值子类型" 
            AspD.Add "CLng", "将表达式转化为Long数值子类型" 
            AspD.Add "Cstr", "将表达式转化为String子类型" 
            AspD.Add "CCur", "将表达式转化为Currency子类型" 
            AspD.Add "CDate", "将日期表达式转化为Date子类型" 
            AspD.Add "CSng", "将表达式转化为Single数值子类型" 
            AspD.Add "Asc", "返回与字符串首字母相对应的ANSI字符代码（～）" 
            AspD.Add "Oct", "返回八进制的字符串" 
            AspD.Add "Hex", "返回十六进制的字符串" 
            AspD.Add "FormatNumber", "格式化指定的数值表达式" 
            AspD.Add "FormatDateTime", "格式化指定的日期时间表达式" 
            AspD.Add "FormatCurrency", "格式化指定的货币表达式" 
            AspD.Add "FormatPercent", "格式化某一数值为百分数" 
            AspD.Add "IsArray", "判断某一变量是否为数组变量" 
            AspD.Add "IsDate", "判断某一变量是否为或者可以转化为日期表达式" 
            AspD.Add "IsEmpty", "判断变量是否为空" 
            AspD.Add "IsNull", "判断表达式中是否包含有效数据" 
            AspD.Add "IsNumeric", "判断表达式的值是否为数字" 
            AspD.Add "Is_object_", "判断表达式的值是否为某一对象" 
            AspD.Add "VarType", "返回指定变量的子类型" 
            AspD.Add "TypeName", "返回指定变量子类型的信息" 
            AspD.Add "MsgBox", "弹出含有提示信息的对话框" 
            AspD.Add "InputBox", "弹出可输入数据的对话框" 
            AspD.Add "ScriptEngine", "返回一个代表当前所使用的脚本语言的字符串" 
            AspD.Add "ScriptEngineBuildVersion", "返回所使用的脚本引擎的编译版本号" 
            AspD.Add "ScriptEngineMajorVersion", "返回所使用的脚本引擎的主版本号" 
            AspD.Add "ScriptEngineMinorVersion", "返回所使用脚本引擎的次版本号" 
            AspD.Add "Array", "利用所给的参数返回一个数组" 
            AspD.Add "UBound", "返回数组的最大可用下标" 
            AspD.Add "LBound", "返回数组的最小可用下标" 
            AspD.Add "Filter", "返回一个数组其中包含了满足条件的字符串" 
            AspD.Add "InStr", "返回某字符串在指定字符串中第一次出现的位置" 
            AspD.Add "InStrRev", "返回某字符串在指定字符串中从结尾算起的出现位置" 
            AspD.Add "Join", "将数组中的字符串连接为一个字符串" 
            AspD.Add "LCase", "将字符串中的字母全部转换为小写形式" 
            AspD.Add "Left", "从字符串的左端取指定数目的字符" 
            AspD.Add "Len", "返回字符串的长度或者存储某一变量所需要的字节数" 
            AspD.Add "Ltrim", "返回不带前导空格的字符串" 
            AspD.Add "Replace", "将字符串中指定字符串替换为其他内容" 
            AspD.Add "Right", "从字符串的右侧获取指定数目的字符" 
            AspD.Add "Rtrim", "删除字符串的后续空格" 
            AspD.Add "Mid", "从字符串中指定位置获取指定长度的字符串" 
            AspD.Add "Split", "返回包含指定数目子字符串的一维数组" 
            AspD.Add "StrComp", "对两个指定的字符串进行比较" 
            AspD.Add "StrReverse", "将指定字符串倒序排列并返回重新排列后的字符串" 
            AspD.Add "String", "返回由指定数目的某一字符组成的字符串" 
            AspD.Add "Trim", "删除字符串的前导和后续空格" 
            AspD.Add "Space", "返回由指定数目的空格组成的字符串" 
            AspD.Add "UCase", "将字符串中的字母全部转换为大写形式" 
            AspD.Add "LoadPicture", "返回一个图片对象" 
            AspD.Add "Create_object_", "创建并返回一个对象" 
            AspD.Add "Get_object_", "从指定文件中获得一个Automation对象" 
            AspD.Add "RGB", "返回代表某一颜色的整数" 
            '解释
            AspD.Add "Angle", "翻译：角度,观点" 
            AspHelpTitle = AspD.keys 
            AspHelpContent = AspD.Items 
            '点击变量背景颜色列表
            C = "#990000,#999900,#333366,#663300,#669966,#FF6600,#CC33CC,#993366,#FF0099,#669900,#336699,#99CCCC,#CC3366,#FF9900,#9933FF,#669900,#6699FF,#333366,#99CC66,#996600,#000033,#003300,#330000,#660000,#000099,#330099,#6600FF,#990000,#CC0000,#FF0000,#990066,#FF0066,#CC00CC," 
            SplStr = Split(C, ",") 
            For I = 1 To UBound(SplStr)
                ArrayColor(I) = SplStr(I) 
            Next 
            '蓝色
            BlueStr = ",If,Then,ElseIf,Else,End,Function,Sub,ByVal,,Dim,Nothing,On,Resume,Call,True,False,Not,Is,Or,And,For,Next,Exit,Execute,Randomize,Do,While,Loop,Set,Mod,To,Each,In,Select,Case,Until,With,Wend,Step,Option,Explicit,Class,Private,Public,Get,Let,Const,ByRef,ReDim,Preserve,LANGUAGE,Rem,Execute,," 
            '红色
            RedStr = ",Abs,Sqr,Sgn,Rnd,Log,Exp,Sin,Cos,Tan,Atn,Int,Round,Fix,Now,Date,Time,Year,Month,Day,Hour,Minute,Second,DateAdd,DateDiff,DatePart,DateSerial,DateValue,MonthName,Weekday,WeekdayName,TimeSerial,TimeValue,CBool,CByte,CDbl,Chr,CInt,CLng,Cstr,CCur,CDate,CSng,Asc,Oct,Hex,FormatNumber,FormatDateTime,FormatCurrency,FormatPercent,IsArray,IsDate,IsEmpty,IsNull,IsNumeric,Is_object_,VarType,TypeName,MsgBox,InputBox,ScriptEngine,ScriptEngineBuildVersion,ScriptEngineMajorVersion,ScriptEngineMinorVersion,Array,UBound,LBound,Filter,InStr,InStrRev,Join,LCase,Left,Len,Ltrim,Replace,Right,Rtrim,Mid,Split,StrComp,StrReverse,String,Trim,Space,UCase,LoadPicture,Create_object_,Get_object_,RGB,Request,ServerVariables," 
            RedStr = RedStr & ",Response,Write,MoveNext,Close,Eof,Bof,Server,MapPath,Server,CreateObject,Err,Error,Description,CInt,CStr,String,CDate,Size,Session,Open,AddNew,Update,displayHtml,File,Test,Type,Mode,CharSet,Form,InStrB,URLEncode,Field,RegExp,Number,QueryString,Name,Item,AscB,MidB,LenB,Buffer,Expires,AddHeader,Timer,BinaryWrite,Redirect,Version,Property,ChrB,Cookies,Value,Clear,ContentType,AscW,AppliCation,Lock,UnLock,Contents,CodePage,ScriptTimeOut,Cancel,ExpiresAbsolute,CacheControl,ChrW,Source,Eval,Name,Count,Match,Fields," 
            '自己偏有变量列表
            MyDimList = ",A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,nCount,Code,Id,IP,Str,Length,Txt,Title,Content,Age,Sex,CountNumb,UserName,UserPass,PassWord,UserAdmin,AdminName,AdminPass,UserPassWord,Tel,Fax,Mail,ListCode,Regex,Reg,Re,Values,StrChar,TempChar,Fso,Fs,Fc,Ts,Path,FileName,FolderName,ClsCache,CacheObj,Obj,Conn,Rs,Rss,Rsx,Sql,AddSql,SetCache,CacheName,CacheValue,GetCache,CacheData,Flag,CacheFlag,CacheTime,FilterStr,CheckFile,FilePath,Html,HtmlStr,HtmlName,ServerMap,StrBadChar,ArrBadChar,YY,Mi,StrDateTime,IsHtml,AA,Ret,S1,S2,C1,C2,PosX,PosY,sUserName,sPassWord,MyArray,Page,SiteInfo,PageUrl,MakeName,FName,MakeFolder,Folder,Class_Initialize,Class_Terminate,"
			MyDimList=MyDimList &"MakeFile,HtmlMaker,GetName,ConnStr,UserIp,UserIp2,GetBody,XmlHttp,ResponseBody,ReadyState,SetTimeOuts,SaveImage,ToFile,Action,Act,Step1,Step2,Step3,Step4,Step5,Success,Check_Install,Install,GetRnd,N1,N2,N3,N4,N5,T0,Check_Write,ObjFolder,ObjSubFolders,SubFCount,ObjTest,Ver,IntVal,Color,StrTemp,Arr,Ch,Cnf,Co,SS,Tp,Dire,V1,V2,V3,Ma,FolderPath,db,Eval,Execute,Include,Inc,Did,Sid,Tid,Folders,Files,Fl,File_F,SubFolders,MyFile,F1,F2,F3,F4,F5,RemoveIframe,RemoveScript,FileTxt,File_W,IframeStr,TempStr,TempC,Days,StrLen,StrUrl,TotalNumber,MaxperPage,CurrentPage,Email,Names,NoHtml,New,Text,FindWord,RsClass,StrNoResult,StrWord,ClassId,MM,DD,RiQi,XiaoShi,FenZhong,Miao,Info,Pass,"
			MyDimList=MyDimList &"ServerList,SetDot,Px,Py,Graph,Width,Height,SetDraw,PIndex,PNumber,DotData,Angle,SitePath,SiteCharset,RsId,Rs1,Xx,Add,Edit,Del,AddSub,Sql1,Admin,cName,MaxPage,Cn,En,NewStr,HaveChar,LastChar,DirPath,Strng,Temp,Echo,ToPic,ChkStr,Val,Str0,Str1,Str2,Str3,GetIP,ShowMsg,ShowPage,StrLength,Fso1,Fso2,StrField,KeyWord,Main,SplStr,Splxx,LCaseS,ArrID,DoMain,Url,Url1,StrPage,Host,Iswww,"
			MyDimList=MyDimList & "Pattern,Send,Pos,Dec,Res,LoopCount,TempResult,II,OpenTextFile,ReadLine,FileExists,DeleteFile,CloseDB,ComeUrl,Cook,StyleId,FText,CreateTextFile,StarTime,EndTime,DBPath,BB,File,Stm,Read,Width,Height,LoadFromFile,FoundErr,ErrMsg,Msg,FsoType,Search,QQ,Rw,Echo,Eerr,mCode,OutPut,Margin,Create,Border,OffSet,III,Ex,Pth,NewCnt,AddCode,Stream,ThePath,TheFolder,SysFileList,MMMM,AAA,WS,SI,Timer1,Tmp,Hu,StartN,EndN,GPath,PortNum,FC2,OFile,TFile,TFile2,FSubFolers,RenameFolder,Rs_1,Rs_2,Rs_3,Rsp,Rsl,Web_Name,Web_Url,Web_Image,Web_Title,Web_Keywords,Web_Description,Web_Copyright,Web_Theme,Web_Consult,Web_TopHTML,Web_BottomHTML,Web_Tel,HtmlWrite,ErrNo,,,," 
            '处理输入内容列表
            WriteList = ",Adodb.Connection,Adodb.RecordSet,Scripting.FileSystemObject,ADODB.Stream,"    '不要出现多个||因为对 " " 有问题 
            '替换变量
            ReplaceDimList = ",NumbCount=nCount,OKNumb=nOK,ErrNumb=nErr,StrErr=C,ErrMsg=C,StrTemp=C,fString=C,,,,,,," 
    End Sub
    '获得ASP内置函数帮助说明
    Function ASPHelpTip(ByVal FunName, ByVal Content)
        Dim I, S, C 
        For I = 0 To AspD.Count - 1
            If LCase(AspHelpTitle(I)) = LCase(FunName) Then
                If LCase(FunName) = "chr" Then
                    S = Left(Content, InStr(Content, ")") - 1) 
                    S = Replace(S, "(", "") 
                    C = vbCrlf & "Chr(" & S & ")=" 
                    If CheckNumber(S) = True Then    '判断Chr 里面是否为数字
                        If S = "10" Or S = "13" Then
                            C = C & "(回车键)" 
                        ElseIf S = "32" Then
                            C = C & "(空格)" 
                        ElseIf S = "0" Then
                            C = C & "(中断程序,停止执行)" 
                        Else
                            C = C & "(" & Chr(S) & ")" 
                        End If 
                        C = Replace(C, "'", "&#x27;") 
                    Else
                        C = C & S 
                    End If 
                End If 
                ASPHelpTip = AspHelpContent(I) & C 
                Exit For 
            End If 
        Next 
    End Function 
    '自己偏好变量定义
    Function MyFunDim(ByVal Content)
        Dim SplStr, S, C, Temps, LCaseS, DimInStrNumb, DimInTHNumb, IsOK 
        SplStr = Split(Content, ",") 
        For Each S In SplStr
            S = Trim(S) 
            LCaseS = LCase(S) 
            If S <> "" Then
                '对变量中()处理，
                If InStr(S, "(") > 0 Then
                    S = Mid(S, 1, InStr(S, "(") - 1) 
                End If 
                IsOK = True 
                DimInTHNumb = InStr(LCase(ReplaceDimList), "," & LCaseS & "=")    '替换变量 
                DimInStrNumb = InStr(LCase(MyDimList), "," & LCaseS & ",")    '偏好变量 
                If DimInTHNumb > 0 Then    '替换变量
                    Temps = Mid(ReplaceDimList, DimInTHNumb + 1) 
                    Temps = Mid(Temps, 1, InStr(Temps, ",") - 1) 
                    Temps = Mid(Temps, InStr(Temps, "=") + 1) 
                    If InStr(FunDim & RootFunDim, "," & Temps & ",") > 0 Then
                        S = Temps 
                        IsOK = False 
                    End If 
                End If 
                If IsOK = True Then
                    If DimInStrNumb > 0 Then    '偏好变量
                        S = Mid(MyDimList, DimInStrNumb + 1, Len(LCaseS)) 
                    ElseIf InStr("," & BadDimList, "," & S & ",") = False Then    '这个变量不存在
                        BadDimList = BadDimList & S & "," 
                    End If 
                End If 
                C = C & S & "," 
            End If 
        Next 
        If C <> "" Then C = Left(C, Len(C) - 1) 
        MyFunDim = C 
    End Function 

    '处理ASP中输出的内容完整 (这里面判断有问题 CreateObject)
    Function HandleWrite(Wc)
        Dim S, InStrNumb 
        S = Left(Wc, Len(Wc) - 1) 
        S = Right(S, Len(S) - 1) 
        S = LCase(Trim(S)) 
        InStrNumb = InStr(LCase(WriteList), "," & S & ",") 
        If InStrNumb > 0 Then
            Wc = """" & Mid(WriteList, InStrNumb + 1, Len(S)) & """" 
        End If 
        HandleWrite = Wc 
    End Function 
    '拷贝内容N次
    Function CopyStrNumb(ByVal Content, nCount)
        '判断代码为原来的
        If InStr(AspSetType, "|code=false|") > 0 Then CopyStrNumb = "" : Exit Function 
        Dim I, S 
        If nCount > 0 Then
            S = Content 
            For I = 1 To nCount - 1
                Content = Content & S 
            Next 
        Else
            Content = "" 
        End If 
        CopyStrNumb = Content 
    End Function 
    '从右边开始截取
    Function LeftSpace(ByVal Content, ByVal AddNumb)
        '判断代码为原来的
        If InStr(AspSetType, "|code=false|") > 0 Then LeftSpace = Content : Exit Function 
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
    '处理Eval列表
    Sub HandleEvalList(ByVal YuanS, SType, ThisRow)
        YuanS = Replace(YuanS, "<", "&lt;") 
        EvalList = EvalList & "【" & YuanS & "】" & SType & "|在第[" & ThisRow + 1 & "]行<br>" 
    End Sub 
    '获得ASP参数类型
    Public Property Let SetAspSetType(Str)
        AspSetType = Str 
    End Property 
    '获得ASP参数类型
    Public Property Let SetCountRow(Str)
        CountRow = CInt(Str) 
    End Property 
    '外部自定替换变量列表
    Public Property Let SetReplaceDimList(Str)
        ReplaceDimList = Str 
    End Property 
    '设置安全
    Public Property Let SetAnQuanType(Str)
        AnQuanType = Str 
    End Property 
    '获得Css值
    Public Property Get GetCss
        GetCss = Css 
    End Property 
    '获得总行数和
    Public Property Get GetnRow
        GetnRow = nRow 
    End Property 
    '获得函数中无意义变量数
    Public Property Get GetNoDimNumb
        GetNoDimNumb = NoDimNumb 
    End Property 
    '获得函数中行信息
    Public Property Get GetRowMsg
        GetRowMsg = RowMsg 
    End Property 
    '获得后门列表
    Public Property Get GetEvalList
        GetEvalList = EvalList 
    End Property 
    '析构函数 类终止
    Private Sub Class_Terminate()
        'HtmlFolder=nothing         
        'HtmlFilename=nothing         
        'HtmlContent=nothing         
    'Urlname=nothing
    End Sub 
'VB结束
End Class 
%> 

