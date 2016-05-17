<% 
'Css

'Css压缩 cssCompression(Content,0)
Function cssCompression(content, Level)
    Level = CStr(Level)                                                             '转成字符好判断
    'Css高级压缩
    If Level = "1" Then
        content = RegExp_Replace(content, "\/\*(.|" & vbCrLf & ")*?\*\/", "") 
        content = RegExp_Replace(content, "\s*([\{\}\:\;\,])\s*", "$1") 
        content = RegExp_Replace(content, "\,[\s\.\#\d]*\{", "{") 
        content = RegExp_Replace(content, ";\s*;", ";") 
    'Css简单压缩
    Else
        If Level >= 2 Then
            content = RegExp_Replace(content, "\/\*(.|" & vbCrLf & ")*?\*\/", "")           '//删除注释
        End If 
        content = RegExp_Replace(content, "\s*([\{\}\:\;\,])\s*", "$1") 
        content = RegExp_Replace(content, "\,[\s\.\#\d]*\{", "{")                  '//容错处理
        content = RegExp_Replace(content, ";\s*;", ";")                                 '//清除连续分号
        content = RegExp_Replace(content, ";\s*}", "}")                                 '//清除末尾分号和大括号
        content = RegExp_Replace(content, "([^\s])\{([^\s])", "$1{$2") 
        content = RegExp_Replace(content, "([^\s])\}([^" & vbCrLf & "]s*)", "$1}" & vbCrLf & "$2") 
 
    End If 
    content = trimVBcrlf(content) 
    cssCompression = content 
End Function

'删除Css里注释
Function deleteCssNote(content)
    deleteCssNote = RegExp_Replace(content, "\/\*(.|" & vbCrLf & ")*?\*\/", "")     '//删除注释
End Function 

'Css格式化  展开CSS
Function unCssCompression(content)
    content = RegExp_Replace(content, "\s*([\{\}\:\;\,])\s*", "$1") 
    content = RegExp_Replace(content, ";\s*;", ";")                                 '//清除连续分号
    content = RegExp_Replace(content, "\,[\s\.\#\d]*{", "{") 
    content = RegExp_Replace(content, "([^\s])\{([^\s])", "$1 {" & vbCrLf & "" & vbTab & "$2") 
    content = RegExp_Replace(content, "([^\s])\}([^" & vbCrLf & "]*)", "$1" & vbCrLf & "}" & vbCrLf & "$2") 
    content = RegExp_Replace(content, "([^\s]);([^\s\}])", "$1;" & vbCrLf & "" & vbTab & "$2") 
    unCssCompression = content 
End Function 

'去掉字符串头尾的连续的回车和空格
Function trimVbCrlf(str)
    trimVbCrlf = PHPRTrim(PHPLTrim(str)) 
End Function 

'PHP里Trim方法
Function PHPTrim(str)
    PHPTrim = trimVbCrlf(str) 
End Function 

'去掉字符串开头的连续的回车和空格
Function PHPLTrim(str)
    Dim pos, isBlankChar 
    pos = 1 
    isBlankChar = True 
    While isBlankChar
        If Mid(str, pos, 1) = " " Or Mid(str, pos, 1) = vbTab Then                      '把vbTab也去掉
            pos = pos + 1 
        ElseIf Mid(str, pos, 2) = vbCrLf Then
            pos = pos + 2 
        Else
            isBlankChar = False 
        End If 
    Wend 
    PHPLTrim = Right(str, Len(str) - pos + 1) 
End Function 

'去掉字符串末尾的连续的回车和空格
Function PHPRTrim(str)
    Dim pos, isBlankChar 
    pos = Len(str) 
    isBlankChar = True 
    While isBlankChar And pos >= 2
        If Mid(str, pos, 1) = " " Or Mid(str, pos, 1) = vbTab Then                      '把vbTab也去掉
            pos = pos - 1 
        ElseIf Mid(str, pos - 1, 2) = vbCrLf Then
            pos = pos - 2 
        Else
            isBlankChar = False 
        End If 
    Wend 
    PHPRTrim = RTrim(Left(str, pos)) 
End Function 


'--------------- 有用 暂时用这文件里 ------------------
'去掉字符串头尾的连续的Tab退格和空格
Function trimVbTab(str)
    trimVbTab = RTrimVBTab(LTrimVbTab(str)) 
End Function 


'去掉字符串开头的连续的Tab退格和空格
Function lTrimVbTab(str)
    Dim pos, isBlankChar 
    pos = 1 
    isBlankChar = True 
    While isBlankChar
        If Mid(str, pos, 1) = " " Then
            pos = pos + 1 
        ElseIf Mid(str, pos, 1) = vbTab Then
            pos = pos + 1 
        Else
            isBlankChar = False 
        End If 
    Wend 
    lTrimVbTab = Right(str, Len(str) - pos + 1) 
End Function 

'去掉字符串末尾的连续的Tab退格和空格
Function rTrimVBTab(str)
    Dim pos, isBlankChar 
    pos = Len(str) 
    isBlankChar = True 
    While isBlankChar And pos >= 2
        If Mid(str, pos, 1) = " " Then
            pos = pos - 1 
        ElseIf Mid(str, pos - 1, 1) = vbTab Then
            pos = pos - 1 
        Else
            isBlankChar = False 
        End If 
    Wend 
    rTrimVBTab = RTrim(Left(str, pos)) 
End Function 


'找Html页中Css Style
Function getHtmlCssStyle(ByVal content)
    Dim splStr, s, StyleYes, StyleStartStr, StyleEndStr, StyleStr, CssStr 
    splStr = Split(content, vbCrLf)                                                 '分割行
    StyleYes = False                                                                'Css样式默认为假
    '循环分行
    For Each s In splStr
        If StyleYes = False Then
            If InStr(LCase(s), "<style") > 0 Then
                StyleStartStr = Mid(s, InStr(LCase(s), "<style")) 
                StyleStartStr = Mid(StyleStartStr, 1, InStr(StyleStartStr, ">")) 
                StyleEndStr = Mid(s, InStr(LCase(s), StyleStartStr) + Len(StyleStartStr)) 
                'HTML中定义的Css在一行
                If InStr(StyleEndStr, "</style>") > 0 Then
                    StyleStr = Mid(StyleEndStr, 1, InStr(StyleEndStr, "</style>") - 1) 
                    CssStr = CssStr & StyleStr & vbCrLf 
                Else
                    CssStr = CssStr & StyleEndStr & vbCrLf 
                    StyleYes = True                                                       '收集CssStyle样式开始
                End If 
                'Call Echo("StyleStartStr",ShowHtml(StyleStartStr))
                'Call Echo("StyleEndStr",ShowHtml(StyleEndStr))
                'Call Echo("StyleStr",ShowHtml(StyleStr))
                'Call Echo("CssStr",ShowHtml(CssStr))
            'Call RwEnd("")
            End If 
        ElseIf StyleYes = True Then
            If InStr(LCase(s), "</style>") > 0 Then
                StyleStr = Mid(s, 1, InStr(LCase(s), "</style>") - 1) 
                CssStr = CssStr & StyleStr & vbCrLf 
                'Call Echo("StyleStr2",ShowHtml(StyleStr))
                'Call Echo("CssStr2",ShowHtml(CssStr))
                'Call RwEnd("")
                StyleYes = False                                                                '收集CssStyle样式结束
            Else
                CssStr = CssStr & s & vbCrLf 
            End If 
        End If 
    Next 
    getHtmlCssStyle = CssStr 
End Function 

'处理成干净的Css内容  CSS格式化
Function handleCleanCss(ByVal content)
    Dim splStr, s, c, AddStrYes, CustomS 
    content = Replace(content, "{", vbCrLf & "{" & vbCrLf) 
    content = Replace(content, "}", vbCrLf & "}" & vbCrLf) 
    content = Replace(content, ";", ";" & vbCrLf) 

    splStr = Split(content, vbCrLf) 
    AddStrYes = False                                                               '追加字符默认为假
    For Each s In splStr
        s = trimVbCrlf(s) 
        CustomS = ""                                                                    '自定义S值
        If s <> "" Then
            If InStr(s, "{") > 0 And InStr(s, "}") = False Then
                AddStrYes = True 
                CustomS = s 
            ElseIf InStr(s, "}") > 0 Then
                AddStrYes = False 
            End If 
            If Left(s, 1) <> "{" Then c = c & vbCrLf 
            If AddStrYes = True Then s = "    " & s 
            If CustomS <> "" Then s = CustomS                                               '自定义值不为空则用自定义内容
            c = c & s 

        End If 
    Next 
    c = trimVbCrlf(c) 
	c=replace(c,"    ;"&vbcrlf,"")			'清除多于的分号
	c=replace(c,";"& vbcrlf &"}", vbcrlf & "}")			'最后一个参数不要分号
    handleCleanCss = c 
End Function 



'移除内容中多除行
Function removeExcessRow(content)
    Dim splStr, s, c, TempS 
    splStr = Split(content, vbCrLf)                                                 '分割行
    For Each s In splStr
        TempS = Replace(Replace(s, " ", ""), vbTab, "") 
        If TempS <> "" Then
            c = c & s & vbCrLf 
        End If 
    Next 
    If c <> "" Then c = Left(c, Len(c) - 2) 
    removeExcessRow = c 
End Function 


'2014 11 30
'向Css里追加样式  a=CssAddToStyle(GetFText("1.html")," .test {color:#FF0f000; font-size:10px; float:left}")
Function cssAddToStyle(content, AddToStyle)
    Dim StyleName, YunStyleStr, ReplaceStyleStr, c 
    If InStr(AddToStyle, "{") > 0 Then
        StyleName = Trim(Mid(AddToStyle, 1, InStr(AddToStyle, "{") - 1)) 
    End If 
    YunStyleStr = FindCssStyle(content, StyleName) 
    ReplaceStyleStr = CssStyleAddToParam(YunStyleStr, AddToStyle)                   'Css样式累加参数
    content = Replace(content, YunStyleStr, ReplaceStyleStr) 
    'C = C & "<hr>Content=" & Content
    cssAddToStyle = content 
    'CssAddToStyle = YunStyleStr
'CssAddToStyle = "StyleName=" & StyleName & "<hr>YunStyleStr=" & YunStyleStr & "<hr>ReplaceStyleStr=" & ReplaceStyleStr
End Function 

'检测Css内容中是否有指定样式
Function checkCssStyle(content, StyleStr)
    Dim StyleName 
    checkCssStyle = True 
    If InStr(StyleStr, "{") > 0 Then
        StyleName = Trim(Mid(StyleStr, 1, InStr(StyleStr, "{") - 1)) 
    End If 
    If StyleName = "" Then
        checkCssStyle = False 
    ElseIf FindCssStyle(content, StyleName) = "" Then
        checkCssStyle = False 
    End If 
End Function 


'Css样式累加参数
Function cssStyleAddToParam(ByVal CssStyleStr, ByVal CssStyleStrTwo)
    Dim splStr, CssStr, s, ParamList, ParamName, CssStyleName 
    CssStyleName = Mid(CssStyleStr, 1, InStr(CssStyleStr, "{")) 
    If InStr(CssStyleStr, "{") > 0 Then
        CssStyleStr = Mid(CssStyleStr, InStr(CssStyleStr, "{") + 1) 
    End If 
    If InStr(CssStyleStr, "}") > 0 Then
        CssStyleStr = Mid(CssStyleStr, 1, InStr(CssStyleStr, "}") - 1) 
    End If 
    If InStr(CssStyleStrTwo, "{") > 0 Then
        CssStyleStrTwo = Mid(CssStyleStrTwo, InStr(CssStyleStrTwo, "{") + 1) 
    End If 
    If InStr(CssStyleStrTwo, "}") > 0 Then
        CssStyleStrTwo = Mid(CssStyleStrTwo, 1, InStr(CssStyleStrTwo, "}") - 1) 
    End If 
    splStr = Split(Replace(CssStyleStr & ";" & CssStyleStrTwo, vbCrLf, ""), ";") 
    For Each s In splStr
        s = Trim(s) 
        If InStr(s, ":") > 0 And s <> "" Then
            ParamName = Trim(Mid(s, 1, InStr(s, ":") - 1)) 
            If InStr("|" & ParamList & "|", "|" & ParamName & "|") = False Then
                ParamList = ParamList & ParamName & "|" 
                'Call Echo("ParamName",ParamName)
                CssStr = CssStr & "    " & s & ";" & vbCrLf 
            End If 
        End If 
    Next 
    If CssStyleName <> "" Then
        CssStr = CssStyleName & vbCrLf & CssStr & "}" 
    End If 
    cssStyleAddToParam = CssStr 
'Call Echo(CssStyleStr,CssStyleStrTwo)
End Function 

'根据Css名称找到对应Css块
Function findCssStyle(ByVal content, ByVal StyleName)
    Dim splStr, s, TempS, FindStyleName 
    'CAll Echo("StyleName",StyleName)
    'CAll Echo("Content",Content)
    StyleName = Trim(StyleName) 
    splStr = Split(content, vbCrLf) 
    For Each s In splStr
        If InStr(s, StyleName) > 0 Then
            FindStyleName = Trim(s) 
            If InStr(FindStyleName, "{") > 0 Then
                FindStyleName = Trim(Mid(FindStyleName, 1, InStr(FindStyleName, "{") - 1)) 
            End If 
            If FindStyleName = StyleName Then
                'Call Eerr( FindStyleName , StyleName)
                If InStr(s, "}") > 0 Then
                    findCssStyle = Mid(s, 1, InStr(s, "}") + 1) 
                    'Call EErr(s,FindCssStyle)
                    Exit Function 
                Else
                    TempS = Mid(content, InStr(content, s & vbCrLf) + 1) 
                    TempS = Mid(TempS, 1, InStr(TempS, "}") + 1) 
                    findCssStyle = TempS 
                    Exit Function 
                End If 
            'Call Eerr("temps",Temps)
            End If 
        'Call Echo(FindStyleName,StyleName)
        End If 
    Next 
End Function 

'================================================
'导入网站前端要用到
'================================================
'处理截取到的Css
Function handleCutCssCode(dirPath, CssStr)
    Dim content, startStr, endStr, splStr, ImageFile, fileName, listStr 
    startStr = "url\(" : endStr = "\)" 
    content = getArray(CssStr, startStr, endStr, False, False) 
    splStr = Split(content, "$Array$") 
    For Each ImageFile In splStr
        If ImageFile <> "" And InStr(ImageFile, ".") > 0 And InStr(vbCrLf & listStr & vbCrLf, vbCrLf & ImageFile & vbCrLf) = False Then'对重复使用的图片处理
            listStr = listStr & ImageFile & vbCrLf 
            fileName = Replace(Replace(Replace(ImageFile, """", ""), "'", ""), "\", "/") 
            If InStr(fileName, "/") > 0 Then
                fileName = Mid(fileName, InStrRev(fileName, "/") + 1) 
            End If 
            CssStr = Replace(CssStr, ImageFile, dirPath & fileName) 
        End If 
    Next 
    handleCutCssCode = CssStr 
End Function 

'处理截取到的HtmlDiv
Function handleCutDivCode(dirPath, DivStr)
    Dim content, startStr, endStr, splStr, ImageFile, ToImageFile, fileName, isHandle 
    startStr = "url\(" : endStr = "\)" 
    content = GetArray(DivStr, startStr, endStr, False, False) 
    splStr = Split(content, "$Array$") 
    For Each ImageFile In splStr

        If ImageFile <> "" And InStr(ImageFile, ".") > 0 And InStr(ImageFile, "{$#") = False Then
            '判断是否有域名 20150202
            If GetWebSite(ImageFile) = "" Then
                fileName = Replace(Replace(Replace(ImageFile, """", ""), "'", ""), "\", "/") 
                If InStr(fileName, "/") > 0 Then
                    fileName = Mid(fileName, InStrRev(fileName, "/") + 1) 
                End If 
                DivStr = Replace(DivStr, ImageFile, dirPath & ImageFile) 
            End If 
        End If 
    Next 
    '图片处理
    'Content = GetIMG(DivStr) & vbCrlf & GetHtmlBackGroundImgList(DivStr)        '再加个Html背景图片
    content = GetImgJsUrl(DivStr, "不重复") & vbCrLf & GetHtmlBackGroundImgList(DivStr) '再加个Html背景图片  加强版20150126
    splStr = Split(content, vbCrLf) 
    For Each ImageFile In splStr
        If ImageFile <> "" Then                                                         '当链接地址当前为HTTP:时则不处理20150313
            isHandle = False 

            If Left(ImageFile, 1) = "\" Then
            '等处理20150817
            ElseIf InStr(ImageFile, ".") > 0 And Left(ImageFile, 5) <> "HTTP:" And InStr(ImageFile, "{$#") = False Then
                isHandle = True 
            End If 
            If isHandle = True Then
                ToImageFile = dirPath & RemoveFileDir(ImageFile)                                '移除文件路径目录
                'html中图片路径替换
                DivStr = Replace(DivStr, """" & ImageFile & """", """" & ToImageFile & """") 
                DivStr = Replace(DivStr, "'" & ImageFile & "'", """" & ToImageFile & """") 
                DivStr = Replace(DivStr, "=" & ImageFile & " ", """" & ToImageFile & """") 
                DivStr = Replace(DivStr, "=" & ImageFile & ">", """" & ToImageFile & """") 
            End If 
        End If 
    Next 
    handleCutDivCode = DivStr 
End Function 

'获得HTMl里背景图片 20150116  如： <td width="980" height="169" background="kslx3bg.jpg">
Function getHtmlBackGroundImgList(ByVal content)
    content = GetArray(content, " background=""", """", False, False) 
    content = Replace(content, "$Array$", vbCrLf) 
    getHtmlBackGroundImgList = content 
End Function 


'处理网站HTML中Css链接    写得不是特别的完善好  Content = HandleWebHtmlImg("/aa/bb/",Content)  外部调用
Function getHandleWebHtmlLink(RootPath, content)
    Dim startStr, endStr, ImgList, splStr, c, CssUrl, NewCssUrl, CssStr 
    startStr = "<link " 
    CssStr = "" 
    endStr = ">" 
    ImgList = GetArray(content, startStr, endStr, False, False) 
    'Call RwEnd(ImgList)
    splStr = Split(ImgList, "$Array$") 
    For Each CssUrl In splStr
        If CssUrl <> "" And InStr(LCase(CssUrl), "stylesheet") > 0 Then
            '获得Css加强版，改于20141125
            CssUrl = LCase(Replace(Replace(Replace(CssUrl, """", ""), "'", ""), ">", " ")) & " " 
            startStr = "href=" : endStr = " " 
            If InStr(CssUrl, startStr) > 0 And InStr(CssUrl, endStr) > 0 Then
                CssUrl = StrCut(CssUrl, startStr, endStr, 2) 
            End If 
            NewCssUrl = HandleHttpUrl(CssUrl) 
            If InStr(NewCssUrl, "/") > 0 Then
                NewCssUrl = Mid(NewCssUrl, InStrRev(NewCssUrl, "/") + 1) 
            End If 
            If LCase(NewCssUrl) <> "common.css" And LCase(NewCssUrl) <> "public.css" Then
                NewCssUrl = RootPath & NewCssUrl 
                CssStr = CssStr & "<link href=""" & NewCssUrl & """ rel=""stylesheet"" type=""text/css"" />" & vbCrLf 
            End If 
        End If 
    Next 
    If CssStr <> "" Then CssStr = Left(CssStr, Len(CssStr) - 2) 
    getHandleWebHtmlLink = CssStr 
End Function 


'获得css链接地址列表(20150824)
Function getCssListUrlList(content)
    Dim startStr, endStr, ImgList, splStr, c, CssUrl, CssStr, urlList 
    startStr = "<link " 
    CssStr = "" 
    endStr = ">" 
    ImgList = GetArray(content, startStr, endStr, False, False) 
    'Call RwEnd(ImgList)
    splStr = Split(ImgList, "$Array$") 
    For Each CssUrl In splStr
        If CssUrl <> "" And InStr(LCase(CssUrl), "stylesheet") > 0 Then
            '获得Css加强版，改于20141125
            CssUrl = LCase(Replace(Replace(Replace(CssUrl, """", ""), "'", ""), ">", " ")) & " " 
            startStr = "href=" : endStr = " " 
            If InStr(CssUrl, startStr) > 0 And InStr(CssUrl, endStr) > 0 Then
                CssUrl = StrCut(CssUrl, startStr, endStr, 2) 
            End If 
            If InStr(vbCrLf & urlList & vbCrLf, vbCrLf & CssUrl & vbCrLf) = False Then
                If urlList <> "" Then urlList = urlList & vbCrLf 
                urlList = urlList & CssUrl & vbCrLf 
            End If 
        End If 
    Next 
    getCssListUrlList = urlList 
End Function 

'读Css文件内容并处理(20150824) 如call rwend(handleReadCssContent("E:\E盘\WEB网站\至前网站\DataDir\VB模块\服务器\Template\模块功能列表\B站页面设计\home\home.css","aa",true))
Function handleReadCssContent(cssFilePath, LabelName, isHandleCss)
    Dim c, startStr, endStr 
    c = getFText(cssFilePath) 
    '截取CSS
    startStr = "/*CssCodeStart*/" 
    endStr = "/*CssCodeEnd*/" 
    If InStr(c, startStr) > 0 And InStr(c, endStr) > 0 Then
        c = StrCut(c, startStr, endStr, 2) 
    End If 
    '处理CSS
    If isHandleCss = True Then
        c = cssCompression(c, 0) 
    End If 
    If LabelName <> "" Then
        c = "/*" & LabelName & " start*/" & c & "/*" & LabelName & " end*/" 
    End If 
    handleReadCssContent = c 
End Function 



'处理Css样式里PX或T
Function handleCssPX(ByVal nValue)
    nValue = LCase(Trim(nValue)) 
    If Right(nValue, 1) <> "%" And Right(nValue, 2) <> "px" Then
        nValue = nValue & "px" 
    End If 
    handleCssPX = nValue 
End Function 

'call rw(handleHtmlStyle(getftext("1.html")))
'压缩html里的style样式 (20151008)
Function handleHtmlStyle(content)
    Dim serchS, replaceS, nLength 
    serchS = content 
    nLength = InStr(LCase(serchS), "</style>") + 7 
    serchS = Mid(serchS, 1, nLength) 

    nLength = InStrRev(LCase(serchS), "<style") 
    If nLength > 0 Then
        serchS = Mid(serchS, nLength) 
    End If 
    replaceS = serchS 
    replaceS = cssCompression(replaceS, 0) & vbCrLf                                 '格式化CSS
    replaceS = removeBlankLines(replaceS) 

    content = Replace(content, serchS, replaceS) 
    handleHtmlStyle = content 
End Function 

%>  

