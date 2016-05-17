<% 
'动态生静态(2013,12,17)

'================ 快速获得网站操作 ==================
'在线修改 修改单文本
'MainStr = DisplayOnlineED2(WEB_ADMINURL &"MainInfo.Asp?act=ShowEdit&Id=" & TempRs("Id") & "&n=" & GetRnd(11), MainStr, "<li|<a ")
'在线修改 产品大类
'DidStr = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditBigClass&Id=" & TempRs("Id") & "&n=" & GetRnd(11), DidStr, "<li|<a ")
'在线修改 产品小类
'SidStr = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditSmallClass&Id=" & TempRs("Id") & "&n=" & GetRnd(11), SidStr, "<li|<a ")
'在线修改 产品子类
'S = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditThreeClass&Id=" & TempRs("Id") & "&n=" & GetRnd(11), S, "<li|<a ")
'在线修改  文章
'ProStr = DisplayOnlineED2(WEB_ADMINURL &"Product.Asp?act=ShowEditProduct&Id=" & TempRs("Id") & "&n=" & GetRnd(11), ProStr, "<li|<a ")
'在线修改 导航大类
'NavDidStr = DisplayOnlineED2(WEB_ADMINURL &"NavManage.Asp?act=EditNavBig&Id=" & TempRs("Id") & "&n=" & GetRnd(11), NavDidStr, "<li|<a ")
'在线修改 导航小类
'NavSidStr = DisplayOnlineED2(WEB_ADMINURL &"NavManage.Asp?act=EditNavSmall&Id=" & TempRs("Id") & "&n=" & GetRnd(11), NavSidStr, "<li|<a ")

'-------------------------------- 下面为网站后台常用快捷标签代码区 -------------------------------------------

'符加文字颜色
Function infoColor(str, color)
    If color <> "" Then str = "<font color=" & color & ">" & str & "</font>" 
    infoColor = str 
End Function 
'图片加载失败显示默认图片
Function imgError()
    imgError = " onerror=""this.src='/UploadFiles/NoImg.jpg'""" 
End Function 
'获得target样式
Function targetStr(ByVal sType)
    If sType <> "" Then
        targetStr = " target='" & sType & "'" 
    End If 
End Function 
'打开方式  (辅助)
Function aTarget(sType)
    aTarget = targetStr(sType) 
End Function 
'获得链接Title样式
Function aTitle(ByVal title)
    If title <> "" Then
        aTitle = " Title='" & title & "'" 
    End If 
End Function 
'获得链接Title
Function imgAlt(ByVal Alt)
    If Alt <> "" Then
        imgAlt = " alt='" & Alt & "'" 
    End If 
End Function 
'图片标题与Alt
Function imgTitleAlt(ByVal str)
    If str <> "" Then
        imgTitleAlt = " alt='" & str & "' title='" & str & "'" 
    End If 
End Function 
'获得A Rel值
Function aRel(ByVal sType)
    If sType = True Then
        aRel = " rel='nofollow'" 
    End If 
End Function 
'获得target样式
Function styleClass(ByVal ClassName)
    If ClassName <> "" Then
        styleClass = " class='" & ClassName & "'" 
    End If 
End Function 
'文本加粗
Function textFontB(ByVal text, BYes)
    If BYes = True Then
        text = "<strong>" & text & "</strong>" 
    End If 
    textFontB = text 
End Function 
'文本加颜色
Function textFontColor(ByVal text, color)
    If color <> "" Then
        text = "<font color='" & color & "'>" & text & "</font>" 
    End If 
    textFontColor = text 
End Function 
'处理文本颜色与加粗
Function fontColorFontB(title, FontB, FontColor)
    fontColorFontB = textFontColor(textFontB(title, FontB), FontColor) 
End Function 
'获得默认文章信息文件名称
Function getDefaultFileName()
    getDefaultFileName = format_Time(Now(), 6) 
End Function 
'获得链接  例：'"<a " & AHref(Url, TempRs("BigClassName"), TempRs("Target")) & ">" & TempRs("BigClassName") & "</a>"
Function aHref(url, title, target)
    url = handleHttpUrl(url)                                                        '处理一下URL 让之完整
    aHref = "href='" & url & "'" & aTitle(title) & aTarget(target) 
End Function 
'获得图片路径
Function imgSrc(url, title, target)
    url = handleHttpUrl(url)                                                        '处理一下URL 让之完整
    imgSrc = "src='" & url & "'" & aTitle(title) & imgAlt(title) & aTarget(target) 
End Function 

'============== 网站后台使用 ==================

'选择Target打开方式
Function selectTarget(target)
    Dim c, sel 
    c = c & "<select name=""Target"" id=""Target"">" & vbCrLf 
    c = c & "  <option value=''>链接打开方式</option>" & vbCrLf 
    If target = "" Then sel = " selected" Else sel = "" 
    c = c & "  <option" & sel & " value=''>本页打开</option>" & vbCrLf 
    If target = "_blank" Then sel = " selected" Else sel = "" 
    c = c & "  <option value=""_blank""" & sel & ">新页打开</option>" & vbCrLf 
    If target = "Index" Then sel = " selected" Else sel = "" 
    c = c & "  <option value=""Index""" & sel & ">Index页打开</option>" & vbCrLf 
    If target = "Main" Then sel = " selected" Else sel = "" 
    c = c & "  <option value=""Main""" & sel & ">Main页打开</option>" & vbCrLf 
    c = c & "</select>" & vbCrLf 
    selectTarget = c 
End Function 
'选择文本颜色
Function selectFontColor(FontColor)
    Dim c, sel 
    c = c & "  <select name=""FontColor"" id=""FontColor"">" & vbCrLf 
    c = c & "    <option value=''>文本颜色</option>" & vbCrLf 
    If FontColor = "Red" Then sel = " selected" Else sel = "" 
    c = c & "    <option value=""Red"" class=""FontColor_Red""" & sel & ">红色</option>" & vbCrLf 
    If FontColor = "Blue" Then sel = " selected" Else sel = "" 
    c = c & "    <option value=""Blue"" class=""FontColor_Blue""" & sel & ">蓝色</option>" & vbCrLf 
    If FontColor = "Green" Then sel = " selected" Else sel = "" 
    c = c & "    <option value=""Green"" class=""FontColor_Green""" & sel & ">绿色</option>" & vbCrLf 
    If FontColor = "Black" Then sel = " selected" Else sel = "" 
    c = c & "    <option value=""Black"" class=""FontColor_Black""" & sel & ">黑色</option>" & vbCrLf 
    If FontColor = "White" Then sel = " selected" Else sel = "" 
    c = c & "    <option value=""White"" class=""FontColor_White""" & sel & ">白色</option>" & vbCrLf 
    c = c & "  </select>" & vbCrLf 
    selectFontColor = c 
End Function 
'选择男女
Function selectSex(sex)
    Dim c, sel 
    c = c & "  <select name=""FontColor"" id=""FontColor"">" & vbCrLf 
    c = c & "    <option value=""男"">男</option>" & vbCrLf 
    sel = IIf(sex = "女", " selected", "") 
    c = c & "    <option value=""女""" & sel & ">女</option>" & vbCrLf 
    c = c & "  </select>" & vbCrLf 
    selectSex = c 
End Function 
'选择Session或Cookies验证
Function selectSessionCookies(VerificationMode)
    Dim c, sel 
    c = c & "  <select name=""VerificationMode"" id=""VerificationMode"">" & vbCrLf 
    c = c & "    <option value=""1"">Session验证</option>" & vbCrLf 
    sel = IIf(VerificationMode = "0", " selected", "") 
    c = c & "    <option value=""0""" & sel & ">Cookies验证</option>" & vbCrLf 
    c = c & "  </select>" & vbCrLf 
    selectSessionCookies = c 
End Function 
'显示选择分割内容  showSelectList("aa","aa|bb|cc","|","bb")
Function showSelectList(IDName, content, SplType, ThisValue)
    Dim c, sel, splStr, s 
    IDName = Trim(IDName) 
    If SplType = "" Then SplType = "|_-|" 
    If IDName <> "" Then c = c & "  <select name=""" & IDName & """ id=""" & IDName & """>" & vbCrLf 

    splStr = Split(content, SplType) 
    For Each s In splStr
        sel = "" 
        If s = ThisValue Then sel = " selected" 
        c = c & "    <option value=""" & s & """" & sel & ">" & s & "</option>" & vbCrLf 
    Next 
    If IDName <> "" Then c = c & "  </select>" & vbCrLf 
    showSelectList = c 
End Function 

'显示文章展示列表样式 20150114   例 Call Rw(ShowArticleListStyle("下载列表二.html"))
Function showArticleListStyle(ByVal ThisValue)
    showArticleListStyle = handleArticleListStyleOrInfoStyle("文章展示样式", "ArticleListStyle", ThisValue) 
End Function 
'显示文章信息展示样式 20150114   例 Call Rw(ShowArticleInfoStyle("下载列表二.html"))
Function showArticleInfoStyle(ByVal ThisValue)
    showArticleInfoStyle = handleArticleListStyleOrInfoStyle("文章信息展示样式", "ArticleInfoStyle", ThisValue) 
End Function 
'处理文章展示列表样式和文章信息样式
Function handleArticleListStyleOrInfoStyle(folderName, InputName, ByVal ThisValue)
    Dim ResourceDir, content, c, splStr, fileName, sel 
    'ResourceDir = GetWebSkins() & "\Index\"& FolderName &"\"

    ResourceDir = getWebImages() & "\" & folderName & "\" 

    content = getFileFolderList(ResourceDir, True, "html", "名称", "", "", "") 

    ThisValue = LCase(ThisValue)                                                    '转成小写 好对比

    c = c & "  <select name=""" & InputName & """ id=""" & InputName & """>" & vbCrLf 
    c = c & "    <option value=""""></option>" & vbCrLf 
    splStr = Split(content, vbCrLf) 
    For Each fileName In splStr
        If fileName <> "" Then
            sel = IIF(LCase(fileName) = ThisValue, " selected", "") 
            c = c & "    <option value=""" & fileName & """" & sel & ">" & fileName & "</option>" & vbCrLf 
        End If 
    Next 
    c = c & "  </select>" & vbCrLf 

    handleArticleListStyleOrInfoStyle = c 
End Function 

'获得模块皮肤 ShowWebModuleSkins("ModuleSkins", ModuleSkins)
Function showWebModuleSkins(InputName, ByVal ThisValue)
    Dim ResourceDir, content, c, splStr, fileName, sel 
    ResourceDir = getWebSkins() & "\Index\column" 
    'Call Echo("ResourceDir",ResourceDir)
    content = getDirFolderNameList(ResourceDir) 
    'Call Echo("Content",Content)

    ThisValue = LCase(ThisValue)                                                    '转成小写 好对比

    c = c & "  <select name=""" & InputName & """ id=""" & InputName & """>" & vbCrLf 
    c = c & "    <option value=""""></option>" & vbCrLf 
    splStr = Split(content, vbCrLf) 
    For Each fileName In splStr
        If fileName <> "" Then
            sel = IIF(LCase(fileName) = ThisValue, " selected", "") 
            c = c & "    <option value=""" & fileName & """" & sel & ">" & fileName & "</option>" & vbCrLf 
        End If 
    Next 
    c = c & "  </select>" & vbCrLf 

    showWebModuleSkins = c 
End Function 

'显示单选项列表
Function showRadioList(IDName, content, SplType, ThisValue)
    Dim c, sel, splStr, s, i 
    IDName = Trim(IDName) 
    If SplType = "" Then SplType = "|_-|" 
    i = 0 
    splStr = Split(content, SplType) 
    For Each s In splStr
        sel = "" : i = i + 1 
        If s = ThisValue Then sel = " checked" 
        c = c & "<input type=""radio"" name=""" & IDName & """ id=""" & IDName & i & """ value=""radio"" " & sel & "><label for=""" & IDName & i & """>" & s & "</label>" & vbCrLf 
    Next 

    showRadioList = c 
End Function 
'显示Input复选 InputCheckBox("Id",ID,"")
Function inputCheckBox(textName, checked, helpStr)
    'Dim sel
    'If CStr(valueStr) = "True" Or CStr(checked) = "1" Then sel = " checked" Else sel = ""
    'inputCheckBox = "<input type='checkbox' name='" & textName & "' id='" & textName & "'" & sel & " value='1'>"
    'If helpStr <> "" Then inputCheckBox = "<label for='" & textName & "'>" & inputCheckBox & helpStr & "</label> "
    inputCheckBox = handleInputCheckBox(textName, checked, 1, helpStr, "") 
End Function 
'显示Input复选 InputCheckBox("Id",ID,"")
Function inputCheckBox3(textName, checked, valueStr, helpStr)
    inputCheckBox3 = handleInputCheckBox(textName, checked, valueStr, helpStr, "newidname") 
End Function 
Function handleInputCheckBox(textName, checked, valueStr, helpStr, sType)
    Dim s, sel, idName 
    If CStr(valueStr) = "True" Or CStr(checked) = "1" Then sel = " checked" Else sel = "" 
    idName = textName                                                               'id名等于文件名称
    sType = "|" & sType & "|" 
    If InStr(sType, "|newidname|") > 0 Then
        idName = textName & phprand(1, 9999) 
    End If 
    s = "<input type='checkbox' name='" & textName & "' id='" & idName & "'" & sel & " value='" & valueStr & "'>" 
    If helpStr <> "" Then s = "<label for='" & idName & "'>" & s & helpStr & "</label> " 
    handleInputCheckBox = s 
End Function 

'显示Input文本  InputText("FolderName", FolderName, "40px", "帮助文字")
Function inputText(textName, valueStr, width, helpStr)
    Dim Css 

    width = Trim(LCase(width)) 
    If width <> "" Then
        If Right(width, 1) <> "%" And Right(width, 2) <> "px" Then
            width = width & "px" 
        End If 
        Css = " style='width:" & width & ";'" 
    End If 
    inputText = "<input name=""" & textName & """ type=""text"" id=""" & textName & """ value=""" & valueStr & """" & Css & " />" & helpStr 
End Function 
'显示Input文本  InputText("FolderName", FolderName, "40px", "帮助文字")
Function inputText2(textName, valueStr, width, className, helpStr)
    Dim Css 
    If className <> "" Then


        className = " class=""" & className & """" 
    End If 
    width = Trim(LCase(width)) 
    If width <> "" Then
        If Right(width, 1) <> "%" And Right(width, 2) <> "px" Then
            width = width & "px" 
        End If 
        Css = " style='width:" & width & ";'" 
    End If 
    inputText2 = "<input name=""" & textName & """ type=""text"" id=""" & textName & """ value=""" & valueStr & """" & Css & className & " />" & helpStr 
End Function 
'显示Input文本在左边  InputLeftText(TextName, ValueStr, "98%", "")
Function inputLeftText(textName, valueStr, width, helpStr)
    Dim Css 
    width = Trim(LCase(width)) 
    If width <> "" Then
        If Right(width, 1) <> "%" And Right(width, 2) <> "px" Then
            width = width & "px" 
        End If 
        Css = " style='width:" & width & ";'" 
    End If 
    inputLeftText = helpStr & "<input name=""" & textName & """ type=""text"" id=""" & textName & """ value=""" & valueStr & """" & Css & " />" & vbCrLf 
End Function 
'显示Input文本在左边 帮助文字在右边
Function inputLeftTextHelpTextRight(textName, valueStr, width, helpStr)
    Dim Css 
    width = Trim(LCase(width)) 
    If width <> "" Then
        If Right(width, 1) <> "%" And Right(width, 2) <> "px" Then
            width = width & "px" 
        End If 
        Css = " style='width:" & width & ";'" 
    End If 
    inputLeftTextHelpTextRight = "<input name=""" & textName & """ type=""text"" id=""" & textName & """ value=""" & valueStr & """" & Css & " />" & helpStr 
End Function 
'显示Input文本在中边 提示文本在左边
Function inputLeftTextContent(textName, valueStr, width, helpStr)
    inputLeftTextContent = handleInputLeftRightTextContent("左边", textName, valueStr, width, helpStr) 
End Function 
'显示Input文本在中边 提示文本在右边
Function inputRightTextContent(textName, valueStr, width, helpStr)
    inputRightTextContent = handleInputLeftRightTextContent("右边", textName, valueStr, width, helpStr) 
End Function 
'显示Input文本在中边 提示文本在左边 或 提示文本在右边 20150114
Function handleInputLeftRightTextContent(sType, textName, valueStr, width, helpStr)
    Dim Css 
    width = Trim(LCase(width)) 
    If width <> "" Then
        If Right(width, 1) <> "%" And Right(width, 2) <> "px" Then
            width = width & "px" 
        End If 
        Css = " style='width:" & width & ";'" 
    End If 
    If Css = "" Then
        Css = " style='text-align:center;'" 
    Else
        Css = Replace(Css, ";'", ";text-align:center;'") 
    End If 
    handleInputLeftRightTextContent = "<input name=""" & textName & """ type=""text"" id=""" & textName & """ value=""" & valueStr & """" & Css & " />" 

    If sType = "左边" Then
        handleInputLeftRightTextContent = helpStr & handleInputLeftRightTextContent & vbCrLf 
    Else
        handleInputLeftRightTextContent = handleInputLeftRightTextContent & helpStr 
    End If 

End Function 

'显示Input文本在左边密码
Function inputLeftPassText(textName, valueStr, width, helpStr)
    Dim Css 
    width = Trim(LCase(width)) 
    If width <> "" Then
        If Right(width, 1) <> "%" And Right(width, 2) <> "px" Then
            width = width & "px" 
        End If 
        Css = " style='width:" & width & ";'" 
    End If 
    inputLeftPassText = helpStr & "<input name=""" & textName & """ type=""password"" id=""" & textName & """ value=""" & valueStr & """" & Css & " />" & vbCrLf 
End Function 
'显示Input文本在左边密码类型
Function inputLeftPassTextContent(textName, valueStr, width, helpStr)
    Dim Css 
    width = Trim(LCase(width)) 
    If width <> "" Then
        If Right(width, 1) <> "%" And Right(width, 2) <> "px" Then
            width = width & "px" 
        End If 
        Css = " style='width:" & width & ";'" 
    End If 
    If Css = "" Then
        Css = " style='text-align:center;'" 
    Else
        Css = Replace(Css, ";'", ";text-align:center;'") 
    End If 
    inputLeftPassTextContent = helpStr & "<input name=""" & textName & """ type=""password"" id=""" & textName & """ value=""" & valueStr & """" & Css & " />" & vbCrLf 
End Function 
'显示Input隐藏文本
Function inputHiddenText(textName, valueStr)
    inputHiddenText = "<input name=""" & textName & """ type=""hidden"" id=""" & textName & """ value=""" & valueStr & """ />" & vbCrLf 
End Function 
'显示Input文本域 InputTextArea("FindTpl", FindTpl, "60%" , "120px", "")
Function inputTextArea(textName, valueStr, width, height, helpStr)
    Dim Css, HeightStr 
    width = Trim(LCase(width)) 
    If width <> "" Then
        If Right(width, 1) <> "%" And Right(width, 2) <> "px" Then
            width = width & "px" 
        End If 
        Css = " style='width:" & width & ";'" 
    End If 
    If height <> "" Then
        If checkNumber(height) Then                                                     '自动加个px像素
            height = height & "px" 
        End If 
        HeightStr = "height:" & height & ";" 
        If Css <> "" Then
            Css = Replace(Css, ";'", ";" & HeightStr & ";'") 
        Else
            Css = " style='height:" & height & ";'" 
        End If 
    End If 
    Css = Replace(Css, ";;", ";")                                                   '去掉多余的值
    inputTextArea = "<textarea name=""" & textName & """ type=""text"" id=""" & textName & """" & Css & ">" & valueStr & "</textarea>" & helpStr 
End Function 
'显示隐藏Input文本域 InputTextArea("WebDescription", WebDescription, "99%", "100px", "")
Function inputHiddenTextArea(textName, valueStr, width, height, helpStr)
    inputHiddenTextArea = handleInputHiddenTextArea(textName, valueStr, width, height, "", helpStr) 
End Function 
'显示隐藏Input文本域 InputTextArea("WebDescription", WebDescription, "99%", "100px", "")
Function handleInputHiddenTextArea(textName, valueStr, width, height, className, helpStr)
    Dim Css, HeightStr 
    If className <> "" Then
        className = " class=""" & className & """" 
    End If 
    If width <> "" Then Css = " style='width:" & width & ";'" 
    If height <> "" Then
        HeightStr = "height:" & height & ";" 
        If Css <> "" Then
            Css = Replace(Css, ";'", ";" & HeightStr & ";'") 
        Else
            Css = " style='height:" & height & ";display:none;'" 
        End If 
    End If 
    handleInputHiddenTextArea = "<textarea name=""" & textName & """ type=""text"" id=""" & textName & """" & Css & className & ">" & valueStr & "</textarea>" & helpStr 
End Function 
'显示目录列表 以Select方式显示
Function showSelectDirList(folderPath, valueStr)
    Dim splStr, c, fileName, sel 
    splStr = Split(getDirFileSort(folderPath), vbCrLf) 
    For Each fileName In splStr
        If fileName <> "" Then
            sel = IIF(valueStr = fileName, " selected", "") 
            c = c & "<option value=""" & folderPath & fileName & """ " & sel & ">" & fileName & "</option>" & vbCrLf 
        End If 
    Next 
    showSelectDirList = c 
End Function 
'给Input加个Disabled不可操作
Function inputDisabled(ByVal content)
    inputDisabled = Replace(content, "<input ", "<input disabled=""disabled"" ") 
End Function 

'给Input加个rel关系内容
Function inputAddAlt(ByVal content, AltStr)
    Dim SearchStr, replaceStr 
    SearchStr = "<input " 
    replaceStr = SearchStr & "alt=""" & AltStr & """ " 
    If InStr(content, SearchStr) > 0 Then
        content = Replace(content, SearchStr, replaceStr) 
    Else
        SearchStr = "<textarea " 
        replaceStr = SearchStr & "alt=""" & AltStr & """ " 
        If InStr(content, SearchStr) > 0 Then
            content = Replace(content, SearchStr, replaceStr) 
        End If 
    End If 
    inputAddAlt = content 
End Function 



'快速调用设置====================================================

'网站描述
Function webTitle_InputTextArea(WebTitle)
    webTitle_InputTextArea = inputText("WebTitle", WebTitle, "70%", "  多个关键词用-隔开") '不填为网站默认标题
End Function 
'网站关键词
Function webKeywords_InputText(WebKeywords)
    webKeywords_InputText = inputText("WebKeywords", WebKeywords, "70%", " 请以，隔开(中文逗号)") 
End Function 
'网站描述
Function webDescription_InputTextArea(WebDescription)
    webDescription_InputTextArea = inputTextArea("WebDescription", WebDescription, "99%", "100px", "") 
End Function 
'静态文件夹名
Function folderName_InputText(folderName)
    folderName_InputText = inputText("FolderName", folderName, "40%", "") 
End Function 
'静态文件名
Function fileName_InputText(fileName)
    fileName_InputText = inputText("FileName", fileName, "40%", ".html 也可以是网络上的链接地址") 
End Function 
'模板文件名

Function templatePath_InputText(TemplatePath)
    templatePath_InputText = inputText("TemplatePath", TemplatePath, "40%", " 不填为默认") 
End Function 
'获得拼音按钮内容
Function clickPinYinHTMLStr(did)
    clickPinYinHTMLStr = "<a href=""javascript:GetPinYin('FolderName','" & did & "','AjAx.Asp?act=GetPinYin')"" >获得拼音</a>" 
End Function 
'选择文本颜色与文本加粗
Function showFontColorFontB(FontColor, FontB)
    showFontColorFontB = selectFontColor(FontColor) & inputCheckBox("FontB", FontB, "加粗") 
End Function 
'显示文本TEXT排序
Function showSort(sort)
    showSort = inputText("Sort", sort, "30px", "") 
    showSort = Replace(showSort, ";'", ";text-align:center;'") 
End Function 
'网站导航类型顶部底部等
Function showWebNavType(NavTop, NavButtom, NavLeft, NavContent, NavRight, NavOthre)
    Dim c 
    c = c & inputCheckBox("NavTop", NavTop, "顶部导航") 
    c = c & inputCheckBox("NavButtom", NavButtom, "底部导航") 
    c = c & inputCheckBox("NavLeft", NavLeft, "左边导航") 
    c = c & inputCheckBox("NavContent", NavContent, "中间导航") 
    c = c & inputCheckBox("NavRight", NavRight, "右边导航") 
    c = c & inputCheckBox("NavOthre", NavOthre, "其它导航") 
    showWebNavType = c 
End Function 
Function showOnHtml(OnHtml)
    showOnHtml = inputCheckBox("OnHtml", OnHtml, "生成HTML") 
End Function 
Function showThrough(Through)
    showThrough = inputCheckBox("Through", Through, "审核") 
End Function 
Function showRecommend(Recommend)
    showRecommend = inputCheckBox("Recommend", Recommend, "推荐") 
End Function 
'显示开户与关闭图片
Function showOnOffImg(id, Table, fieldName, Recommend, url)
    Dim temp, Img, aUrl 
    If rq("page") <> "" Then temp = "&page=" & rq("page") Else temp = "" 
    If Recommend = True Then
        Img = "<img src=""" & adminDir & "Images/yes.gif"">" 
    Else
        Img = "<img src=""" & adminDir & "Images/webno.gif"">" 
    End If 
    'Call Echo(GetUrl(),""& adminDir &"HandleDatabase.Asp?act=SetTrueFalse&Table=" & Table & "&FieldName=" & FieldName & "&Url=" & Url & "&Id=" & Id)
    aUrl = getUrlAddToParam(getUrl(), "" & adminDir & "HandleDatabase.Asp?act=SetTrueFalse&Table=" & Table & "&FieldName=" & fieldName & "&Url=" & url & "&Id=" & id, "replace") 
    showOnOffImg = "<a href=""" & aUrl & """>" & Img & "</a>" 
    '旧版
'ShowOnOffImg = "<a href="& adminDir &"HandleDatabase.Asp?act=SetTrueFalse&Table=" & Table & "&FieldName=" & FieldName & "&Url=" & Url & "&Id=" & Id & Temp & ">" & Img & "</a>"
End Function 
'显示开户与关闭图片
Function newShowOnOffImg(id, Table, fieldName, Recommend, url)
    Dim temp, Img 
    If rq("page") <> "" Then temp = "&page=" & rq("page") Else temp = "" 
    If Recommend = 1 Then
        Img = "<img src=""/Images/yes.gif"">" 
    Else
        Img = "<img src=""/Images/webno.gif"">" 
    End If 
    newShowOnOffImg = "<a href=/WebAdmin/ZAction.Asp?act=Through&Table=" & Table & "&FieldName=" & fieldName & "&Url=" & url & "&Id=" & id & temp & ">" & Img & "</a>" 
End Function 


'获得控制Css样式 20150128  暂时不用
Function controlDialogCss()
    Dim c 
    c = "<style>" & vbCrLf 
    c = c & "/*控制Css20150128*/" & vbCrLf 
    c = c & ".controlDialog{" & vbCrLf 
    c = c & "    position:relative;" & vbCrLf 
    c = c & "    height:50px;" & vbCrLf 
    c = c & "    width:auto;" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & ".controlDialog .menu{" & vbCrLf 
    c = c & "    position:absolute;" & vbCrLf 
    c = c & "    right:0px;" & vbCrLf 
    c = c & "    top:0px;" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & ".controlDialog .menu a{" & vbCrLf 
    c = c & "    color:#FF0000;" & vbCrLf 
    c = c & "    font-size:14px;" & vbCrLf 
    c = c & "    text-decoration:none;" & vbCrLf 
    c = c & "    background-color:#FFFFFF;" & vbCrLf 
    c = c & "    border:1px solid #003300;" & vbCrLf 
    c = c & "    padding:4px;" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & ".controlDialog .menu a:hover{" & vbCrLf 
    c = c & "    color:#C60000;" & vbCrLf 
    c = c & "    text-decoration:underline;" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "</style>" & vbCrLf 
    controlDialogCss = c 
End Function 


'删除里暂存代码
Function batchDeleteTempStr(content, startStr, endStr)
    Dim i, s 
    For i = 1 To 9
        If InStr(content, startStr) = False Then
            Exit For 
        End If 
        s = getStrCut(content, startStr, endStr, 1) 
        content = Replace(content, s, "") 
    Next 
    batchDeleteTempStr = content 
End Function 
%>    