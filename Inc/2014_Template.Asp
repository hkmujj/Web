<% 
'删除这种多余标签 <R#读出内容BlockName网站公告 start#>  <R#读出内容BlockName网站公告 end#>
'对模板处理



'读模块内容
Function XY_ReadTemplateModule(action)
    Dim moduleId, filePath, c, i 
    Dim sourceList                                                                  '源内容列表 20150109
    Dim replaceList                                                                 '替换内容列表
    Dim splSource, splReplace, sourceStr, replaceStr 
    filePath = RParam(action, "File") 
    moduleId = RParam(action, "ModuleId") 
    sourceList = RParam(action, "SourceList") 
    replaceList = RParam(action, "ReplaceList") 
    'Call Echo(SourceList,ReplaceList)

    If moduleId = "" Then moduleId = RParam(action, "ModuleName")                   '用块名称
    filePath = filePath & ".html" 
    'Call Echo("FilePath",FilePath)
    'Call Echo("ModuleId",ModuleId)
    c = readTemplateModuleStr(filePath, "", moduleId) 
    '加替换于20160331
    If sourceList <> "" And replaceList <> "" Then
        splSource = Split(sourceList, "[Array]") 
        splReplace = Split(replaceList, "[Array]") 
        For i = 0 To UBound(splSource)
            sourceStr = splSource(i) 
            replaceStr = splReplace(i) 
            c = Replace(c, sourceStr, replaceStr) 
        Next 
    End If 
    XY_ReadTemplateModule = c 
End Function 

'读模块内容
Function readTemplateModuleStr(filePath, defaultContent, moduleId)
    Dim startStr, endStr, content 
    startStr = "<!--#Module " & moduleId & " start#-->" 
    endStr = "<!--#Module " & moduleId & " end#-->" 
    'FilePath = ReplaceGlobleLable(FilePath)                '替换全部标签        '添加于2014 12 11

    '文件不存在，则追加模板路径 20150616 给VB软件里用
    If checkFile(filePath) = False Then
        filePath = webTemplate & filePath 
    End If 

    If defaultContent <> "" Then
        content = defaultContent 
    ElseIf checkFile(filePath) = True Then
        content = getFText(filePath) 
    Else
        content = code                                                                  '默认用内容指定内容
    End If 
    'Call Die("显示" & ModuleId & "," & Content)
    'Call Eerr(filepath & checkfile(filepath), Content)
    If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
        readTemplateModuleStr = strCut(content, startStr, endStr, 2) 
    Else
        readTemplateModuleStr = "模块[" & moduleId & "]不存在,路径=" & filePath 
    End If 
End Function 
'找模块对应内容
Function findModuleStr(content, valueStr)
    Dim startStr, endStr, YuanStr, replaceStr, i, Block, BlockFile, action 
    For i = 1 To 9
        startStr = "[$读出内容 " : endStr = "$]" 
        If InStr(valueStr, startStr) > 0 And InStr(valueStr, endStr) > 0 Then
            action = strCut(valueStr, startStr, endStr, 2) 
            Block = RParam(action, "Block") 
            BlockFile = RParam(action, "File") 
            If InStr(vbCrLf & ReadBlockList & vbCrLf, vbCrLf & Block & vbCrLf) = False Then
                ReadBlockList = ReadBlockList & Block & vbCrLf 
            End If 
            '块文件存在 则读出内容
            If BlockFile <> "" Then
                content = getFText(BlockFile) 
            End If 
            YuanStr = startStr & action & endStr 
            replaceStr = "" 

            startStr = "<R#读出内容" & Block & " start#>" : endStr = "<R#读出内容" & Block & " end#>" 
            If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
                replaceStr = strCut(content, startStr, endStr, 2) 
            Else
                startStr = "<!--#读出内容" & Block : endStr = "#-->" 
                If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
                    replaceStr = strCut(content, startStr, endStr, 2) 
                End If 
            End If 
            'Call Eerr(YuanStr,ReplaceStr)
            valueStr = Replace(valueStr, YuanStr, replaceStr) 
        'Call Echo("ValueStr",ValueStr)
        Else
            '没有模块要处理了 则退出
            Exit For 
        End If 
    Next 
    findModuleStr = valueStr 
End Function 

'读出Left模板样式        这里面字符 ' 来回复制会出错，所以我们用 \|*|\ 代替处理
Function XY_ReadColumeSetTitle(action)
    Dim startStr, endStr, Style, title, valueStr, MoreClass, MoreUrl, MoreStr, aStr,c
    action = handleInModule(action, "start") 
    Style = RParam(action, "style") 
    title = RParam(action, "Title") 
    'Call Echo("ContentHeight",ContentHeight)
    'ValueStr = RParam(Action,"value")
    '根据模块找内容
    valueStr = moduleFindContent(action, "value") 
    'Call Eerr("ValueStr",ValueStr)
    valueStr = findModuleStr(code, valueStr)                                        '找模块对应内容

    MoreClass = RParam(action, "MoreClass") 
    MoreUrl = PHPTrim(RParam(action, "MoreUrl")) 
    MoreStr = RParam(action, "MoreStr") 
    valueStr = handleInModule(valueStr, "end") 
    c = readColumeSetTitle(action, Style, title, valueStr) 

    If MoreClass = "" Then MoreClass = "more"                                       'More链接为空 则用默认代替
    'If MoreUrl="" Then MoreUrl="#"                    'More链接网址为空 则用默认#代替
    'More链接样式不能为空，因为没有样式它就不能让More在最近边
    If MoreUrl <> "" And MoreStr <> "" Then
        'AStr = "<a href='"& MoreUrl &"' class='"& MoreClass &"'>"& MoreStr &"</a>"
        aStr = "<a " & AHref(MoreUrl, title, "") & " class='" & MoreClass & "'>" & MoreStr & "</a>" 
        c = Replace(c, "<!--#AMore#-->", aStr) 
    End If 
	XY_ReadColumeSetTitle=c
End Function 

'读栏目并赋标题与内容值
Function readColumeSetTitle(action, id, ColumeTitle, ColumeContent)
    Dim TitleWidth                                                                  '标题宽度
    Dim TitleHeight                                                                 '标题高度
    Dim ContentHeight                                                               '内容高度
    Dim ContentWidth                                                                '内容宽度
    Dim ContentCss 

    TitleWidth = RParam(action, "TitleWidth")                                       '获得标题高度    待应用20150715
    TitleHeight = RParam(action, "TitleHeight")                                     '获得标题宽度
    ContentWidth = RParam(action, "ContentWidth")                                   '获得内容宽度
    ContentHeight = RParam(action, "ContentHeight")                                 '获得内容高度

    '标题宽
    TitleWidth = Trim(TitleWidth) 
    '自动加px单位，不加会无效果 20150115
    If Right(TitleHeight, 1) <> "%" And Right(TitleHeight, 2) <> "px" And TitleHeight <> "" And TitleHeight <> "auto" Then
        TitleHeight = TitleHeight & "px" 
    End If 
    If Right(TitleWidth, 1) <> "%" And Right(TitleWidth, 2) <> "px" And TitleWidth <> "" And TitleWidth <> "auto" Then
        TitleWidth = TitleWidth & "px" 
    End If 
    '内容高
    ContentHeight = Trim(ContentHeight) 
    '自动加px单位，不加会无效果 20150115
    If Right(ContentHeight, 1) <> "%" And Right(ContentHeight, 2) <> "px" And ContentHeight <> "" And ContentHeight <> "auto" Then
        ContentHeight = ContentHeight & "px" 
    End If 
    '内容宽
    ContentWidth = Trim(ContentWidth) 
    '自动加px单位，不加会无效果 20150115
    If Right(ContentWidth, 1) <> "%" And Right(ContentWidth, 2) <> "px" And ContentWidth <> "" And ContentWidth <> "auto" Then
        ContentWidth = ContentWidth & "px" 
    End If 

    If ContentHeight <> "" Then
        ContentCss = "height:" & ContentHeight & ";" 
    End If 
    If ContentWidth <> "" Then
        ContentCss = ContentCss & "width:" & ContentWidth & ";" 
    End If 

    Dim content 
    content = readColumn(id) 
    '标题宽
    If TitleWidth <> "" Then
        content = Replace(content, "<div class=""tvalue"">", "<div class=""tvalue"" style='width:" & TitleWidth & ";'>") 
    End If 
    '内容高
    If ContentCss <> "" Then
        content = Replace(content, "<div class=""ccontent"">", "<div class=""ccontent"" style='" & ContentCss & "'>") 
    End If 
    'call echo(ContentWidth,ContentCss)

    content = Replace(content, "栏目标题", ColumeTitle) 
    content = Replace(content, "栏目内容", ColumeContent) 
    readColumeSetTitle = content 
End Function 

'读栏目模块
Function readColumn(id)
    Dim templateFilePath, startStr, endStr, s 
    'Call Echo("WebTemplate",WebTemplate)
    templateFilePath = webTemplate & "\Template_Left.html" 
    startStr = "/*columnlist" & id & "Start*/" 
    endStr = "/*columnlist" & id & "End*/" 
    s = readTemplateFileModular(templateFilePath, startStr, endStr) 
    If s = "[$NO$]" Then
        s = "Left样式ID[" & id & "]不存在" 
    End If 
    readColumn = s 
End Function 


'读模板素材
Function readTemplateSource(id)
    Dim templateFilePath, startStr, endStr, s 
    templateFilePath = webTemplate & "\TemplateSource.html" 
    startStr = "<!--#sourceHtml" & id & "Start#-->" 
    endStr = "<!--#sourceHtml" & id & "End#-->" 
    s = readTemplateFileModular(templateFilePath, startStr, endStr) 
    If s = "[$NO$]" Then
        s = "模板资源ID[" & id & "]不存在" 
    End If 
    readTemplateSource = s 
End Function 



'读模板文件中某模块
Function readTemplateFileModular(templateFilePath, startStr, endStr)
    Dim content 
    readTemplateFileModular = "" 
    content = getFText(templateFilePath) 
    If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
        readTemplateFileModular = strCut(content, startStr, endStr, 2) 
    Else
        readTemplateFileModular = "[$NO$]" 
    End If 
End Function 

'读文件模板素材
Function readTemplateFileSource(templateFilePath, id)
    Dim startStr, endStr, s, c 
    startStr = "<!--#sourceHtml" & Replace(id, ".html", "") & "Start#-->" 
    endStr = "<!--#sourceHtml" & Replace(id, ".html", "") & "End#-->" 
    s = readTemplateFileModular(templateFilePath, startStr, endStr) 
    If s = "[$NO$]" Then
        '加一个读取本文本里配置列表(20150815)
        c = getStrCut(pubCode, startStr, endStr, 2) 
        If c <> "" Then
            readTemplateFileSource = c 
            'call rwend(c)
            Exit Function 
        End If 
        c = getftext(templateFilePath) 
        '存在 <!--#TemplateSplitStart#-->  就返回当前全部内容
        If InStr(c, "<!--#DialogStart#-->") > 0 Then
            readTemplateFileSource = c 
            Exit Function 
        End If 

        s = "模板资源ID[" & id & "]不存在,路径TemplateFilePath=" & handlePath(templateFilePath) 
    End If 
    readTemplateFileSource = s 
End Function 
'读出文件展示列表资源
Function readArticleListStyleSource(id)
    Dim filePath 
    filePath = getWebImages() & "\文章展示样式\" & id 
    If checkFile(filePath) = False Then
        filePath = webTemplate & "\Resources\" & id 
    End If 
    'call echo(checkfile(filePath),filePath)
    readArticleListStyleSource = readTemplateFileSource(filePath, id) 

End Function 
'读出文件信息列表资源
Function readArticleInfoStyleSource(id)
    Dim filePath 
    filePath = getWebImages() & "\文章信息展示样式\" & id 
    If checkFile(filePath) = False Then
        filePath = webTemplate & "\Resources\" & id 
    End If 
    readArticleInfoStyleSource = readTemplateFileSource(filePath, id) 
End Function 


%>   
