<!--#Include File = "Inc/_Config.Asp"-->          
<!--#Include File = "web/function.asp"-->    
<% 
'asp������

Call openconn() 
'=========

Dim code                                                                        'html����
Dim templateName                                                                'ģ������
Dim cfg_webSiteUrl, cfg_webTemplate, cfg_webImages, cfg_webCss, cfg_webJs, cfg_webTitle, cfg_webKeywords, cfg_webDescription, cfg_webSiteBottom, cfg_flags 
Dim glb_columnName, glb_columnId, glb_id, glb_columnType, glb_columnENType, glb_table, glb_detailTitle, glb_flags 
Dim webTemplate                                                                 '��վģ��·��
Dim glb_url, glb_filePath                                                       '��ǰ������ַ,���ļ�·��
Dim glb_isonhtml                                                                '�Ƿ����ɾ�̬��ҳ
Dim glb_locationType                                                            'λ������

Dim glb_bodyContent                                                             '��������
Dim glb_artitleAuthor                                                           '��������
Dim glb_artitleAdddatetime                                                      '�������ʱ��
Dim glb_upArticle                                                               '��һƪ����
Dim glb_downArticle                                                             '��һƪ����
Dim glb_aritcleRelatedTags                                                      '���±�ǩ��
Dim glb_aritcleSmallImage, glb_aritcleBigImage                                  '����Сͼ�����´�ͼ
Dim glb_searchKeyWord                                                           '�����ؼ���
dim cacheHtmlFilePath															'����html�ļ�·��

Dim isMakeHtml                                                                  '�Ƿ�������ҳ
'������   ReplaceValueParamΪ�����ַ���ʾ��ʽ
Function handleAction(content)
    Dim startStr, endStr, ActionList, splStr, action, s, HandYes 
    startStr = "{\$" : endStr = "\$}" 
    ActionList = getArray(content, startStr, endStr, True, True) 
    'Call echo("ActionList ", ActionList)
    splStr = Split(ActionList, "$Array$") 
    For Each s In splStr
        action = Trim(s) 
        action = handleInModule(action, "start")                                        '����\'�滻��
        If action <> "" Then
            action = Trim(Mid(action, 3, Len(action) - 4)) & " " 
            'call echo("s",s)
            HandYes = True                                                                  '����Ϊ��
            '{VB #} �����Ƿ���ͼƬ·���Ŀ����Ϊ����VB�ﲻ�������·��
            If checkFunValue(action, "# ") = True Then
                action = "" 
            '����
            ElseIf checkFunValue(action, "GetLableValue ") = True Then
                action = XY_getLableValue(action) 
            '�����������������б�
            ElseIf checkFunValue(action, "TitleInSearchEngineList ") = True Then
                action = XY_TitleInSearchEngineList(action) 

            '�����ļ�
            ElseIf checkFunValue(action, "Include ") = True Then
                action = XY_Include(action) 
            '��Ŀ�б�
            ElseIf checkFunValue(action, "ColumnList ") = True Then
                action = XY_AP_ColumnList(action) 
            '�����б�
            ElseIf checkFunValue(action, "ArticleList ") = True Or checkFunValue(action, "CustomInfoList ") = True Then
                action = XY_AP_ArticleList(action) 
            '�����б�
            ElseIf checkFunValue(action, "CommentList ") = True Then
                action = XY_AP_CommentList(action) 
            '����ͳ���б�
            ElseIf checkFunValue(action, "SearchStatList ") = True Then
                action = XY_AP_SearchStatList(action) 
            '���������б�
            ElseIf checkFunValue(action, "Links ") = True Then
                action = XY_AP_Links(action)

            '��ʾ��ҳ����
            ElseIf checkFunValue(action, "GetOnePageBody ") = True Or checkFunValue(action, "MainInfo ") = True Then
                action = XY_AP_GetOnePageBody(action) 
            '��ʾ��������
            ElseIf checkFunValue(action, "GetArticleBody ") = True Then
                action = XY_AP_GetArticleBody(action) 
            '��ʾ��Ŀ����
            ElseIf checkFunValue(action, "GetColumnBody ") = True Then
                action = XY_AP_GetColumnBody(action) 

            '�����ĿURL
            ElseIf checkFunValue(action, "GetColumnUrl ") = True Then
                action = XY_GetColumnUrl(action) 
            '�������URL
            ElseIf checkFunValue(action, "GetArticleUrl ") = True Then
                action = XY_GetArticleUrl(action) 
            '��õ�ҳURL
            ElseIf checkFunValue(action, "GetOnePageUrl ") = True Then
                action = XY_GetOnePageUrl(action) 

            '��ʾ������ ���ò���
            ElseIf checkFunValue(action, "DisplayWrap ") = True Then
                action = XY_DisplayWrap(action) 
            '��ʾ����
            ElseIf checkFunValue(action, "Layout ") = True Then
                action = XY_Layout(action) 
            '��ʾģ��
            ElseIf checkFunValue(action, "Module ") = True Then
                action = XY_Module(action) 
            '�������ģ�� 20150108
            ElseIf checkFunValue(action, "GetContentModule ") = True Then
                action = XY_ReadTemplateModule(action) 
            '��ģ����ʽ�����ñ���������   ������и���ĿStyle��������
            ElseIf checkFunValue(action, "ReadColumeSetTitle ") = True Then
                action = XY_ReadColumeSetTitle(action) 

            '��ʾJS��ȾASP/PHP/VB�ȳ���ı༭��
            ElseIf checkFunValue(action, "displayEditor ") = True Then
                action = displayEditor(action) 

            'Js����վͳ��
            ElseIf checkFunValue(action, "JsWebStat ") = True Then
                action = XY_JsWebStat(action) 

                '------------------- ������ -----------------------
            '��ͨ����A
            ElseIf checkFunValue(action, "HrefA ") = True Then
                action = XY_HrefA(action) 

            '��Ŀ�˵�(���ú�̨��Ŀ����)
            ElseIf checkFunValue(action, "ColumnMenu ") = True Then
                action = XY_AP_ColumnMenu(action) 

            '��վ�ײ�
            ElseIf checkFunValue(action, "WebSiteBottom ") = True Then
                action = XY_AP_WebSiteBottom(action) 
            '��ʾ��վ��Ŀ 20160331
            ElseIf checkFunValue(action, "DisplayWebColumn ") = True Then
                action = XY_DisplayWebColumn(action) 
            'URL����
            ElseIf checkFunValue(action, "escape ") = True Then
                action = XY_escape(action) 
            'URL����
            ElseIf checkFunValue(action, "unescape ") = True Then
                action = XY_unescape(action) 
            'asp��php�汾
            ElseIf checkFunValue(action, "EDITORTYPE ") = True Then
                action = XY_EDITORTYPE(action)


            '��ʱ������
            ElseIf checkFunValue(action, "copyTemplateMaterial ") = True Then
                action = "" 
            ElseIf checkFunValue(action, "clearCache ") = True Then
                action = "" 
            Else
                HandYes = False                                                                 '����Ϊ��
            End If 
            'ע���������е�����ʾ �� And IsNul(action)=False
            If isNul(action) = True Then action = "" 
            If HandYes = True Then
                content = Replace(content, s, action) 
            End If 
        End If 
    Next 
    handleAction = content 
End Function 

'��ʾ��վ��Ŀ �°� ��֮ǰ��վ ��������Ľ�������
Function XY_DisplayWebColumn(action)
    Dim i, c, s, url, sql, dropDownMenu, focusType, addSql 
    Dim isConcise                                                                   '�����ʾ20150212
    Dim styleId, styleValue                                                         '��ʽID����ʽ����
    Dim cssNameAddId 
    Dim shopnavidwrap                                                               '�Ƿ���ʾ��ĿID��

    styleId = PHPTrim(RParam(action, "styleID"))
    styleValue = PHPTrim(RParam(action, "styleValue"))
    addSql = PHPTrim(RParam(action, "addSql"))
    shopnavidwrap = PHPTrim(RParam(action, "shopnavidwrap"))
    'If styleId <> "" Then
    'Call ReadNavCSS(styleId, styleValue)
    'End If

    'Ϊ�������� ���Զ���ȡ��ʽ����  20150615
    If checkStrIsNumberType(styleValue) Then
        cssNameAddId = "_" & styleValue                                                 'Css����׷��Id���
    End If 
    sql = "select * from " & db_PREFIX & "webcolumn" 
    '׷��sql
    If addSql <> "" Then
        sql = getWhereAnd(sql, addSql) 
    End If 
    If checkSql(sql) = False Then Call eerr("Sql", sql) 
    rs.Open sql, conn, 1, 1 
    dropDownMenu = LCase(RParam(action, "DropDownMenu")) 
    focusType = LCase(RParam(action, "FocusType")) 
    isConcise = LCase(RParam(action, "isConcise")) 
    If isConcise = "true" Then
        isConcise = False 
    Else
        isConcise = True 
    End If 
    If isConcise = True Then c = c & copyStr(" ", 4) & "<li class=left></li>" & vbCrLf 
    For i = 1 To rs.RecordCount

        '��PHP��$rs=mysql_fetch_array($rsObj);                                            //��PHP�ã���Ϊ�� asptophpת��������
        url = getColumnUrl(rs("columnname"), "name") 
        If rs("columnName") = glb_columnName Then
            If focusType = "a" Then
                s = copyStr(" ", 8) & "<li class=focus><a href=""" & url & """>" & rs("columname") & "</a>" 
            Else
                s = copyStr(" ", 8) & "<li class=focus>" & rs("columnname") 
            End If 
        Else
            s = copyStr(" ", 8) & "<li><a href=""" & url & """>" & rs("columnname") & "</a>" 
        End If 
        url = WEB_ADMINURL & "?act=addEditHandle&actionType=WebColumn&lableTitle=��վ��Ŀ&nPageSize=10&page=&id=" & rs("id") & "&n=" & getRnd(11) 
        s = handleDisplayOnlineEditDialog(url, s, "", "div|li|span")                    '�����Ƿ���������޸Ĺ�����

        c = c & s 

        'С��

        c = c & copyStr(" ", 8) & "</li>" & vbCrLf 

        If isConcise = True Then c = c & copyStr(" ", 8) & "<li class=line></li>" & vbCrLf 
    rs.MoveNext : Next : rs.Close 
    If isConcise = True Then c = c & copyStr(" ", 8) & "<li class=right></li>" & vbCrLf 

    If styleId <> "" Then
        c = "<ul class='nav" & styleId & cssNameAddId & "'>" & vbCrLf & c & vbCrLf & "</ul>" & vbCrLf 
    End If 
    If shopnavidwrap = "1" Or shopnavidwrap = "true" Then
        c = "<div id='nav" & styleId & cssNameAddId & "'>" & vbCrLf & c & vbCrLf & "</div>" & vbCrLf 
    End If 

    XY_DisplayWebColumn = c 
End Function 

'�滻ȫ�ֱ��� {$cfg_websiteurl$}
Function replaceGlobleVariable(ByVal content)
    content = handleRGV(content, "{$cfg_webSiteUrl$}", cfg_webSiteUrl)              '��ַ
    content = handleRGV(content, "{$cfg_webTemplate$}", cfg_webTemplate)            'ģ��
    content = handleRGV(content, "{$cfg_webImages$}", cfg_webImages)                'ͼƬ·��
    content = handleRGV(content, "{$cfg_webCss$}", cfg_webCss)                      'css·��
    content = handleRGV(content, "{$cfg_webJs$}", cfg_webJs)                        'js·��
    content = handleRGV(content, "{$cfg_webTitle$}", cfg_webTitle)                  '��վ����
    content = handleRGV(content, "{$cfg_webKeywords$}", cfg_webKeywords)            '��վ�ؼ���
    content = handleRGV(content, "{$cfg_webDescription$}", cfg_webDescription)      '��վ����
    content = handleRGV(content, "{$cfg_webSiteBottom$}", cfg_webSiteBottom)        '��վ����

    content = handleRGV(content, "{$glb_columnId$}", glb_columnId)                  '��ĿId
    content = handleRGV(content, "{$glb_columnName$}", glb_columnName)              '��Ŀ����
    content = handleRGV(content, "{$glb_columnType$}", glb_columnType)              '��Ŀ����
    content = handleRGV(content, "{$glb_columnENType$}", glb_columnENType)          '��ĿӢ������

    content = handleRGV(content, "{$glb_Table$}", glb_table)                        '��
    content = handleRGV(content, "{$glb_Id$}", glb_id)                              'id


    '���ݾɰ汾 ��������ȥ��
    content = handleRGV(content, "{$WebImages$}", cfg_webImages)                    'ͼƬ·��
    content = handleRGV(content, "{$WebCss$}", cfg_webCss)                          'css·��
    content = handleRGV(content, "{$WebJs$}", cfg_webJs)                            'js·��
    content = handleRGV(content, "{$Web_Title$}", cfg_webTitle) 
    content = handleRGV(content, "{$Web_KeyWords$}", cfg_webKeywords) 
    content = handleRGV(content, "{$Web_Description$}", cfg_webDescription) 


    content = handleRGV(content, "{$EDITORTYPE$}", EDITORTYPE)                      '��׺
    content = handleRGV(content, "{$WEB_VIEWURL$}", WEB_VIEWURL)                    '��ҳ��ʾ��ַ
    '�����õ�
    content = handleRGV(content, "{$glb_artitleAuthor$}", glb_artitleAuthor)        '��������
    content = handleRGV(content, "{$glb_artitleAdddatetime$}", glb_artitleAdddatetime) '�������ʱ��
    content = handleRGV(content, "{$glb_upArticle$}", glb_upArticle)                '��һƪ����
    content = handleRGV(content, "{$glb_downArticle$}", glb_downArticle)            '��һƪ����
    content = handleRGV(content, "{$glb_aritcleRelatedTags$}", glb_aritcleRelatedTags) '���±�ǩ��
    content = handleRGV(content, "{$glb_aritcleBigImage$}", glb_aritcleBigImage)    '���´�ͼ
    content = handleRGV(content, "{$glb_aritcleSmallImage$}", glb_aritcleSmallImage) '����Сͼ
    content = handleRGV(content, "{$glb_searchKeyWord$}", glb_searchKeyWord)        '��ҳ��ʾ��ַ


    replaceGlobleVariable = content 
End Function 

'�����滻
Function handleRGV(ByVal content, findStr, replaceStr)
    Dim lableName 
    '��[$$]����
    lableName = Mid(findStr, 3, Len(findStr) - 4) & " " 
    lableName = Mid(lableName, 1, InStr(lableName, " ") - 1) 
    content = replaceValueParam(content, lableName, replaceStr) 
    content = replaceValueParam(content, LCase(lableName), replaceStr) 
    'ֱ���滻{$$}���ַ�ʽ������֮ǰ��վ
    content = Replace(content, findStr, replaceStr) 
    content = Replace(content, LCase(findStr), replaceStr) 
    handleRGV = content 
End Function 

'������ַ����
Sub loadWebConfig()
    Dim templatedir 
    Call openconn() 
    rs.Open "select * from " & db_PREFIX & "website", conn, 1, 1 
    If Not rs.EOF Then
        cfg_webSiteUrl = phptrim(rs("webSiteUrl"))                                      '��ַ
        cfg_webTemplate = webDir & phptrim(rs("webTemplate"))                           'ģ��·��
        cfg_webImages = webDir & phptrim(rs("webImages"))                               'ͼƬ·��
        cfg_webCss = webDir & phptrim(rs("webCss"))                                     'css·��
        cfg_webJs = webDir & phptrim(rs("webJs"))                                       'js·��
        cfg_webTitle = rs("webTitle")                                                   '��ַ����
        cfg_webKeywords = rs("webKeywords")                                             '��վ�ؼ���
        cfg_webDescription = rs("webDescription")                                       '��վ����
        cfg_webSiteBottom = rs("webSiteBottom")                                         '��վ�ص�
        cfg_flags = rs("flags")                                                         '��

        '�Ļ�ģ��20160202
        If Request("templatedir") <> "" Then
			'ɾ������Ŀ¼ǰ���Ŀ¼������Ҫ�Ǹ�����20160414
            templatedir = replace(handlePath(Request("templatedir")),handlePath("/"),"/")
			'call eerr("templatedir",templatedir)
			
            If(InStr(templatedir, ":") > 0 Or InStr(templatedir, "..") > 0) And getIP() <> "127.0.0.1" Then
                Call eerr("��ʾ", "ģ��Ŀ¼�зǷ��ַ�") 
            End If 
            templatedir = handlehttpurl(Replace(templatedir, handlePath("/"), "/")) 

            cfg_webImages = Replace(cfg_webImages, cfg_webTemplate, templatedir) 
            cfg_webCss = Replace(cfg_webCss, cfg_webTemplate, templatedir) 
            cfg_webJs = Replace(cfg_webJs, cfg_webTemplate, templatedir) 
            cfg_webTemplate = templatedir 
        End If 
        webTemplate = cfg_webTemplate 
    End If : rs.Close 
End Sub 

'��վλ�� ������
Function thisPosition(content)
    Dim c 
    c = "<a href=""/"">��ҳ</a>" 
    If glb_columnName <> "" Then
        c = c & " >> <a href=""" & getColumnUrl(glb_columnName, "name") & """>" & glb_columnName & "</a>" 
    End If 
    '20160330
    If glb_locationType = "detail" Then
        c = c & " >> �鿴����" 
    End If 
    'call echo("glb_locationType",glb_locationType)

    content = Replace(content, "[$detailPosition$]", c) 
    content = Replace(content, "[$detailTitle$]", glb_detailTitle) 
    content = Replace(content, "[$detailContent$]", glb_bodyContent) 

    thisPosition = content 
End Function 

'��ʾ�����б�
Function getDetailList(action, content, actionName, lableTitle, ByVal fieldNameList, nPageSize, nPage, addSql)
    Call openconn() 
    Dim defaultList, i, s, c, tableName, j, splxx, sql 
    Dim x, url, nCount 
    Dim pageInfo 

    Dim fieldName                                                                   '�ֶ�����
    Dim splFieldName                                                                '�ָ��ֶ�

    Dim replaceStr                                                                  '�滻�ַ�
    tableName = LCase(actionName)                                                   '������
    Dim listFileName                                                                '�б��ļ�����
    listFileName = RParam(action, "listFileName") 
    Dim abcolorStr                                                                  'A�Ӵֺ���ɫ
    Dim atargetStr                                                                  'A���Ӵ򿪷�ʽ
    Dim atitleStr                                                                   'A���ӵ�title20160407
    Dim anofollowStr                                                                'A���ӵ�nofollow

    Dim id 
    id = rq("id") 
    Call checkIDSQL(Request("id")) 

    If fieldNameList = "*" Then
        fieldNameList = getHandleFieldList(db_PREFIX & tableName, "�ֶ��б�") 
    End If 

    fieldNameList = specialStrReplace(fieldNameList)                                '�����ַ�����
    splFieldName = Split(fieldNameList, ",")                                        '�ֶηָ������


    defaultList = getStrCut(content, "[list]", "[/list]", 2) 
    pageInfo = getStrCut(content, "[page]", "[/page]", 1) 
    If pageInfo <> "" Then
        content = Replace(content, pageInfo, "") 
    End If 

    sql = "select * from " & db_PREFIX & tableName & " " & addSql 
    '���SQL
    If checksql(sql) = False Then
        Call errorLog("������ʾ��<br>sql=" & sql & "<br>") 
        Exit Function 
    End If 
    rs.Open sql, conn, 1, 1 
    nCount = rs.RecordCount 

    'Ϊ��̬��ҳ��ַ
    If isMakeHtml = True Then
        url = "" 
        If Len(listFileName) > 5 Then
            url = Mid(listFileName, 1, Len(listFileName) - 5) & "[id].html" 
            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
        End If 
    Else
        url = getUrlAddToParam(getUrl(), "?page=[id]", "replace") 
    End If 
    content = Replace(content, "[$pageInfo$]", webPageControl(nCount, nPageSize, nPage, url, pageInfo)) 

    If EDITORTYPE = "asp" Then
        x = getRsPageNumber(rs, nCount, nPageSize, nPage)                               '���Rsҳ��                                                  '��¼����
    Else
        If nPage <> "" Then
            nPage = nPage - 1 
        End If 
        sql = "select * from " & db_PREFIX & "" & tableName & " " & addSql & " limit " & nPageSize * nPage & "," & nPageSize 
        rs.Open sql, conn, 1, 1 
        x = rs.RecordCount 
    End If 
    For i = 1 To x
        '��PHP��$rs=mysql_fetch_array($rsObj);                                            //��PHP�ã���Ϊ�� asptophpת��������

        s = defaultList 
        s = Replace(s, "[$id$]", rs("id")) 
        For j = 0 To UBound(splFieldName)
            If splFieldName(j) <> "" Then
                splxx = Split(splFieldName(j) & "|||", "|") 
                fieldName = splxx(0) 
                replaceStr = rs(fieldName) & ""
                s = replaceValueParam(s, fieldName, replaceStr)
            End If 

            If isMakeHtml = True Then
                url = getHandleRsUrl(rs("fileName"), rs("customAUrl"), "/detail/detail" & rs("id")) 
            Else
                url = handleWebUrl("?act=detail&id=" & rs("id")) 
                If rs("customAUrl") <> "" Then
                    url = rs("customAUrl") 					 
                End If 
            End If

            'A���������ɫ
            abcolorStr = "" 
            If InStr(fieldNameList, ",titlecolor,") > 0 Then
                'A������ɫ
                If rs("titlecolor") <> "" Then
                    abcolorStr = "color:" & rs("titlecolor") & ";" 
                End If 
            End If
            If InStr(fieldNameList, ",flags,") > 0 Then
                'A���ӼӴ�
                If InStr(rs("flags"), "|b|") > 0 Then
                    abcolorStr = abcolorStr & "font-weight:bold;" 
                End If 
            End If 
            If abcolorStr <> "" Then
                abcolorStr = " style=""" & abcolorStr & """" 
            End If 

            '�򿪷�ʽ2016
            If InStr(fieldNameList, ",target,") > 0 Then
                atargetStr = IIF(rs("target") <> "", " target=""" & rs("target") & """", "")
            End If 

            'A��title
            If InStr(fieldNameList, ",title,") > 0 Then
                atitleStr = IIF(rs("title") <> "", " title=""" & rs("title") & """", "") 
            End If 

            'A��nofollow
            If InStr(fieldNameList, ",nofollow,") > 0 Then
                anofollowStr = IIF(rs("nofollow") <> 0, " rel=""nofollow""", "") 
            End If 



            s = replaceValueParam(s, "url", url) 
            s = replaceValueParam(s, "abcolor", abcolorStr)                                 'A���Ӽ���ɫ��Ӵ�
            s = replaceValueParam(s, "atitle", atitleStr)                                   'A����title
            s = replaceValueParam(s, "anofollow", anofollowStr)                             'A����nofollow
            s = replaceValueParam(s, "atarget", atargetStr)                                 'A���Ӵ򿪷�ʽ


        Next 
        '�����б�����߱༭
        url = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page=&parentid=&id=" & rs("id") & "&n=" & getRnd(11) 
        s = handleDisplayOnlineEditDialog(url, s, "", "div|li|span") 

        c = c & s 
    rs.MoveNext : Next : rs.Close 
    content = Replace(content, "[list]" & defaultList & "[/list]", c) 

    If isMakeHtml = True Then
        url = "" 
        If Len(listFileName) > 5 Then
            url = Mid(listFileName, 1, Len(listFileName) - 5) & "[id].html" 
            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
        End If 
    Else
        url = getUrlAddToParam(getUrl(), "?page=[id]", "replace") 
    End If 

    getDetailList = content 
End Function 


'****************************************************
'Ĭ���б�ģ��
Function defaultListTemplate()
    Dim c, templateHtml, listTemplate, lableName, startStr, endStr 

    templateHtml = getFText(cfg_webTemplate & "/" & templateName) 

    lableName = "list" 
    startStr = "<!--#" & lableName & " start#-->" 
    endStr = "<!--#" & lableName & " end#-->" 
    'call rwend(templateHtml)
    If InStr(templateHtml, startStr) > 0 And InStr(templateHtml, endStr) > 0 Then
        listTemplate = strCut(templateHtml, startStr, endStr, 2) 
    Else
        startStr = "<!--#" & lableName 
        endStr = "#-->" 
        If InStr(templateHtml, startStr) > 0 And InStr(templateHtml, endStr) > 0 Then
            listTemplate = strCut(templateHtml, startStr, endStr, 2) 
        End If 
    End If 
    If listTemplate = "" Then
        c = "<ul class=""list"">" & vbCrLf 
        c = c & "[list]    <li><a href=""[$url$]""[$atitle$][$atarget$][$abcolor$][$anofollow$]>[$title$]</a><span class=""time"">[$adddatetime format_time='7'$]</span></li>" & vbCrLf 
        c = c & "[/list]" & vbCrLf 
        c = c & "</ul>" & vbCrLf 
        c = c & "<div class=""clear10""></div>" & vbCrLf 
        c = c & "<div>[$pageInfo$]</div>" & vbCrLf 
        listTemplate = c 
    End If 

    defaultListTemplate = listTemplate 
End Function 

'���崦��20160622
cacheHtmlFilePath="/cache/html/" & setFileName(getThisUrlFileParam()) & ".html"
'���û���
if request("cache")<>"false" and onCacheHtml=true then
	if checkFile(cacheHtmlFilePath)=true then
		'call echo("��ȡ�����ļ�","OK")
		call rwend(getftext(cacheHtmlFilePath))
	end if
end if

'��¼��ǰ׺
If Request("db_PREFIX") <> "" Then
    db_PREFIX = Request("db_PREFIX") 
ElseIf Session("db_PREFIX") <> "" Then
    db_PREFIX = Session("db_PREFIX") 
End If 
'������ַ����
Call loadWebConfig() 
isMakeHtml = False                                                              'Ĭ������HTMLΪ�ر�
If Request("isMakeHtml") = "1" Or Request("isMakeHtml") = "true" Then
    isMakeHtml = True 
End If 
templateName = Request("templateName")                                          'ģ������

'�������ݴ���ҳ
Select Case Request("act")
    Case "savedata" : saveData(Request("stype")) : Response.End()                   '��������
    ''վ��ͳ�� | ����IP[653] | ����PV[9865] | ��ǰ����[65]')
    Case "webstat" : webStat(adminDir & "/Data/Stat/"):Response.End()    '��վͳ��
	
    Case "saveSiteMap" : isMakeHtml=true:saveSiteMap() :response.End()                                             '����sitemap.xml
	
	case "handleAction" 
	if request("ishtml")="1" then
		isMakeHtml = True
	end if
	rwend(handleAction(request("content")))		'������
End Select


'����html
If Request("act") = "makehtml" Then
    Call echo("makehtml", "makehtml") 
    isMakeHtml = True 
    Call makeWebHtml(" action actionType='" & Request("act") & "' columnName='" & Request("columnName") & "' id='" & Request("id") & "' ") 
    Call createFileGBK("index.html", code) 

'����Html����վ
ElseIf Request("act") = "copyHtmlToWeb" Then
    Call copyHtmlToWeb() 
'ȫ������
ElseIf Request("act") = "makeallhtml" Then
    Call makeAllHtml("", "", Request("id")) 

'���ɵ�ǰҳ��
ElseIf Request("isMakeHtml") <> "" And Request("isSave") <> "" Then

    Call handlePower("���ɵ�ǰHTMLҳ��")                                            '����Ȩ�޴���
    Call writeSystemLog("", "���ɵ�ǰHTMLҳ��")                                     'ϵͳ��־

    isMakeHtml = True 


    Call checkIDSQL(Request("id")) 
    Call rw(makeWebHtml(" action actionType='" & Request("act") & "' columnName='" & Request("columnName") & "' columnType='" & Request("columnType") & "' id='" & Request("id") & "' npage='" & Request("page") & "' ")) 
    glb_filePath = Replace(glb_url, cfg_webSiteUrl, "") 
    If Right(glb_filePath, 1) = "/" Then
        glb_filePath = glb_filePath & "index.html" 
    ElseIf glb_filePath = "" And glb_columnType = "��ҳ" Then
        glb_filePath = "index.html" 
    End If 
    '�ļ���Ϊ��  ���ҿ�������html
    If glb_filePath <> "" And glb_isonhtml = True Then
        Call createDirFolder(getFileAttr(glb_filePath, "1")) 
        Call createFileGBK(glb_filePath, code) 
        If Request("act") = "detail" Then
            conn.Execute("update " & db_PREFIX & "ArticleDetail set ishtml=true where id=" & Request("id")) 
        ElseIf Request("act") = "nav" Then
            If Request("id") <> "" Then
                conn.Execute("update " & db_PREFIX & "WebColumn set ishtml=true where id=" & Request("id")) 
            Else
                conn.Execute("update " & db_PREFIX & "WebColumn set ishtml=true where columnname='" & Request("columnName") & "'") 
            End If 
        End If 
        Call echo("�����ļ�·��", "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>") 

        '�������������� 20160216
        If glb_columnType = "����" Then
            Call makeAllHtml("", "", glb_columnId) 
        End If 

    End If 

'ȫ������
ElseIf Request("act") = "Search" Then
    Call rw(makeWebHtml("actionType='Search' npage='1' ")) 
Else
    If LCase(Request("issave")) = "1" Then
        Call makeAllHtml(Request("columnType"), Request("columnName"), Request("columnId")) 
    Else
        Call checkIDSQL(Request("id")) 
        Call rw(makeWebHtml(" action actionType='" & Request("act") & "' columnName='" & Request("columnName") & "' columnType='" & Request("columnType") & "' id='" & Request("id") & "' npage='" & Request("page") & "' ")) 
    End If 
End If
'��������html 
if onCacheHtml=true then
	call createFile(cacheHtmlFilePath,code)		'���浽�����ļ���20160622
end if
'���ID�Ƿ�SQL��ȫ
Function checkIDSQL(id)
    If checkNumber(id) = False And id <> "" Then
        Call eerr("��ʾ", "id���зǷ��ַ�") 
    End If 
End Function 




'http://127.0.0.1/aspweb.asp?act=nav&columnName=ASP
'http://127.0.0.1/aspweb.asp?act=detail&id=75
'����html��̬ҳ
Function makeWebHtml(action)
    Dim actionType, npagesize, npage, url, addSql, sortSql 
    actionType = RParam(action, "actionType") 
    npage = RParam(action, "npage") 
    npage = getnumber(npage) 
    If npage = "" Then
        npage = 1 
    Else
        npage = CInt(npage) 
    End If 
    '����
    If actionType = "nav" Then
        glb_columnType = RParam(action, "columnType") 
        glb_columnName = RParam(action, "columnName") 
        glb_columnId = RParam(action, "columnId") 
        If glb_columnId = "" Then
            glb_columnId = RParam(action, "id") 
        End If 
        If glb_columnType <> "" Then
            addSql = "where columnType='" & glb_columnType & "'" 
        End If 
        If glb_columnName <> "" Then
            addSql = getWhereAnd(addSql, "where columnName='" & glb_columnName & "'") 
        End If 
        If glb_columnId <> "" Then
            addSql = getWhereAnd(addSql, "where id=" & glb_columnId & "") 
        End If 
        'call echo("addsql",addsql)
        rs.Open "Select * from " & db_PREFIX & "webcolumn " & addSql, conn, 1, 1 
        If Not rs.EOF Then
            glb_columnId = rs("id") 
            glb_columnName = rs("columnname") 
            glb_columnType = rs("columntype") 
            glb_bodyContent = rs("bodycontent") 
            glb_detailTitle = glb_columnName 
            glb_flags = rs("flags") 
            npagesize = rs("npagesize")                                                     'ÿҳ��ʾ����
            glb_isonhtml = rs("isonhtml")                                                   '�Ƿ����ɾ�̬��ҳ
            sortSql = " " & rs("sortsql")                                                   '����SQL

            If rs("webTitle") <> "" Then
                cfg_webTitle = rs("webTitle")                                                   '��ַ����
            End If 
            If rs("webKeywords") <> "" Then
                cfg_webKeywords = rs("webKeywords")                                             '��վ�ؼ���
            End If 
            If rs("webDescription") <> "" Then
                cfg_webDescription = rs("webDescription")                                       '��վ����
            End If 
            If templateName = "" Then
                If Trim(rs("templatePath")) <> "" Then
                    templateName = rs("templatePath") 
                ElseIf rs("columntype") <> "��ҳ" Then
                    templateName = getDateilTemplate(rs("id"), "List") 
                End If 
            End If 
        End If : rs.Close 
        glb_columnENType = handleColumnType(glb_columnType) 
        glb_url = getColumnUrl(glb_columnName, "name") 

        '�������б�
        If InStr("|��Ʒ|����|��Ƶ|����|����|", "|" & glb_columnType & "|") > 0 Then
            glb_bodyContent = getDetailList(action, defaultListTemplate(), "ArticleDetail", "��Ŀ�б�", "*", npagesize, npage, "where parentid=" & glb_columnId & sortSql) 
        '�������б�
        ElseIf InStr("|����|", "|" & glb_columnType & "|") > 0 Then
            glb_bodyContent = getDetailList(action, defaultListTemplate(), "GuestBook", "�����б�", "*", npagesize, npage, " where isthrough<>0 " & sortSql) 
        ElseIf glb_columnType = "�ı�" Then
            '������Ŀ�ӹ���
            If Request("gl") = "edit" Then
                glb_bodyContent = "<span>" & glb_bodyContent & "</span>" 
            End If 
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=WebColumn&lableTitle=��վ��Ŀ&nPageSize=10&page=&id=" & glb_columnId & "&n=" & getRnd(11) 
            glb_bodyContent = handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span") 

        End If 
    'ϸ��
    ElseIf actionType = "detail" Then
        glb_locationType = "detail" 
        rs.Open "Select * from " & db_PREFIX & "articledetail where id=" & RParam(action, "id"), conn, 1, 1 
        If Not rs.EOF Then
            glb_columnName = getColumnName(rs("parentid")) 
            glb_detailTitle = rs("title") 
            glb_flags = rs("flags") 
            glb_isonhtml = rs("isonhtml")                                                   '�Ƿ����ɾ�̬��ҳ
            glb_id = rs("id")                                                               '����ID
            If isMakeHtml = True Then
                glb_url = getHandleRsUrl(rs("fileName"), rs("customAUrl"), "/detail/detail" & rs("id")) 
            Else
                glb_url = handleWebUrl("?act=detail&id=" & rs("id")) 
            End If 

            If rs("webTitle") <> "" Then
                cfg_webTitle = rs("webTitle")                                                   '��ַ����
            End If 
            If rs("webKeywords") <> "" Then
                cfg_webKeywords = rs("webKeywords")                                             '��վ�ؼ���
            End If 
            If rs("webDescription") <> "" Then
                cfg_webDescription = rs("webDescription")                                       '��վ����
            End If 

            glb_artitleAuthor = rs("author") 
            glb_artitleAdddatetime = rs("adddatetime") 
            glb_upArticle = upArticle(rs("parentid"), "sortrank", rs("sortrank")) 
            glb_downArticle = downArticle(rs("parentid"), "sortrank", rs("sortrank")) 
            glb_aritcleRelatedTags = aritcleRelatedTags(rs("relatedtags")) 
            glb_aritcleSmallImage = rs("smallimage") 
            glb_aritcleBigImage = rs("bigimage") 

            '��������
            'glb_bodyContent = "<div class=""articleinfowrap"">[$articleinfowrap$]</div>" & rs("bodycontent") & "[$relatedtags$]<ul class=""updownarticlewrap"">[$updownArticle$]</ul>"
            '��һƪ���£���һƪ����
            'glb_bodyContent = Replace(glb_bodyContent, "[$updownArticle$]", upArticle(rs("parentid"), "sortrank", rs("sortrank")) & downArticle(rs("parentid"), "sortrank", rs("sortrank")))
            'glb_bodyContent = Replace(glb_bodyContent, "[$articleinfowrap$]", "��Դ��" & rs("author") & " &nbsp; ����ʱ�䣺" & format_Time(rs("adddatetime"), 1))
            'glb_bodyContent = Replace(glb_bodyContent, "[$relatedtags$]", aritcleRelatedTags(rs("relatedtags")))

            glb_bodyContent = rs("bodycontent") 

            '������ϸ�ӿ���
            If Request("gl") = "edit" Then
                glb_bodyContent = "<span>" & glb_bodyContent & "</span>" 
            End If 
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page=&parentid=&id=" & RParam(action, "id") & "&n=" & getRnd(11) 
            glb_bodyContent = handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span") 

            If templateName = "" Then
                If Trim(rs("templatePath")) <> "" Then
                    templateName = rs("templatePath") 
                Else
                    templateName = getDateilTemplate(rs("parentid"), "Detail") 
                End If 
            End If 

        End If : rs.Close 

    '��ҳ
    ElseIf actionType = "onepage" Then
        rs.Open "Select * from " & db_PREFIX & "onepage where id=" & RParam(action, "id"), conn, 1, 1 
        If Not rs.EOF Then
            glb_detailTitle = rs("title") 
            glb_isonhtml = rs("isonhtml")                                                   '�Ƿ����ɾ�̬��ҳ
            If isMakeHtml = True Then
                glb_url = getHandleRsUrl(rs("fileName"), rs("customAUrl"), "/page/page" & rs("id")) 
            Else
                glb_url = handleWebUrl("?act=detail&id=" & rs("id")) 
            End If 

            If rs("webTitle") <> "" Then
                cfg_webTitle = rs("webTitle")                                                   '��ַ����
            End If 
            If rs("webKeywords") <> "" Then
                cfg_webKeywords = rs("webKeywords")                                             '��վ�ؼ���
            End If 
            If rs("webDescription") <> "" Then
                cfg_webDescription = rs("webDescription")                                       '��վ����
            End If 
            '����
            glb_bodyContent = rs("bodycontent") 


            '������ϸ�ӿ���
            If Request("gl") = "edit" Then
                glb_bodyContent = "<span>" & glb_bodyContent & "</span>" 
            End If 
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page=&parentid=&id=" & RParam(action, "id") & "&n=" & getRnd(11) 
            glb_bodyContent = handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span") 


            If templateName = "" Then
                If Trim(rs("templatePath")) <> "" Then
                    templateName = rs("templatePath") 
                Else
                    templateName = "Main_Model.html" 
                'call echo(templateName,"templateName")
                End If 
            End If 

        End If : rs.Close 

    '����
    ElseIf actionType = "Search" Then
        templateName = "Main_Model.html" 
        glb_searchKeyWord = Request("wd") 
        addSql = " where title like '%" & glb_searchKeyWord & "%'" 
        npagesize = 20 
        'call echo(npagesize, npage)
        glb_bodyContent = getDetailList(action, defaultListTemplate(), "ArticleDetail", "��վ��Ŀ", "*", npagesize, npage, addSql) 

    '���صȴ�
    ElseIf actionType = "loading" Then
        Call rwend("ҳ�����ڼ����С�����") 
    End If 
    'ģ��Ϊ�գ�����Ĭ����ҳģ��
    If templateName = "" Then
        templateName = "Index_Model.html"                                               'Ĭ��ģ��
    End If 
    '��⵱ǰ·���Ƿ���ģ��
    If InStr(templateName, "/") = False Then
        templateName = cfg_webTemplate & "/" & templateName 
    End If 
    'call echo("templateName",templateName)
    code = getftext(templateName) 


    code = handleAction(code)                                                       '������
    code = thisPosition(code)                                                       'λ��
    code = replaceGlobleVariable(code)                                              '�滻ȫ�ֱ�ǩ
    code = handleAction(code)                                                       '������    '����һ�Σ��������������ﶯ��

    code = handleAction(code)                                                       '������
    code = handleAction(code)                                                       '������
    code = thisPosition(code)                                                       'λ��
    code = replaceGlobleVariable(code)                                              '�滻ȫ�ֱ�ǩ
    code = delTemplateMyNote(code)                                                  'ɾ����������

    '��ʽ��HTML
    If InStr(cfg_flags, "|formattinghtml|") > 0 Then
        'code = HtmlFormatting(code)        '��
        code = handleHtmlFormatting(code, False, 0, "ɾ������")                         '�Զ���
    '��ʽ��HTML�ڶ���
    ElseIf InStr(cfg_flags, "|formattinghtmltow|") > 0 Then
        code = htmlFormatting(code)                                                     '��
        code = handleHtmlFormatting(code, False, 0, "ɾ������")                         '�Զ���
    'ѹ��HTML
    ElseIf InStr(cfg_flags, "|ziphtml|") > 0 Then
        code = ziphtml(code) 

    End If 
    '�պϱ�ǩ
    If InStr(cfg_flags, "|labelclose|") > 0 Then
        code = handleCloseHtml(code, True, "")                                          'ͼƬ�Զ���alt  "|*|",
    End If 

    '���߱༭20160127
    If rq("gl") = "edit" Then
        If InStr(code, "</head>") > 0 Then
            If InStr(code, "jquery.Min.js") = False Then
                code = Replace(code, "</head>", "<script src=""/Jquery/jquery.Min.js""></script></head>") 
            End If 
            code = Replace(code, "</head>", "<script src=""/Jquery/Callcontext_menu.js""></script></head>") 
        End If 
        If InStr(code, "<body>") > 0 Then
        'Code = Replace(Code,"<body>", "<body onLoad=""ContextMenu.intializeContextMenu()"">")
        End If 
    End If 
    'call echo(templateName,templateName)
    makeWebHtml = code 
End Function 

'���Ĭ��ϸ��ģ��ҳ
Function getDateilTemplate(parentid, templateType)
    Dim templateName 
    templateName = "Main_Model.html" 
    rsx.Open "select * from " & db_PREFIX & "webcolumn where id=" & parentid, conn, 1, 1 
    If Not rsx.EOF Then
        'call echo("columntype",rsx("columntype"))
        If rsx("columntype") = "����" Then
            '����ϸ��ҳ
            If checkFile(cfg_webTemplate & "/News_" & templateType & ".html") = True Then
                templateName = "News_" & templateType & ".html" 
            End If 
        ElseIf rsx("columntype") = "��Ʒ" Then
            '��Ʒϸ��ҳ
            If checkFile(cfg_webTemplate & "/Product_" & templateType & ".html") = True Then
                templateName = "Product_" & templateType & ".html" 
            End If 
        ElseIf rsx("columntype") = "����" Then
            '����ϸ��ҳ
            If checkFile(cfg_webTemplate & "/Down_" & templateType & ".html") = True Then
                templateName = "Down_" & templateType & ".html" 
            End If 

        ElseIf rsx("columntype") = "��Ƶ" Then
            '��Ƶϸ��ҳ
            If checkFile(cfg_webTemplate & "/Video_" & templateType & ".html") = True Then
                templateName = "Video_" & templateType & ".html" 
            End If 
        ElseIf rsx("columntype") = "����" Then
            '��Ƶϸ��ҳ
            If checkFile(cfg_webTemplate & "/GuestBook_" & templateType & ".html") = True Then
                templateName = "Video_" & templateType & ".html" 
            End If 
        ElseIf rsx("columntype") = "�ı�" Then
            '��Ƶϸ��ҳ
            If checkFile(cfg_webTemplate & "/Page_" & templateType & ".html") = True Then
                templateName = "Page_" & templateType & ".html" 
            End If 
        End If 
    End If : rsx.Close 
    'call echo(templateType,templateName)
    getDateilTemplate = templateName 

End Function 


'����ȫ��htmlҳ��
Sub makeAllHtml(columnType, columnName, columnId)
    Dim action, s, i, nPageSize, nCountSize, nPage, addSql, url, articleSql 
    Call handlePower("����ȫ��HTMLҳ��")                                            '����Ȩ�޴���
    Call writeSystemLog("", "����ȫ��HTMLҳ��")                                     'ϵͳ��־

    isMakeHtml = True 
    '��Ŀ
    Call echo("��Ŀ", "") 
    If columnType <> "" Then
        addSql = "where columnType='" & columnType & "'" 
    End If 
    If columnName <> "" Then
        addSql = getWhereAnd(addSql, "where columnName='" & columnName & "'") 
    End If 
    If columnId <> "" Then
        addSql = getWhereAnd(addSql, "where id in(" & columnId & ")") 
    End If 
    rss.Open "select * from " & db_PREFIX & "webcolumn " & addSql & " order by sortrank asc", conn, 1, 1 
    While Not rss.EOF
        glb_columnName = "" 
        '��������html
        If rss("isonhtml") = True Then
            If InStr("|��Ʒ|����|��Ƶ|����|����|����|����|��Ƹ|����|", "|" & rss("columntype") & "|") > 0 Then
                If rss("columntype") = "����" Then
                    nCountSize = getRecordCount(db_PREFIX & "guestbook", "")                        '��¼��
                Else
                    nCountSize = getRecordCount(db_PREFIX & "articledetail", " where parentid=" & rss("id")) '��¼��
                End If 
                nPageSize = rss("npagesize") 
                nPage = getPageNumb(CInt(nCountSize), CInt(nPageSize)) 
                If nPage <= 0 Then
                    nPage = 1 
                End If 
                For i = 1 To nPage
                    url = getHandleRsUrl(rss("fileName"), rss("customAUrl"), "/nav" & rss("id")) 
                    glb_filePath = Replace(url, cfg_webSiteUrl, "") 
                    If Right(glb_filePath, 1) = "/" Or glb_filePath = "" Then
                        glb_filePath = glb_filePath & "index.html" 
                    End If 
                    'call echo("glb_filePath",glb_filePath)
                    action = " action actionType='nav' columnName='" & rss("columnname") & "' npage='" & i & "' listfilename='" & glb_filePath & "' " 
                    'call echo("action",action)
                    Call makeWebHtml(action) 
                    If i > 1 Then
                        glb_filePath = Mid(glb_filePath, 1, Len(glb_filePath) - 5) & i & ".html" 
                    End If 
                    s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
                    Call echo(action, s) 
                    If glb_filePath <> "" Then
                        Call createDirFolder(getFileAttr(glb_filePath, "1")) 
                        Call createFileGBK(glb_filePath, code) 
                    End If 
                    doevents() 
                    templateName = ""                                                               '���ģ���ļ�����
                Next 
            Else
                action = " action actionType='nav' columnName='" & rss("columnname") & "'" 
                Call makeWebHtml(action) 
                glb_filePath = Replace(getColumnUrl(rss("columnname"), "name"), cfg_webSiteUrl, "") 
                If Right(glb_filePath, 1) = "/" Or glb_filePath = "" Then
                    glb_filePath = glb_filePath & "index.html" 
                End If 
                s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
                Call echo(action, s) 
                If glb_filePath <> "" Then
                    Call createDirFolder(getFileAttr(glb_filePath, "1")) 
                    Call createFileGBK(glb_filePath, code) 
                End If 
                doevents() 
                templateName = "" 
            End If 
            conn.Execute("update " & db_PREFIX & "WebColumn set ishtml=true where id=" & rss("id")) '���µ���Ϊ����״̬
        End If 
    rss.MoveNext : Wend : rss.Close 

    '��������ָ����Ŀ��Ӧ����
    If columnId <> "" Then
        articleSql = "select * from " & db_PREFIX & "articledetail where parentid=" & columnId & " order by sortrank asc" 
    '������������
    ElseIf addSql = "" Then
        articleSql = "select * from " & db_PREFIX & "articledetail order by sortrank asc" 
    End If 
    If articleSql <> "" Then
        '����
        Call echo("����", "") 
        rss.Open articleSql, conn, 1, 1 
        While Not rss.EOF
            glb_columnName = "" 
            action = " action actionType='detail' columnName='" & rss("parentid") & "' id='" & rss("id") & "'" 
            'call echo("action",action)
            Call makeWebHtml(action) 
            glb_filePath = Replace(glb_url, cfg_webSiteUrl, "") 
            If Right(glb_filePath, 1) = "/" Then
                glb_filePath = glb_filePath & "index.html" 
            End If 
            s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
            Call echo(action, s) 
            '�ļ���Ϊ��  ���ҿ�������html
            If glb_filePath <> "" And rss("isonhtml") = True Then
                Call createDirFolder(getFileAttr(glb_filePath, "1")) 
                Call createFileGBK(glb_filePath, code) 
                conn.Execute("update " & db_PREFIX & "ArticleDetail set ishtml=true where id=" & rss("id")) '��������Ϊ����״̬
            End If 
            templateName = ""                                                               '���ģ���ļ�����
        rss.MoveNext : Wend : rss.Close 
    End If 

    If addSql = "" Then
        '��ҳ
        Call echo("��ҳ", "") 
        rss.Open "select * from " & db_PREFIX & "onepage order by sortrank asc", conn, 1, 1 
        While Not rss.EOF
            glb_columnName = "" 
            action = " action actionType='onepage' id='" & rss("id") & "'" 
            'call echo("action",action)
            Call makeWebHtml(action) 
            glb_filePath = Replace(glb_url, cfg_webSiteUrl, "") 
            If Right(glb_filePath, 1) = "/" Then
                glb_filePath = glb_filePath & "index.html" 
            End If 
            s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
            Call echo(action, s) 
            '�ļ���Ϊ��  ���ҿ�������html
            If glb_filePath <> "" And rss("isonhtml") = True Then
                Call createDirFolder(getFileAttr(glb_filePath, "1")) 
                Call createFileGBK(glb_filePath, code) 
                conn.Execute("update " & db_PREFIX & "onepage set ishtml=true where id=" & rss("id")) '���µ�ҳΪ����״̬
            End If 
            templateName = ""                                                               '���ģ���ļ�����
        rss.MoveNext : Wend : rss.Close 

    End If 


End Sub 

'����html����վ
Sub copyHtmlToWeb()
    Dim webDir,toWebDir, toFilePath, filePath, fileName, fileList, splStr, content, s, s1, c, webImages, webCss, webJs, splJs 
    Dim webFolderName, jsFileList, setFileCode, nErrLevel, jsFilePath

    setFileCode = Request("setcode")                                                '�����ļ��������

    Call handlePower("��������HTMLҳ��")                                            '����Ȩ�޴���
    Call writeSystemLog("", "��������HTMLҳ��")                                     'ϵͳ��־

    webFolderName = cfg_webTemplate 
    If Left(webFolderName, 1) = "/" Then
        webFolderName = Mid(webFolderName, 2) 
    End If 
    If Right(webFolderName, 1) = "/" Then
        webFolderName = Mid(webFolderName, 1, Len(webFolderName) - 1) 
    End If 
    If InStr(webFolderName, "/") > 0 Then
        webFolderName = Mid(webFolderName, InStr(webFolderName, "/") + 1) 
    End If 
    webDir = "/htmladmin/" & webFolderName & "/"
	toWebDir="/htmlw" & "eb/viewweb/" 
	call createDirFolder(toWebDir)
	toWebDir = toWebDir & pinYin2(webFolderName) & "/" 
	
	call deleteFolder(toWebDir)				'ɾ��
	call createFolder("/htmlweb/web")		'�����ļ��� ��ֹweb�ļ��в�����20160504
    Call deleteFolder(webDir) 
    Call createDirFolder(webDir) 
    webImages = webDir & "Images/"
    webCss = webDir & "Css/" 
    webJs = webDir & "Js/"
    Call copyFolder(cfg_webImages, webImages) 
    Call copyFolder(cfg_webCss, webCss) 
    Call createFolder(webJs)                                                        '����Js�ļ���


    '����Js�ļ���
    splJs = Split(getDirJsList(webJs), vbCrLf) 
    For Each filePath In splJs
        If filePath <> "" Then
            toFilePath = webJs & getFileName(filePath) 
            Call echo("js", filePath) 
            Call moveFile(filePath, toFilePath) 
        End If 
    Next 
    '����Css�ļ���
    splStr = Split(getDirCssList(webCss), vbCrLf) 
    For Each filePath In splStr
        If filePath <> "" Then
            content = getftext(filePath) 
            content = Replace(content, cfg_webImages, "../images/") 
			
			content = deleteCssNote(content)			
			content=phptrim(content)
			'����Ϊutf-8���� 20160527
			if lcase(setFileCode)="utf-8" then
				content=replace(content,"gb2312","utf-8")
			end if
            Call writeToFile(filePath, content, setFileCode) 
            Call echo("css", cfg_webImages) 
        End If 
    Next 
    '������ĿHTML
    isMakeHtml = True 
    rss.Open "select * from " & db_PREFIX & "webcolumn where isonhtml=true", conn, 1, 1 
    While Not rss.EOF
        glb_filePath = Replace(getColumnUrl(rss("columnname"), "name"), cfg_webSiteUrl, "") 
        If Right(glb_filePath, 1) = "/" Or Right(glb_filePath, 1) = "" Then
            glb_filePath = glb_filePath & "index.html" 
        End If 
        If Right(glb_filePath, 5) = ".html" Then
            If Right(glb_filePath, 11) = "/index.html" Then
                fileList = fileList & glb_filePath & vbCrLf 
            Else
                fileList = glb_filePath & vbCrLf & fileList 
            End If 
            fileName = Replace(glb_filePath, "/", "_") 
            toFilePath = webDir & fileName 
            Call copyfile(glb_filePath, toFilePath) 
            Call echo("����", glb_filePath) 
        End If 
    rss.MoveNext : Wend : rss.Close 
    '��������HTML
    rss.Open "select * from " & db_PREFIX & "articledetail where isonhtml=true", conn, 1, 1 
    While Not rss.EOF
        glb_url = getHandleRsUrl(rss("fileName"), rss("customAUrl"), "/detail/detail" & rss("id")) 
        glb_filePath = Replace(glb_url, cfg_webSiteUrl, "") 
        If Right(glb_filePath, 1) = "/" Or Right(glb_filePath, 1) = "" Then
            glb_filePath = glb_filePath & "index.html" 
        End If 
        If Right(glb_filePath, 5) = ".html" Then
            If Right(glb_filePath, 11) = "/index.html" Then
                fileList = fileList & glb_filePath & vbCrLf 
            Else
                fileList = glb_filePath & vbCrLf & fileList 
            End If 
            fileName = Replace(glb_filePath, "/", "_") 
            toFilePath = webDir & fileName 
            Call copyfile(glb_filePath, toFilePath) 
            Call echo("����" & rss("title"), glb_filePath) 
        End If 
    rss.MoveNext : Wend : rss.Close 
    '���Ƶ���HTML
    rss.Open "select * from " & db_PREFIX & "onepage where isonhtml=true", conn, 1, 1 
    While Not rss.EOF
        glb_url = getHandleRsUrl(rss("fileName"), rss("customAUrl"), "/page/page" & rss("id")) 
        glb_filePath = Replace(glb_url, cfg_webSiteUrl, "") 
        If Right(glb_filePath, 1) = "/" Or Right(glb_filePath, 1) = "" Then
            glb_filePath = glb_filePath & "index.html" 
        End If 
        If Right(glb_filePath, 5) = ".html" Then
            If Right(glb_filePath, 11) = "/index.html" Then
                fileList = fileList & glb_filePath & vbCrLf 
            Else
                fileList = glb_filePath & vbCrLf & fileList 
            End If 
            fileName = Replace(glb_filePath, "/", "_") 
            toFilePath = webDir & fileName 
            Call copyfile(glb_filePath, toFilePath) 
            Call echo("��ҳ" & rss("title"), glb_filePath) 
        End If 
    rss.MoveNext : Wend : rss.Close 
    '��������html�ļ��б�
    'call echo(cfg_webSiteUrl,cfg_webTemplate)
    'call rwend(fileList)
    Dim sourceUrl, replaceUrl 
    splStr = Split(fileList, vbCrLf) 
    For Each filePath In splStr
        If filePath <> "" Then
            filePath = webDir & Replace(filePath, "/", "_") 
            Call echo("filePath", filePath) 
            content = getftext(filePath) 

            For Each s In splStr
                s1 = s 
                If Right(s1, 11) = "/index.html" Then
                    s1 = Left(s1, Len(s1) - 11) & "/" 
                End If 
                sourceUrl = cfg_webSiteUrl & s1 
                replaceUrl = cfg_webSiteUrl & Replace(s, "/", "_") 
                'Call echo(sourceUrl, replaceUrl) 							'����  ���������ʾ20160613
                content = Replace(content, sourceUrl, replaceUrl) 
            Next 
            content = Replace(content, cfg_webSiteUrl, "")                                  'ɾ����ַ
            content = Replace(content, cfg_webTemplate, "")                                 'ɾ��ģ��·��
            'content=nullLinkAddDefaultName(content)
            For Each s In splJs
                If s <> "" Then
                    fileName = getFileName(s) 
                    content = Replace(content, "Images/" & fileName, "js/" & fileName) 
                End If 
            Next 
            If InStr(content, "/Jquery/Jquery.Min.js") > 0 Then
                content = Replace(content, "/Jquery/Jquery.Min.js", "js/Jquery.Min.js") 
                Call copyfile("/Jquery/Jquery.Min.js", webJs & "/Jquery.Min.js") 
            End If 
            content = Replace(content, "<a href="""" ", "<a href=""index.html"" ")    '����ҳ��index.html
            Call createFileGBK(filePath, content) 
        End If 
    Next 

    '�Ѹ�����վ���µ�images/�ļ����µ�js�Ƶ�js/�ļ�����  20160315
    Dim htmlFileList, splHtmlFile, splJsFile, htmlFilePath, jsFileName 
    jsFileList = getDirJsNameList(webImages) 
    htmlFileList = getDirHtmlList(webDir) 
    splHtmlFile = Split(htmlFileList, vbCrLf) 
    splJsFile = Split(jsFileList, vbCrLf) 
    For Each htmlFilePath In splHtmlFile
        content = getftext(htmlFilePath) 
        For Each jsFileName In splJsFile
            content = regExp_Replace(content, "Images/" & jsFileName, "js/" & jsFileName) 
        Next 

        nErrLevel = 0 
        content = handleHtmlFormatting(content, False, nErrLevel, "|ɾ������|")         '|ɾ������|
        content = handleCloseHtml(content, True, "")                                    '�պϱ�ǩ
        nErrLevel = checkHtmlFormatting(content) 
        If checkHtmlFormatting(content) = False Then
            Call eerr(htmlFilePath & "(��ʽ������)", nErrLevel) 		'ע��
        End If 
		'����Ϊutf-8����
		if lcase(setFileCode)="utf-8" then
			content=replace(content,"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />","<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />")
		end if
		content=phptrim(content)
        Call writeToFile(htmlFilePath, content, setFileCode)
    Next 
    'images��js�ƶ���js��
    For Each jsFileName In splJsFile
		jsFilePath=webImages & jsFileName
        content = getftext(jsFilePath) 		
		content=phptrim(content)
        Call writeToFile(webJs & jsFileName, content, setFileCode) 
		call deleteFile(jsFilePath)
    Next 
	
	call copyFolder(webDir,toWebDir)
	'ʹhtmlWeb�ļ�����phpѹ��
	if request("isMakeZip")="1" then
	    Call makeHtmlWebToZip(webDir) 
	end if
	'ʹ��վ��xml���20160612
	if request("isMakeXml")="1" then
		call makeHtmlWebToXmlZip("/htmladmin/", webFolderName)
	end if
End Sub 

'ʹhtmlWeb�ļ�����phpѹ��
Function makeHtmlWebToZip(webDir)
    Dim content, splStr, filePath, c, fileArray, fileName, fileType, isTrue 
    Dim webFolderName 
    Dim cleanFileList 
    splStr = Split(webDir, "/") 
    webFolderName = splStr(2) 
    'call eerr(webFolderName,webDir)
    content = getFileFolderList(webDir, True, "ȫ��", "", "ȫ���ļ���", "", "") 
    splStr = Split(content, vbCrLf) 
    For Each filePath In splStr
        If checkfolder(filePath) = False Then
            fileArray = handleFilePathArray(filePath) 
            fileName = LCase(fileArray(2)) 
            fileType = LCase(fileArray(4)) 
            fileName = remoteNumber(fileName) 
            isTrue = True 

            If InStr("|" & cleanFileList & "|", "|" & fileName & "|") > 0 And fileType = "html" Then
                isTrue = False 
            End If 
            If isTrue = True Then
                'call echo(fileType,fileName)
                If c <> "" Then c = c & "|" 
                c = c & Replace(filePath, handlePath("/"), "") 
                cleanFileList = cleanFileList & fileName & "|" 
            End If 
        End If 
    Next 
    Call rw(c) 
    c = c & "|||||" 
    Call createFileGBK("htmlweb/1.txt", c) 
    Call echo("<hr>cccccccccccc", c) 
    '���ж�����ļ�����20160309
    If checkFile("/myZIP.php") = True Then
        Call echo("", XMLPost(getHost() & "/myZIP.php?webFolderName=" & webFolderName , "content=" & escape(c))) 
    End If 

End Function 
'ʹ��վ��xml���20160612
function makeHtmlWebToXmlZip(newWebDir,rootDir)
        Dim xmlFileName,xmlSize
        xmlFileName = getIP() & "_update.xml" 

		'newWebDir="\Templates2015\"
		'rootDir="\sharembweb\"
	
        Dim objXmlZIP : Set objXmlZIP = new xmlZIP
            Call objXmlZIP.run(handlePath(newWebDir), handlePath(newWebDir & rootDir), False, xmlFileName)  
			call echo(handlePath(newWebDir), handlePath(newWebDir & rootDir))
        Set objXmlZIP = Nothing 
		doevents
		xmlSize=getFSize(xmlFileName)
		xmlSize=printSpaceValue(xmlSize)
        Call echo("����xml����ļ�", "<a href=/tools/downfile.asp?act=download&downfile=" & xorEnc("/" & xmlFileName, 31380) & " title='�������'>�������" & xmlFileName & "("& xmlSize &")</a>") 
end function


'���ɸ���sitemap.xml 20160118
Sub saveSiteMap()
    Dim isWebRunHtml                                                                '�Ƿ�Ϊhtml��ʽ��ʾ��վ
    Dim changefreg                                                                  '����Ƶ��
    Dim priority                                                                    '���ȼ�
    Dim c, url 
    Call handlePower("�޸�����SiteMap")                                             '����Ȩ�޴���

    changefreg = Request("changefreg") 
    priority = Request("priority") 
    Call loadWebConfig()                                                            '��������
    'call eerr("cfg_flags",cfg_flags)
    If InStr(cfg_flags, "|htmlrun|") > 0 Then
        isWebRunHtml = True 
    Else
        isWebRunHtml = False 
    End If 

    c = c & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf 
    c = c & vbTab & "<urlset xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"">" & vbCrLf 
	dim rsx:Set rsx = CreateObject("Adodb.RecordSet")
    '��Ŀ
    rsx.Open "select * from " & db_PREFIX & "webcolumn order by sortrank asc", conn, 1, 1 
    While Not rsx.EOF
        If rsx("nofollow") = False Then
            c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
            If isWebRunHtml = True Then
                url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/nav" & rsx("id")) 
				url=handleAction(url)
            Else
                url = escape("?act=nav&columnName=" & rsx("columnname")) 
            End If 
            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
            'call echo(cfg_webSiteUrl,url)

            c = c & copystr(vbTab, 3) & "<loc>" & url & "</loc>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<lastmod>" & format_Time(rsx("updatetime"), 2) & "</lastmod>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<changefreq>" & changefreg & "</changefreq>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<priority>" & priority & "</priority>" & vbCrLf 
            c = c & copystr(vbTab, 2) & "</url>" & vbCrLf 
            Call echo("��Ŀ", "<a href=""" & url & """ target='_blank'>" & url & "</a>") 
        End If 
    rsx.MoveNext : Wend : rsx.Close 

    '����
    rsx.Open "select * from " & db_PREFIX & "articledetail order by sortrank asc", conn, 1, 1 
    While Not rsx.EOF
        If rsx("nofollow") = False Then
            c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
            If isWebRunHtml = True Then
                url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/detail/detail" & rsx("id"))
				url=handleAction(url)
            Else
                url = "?act=detail&id=" & rsx("id") 
            End If
            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
            'call echo(cfg_webSiteUrl,url)

            c = c & copystr(vbTab, 3) & "<loc>" & url & "</loc>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<lastmod>" & format_Time(rsx("updatetime"), 2) & "</lastmod>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<changefreq>" & changefreg & "</changefreq>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<priority>" & priority & "</priority>" & vbCrLf 
            c = c & copystr(vbTab, 2) & "</url>" & vbCrLf 
            Call echo("����", "<a href=""" & url & """ target='_blank'>" & url & "</a>") 
        End If 
    rsx.MoveNext : Wend : rsx.Close 

    '��ҳ
    rsx.Open "select * from " & db_PREFIX & "onepage order by sortrank asc", conn, 1, 1 
    While Not rsx.EOF
        If rsx("nofollow") = False Then
            c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
            If isWebRunHtml = True Then
                url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/page/detail" & rsx("id")) 
				url=handleAction(url)				
            Else
                url = "?act=onepage&id=" & rsx("id") 
            End If
            url = urlAddHttpUrl(cfg_webSiteUrl, url)
            'call echo(cfg_webSiteUrl,url)

            c = c & copystr(vbTab, 3) & "<loc>" & url & "</loc>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<lastmod>" & format_Time(rsx("updatetime"), 2) & "</lastmod>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<changefreq>" & changefreg & "</changefreq>" & vbCrLf 
            c = c & copystr(vbTab, 3) & "<priority>" & priority & "</priority>" & vbCrLf 
            c = c & copystr(vbTab, 2) & "</url>" & vbCrLf 
            Call echo("��ҳ", "<a href=""" & url & """ target='_blank'>" & url & "</a>") 
        End If 
    rsx.MoveNext : Wend : rsx.Close 


    c = c & vbTab & "</urlset>" & vbCrLf 

    Call loadWebConfig() 
	
    Call createfile("sitemap.xml", c) 
    Call echo("����sitemap.xml�ļ��ɹ�", "<a href='/sitemap.xml' target='_blank'>���Ԥ��sitemap.xml</a>") 

    '�ж��Ƿ�����sitemap.html
    If Request("issitemaphtml") = "1" Then
        c = "" 
        '�ڶ���
        '��Ŀ
        rsx.Open "select * from " & db_PREFIX & "webcolumn order by sortrank asc", conn, 1, 1 
        While Not rsx.EOF
            If rsx("nofollow") = False Then
                If isWebRunHtml = True Then
                    url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/nav" & rsx("id")) 
					url=handleAction(url)
                Else
                    url = escape("?act=nav&columnName=" & rsx("columnname")) 
                End If 
                url = urlAddHttpUrl(cfg_webSiteUrl, url) 

                c = c & "<li style=""width:20%;""><a href=""" & url & """>" & rsx("columnname") & "</a><ul>" & vbCrLf 

                '����
                rss.Open "select * from " & db_PREFIX & "articledetail where parentId=" & rsx("id") & " order by sortrank asc", conn, 1, 1 
                While Not rss.EOF
                    If rss("nofollow") = False Then
                        If isWebRunHtml = True Then
                            url = getRsUrl(rss("fileName"), rss("customAUrl"), "/detail/detail" & rss("id")) 
							url=handleAction(url)
                        Else
                            url = "?act=detail&id=" & rss("id") 
                        End If 
                        url = urlAddHttpUrl(cfg_webSiteUrl, url) 
                        c = c & "<li style=""width:20%;""><a href=""" & url & """ target=""_blank"">" & rss("title") & "</a>" & vbCrLf 
                    End If 
                rss.MoveNext : Wend : rss.Close 
                c = c & "</ul></li>" & vbCrLf 


            End If 
        rsx.MoveNext : Wend : rsx.Close 

        '����
        c = c & "<li style=""width:20%;""><a href=""javascript:;"">�����б�</a><ul>" & vbCrLf 
        rsx.Open "select * from " & db_PREFIX & "onepage order by sortrank asc", conn, 1, 1 
        While Not rsx.EOF
            If rsx("nofollow") = False Then
                c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
                If isWebRunHtml = True Then
                    url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/page/detail" & rsx("id")) 
					url=handleAction(url)
                Else
                    url = "?act=onepage&id=" & rsx("id") 
                End If 
                c = c & "<li style=""width:20%;""><a href=""" & url & """ target=""_blank"">" & rsx("title") & "</a>" & vbCrLf 
            End If 
        rsx.MoveNext : Wend : rsx.Close 
        c = c & "</ul></li>" & vbCrLf 

        Dim templateContent 
        templateContent = getftext(adminDir & "/template_SiteMap.html") 


        templateContent = Replace(templateContent, "{$content$}", c) 
        templateContent = Replace(templateContent, "{$Web_Title$}", cfg_webTitle) 
		 
		
        Call createfile("sitemap.html", templateContent) 
        Call echo("����sitemap.html�ļ��ɹ�", "<a href='/sitemap.html' target='_blank'>���Ԥ��sitemap.html</a>") 
    End If 
    Call writeSystemLog("", "����sitemap.xml")                                      'ϵͳ��־
End Sub 
%>       