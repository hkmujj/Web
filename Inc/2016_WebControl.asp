<% 
'��վ���� 20160223



'����ģ���滻����
Function handleModuleReplaceArray(ByVal content)
    Dim i, startStr, endStr, s, lableName 
    For i = 1 To UBound(ModuleReplaceArray) - 1
        If ModuleReplaceArray(i, 0) = "" Then
            Exit For 
        End If 
        'call echo(ModuleReplaceArray(i,0),ModuleReplaceArray(0,i))
        lableName = ModuleReplaceArray(i, 0) 
        s = ModuleReplaceArray(0, i) 
        If lableName = "��ɾ����" Then
            content = Replace(content, s, "") 
        Else
            startStr = "<replacestrname " & lableName & ">" : endStr = "</replacestrname " & lableName & ">" 
            If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
                content = replaceContentModule(content, startStr, endStr, s, "") 
            End If 
            startStr = "<replacestrname " & lableName & "/>" 
            If InStr(content, startStr) > 0 Then
                content = replaceContentRowModule(content, "<replacestrname " & lableName & "/>", s, "") 
            End If 
        End If 
    Next 
    handleModuleReplaceArray = content 
End Function 

'ȥ��ģ���ﲻ��Ҫ��ʾ���� ɾ��ģ�����ҵ�ע�ʹ���
Function delTemplateMyNote(code)
    Dim startStr, endStr, i, s, handleNumb, splStr, Block, id 
    Dim content, DragSortCssStr, DragSortStart, DragSortEnd, DragSortValue, c 
    dim lableName,lableStartStr,lableEndStr
	handleNumb = 99                                                                 '���ﶨ�����Ҫ
	
	'��ǿ��  �����Ҳ����<!--#aaa start#--><!--#aaa end#-->
    startStr = "<!--#" : endStr = "#-->" 
    For i = 1 To handleNumb
        If InStr(code, startStr) > 0 And InStr(code, endStr) > 0 Then
            lableName = StrCut(code, startStr, endStr, 2)
			if instr(lableName," start")>0 then
				lableName=mid(lableName,1,len(lableName)-6)
			end if	
			
			s=startStr & lableName & endStr
			lableStartStr=startStr & lableName & " start" & endStr
			lableEndStr=startStr & lableName & " end" & endStr
			If InStr(code, lableStartStr) > 0 And InStr(code, lableEndStr) > 0 Then
            	s = StrCut(code, lableStartStr, lableEndStr, 1)
				'call echo(">>",s)
			end if
			code=replace(code,s,"")
			'call echo("s",s)
			'call echo("lableName",lableName)
			'call echo("lableStartStr",replace(lableStartStr,"<","&lt;"))
			'call echo("lableEndStr",replace(lableEndStr,"<","&lt;"))
		else
			exit for
        End If 
    Next
	
	

    '���ReadBlockList�������б�����  �����и�����ĵط����������ݿ��Դ��ⲿ�������ݣ�����Ժ���
    'Call Eerr("ReadBlockList",ReadBlockList)
    'д��20141118
    'splStr = Split(ReadBlockList, vbCrLf)                 '�������֣�������
    '�޸���20151230
    For i = 1 To handleNumb
        startStr = "<R#��������" : endStr = " start#>" 
        Block = StrCut(code, startStr, endStr, 2) 
        If Block <> "" Then
            startStr = "<R#��������" & Block & " start#>" : endStr = "<R#��������" & Block & " end#>" 
            If InStr(code, startStr) > 0 And InStr(code, endStr) > 0 Then
                s = StrCut(code, startStr, endStr, 1) 
                code = Replace(code, s, "")                                                     '�Ƴ�
            End If 
        Else
            Exit For 
        End If 
    Next

	'ɾ����ҳ����20160309
	startStr = "<!--#list start#-->" 
	endStr = "<!--#list end#-->" 
	If InStr(code, startStr) > 0 And InStr(code, endStr) > 0 Then
		s=StrCut(code, startStr, endStr, 2) 
		code=replace(code,s,"")
	End If 

    If Request("gl") = "yun" Then
        content = GetFText("/Jquery/dragsort/Config.html") 
        content = GetFText("/Jquery/dragsort/ģ����ק.html") 
        'Css��ʽ
        startStr = "<style>" 
        endStr = "</style>" 
        If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
            DragSortCssStr = StrCut(content, startStr, endStr, 1) 
        End If 
        '��ʼ����
        startStr = "<!--#top start#-->" 
        endStr = "<!--#top end#-->" 
        If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
            DragSortStart = StrCut(content, startStr, endStr, 2) 
        End If 
        '��������
        startStr = "<!--#foot start#-->" 
        endStr = "<!--#foot end#-->" 
        If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
            DragSortEnd = StrCut(content, startStr, endStr, 2) 
        End If 
        '��ʾ������
        startStr = "<!--#value start#-->" 
        endStr = "<!--#value end#-->" 
        If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 Then
            DragSortValue = StrCut(content, startStr, endStr, 2) 
        End If 



        '���ƴ���
        startStr = "<dIv datid='" 
        endStr = "</dIv>" 
        content = GetArray(code, startStr, endStr, False, False) 
        splStr = Split(content, "$Array$") 
        For Each s In splStr
            startStr = "��DatId��'" 
            id = Mid(s, 1, InStr(s, startStr) - 1) 
            s = Mid(s, InStr(s, startStr) + Len(startStr)) 
            'C=C & "<li><div title='"& Id &"'>" & vbcrlf & "<div " & S & "</div>"& vbcrlf &"<div class='clear'></div></div><div class='clear'></div></li>"
            s = "<div" & s & "</div>" 
            'Call Die(S)
            c = c & Replace(Replace(DragSortValue, "{$value$}", s), "{$id$", id) 
        Next 
        c = Replace(c, "�����С�", vbCrLf) 
        c = DragSortStart & c & DragSortEnd 
        code = Mid(code, 1, InStr(code, "<body>") - 1) 
        code = Replace(code, "</head>", DragSortCssStr & "</head></body>" & c & "</body></html>") 
    End If 

    'ɾ��VB������ɵ���������
    startStr = "<dIv datid='" : endStr = "��DatId��'" 
    For i = 1 To handleNumb
        If InStr(code, startStr) > 0 And InStr(code, endStr) > 0 Then
            id = StrCut(code, startStr, endStr, 2) 
            code = Replace2(code, startStr & id & endStr, "<div ") 
        Else
            Exit For 
        End If 
    Next 
    code = Replace(code, "</dIv>", "</div>")                                  '�滻���������div

    '����Χ���
    startStr = "<!--#dialogteststart#-->" : endStr = "<!--#dialogtestend#-->" 
    code = Replace(code, "<!--#dialogtest start#-->", startStr) 
    code = Replace(code, "<!--#dialogtest end#-->", endStr) 
    For i = 1 To handleNumb
        If InStr(code, startStr) > 0 And InStr(code, endStr) > 0 Then
            s = StrCut(code, startStr, endStr, 1) 
            code = Replace2(code, s, "") 
        Else
            Exit For 
        End If 
    Next 
    '��ת���
    startStr = "<!--#teststart#-->" : endStr = "<!--#testend#-->" 
    code = Replace(code, "<!--#del start#-->", startStr)                         '������һ��
    code = Replace(code, "<!--#del end#-->", endStr)                             '������һ�� ����ʽ
    code = Replace(code, "<!--#test start#-->", startStr) 
    code = Replace(code, "<!--#test end#-->", endStr) 

    For i = 1 To handleNumb
        If InStr(code, startStr) > 0 And InStr(code, endStr) > 0 Then
            s = StrCut(code, startStr, endStr, 1) 
            code = Replace2(code, s, "") 
        Else
            Exit For 
        End If 
    Next 
    'ɾ��ע�͵�span
    code = Replace(code, "<sPAn class=""testspan"">", "")                        '����Span
    code = Replace(code, "<sPAn class=""testhidde"">", "")                       '����Span
    code = Replace(code, "</sPAn>", "") 

    'delTemplateMyNote = Code:Exit Function

    startStr = "<!--#" : endStr = "#-->" 
    For i = 1 To handleNumb
        If InStr(code, startStr) > 0 And InStr(code, endStr) > 0 Then
            s = StrCut(code, startStr, endStr, 1) 
            code = Replace2(code, s, "") 
        Else
            Exit For 
        End If 
    Next 


    delTemplateMyNote = code 
End Function 

'�����滻����ֵ 20160114
Function handleReplaceValueParam(content, ByVal paramName, replaceStr)
    If InStr(content, "[$" & paramName) = False Then
        paramName = LCase(paramName) 
    End If 
    handleReplaceValueParam = replaceValueParam(content, paramName, replaceStr) 
End Function 

'�滻����ֵ 2014  12 01
Function replaceValueParam(content, paramName, replaceStr)
    Dim startStr, endStr, labelStr, tempLabelStr, nLen, nTimeFormat, delHtmlYes, funStr, trimYes,isEscape, s ,i
    Dim ifStr                                                                       '�ж��ַ�
    Dim elseIfStr                                                                   '�ڶ��ж��ַ�
    Dim valueStr                                                                    '��ʾ�ַ�
    Dim elseStr                                                                     '�����ַ�
	dim elseIfValue,elseValue																	'�ڶ��ж�ֵ
    Dim instrStr,instr2Str                                                                    '�����ַ�
	dim tempReplaceStr																'�ݴ�
    'ReplaceStr = ReplaceStr & "�����������������ʱ̼ѽ��"
    'ReplaceStr = CStr(ReplaceStr)            'ת���ַ�����
    If IsNul(replaceStr) = True Then replaceStr = "" 
	tempReplaceStr=replaceStr
	
	'��ദ��99��  20160225
	for i =1 to 99 
		replaceStr=tempReplaceStr													'�ָ�
		startStr = "[$" & paramName : endStr = "$]" 
		'�ֶ������ϸ��ж� 20160226
		If InStr(content, startStr) > 0 And InStr(content, endStr) > 0 and (InStr(content, startStr & " ") > 0 or InStr(content, startStr & endStr) > 0) Then
			'��ö�Ӧ�ֶμ�ǿ��20151231
			If InStr(content, startStr & endStr) > 0 Then
				labelStr = startStr & endStr 
			ElseIf InStr(content, startStr & " ") > 0 Then
				labelStr = StrCut(content, startStr & " ", endStr, 1) 
			Else
				labelStr = StrCut(content, startStr, endStr, 1) 
			End If 
	 
			tempLabelStr = labelStr 
			labelStr = handleInModule(labelStr, "start") 
			'ɾ��Html
			delHtmlYes = RParam(labelStr, "delHtml")                                        '�Ƿ�ɾ��Html
			If delHtmlYes = "true" Then replaceStr = Replace(DelHtml(replaceStr), "<", "&lt;") 'HTML����
			'ɾ�����߿ո�
			trimYes = RParam(labelStr, "trim")                                              '�Ƿ�ɾ�����߿ո�
			If trimYes = "true" Then replaceStr = TrimVbCrlf(replaceStr) 
	
			'��ȡ�ַ�����
			nLen = RParam(labelStr, "len")                                                  '�ַ�����ֵ
			nLen = HandleNumber(nLen) 
			'If nLen<>"" Then ReplaceStr = CutStr(ReplaceStr,nLen,"null")' Left(ReplaceStr,nLen)
			If nLen <> "" Then replaceStr = CutStr(replaceStr, nLen, "...")                 'Left(ReplaceStr,nLen)
	
			'ʱ�䴦��
			nTimeFormat = RParam(labelStr, "format_time")                                   'ʱ�䴦��ֵ
			If nTimeFormat <> "" Then
				replaceStr = Format_Time(replaceStr, nTimeFormat) 
			End If 
	
			'�����Ŀ����
			s = RParam(labelStr, "getcolumnname") 
			If s <> "" Then
				If s = "@ME" Then
					s = replaceStr 
				End If 
				replaceStr = getcolumnname(s) 
			End If 
			'�����ĿURL
			s = RParam(labelStr, "getcolumnurl") 
			If s <> "" Then
				If s = "@ME" Then
					s = replaceStr 
				End If 
				replaceStr = getcolumnurl(s, "id") 
			End If 
	
			ifStr = RParam(labelStr, "if") 
			elseIfStr = RParam(labelStr, "elseif") 
			valueStr = RParam(labelStr, "value") 
			elseifValue = RParam(labelStr, "elseifvalue") 
			elseValue = RParam(labelStr, "elsevalue") 
			instrStr = RParam(labelStr, "instr")
			instr2Str = RParam(labelStr, "instr2")
			
			'call echo("ifStr",ifStr)
			'call echo("valueStr",valueStr)
			'call echo("elseStr",elseStr)
			'call echo("elseIfStr",elseIfStr)
			'call echo("replaceStr",replaceStr)
			If ifStr <> "" Or instrStr <> "" Then
				If(ifStr = CStr(replaceStr) And ifStr <> "") then
					replaceStr = valueStr  
				elseif elseIfStr = CStr(replaceStr) And elseIfStr <> "" Then
					replaceStr = valueStr  
					if elseifValue<>"" then
						replaceStr = elseifValue
					end if
				ElseIf InStr(CStr(replaceStr), instrStr) > 0 And instrStr <> "" Then
					replaceStr = valueStr 		 
				ElseIf InStr(CStr(replaceStr), instr2Str) > 0 And instr2Str <> "" Then
					replaceStr = valueStr  
					if elseifValue<>"" then
						replaceStr = elseifValue
					end if
				Else
					If elseValue <> "@ME" Then
						replaceStr = elseValue 
					End If 
				End If 
			End If 
	
			'��������20151231    [$title  function='left(@ME,40)'$]
			funStr = RParam(labelStr, "function")                                           '����
			If funStr <> "" Then
				funStr = Replace(funStr, "@ME", replaceStr) 
				replaceStr = handleContentCode(funStr, "") 
			End If 
	
			'Ĭ��ֵ
			s = RParam(labelStr, "default") 
			If s <> "" and s<>"@ME" Then
				If replaceStr = "" Then
					replaceStr = s 
				End If 
			End If 
			'escapeת��
			isEscape=lcase(RParam(labelStr, "escape")) 
			if isEscape="1" or isEscape="true" then
				replaceStr=escape(replaceStr)
			end if
	
			'�ı���ɫ
			s = RParam(labelStr, "fontcolor")                                               '����
			If s <> "" Then
				replaceStr = "<font color=""" & s & """>" & replaceStr & "</font>" 
			End If 
			
 
			
			
			'call echo(tempLabelStr,replaceStr)
			content = Replace(content, tempLabelStr, replaceStr) 
		else
			exit for
		End If 
	next
    replaceValueParam = content 
End Function 
 
  
'��ʾ�༭��20160115
Function displayEditor(action)
    Dim c 
    c = c & "<script type=""text/javascript"" src=""\Jquery\syntaxhighlighter\scripts/shCore.js""></script> " & vbCrLf 
    c = c & "<script type=""text/javascript"" src=""\Jquery\syntaxhighlighter\scripts/shBrushJScript.js""></script>" & vbCrLf 
    c = c & "<script type=""text/javascript"" src=""\Jquery\syntaxhighlighter\scripts/shBrushPhp.js""></script> " & vbCrLf 
    c = c & "<script type=""text/javascript"" src=""\Jquery\syntaxhighlighter\scripts/shBrushVb.js""></script> " & vbCrLf 
    c = c & "<link type=""text/css"" rel=""stylesheet"" href=""\Jquery\syntaxhighlighter\styles/shCore.css""/>" & vbCrLf 
    c = c & "<link type=""text/css"" rel=""stylesheet"" href=""\Jquery\syntaxhighlighter\styles/shThemeDefault.css""/>" & vbCrLf 
    c = c & "<script type=""text/javascript"">" & vbCrLf 
    c = c & "    SyntaxHighlighter.config.clipboardSwf = '\Jquery\syntaxhighlighter\scripts/clipboard.swf';" & vbCrLf 
    c = c & "    SyntaxHighlighter.all();" & vbCrLf 
    c = c & "</script>" & vbCrLf 

    displayEditor = c 
End Function 
'������վurl20160202
Function handleWebUrl(url)
    If Request("gl") <> "" Then
        url = getUrlAddToParam(url, "&gl=" & Request("gl"), "replace") 
    End If 
    If Request("templatedir") <> "" Then
        url = getUrlAddToParam(url, "&templatedir=" & Request("templatedir"), "replace") 
    End If 
    handleWebUrl = url 
End Function 

'
'���������޸�
'MainContent = HandleDisplayOnlineEditDialog(""& adminDir &"NavManage.Asp?act=EditNavBig&Id=" & TempRs("Id") & "&n=" & GetRnd(11), MainContent,"style='float:right;padding:0 4px;'")
Function handleDisplayOnlineEditDialog(url, content, cssStyle, replaceStr)
    Dim controlStr, splStr, s, addOK 
    If Request("gl") = "edit" Then
        If InStr(url, "&") > 0 Then
            url = url & "&vbgl=true" 
        End If 
        addOK = False                                                                   '���Ĭ��Ϊ��
        controlStr = getControlStr(url) & """" & cssStyle 
        If replaceStr <> "" Then
            splStr = Split(replaceStr, "|") 
            For Each s In splStr
                If s <> "" And InStr(content, s) > 0 Then
                    content = Replace2(content, s, s & controlStr) 
                    addOK = True 
                    Exit For 
                End If 
            Next 
        End If 
        If addOK = False Then
            '��һ��
            'C = "<div "& ControlStr &">" & vbCrlf
            'C=C & Content & vbCrlf
            'C = C & "</div>" & vbCrlf
            'Content = C
            '�ڶ���
            content = htmlAddAction(content, controlStr) 

        'Content = "<div "& ControlStr &">" & Content & "</div>"
        End If 
    End If 
    handleDisplayOnlineEditDialog = content 
End Function 
'��ÿ�������
Function getControlStr(url)
    If Request("gl") = "edit" Then
        getControlStr = " onMouseMove=""onColor(this,'#FDFAC6','red')"" onMouseOut=""offColor(this,'','')"" onDblClick=""window1('" & url & "','��Ϣ�޸�')"" title='˫�����Ҽ������޸�' oncontextmenu=""CommonMenu(event,this,'')" 'ɾ����ַΪ��
    End If 
End Function 

'html�Ӷ���(20151103)  call rw(htmlAddAction("  <a href=""javascript:;"">222222</a>", "onclick=""javascript:alert(111);"" "))
Function htmlAddAction(content, jsAction)
    Dim s, startStr, endStr, isHandle, lableName 
    s = content 
    s = phptrim(s) 
    startStr = Mid(s, 1, InStr(s, " ")) 
    endStr = ">" 
    isHandle = True 

    lableName = Trim(LCase(Replace(startStr, "<", ""))) 
    If InStr(s, startStr) = False Or InStr(s, endStr) = False Or InStr("|a|div|span|font|h1|h2|h3|h4|h5|h6|dt|dd|dl|li|ul|table|tr|td|", "|" & lableName & "|") = False Then
        isHandle = False 
    End If 

    If isHandle = True Then
        content = startStr & jsAction & Right(s, Len(s) - Len(startStr)) 
    End If 
    htmlAddAction = content 
End Function 


%> 
