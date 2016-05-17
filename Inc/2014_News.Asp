<%
'��վ������ַ���� (20140408�Ľ�)
Function webDebug(url, DebugUrl)
    If Request.QueryString("Debug") <> "" Then
        If getwebsite(DebugUrl) <> "" Then
            url = DebugUrl 
            webDebug = url 
            Exit Function 
        Else
            url = "/Inc/Create_Html.Asp?Debug=1&" & DebugUrl 
        End If 
        If Request("MackHtml") = "False" Then
            url = url & "&MackHtml=False" 
        End If 
        If Request("gl") <> "" Then
            url = url & "&gl=" & Request("gl") 
        End If 
    Else
        If CheckMakeHtmlFile(url) = True Then
            url = LCase(url)                                                                '��URL����Сдת��
        End If 
    End If 
    If CheckMakeHtmlFile(url) = True Then
        url = fullHttpUrl(host(), url)                                                  '����ַ����
    Else
        '׷����20141231  ԭ����Ϊ�ļ����ƿ��Զ�����ַ url:
        url = Trim(url) 
        If Left(LCase(url), 4) = "url:" Then
            url = Mid(url, 5) 
        End If 
        If Left(LCase(url), 5) = "/url:" Then
            url = Mid(url, 6) 
        End If 
        url = HandleTemplateAction(url, False) 

    End If 

    webDebug = url 
End Function 
'�ж��Ƿ�Ϊ����IP���������Ϊ�����ļ��мӸ���Ӧ��վ����
Function getWebFolderName()
    If getIP = "127.0.0.1" Or InStr(getIP, "192.168.") > 0 Then
        getWebFolderName = "/wwwroot/" & WebFolderName & "/" 
    End If 
End Function 
'�ж�Url�����Ƿ��.html��׺
Function getHandleUrl(ByVal url)
    Dim s 
    url = CStr(Trim(url)) 
    If url <> "" Then
        url = Replace(Trim(url), "\", "/") 
        If InStr(url, "://") = False Then url = Replace(Replace(url, "//", "/"), "//", "/") 
    End If 
    If Left(url, 1) <> "/" And Right(url, 1) <> "/" Then
        url = "/Html/" & url & ".Html" 
    ElseIf Right(LCase(url), 4) <> ".html" Then                                     '����û��.html ������
        s = Mid(url, InStrRev(url, "/") + 1) 
        If InStr(s, ".") = False Then
            If s <> "" Then
                url = url & ".Html" 
            End If 
        End If 
    End If 
    getHandleUrl = url 
End Function 
'��������NewsHandleListGroup
'��  �ã���ʾ�����б�
'ʱ  �䣺2014��10��12��
'��  ����Did ��������
'*       TopNumb ��ʾ����
'*       StartLable ��ʼ��ǩ
'*       EndLable ������ǩ
'*       CutStrNumb ��ȡ�ַ���
'*       DateYes ��ʾʱ��True,False
'����ֵ���ַ���
'��  �ԣ�Call Echo("���Ժ��� NewsHandleListGroup", NewsHandleListGroup("��Ʒ����", 6, "<li>", "</li>", 32, False, ""))
'****************************************************
Function NewsHandleListGroup(Did, TopNumb, StartLable, EndLable, CutStrNumb, DateYes, Action)
    Dim Url, S, C, Sql, ShowTitle, I, AddSql,FileName
    TopNumb = CInt(TopNumb)                        '����ʾ����ת���������ͣ�Ҫ��Ȼ�᲻������
	
	
	Dim Sid,Tid
	Sid = RParam(Action,"Sid")
	Tid = RParam(Action,"Tid")
	
	AddSql = RParam(Action,"AddSql")
    
	'�ϵ�20121224
	If LCase(Left(Did, 4)) = "sql:" Then
        AddSql = Mid(Did, 5)
    ElseIf Did<>"" Then
        AddSql = "Where BigClassName='" & Did & "'"' And Recommend<>0 Order By Sort"
    End If
	
	'������20141224
	AddSql = RParam(Action,"AddSql")
	Sql = "Select * From [Product] " & IIF(Did<>""," Where BigClassName='" & Did & "' ","")
	Sql = GetWhereAnd(Sql, IIF(Sid<>""," Where SmallClassName='" & Sid & "' ",""))
	Sql = GetWhereAnd(Sql, IIF(Tid<>""," Where ThreeClassName='" & Tid & "' ",""))
	
	Sql = GetWhereAnd(Sql,AddSql)
	
	'If AddSql <>"" Then
	'	Sql = GetWhereAnd(Sql,AddSql)
	'Else
	'	Sql = GetWhereAnd(Sql," And Recommend<>0 Order By Sort")
	'End If
	
	'Call echo("Sql",Sql)
    Rs.Open Sql, Conn, 1, 1 
    For I = 1 To Rs.RecordCount
        If I > TopNumb Then Exit For
		 
		'Edit 20141220
		Url = Rs("FileName")
		If CheckMakeHtmlFile(Url)=True Then
			Url = GetHandleFileName(Url)
			Url = DetailPageAddHtmlPrefix(Rs("FolderName")) & "/" & Url
		End If
		Call WebDebug(Url, "act=CreateArticle&ID=" & Rs("Id"))                      '����
		
		
        ShowTitle = CutStr(Rs("Title"), CutStrNumb, "... ")
        ShowTitle = FontColorFontB(ShowTitle, Rs("FontB"), Rs("FontColor"))

        S = "<a " & AHref(Url, Rs("Title"), Rs("Target")) & ">" & ShowTitle & "</a>" & vbCrLf 
        If LCase(DateYes) = "true" Then
            S = S & "<span class=""timeright"">" & Format_Time(Rs("UpDateTime"), 14) & "</span>" & vbCrLf 
        End If 
		S = DisplayOnlineEditDialog(""& adminDir &"Product.Asp?act=ShowEditProduct&Id=" & Rs("Id") & "&n=" & GetRnd(11), S)
        C = C & StartLable & S & EndLable 
    Rs.MoveNext : Next : Rs.Close 
    NewsHandleListGroup = C 
End Function
'�����б�
Function NewsList(Did,TopNumb,StartLable,EndLable,DateYes,Action)
	'Call Echo("NewsList","NewsList")
	NewsList = NewsHandleListGroup(Did,TopNumb,StartLable,EndLable, 28, DateYes,Action)
End Function
'����ļ�����
Function GetArticleParam(ByVal TableName,ByVal FieldName,ByVal AddSql)
	Dim Sql,FileName,Url,BigFolder
	Sql = "Select * From ["& TableName &"] " & AddSql
	TempRs.Open Sql,Conn,1,1
	If Not TempRs.Eof Then
	
		'�������ƻ����ַ			
		If FieldName = "[url]" then		
			Url = TempRs("FileName")
			If CheckRemoteUrl(Url)=False Then
				Url = GetHandleUrl(TempRs("FileName")) 
				Url = FullHttpUrl(Host,Url)
				Call WebDebug(Url, "act=CreateArticle&ID=" & TempRs("Id"))                      '����  
			End If 
	
			GetArticleParam = Url
		ElseIf FieldName<>"" Then
			GetArticleParam = TempRs(FieldName)
		End If
	End If:TempRs.Close
End Function
%>