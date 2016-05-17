<% 
'与phpcms通用   我的后台


'显示导航列表
Function XY_PHP_NavList(action)
    XY_PHP_NavList = XY_PHP_GeneralList(action, "nav", "Select * From [v9_category] where parentid=0 and ismenu=1 order by listorder asc") 
End Function 

'显示细节列表
Function XY_PHP_DetailList(action)
    XY_PHP_DetailList = XY_PHP_GeneralList(action, "detail", "Select top 6 * From [v9_picture]") 
End Function 

'通用信息列表
Function XY_PHP_GeneralList(action, sType, sql)
    Dim did, sid, tid, title, topNumb, cutStrNumb, addSql 
    Call HandleFunParameter(action, did, sid, tid, title, topNumb, cutStrNumb, addSql) '获得这么函数参数
    Dim defaultStr, i, j, s, c, startStr, endStr, url 
    Dim noFollow                                                                    '不追踪 20141222
    defaultStr = GetDefaultValue(action)                                            '获得默认内容
    Dim modI                                                                        '余循环20150112
    noFollow = Trim(LCase(RParam(action, "noFollow")))                              '不追踪

    rs.Open sql, conn, 1, 1 
    'Call eerr("Sql【"& Rs.RecordCount &"】", Sql)
    For i = 1 To rs.RecordCount
        If rs.EOF Then Exit For 
        '【导航】
        If sType = "nav" Then
            url = sType & "_" & rs("catid") & ".html" 
            did = rs("catname") 
            title = rs("catname") 
        '【文章】
        ElseIf sType = "detail" Then
            url = sType & "_" & rs("ID") & ".html" 
            did = "" 
            title = rs("title") 
        End If 
        '网址判断
        If url = WEBURLFILEPATH Then
            startStr = "[list-focus]" : endStr = "[/list-focus]" 
        Else
            startStr = "[list-" & i & "]" : endStr = "[/list-" & i & "]" 
        End If 
        '网址加前缀
        url = WEBURLPREFIX & url 
        '例[list-mod2]  [/list-mod2]    20150112
        For modI = 6 To 2 Step - 1
            If InStr(defaultStr, startStr) = False And i Mod modI = 0 Then
                startStr = "[list-mod" & modI & "]" : endStr = "[/list-mod" & modI & "]" 
                If InStr(defaultStr, startStr) > 0 Then
                    Exit For 
                End If 
            End If 
        Next 

        '没有则用默认
        If InStr(defaultStr, startStr) = False Then
            startStr = "[list]" : endStr = "[/list]" 
        End If 

        If InStr(defaultStr, startStr) > 0 And InStr(defaultStr, endStr) > 0 Then
            s = StrCut(defaultStr, startStr, endStr, 2) 
            For j = 1 To 3
                s = ReplaceValueParam(s, "ni", i)                                               '不对为i，因为i会与imgurl冲突 [$i$]
                s = ReplaceValueParam(s, "编号-1", i - 1)                                       '不对为i，因为i会与imgurl冲突 [$i$]
                s = ReplaceValueParam(s, "编号", i)                                             '不对为i，因为i会与imgurl冲突 [$i$]
                s = ReplaceValueParam(s, "bigclassname", did) 
                s = ReplaceValueParam(s, "title", title) 
                s = ReplaceValueParam(s, "url", url) 
            Next 


            '开始位置加Dialog内容
            startStr = "[list-" & i & " startdialog]" : endStr = "[/list-" & i & " startdialog]" 
            If InStr(defaultStr, startStr) > 0 And InStr(defaultStr, endStr) > 0 Then
                s = StrCut(defaultStr, startStr, endStr, 2) & s 
            End If 
            '结束位置加Dialog内容
            startStr = "[list-" & i & " enddialog]" : endStr = "[/list-" & i & " enddialog]" 
            If InStr(defaultStr, startStr) > 0 And InStr(defaultStr, endStr) > 0 Then
                s = s & StrCut(defaultStr, startStr, endStr, 2) 
            End If 
            c = c & s 
        End If 
    rs.MoveNext : Next : rs.Close 

    '开始内容加Dialog内容
    startStr = "[dialog start]" : endStr = "[/dialog start]" 
    If InStr(defaultStr, startStr) > 0 And InStr(defaultStr, endStr) > 0 Then
        c = StrCut(defaultStr, startStr, endStr, 2) & c 
    End If 
    '结束内容加Dialog内容
    startStr = "[dialog end]" : endStr = "[/dialog end]" 
    If InStr(defaultStr, startStr) > 0 And InStr(defaultStr, endStr) > 0 Then
        c = c & StrCut(defaultStr, startStr, endStr, 2) 
    End If 
    XY_PHP_GeneralList = c
End Function


'获得文本内容 这个是干什么用得？？？20150121
Function XY_PHP_SinglePage(Action)
	XY_PHP_SinglePage=XY_PHP_GetFieldValue(Action,"select * from [v9_page]","content")
End Function
'获得单个字段内容
function XY_PHP_GetFieldValue(Action,sql,fieldName)
	rs.open sql,conn,1,1
	if not rs.eof then
		XY_PHP_GetFieldValue=rs(fieldName)
	else
		XY_PHP_GetFieldValue=GetDefaultValue(Action)
	end if:rs.close
end function
%> 
















