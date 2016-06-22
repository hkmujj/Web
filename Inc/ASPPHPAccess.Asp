<%
'ASP PHP数据操作通用文件

'判断追加Sql是加Where 还是And   sql = getWhereAnd(sql,addSql)        修改于20141007 加强版
Function getWhereAnd(ByVal sql, ByVal addSql)
    Dim LCaseAddSql, AddType 
    '追加SQl为空则退出
    If Trim(addSql) = "" Then getWhereAnd = sql : Exit Function 
    If InStr(LCase(sql), " where ") > 0 Then
        AddType = " And " 
    Else
        AddType = " Where " 
    End If 
    If addSql <> "" Then 
        addSql = Trim(addSql) 
        LCaseAddSql = LCase(addSql) 
        If Left(LCaseAddSql, 6) = "order " Or Left(LCaseAddSql, 6) = "group " Then
            getWhereAnd = sql & " " & addSql : Exit Function 								'改进必需加空格，因为前面已经删除了20160115
        ElseIf Left(LCaseAddSql, 6) = "where " Then
            addSql = Mid(addSql, 7) 
        ElseIf Left(LCaseAddSql, 4) = "and " Then
            addSql = Mid(addSql, 5)
        End If 
        sql = sql & AddType & addSql 
    End If   
    getWhereAnd = sql 
End Function 
'多个查询 Or 与 And        二次修改于20140703
Function orAndSearch(addSql, SeectField, SearchValue)
    Dim splStr, s, c 
    SearchValue = RegExp_Replace(SearchValue, " or ", " Or ") 
    SearchValue = RegExp_Replace(SearchValue, " and ", " And ") 
    If InStr(SearchValue, " Or ") > 0 Then
        splStr = Split(SearchValue, " Or ") 
        For Each s In splStr
            If s <> "" Then
                If c <> "" Then c = c & " Or " 
                c = c & " " & SeectField & " Like '%" & s & "%'" 
            End If 
        Next 
    ElseIf InStr(SearchValue, " And ") > 0 Then
        splStr = Split(SearchValue, " And ") 
        For Each s In splStr
            If s <> "" Then
                If c <> "" Then c = c & " And " 
                c = c & " " & SeectField & " Like '%" & s & "%'" 
            End If 
        Next 
    ElseIf SearchValue <> "" Then
        splStr = Split(SearchValue, " And ") 
        For Each s In splStr
            If s <> "" Then
                If c <> "" Then c = c & " And " 
                c = c & " " & SeectField & " Like '%" & s & "%'" 
            End If 
        Next 
    End If 
    If c <> "" Then
        If InStr(LCase(addSql), " where ") = 0 Then
            c = " Where " & c 
        Else
            c = " And " & c 
        End If 
        addSql = addSql & c 
    End If 
    orAndSearch = addSql 
End Function 


%>