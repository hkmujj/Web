<% 
'读取单个文本显示标题[MainInfo]
Function ReadMainInfoShowTitle(Title)
    Dim AddSql 
    AddSql = "Select * From [MainInfo] Where Title='" & Title & "' And Recommend<>0 " 
    If UserId <> "" Then AddSql = AddSql & " And UserId=" & UserId 
    Rs.Open AddSql & " Order By Sort,ID", Conn, 1, 1 
    If Not Rs.EOF Then
        ReadMainInfoShowTitle = Rs("ShowTitle") 
    End If : Rs.Close 
End Function
 
'读取单个文件内容[MainInfo]
Function ReadMainInfo(Title)
    Dim AddSql 
    AddSql = "Select * From [MainInfo] Where Title='" & Title & "' And Recommend<>0 " 
    If UserId <> "" Then AddSql = AddSql & " And UserId=" & UserId 
    AddSql = AddSql & " Order By Sort,ID" 
    Rs.Open AddSql, Conn, 1, 1 
    If Not Rs.EOF Then
        ReadMainInfo = Rs("Content") 
    End If : Rs.Close 
End Function
 
'根据标题获得MainInfo中ID
Function ReadMainId(Title)
    Dim AddSql 
    AddSql = "Select * From [MainInfo] Where Title='" & Title & "' And Recommend<>0 " 
    If UserId <> "" Then AddSql = AddSql & " And UserId=" & UserId 
    Rs.Open AddSql, Conn, 1, 1 
    If Not Rs.EOF Then
        ReadMainId = Rs("Id") 
    End If : Rs.Close 
End Function
 
'获得文本内容参数
Function GetMainParam(ByVal TableName, ByVal FieldName, ByVal AddSql)
    Dim Sql, FileName, Url 
    Sql = "Select * From [" & TableName & "] " & AddSql 
    TempRs.Open Sql, Conn, 1, 1 
    If Not TempRs.EOF Then
        '根据名称获得网址
        If FieldName = "[url]" Then
            FileName = Trim(TempRs("FileName")) 
            If FileName = "" Then FileName = "MainInfo" & TempRs("Id") '文件名称为空则使用默认
            Url = GetHandleFileName(FileName) 
            Url = WebDebug(Url, "act=MainInfo&Id=" & TempRs("Id")) '调试
            GetMainParam = Url 
        ElseIf FieldName <> "" Then
            GetMainParam = TempRs(FieldName) 
        End If 
    End If : TempRs.Close 
End Function
 
%> 
