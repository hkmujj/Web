<% 
'新网站函数

'特殊字符替换
Function specialStrReplace(ByVal content)
    content = Replace(content, "\|", "[$特殊字符A]$") 
    content = Replace(content, "\-", "[$特殊字符B]$") 
    content = Replace(content, "\,", "[$特殊字符C]$") 
    content = Replace(content, "\'", "[$特殊字符D]$") 
    content = Replace(content, "\""", "[$特殊字符E]$") 
    specialStrReplace = content 
End Function 
'解特殊字符替换
Function unSpecialStrReplace(ByVal content, startStr)
    content = Replace(content, "[$特殊字符A]$", startStr & "|") 
    content = Replace(content, "[$特殊字符B]$", startStr & "-") 
    content = Replace(content, "[$特殊字符C]$", startStr & ",") 
    content = Replace(content, "[$特殊字符D]$", startStr & "'") 
    content = Replace(content, "[$特殊字符E]$", startStr & """") 
    unSpecialStrReplace = content 
End Function 

'栏目类型处理 首页|文本|产品|新闻|视频|下载|案例|留言|反馈|招聘|订单
Function handleColumnType(columnName)
    Dim s 
    Select Case columnName
        Case "首页" : s = "home"
        Case "文本" : s = "text"
        Case "产品" : s = "product"
        Case "新闻" : s = "news"
        Case "视频" : s = "video"
        Case "下载" : s = "download"
        Case "案例" : s = "case"
        Case "留言" : s = "message"
        Case "反馈" : s = "feedback"
        Case "招聘" : s = "job"
        Case "订单" : s = "order"
    End Select
    handleColumnType = s 
End Function

%> 
