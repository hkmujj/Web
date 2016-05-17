<% 
'Check验证 (2013,10,26)

'检测URL文件名称是否带参数，如:.js?  .css?  用法 checkUrlFileNameParam("http://sdfsd.com/aaa.js","js|css|")
function checkUrlFileNameParam(httpurl,sList)
	dim url,splstr,searchStr
	url=lcase(httpurl)
	sList=lcase(sList)
	splstr=split(sList,"|")
	for each searchStr in splstr
		if searchStr<>"" then
			searchStr="."& searchStr &"?"
			'call echo("searchStr",searchStr)
			if InStr(url,searchStr) > 0 then
				checkUrlFileNameParam=true
				exit function
			end if
		end if
	next
	checkUrlFileNameParam=false
end function

'是大写 20160105
Function isUCase(content)
    Dim i, s 
    isUCase = True 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", s) = False Then
            isUCase = False 
            Exit Function 
        End If 
    Next 
End Function 
'是小写 20160105
Function isLCase(content)
    Dim i, s 
    isLCase = True 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        If InStr("abcdefghijklmnopqrstuvwxyz", s) = False Then
            isLCase = False 
            Exit Function 
        End If 
    Next 
End Function 


'检测错误
Sub checkErr(errMsg)
    If Err.Number <> 0 Then
        Call echo(Err.Number, "自定义错误信息：" & errMsg & "<br>系统错误信息：" & Err.Description & "<br>最初生成错误的对象或应用程序的名称：" & Err.Source) 
    End If 
End Sub 

'把字符true转真类型 false转假类型
Function strToTrueFalse(ByVal content)
    content = Trim(LCase(content)) 
    strToTrueFalse = IIF(content = "true", True, False) 
End Function 
'把字符true转1类型 false转0类型
Function strTrueFalseToInt(ByVal content)
    content = Trim(LCase(content)) 
    strTrueFalseToInt = IIF(content = "true", 1, 0) 
End Function 
'检查换行
Function checkVbCrlf(content)
    checkVbCrlf = False 
    If InStr(content, vbCrLf) > 0 Then checkVbCrlf = True 
End Function 
'检查换行    辅助
Function checkBr(content)
    checkBr = checkVbCrlf(content) 
End Function 

'判断数字奇偶
Function isParity(Numb)
    isParity = "" 
    If Not IsNumeric(Numb) Then
        isParity = "这不是一个数字啊" 
        Exit Function 
    End If 
    If Numb Mod 2 = 0 Then
        isParity = 0 
    Else
        isParity = 1 
    End If 
End Function 
'检测eval正确性
Function checkEval(content)
    On Error Resume Next 
    Call Execute(content) 
    If Err.Number = 0 Then
        checkEval = True 
    Else
        checkEval = False 
    End If 
End Function 
'********************************************************
'过滤SQL非法字符并格式化html代码
'********************************************************
Function replace_SQLText(fString)
    If IsNull(fString) Then
        replace_SQLText = "" 
        Exit Function 
    Else
        fString = Trim(fString) 
        fString = Replace(fString, "'", "''") 
        fString = Replace(fString, ";", "；") 
        fString = Replace(fString, "--", "—") 
        fString = displayHtml(fString) 
        replace_SQLText = fString 
    End If 
End Function 
'********************************************************
'检查是否外部提交数据
'********************************************************
Function chkPost()
    Dim Server_v1, Server_v2 
    chkPost = False 
    Server_v1 = CStr(Request.ServerVariables("HTTP_REFERER")) 
    Server_v2 = CStr(Request.ServerVariables("SERVER_NAME")) 
    Call echo(Server_v1, Server_v2) 
    If Mid(Server_v1, 8, Len(Server_v2)) <> Server_v2 Then
        chkPost = False 
    Else
        chkPost = True 
    End If 
End Function 
'Response.Write(IsMail("asdf@sdf.com"))
'正则表达验证邮箱
Function isMail(email)
    isMail = True 
    Dim Rep 
    Set Rep = CreateObject("VBscript.RegExp")
        Rep.pattern = "([\.a-zA-Z0-9_-]){2,10}@([a-zA-Z0-9_-]){2,10}(\.([a-zA-Z0-9]){2,}){1,4}$" 
        isMail = Rep.Test(email) 
    Set Rep = Nothing 
End Function 
'邮箱验证第二种
Function isValidEmail(email)
    Dim names, Name, i, c 
    isValidEmail = True 
    names = Split(email, "@") 
    If UBound(names) <> 1 Then isValidEmail = False : Exit Function 
    For Each Name In names
        If Len(Name) <= 0 Then isValidEmail = False : Exit Function 
        For i = 1 To Len(Name)
            c = LCase(Mid(Name, i, 1)) 
            If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then isValidEmail = False : Exit Function 
        Next 
        If Left(Name, 1) = "." Or Right(Name, 1) = "." Then isValidEmail = False : Exit Function 
    Next 
    If InStr(names(1), ".") <= 0 Then isValidEmail = False : Exit Function 
    i = Len(names(1)) - InStrRev(names(1), ".") 
    If i <> 2 And i <> 3 Then isValidEmail = False : Exit Function 
    If InStr(email, "..") > 0 Then isValidEmail = False 
End Function 
'检测为有效字符
Function isCode(ByVal content)
    Dim c 
    c = Replace(content, " ", "") 
    c = Replace(c, Chr(13), "") 
    c = Replace(c, Chr(10), "") 
    c = Replace(c, vbTab, "") 
    isCode = IIF(c <> "", True, False) 
End Function 
'测试是否为数字
Function checkNumber(ByVal content)
    Dim i, s 
    If Len(content) = 0 Then checkNumber = False : Exit Function 
    checkNumber = True 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        If InStr("0123456789", LCase(s)) = False Then
            checkNumber = False 
            Exit Function 
        End If 
    Next 
End Function
'测试是否为数字 (辅助)
Function isNumber(ByVal content)
    isNumber = checkNumber(content) 
End Function 
'测试是否为字母
Function checkABC(ByVal content)
    Dim i, s 
    checkABC = True 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        If InStr("abcdefghijklmnopqrstuvwxyz", LCase(s)) = False Then
            checkABC = False 
            Exit Function 
        End If 
    Next 
End Function 
'获得字符长度 汉字算两个字符
Function getLen(content)
    Dim i, s, n 
    n = 0 
    For i = 1 To Len(content)
        s = Asc(Mid(CStr(content), i, 1)) 
        If s < 0 Then
            n = n + 2 
        Else
            n = n + 1 
        End If 
    Next 
    getLen = n 
End Function 
'测试是否为时间类型
Function checkTime(DateTime)
    checkTime = IIF(IsDate(DateTime) = False, False, True) 
End Function 
'判断是否为空
Function isNul(ByVal s)
    On Error Resume Next : If Err.Number <> 0 Then Err.Clear 
    isNul = False 
    Select Case VarType(s)
        Case vbEmpty, vbNull
            isNul = True : Exit Function 
        Case vbString
            If s = "" Then isNul = True : Exit Function 
        Case vbObject
            Select Case TypeName(s)
                Case "Nothing", "Empty"
                    isNul = True : Exit Function 
                Case "Recordset"
                    If s.State = 0 Then isNul = True : Exit Function 
                    If s.BOF And s.EOF Then isNul = True : Exit Function 
                Case "Dictionary"
                    If s.Count = 0 Then isNul = True : Exit Function 
            End Select
    Case vbArray, 8194, 8204, 8209
        If UBound(s) = -1 Then isNul = True : Exit Function 
    End Select 
    On Error Goto 0 
End Function 


'****************************************************
'函数名：FoundInArr
'作  用：检查一个数组中所有元素是否包含指定字符串
'时  间：2011年10月13日
'参  数： strArr
'strToFind
'strSplit
'返回值：字符串
'调  试：SHtml=R_("{测试}",Function FoundInArr(strArr, strToFind, strSplit))
'****************************************************
Function foundInArr(strArr, strToFind, strSplit)
    Dim arrTemp, i 
    foundInArr = False 
    If InStr(strArr, strSplit) > 0 Then
        arrTemp = Split(strArr, strSplit) 
        For i = 0 To UBound(arrTemp)
            If LCase(Trim(arrTemp(i))) = LCase(Trim(strToFind)) Then
                foundInArr = True : Exit For 
            End If 
        Next 
    Else
        If LCase(Trim(strArr)) = LCase(Trim(strToFind)) Then foundInArr = True 
    End If 
End Function 
%>  
