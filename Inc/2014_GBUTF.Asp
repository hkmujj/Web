<% 
'编码互换 GB2312与UTF-8转换

'UTF转GB23
Function UTF2GB(ByVal UTFStr)
    Dim dig, GBStr 
    For dig = 1 To Len(UTFStr)
        '如果UTF8编码文字以%开头则进行转换
        If Mid(UTFStr, dig, 1) = "%" Then
            'UTF8编码文字大于8则转换为汉字
            If Len(UTFStr) >= dig + 8 Then
                GBStr = GBStr & convChinese(Mid(UTFStr, dig, 9)) 
                dig = dig + 8 
            Else
                GBStr = GBStr & Mid(UTFStr, dig, 1) 
            End If 
        Else
            GBStr = GBStr & Mid(UTFStr, dig, 1) 
        End If 
    Next 
    UTF2GB = GBStr 
End Function 

'检测是否可以把UTF转GB2312
Function checkUTFToGB2312(ByVal UTFStr)
    Dim dig, GBStr 
    checkUTFToGB2312 = True 
    For dig = 1 To Len(UTFStr)
        '如果UTF8编码文字以%开头则进行转换
        If Mid(UTFStr, dig, 1) = "%" Then
            'UTF8编码文字大于8则转换为汉字
            If Len(UTFStr) >= dig + 8 Then
                If convChinese(Mid(UTFStr, dig, 9)) = "[出错Error]" Then
                    checkUTFToGB2312 = False : Exit Function 
                End If 
                dig = dig + 8 
            Else
                GBStr = GBStr & Mid(UTFStr, dig, 1) 
            End If 
        Else
            GBStr = GBStr & Mid(UTFStr, dig, 1) 
        End If 
    Next 
End Function 

'UTF8编码文字将转换为汉字 配合上面
Function convChinese(x)
    On Error Resume Next 
    Err.Clear 
    Dim arrStr, i, j, digS, unicode 
    arrStr = Split(Mid(x, 2), "%") 
    i = 0 
    j = 0 
    For i = 0 To UBound(arrStr)
        arrStr(i) = c16to2(arrStr(i)) 
    Next 
    For i = 0 To UBound(arrStr) - 1
        digS = InStr(arrStr(i), "0") 
        unicode = "" 
        For j = 1 To digS - 1
            If j = 1 Then
                arrStr(i) = Right(arrStr(i), Len(arrStr(i)) - digS) 
                unicode = unicode & arrStr(i) 
            Else
                i = i + 1 
                arrStr(i) = Right(arrStr(i), Len(arrStr(i)) - 2) 
                unicode = unicode & arrStr(i) 
            End If 
        Next 

        If Len(c2to16(unicode)) = 4 Then
            convChinese = convChinese & ChrW(Int("&H" & c2to16(unicode))) 
        Else
            convChinese = convChinese & Chr(Int("&H" & c2to16(unicode))) 
        End If 
        If Err <> 0 Then convChinese = "[出错Error]" : Exit Function               '出错就退出
    Next 
End Function 



'转UTF8编码，非常精准，好用，(2014 12 11)
Function UTF8_URLEncoding(szInput)
    Dim wch, uch, szRet 
    Dim x 
    Dim nAsc 
    If szInput = "" Then
        UTF8_URLEncoding = szInput 
        Exit Function 
    End If 
    For x = 1 To Len(szInput)
        wch = Mid(szInput, x, 1) 
        nAsc = AscW(wch) 

        If nAsc < 0 Then nAsc = nAsc + 65536 

        If(nAsc And &HFF80) = 0 Then
            szRet = szRet & wch 
        Else
            If(nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80) 
                szRet = szRet & uch 
            Else
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _ 
                Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _ 
                Hex(nAsc And &H3F Or &H80) 
                szRet = szRet & uch 
            End If 
        End If 
    Next 
    UTF8_URLEncoding = szRet 
End Function 

'GB转UTF8--将GB编码文字转换为UTF8编码文字  这个不精准，等以后给废除掉
Function toUTF8(szInput)
    Dim wch, uch, szRet 
    Dim x 
    Dim nAsc 
    '如果输入参数为空，则退出函数
    If szInput = "" Then
        toUTF8 = szInput 
        Exit Function 
    End If 
    '开始转换
    For x = 1 To Len(szInput)
        '利用mid函数分拆GB编码文字
        wch = Mid(szInput, x, 1) 
        '利用ascW函数返回每一个GB编码文字的Unicode字符代码
        '注：asc函数返回的是ANSI 字符代码，注意区别
        nAsc = AscW(wch) 
        If nAsc < 0 Then nAsc = nAsc + 65536 

        If(nAsc And &HFF80) = 0 Then
            szRet = szRet & wch 
        Else
            If(nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc / 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80) 
                szRet = szRet & uch 
            Else
                'GB编码文字的Unicode字符代码在0800 - FFFF之间采用三字节模版
                uch = "%" & Hex((nAsc / 2 ^ 12) Or &HE0) & "%" & _ 
                Hex((nAsc / 2 ^ 6) And &H3F Or &H80) & "%" & _ 
                Hex(nAsc And &H3F Or &H80) 
                szRet = szRet & uch 
            End If 
        End If 
    Next 

    toUTF8 = szRet 
End Function 

'GB转unicode---将GB编码文字转换为unicode编码文字
Function chinese2unicode(str)
    Dim i 
    Dim str_one 
    Dim str_unicode 
    If(IsNull(str)) Then
        Exit Function 
    End If 
    For i = 1 To Len(str)
        str_one = Mid(str, i, 1) 
        str_unicode = str_unicode & Chr(38) 
        str_unicode = str_unicode & Chr(35) 
        str_unicode = str_unicode & Chr(120) 
        str_unicode = str_unicode & Hex(AscW(str_one)) 
        str_unicode = str_unicode & Chr(59) 
    Next 
    chinese2unicode = str_unicode 
End Function 


'判断是否为有效的十六进制代码
Function isvalidhex(str)
    Dim c 
    isvalidhex = True 
    str = UCase(str) 
    If Len(str) <> 3 Then isvalidhex = False : Exit Function 
    If Left(str, 1) <> "%" Then isvalidhex = False : Exit Function 
    c = Mid(str, 2, 1) 
    If Not(((c >= "0") And(c <= "9")) Or((c >= "A") And(c <= "Z"))) Then isvalidhex = False : Exit Function 
    c = Mid(str, 3, 1) 
    If Not(((c >= "0") And(c <= "9")) Or((c >= "A") And(c <= "Z"))) Then isvalidhex = False : Exit Function 
End Function 

'---------------------- 自己文件引用 ----------------------------
'作用同JS的escape一样
Function escape(str)
    Dim i, s, c, a 
    s = "" 
    For i = 1 To Len(str)
        c = Mid(str, i, 1) 
        a = AscW(c) 
        If(a >= 48 And a <= 57) Or(a >= 65 And a <= 90) Or(a >= 97 And a <= 122) Then
            s = s & c 
        ElseIf InStr("@*_-./", c) > 0 Then
            s = s & c 
        ElseIf a > 0 And a < 16 Then
            s = s & "%0" & Hex(a) 
        ElseIf a >= 16 And a < 256 Then
            s = s & "%" & Hex(a) 
        Else
            s = s & "%u" & Hex(a) 
        End If 
    Next 
    escape = s 
End Function 


'作用同JS的unescape一样
Function unescape(str)
    Dim i, newStr, c, haveChar, lastChar, next_1_c, next_1_num 
    haveChar = False 
    lastChar = "" 
    newStr = "" 
    For i = 1 To Len(str)
        c = Mid(str, i, 1) 
        If c = "+" Then
            newStr = newStr & " " 
        ElseIf Mid(str, i, 2) = "%u" And i <= Len(str) - 5 Then
            If IsNumeric("&H" & Mid(str, i + 2, 4)) Then
                newStr = newStr & ChrW(CInt("&H" & Mid(str, i + 2, 4))) 
                i = i + 5 
            Else
                newStr = newStr & c 
            End If 
            'ElseIf c="%" and i<=Len(str)-2 Then
            'If IsNumeric("&H" & Mid(str,i+1,2)) Then
            'newstr = newstr & ChrW(CInt("&H" & Mid(str,i+1,2)))
            'i = i+2
            'Else
            'newstr = newstr & c
        'End If
        ElseIf c = "%" And i <= Len(str) - 2 Then
            next_1_c = Mid(str, i + 1, 2) 
            If IsNumeric("&H" & next_1_c) Then
                next_1_num = CInt("&H" & next_1_c) 
                If haveChar = True Then
                    haveChar = False 
                    newStr = newStr & Chr(CInt("&H" & lastChar & next_1_c)) 
                Else
                    If Abs(next_1_num) <= 127 Then
                        newStr = newStr & Chr(next_1_num) 
                    Else
                        haveChar = True 
                        lastChar = next_1_c 
                    End If 
                End If 
                i = i + 2 
            Else
                newStr = newStr & c 
            End If 
        Else
            newStr = newStr & c 
        End If 
    Next 
    unescape = newStr 
End Function 

'汉字乱码如：（中国，&#x4E2D;&#x56FD;）
Function chineseToUnicode(str)
    Dim i, s, c 
    For i = 1 To Len(str)
        s = Mid(str, i, 1) 
        s = "&#x" & Hex(AscW(s)) & ";" 
        c = c & s 
    Next 
    chineseToUnicode = c 
End Function 

'解汉字乱码
Function unicodeToChinese(content)
    Dim splStr, s, c 
    splStr = Split(content, ";") 
    For Each s In splStr
        If InStr(s, "&#x") > 0 Then
            s = Right(s, Len(s) - 3) 
            c = c & ChrW(CInt("&H" & s)) 
        End If 
    Next 
    unicodeToChinese = c 
End Function 



'Url加码
Function URLEncode(ByRef strUrl)
    Dim i 
    Dim tempStr 
    For i = 1 To Len(strUrl)
        If Asc(Mid(strUrl, i, 1)) < 0 Then
            tempStr = "%" & Right(CStr(Hex(Asc(Mid(strUrl, i, 1)))), 2) 
            tempStr = "%" & Left(CStr(Hex(Asc(Mid(strUrl, i, 1)))), Len(CStr(Hex(Asc(Mid(strUrl, i, 1))))) - 2) & tempStr 
            URLEncode = URLEncode & tempStr 
        ElseIf(Asc(Mid(strUrl, i, 1)) >= 65 And Asc(Mid(strUrl, i, 1)) <= 90) Or(Asc(Mid(strUrl, i, 1)) >= 97 And Asc(Mid(strUrl, i, 1)) <= 122) Then
            URLEncode = URLEncode & Mid(strUrl, i, 1) 
        Else
            URLEncode = URLEncode & "%" & Hex(Asc(Mid(strUrl, i, 1))) 
        End If 
    Next 
End Function 

'URL解码   这个好用
Function URLDecode(ByRef strUrl)
    Dim i 
    If InStr(strUrl, "%") = 0 Then URLDecode = strUrl : Exit Function 
    For i = 1 To Len(strUrl)
        If Mid(strUrl, i, 1) = "%" Then
            If val("&H" & Mid(strUrl, i + 1, 2)) > 127 Then
                URLDecode = URLDecode & Chr(val("&H" & Mid(strUrl, i + 1, 2) & Mid(strUrl, i + 4, 2))) 
                i = i + 5 
            Else
                URLDecode = URLDecode & Chr(val("&H" & Mid(strUrl, i + 1, 2))) 
                i = i + 2 
            End If 
        Else
            URLDecode = URLDecode & Mid(strUrl, i, 1) 
        End If 
    Next 
End Function 

'GB2312到UTF-8
Function GBtoUTF8(szInput)
    Dim wch, uch, szRet 
    Dim x 
    Dim nAsc 
    '如果输入参数为空，则退出函数
    If szInput = "" Then
        GBtoUTF8 = szInput 
        Exit Function 
    End If 
    '开始转换
    For x = 1 To Len(szInput)
        wch = Mid(szInput, x, 1) 
        nAsc = AscW(wch) 
        If nAsc < 0 Then nAsc = nAsc + 65536 
        If(nAsc And &HFF80) = 0 Then
            szRet = szRet & wch 
        Else
            If(nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80) 
                szRet = szRet & uch 
            Else
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _ 
                Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _ 
                Hex(nAsc And &H3F Or &H80) 
                szRet = szRet & uch 
            End If 
        End If 
    Next 
    GBtoUTF8 = szRet 
End Function 

'加密解密URL
Function encDecURL(ByVal content, ByVal sType)
    Dim yStr, tStr, temp, i, s, c 
    yStr = "abcdefghijklmnopqrstuvwxyz0123456789:/.=& ()%" 
    tStr = "9ab1cdefghij234klmnopqrst678uvwxyz:/.05*-$[]@" 
    If sType = "解密" Or sType = "0" Then
        temp = yStr 
        yStr = tStr 
        tStr = temp 
    End If 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        If InStr(yStr, s) > 0 Then
            s = Mid(tStr, InStr(yStr, s), 1) 
        End If 
        c = c & s 
    Next 
    encDecURL = c 
End Function 

'加密URL (辅助)
Function encURL(content)
    encURL = encDecURL(content, "加密") 
End Function 

'解密URL (辅助)
Function decURL(content)
    decURL = encDecURL(content, "解密") 
End Function 

%>  

