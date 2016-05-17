<% 
'编码互换 GB2312与UTF-8转换

'UTF转GB23
Function UTF2GB(ByVal UTFStr)
    Dim Dig, GBStr 
    For Dig = 1 To Len(UTFStr)
        '如果UTF8编码文字以%开头则进行转换
        If Mid(UTFStr, Dig, 1) = "%" Then
            'UTF8编码文字大于8则转换为汉字
            If Len(UTFStr) >= Dig + 8 Then
                GBStr = GBStr & ConvChinese(Mid(UTFStr, Dig, 9)) 
                Dig = Dig + 8 
            Else
                GBStr = GBStr & Mid(UTFStr, Dig, 1) 
            End If 
        Else
            GBStr = GBStr & Mid(UTFStr, Dig, 1) 
        End If 
    Next 
    UTF2GB = GBStr 
End Function
 
'检测是否可以把UTF转GB2312
Function CheckUTFToGB2312(ByVal UTFStr)
    Dim Dig, GBStr 
    CheckUTFToGB2312 = True 
    For Dig = 1 To Len(UTFStr)
        '如果UTF8编码文字以%开头则进行转换
        If Mid(UTFStr, Dig, 1) = "%" Then
            'UTF8编码文字大于8则转换为汉字
            If Len(UTFStr) >= Dig + 8 Then
                If ConvChinese(Mid(UTFStr, Dig, 9)) = "[出错Error]" Then
                    CheckUTFToGB2312 = False : Exit Function 
                End If 
                Dig = Dig + 8 
            Else
                GBStr = GBStr & Mid(UTFStr, Dig, 1) 
            End If 
        Else
            GBStr = GBStr & Mid(UTFStr, Dig, 1) 
        End If 
    Next 
End Function
 
'UTF8编码文字将转换为汉字 配合上面
Function ConvChinese(X)
    On Error Resume Next 
    Err.Clear 
    Dim ArrStr, I, J, DigS, Unicode 
    ArrStr = Split(Mid(X, 2), "%") 
    I = 0 
    J = 0 
    For I = 0 To UBound(ArrStr)
        ArrStr(I) = c16to2(ArrStr(I)) 
    Next 
    For I = 0 To UBound(ArrStr) - 1
        DigS = InStr(ArrStr(I), "0") 
        Unicode = "" 
        For J = 1 To DigS - 1
            If J = 1 Then
                ArrStr(I) = Right(ArrStr(I), Len(ArrStr(I)) - DigS) 
                Unicode = Unicode & ArrStr(I) 
            Else
                I = I + 1 
                ArrStr(I) = Right(ArrStr(I), Len(ArrStr(I)) - 2) 
                Unicode = Unicode & ArrStr(I) 
            End If 
        Next 

        If Len(c2to16(Unicode)) = 4 Then
            ConvChinese = ConvChinese & ChrW(Int("&H" & c2to16(Unicode))) 
        Else
            ConvChinese = ConvChinese & Chr(Int("&H" & c2to16(Unicode))) 
        End If 
        If Err <> 0 Then ConvChinese = "[出错Error]" : Exit Function '出错就退出
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
    Dim X 
    Dim nAsc 
    '如果输入参数为空，则退出函数
    If szInput = "" Then
        toUTF8 = szInput 
        Exit Function 
    End If 
    '开始转换
    For X = 1 To Len(szInput)
        '利用mid函数分拆GB编码文字
        wch = Mid(szInput, X, 1) 
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
Function chinese2unicode(Str)
    Dim I 
    Dim Str_one 
    Dim Str_unicode 
    If(IsNull(Str)) Then
        Exit Function 
    End If 
    For I = 1 To Len(Str)
        Str_one = Mid(Str, I, 1) 
        Str_unicode = Str_unicode & Chr(38) 
        Str_unicode = Str_unicode & Chr(35) 
        Str_unicode = Str_unicode & Chr(120) 
        Str_unicode = Str_unicode & Hex(AscW(Str_one)) 
        Str_unicode = Str_unicode & Chr(59) 
    Next 
    chinese2unicode = Str_unicode 
End Function
 

'判断是否为有效的十六进制代码
Function isvalidhex(Str)
    Dim C 
    isvalidhex = True 
    Str = UCase(Str) 
    If Len(Str) <> 3 Then isvalidhex = False : Exit Function 
    If Left(Str, 1) <> "%" Then isvalidhex = False : Exit Function 
    C = Mid(Str, 2, 1) 
    If Not(((C >= "0") and(C <= "9")) Or((C >= "A") and(C <= "Z"))) Then isvalidhex = False : Exit Function 
    C = Mid(Str, 3, 1) 
    If Not(((C >= "0") and(C <= "9")) Or((C >= "A") and(C <= "Z"))) Then isvalidhex = False : Exit Function 
End Function
 
'---------------------- 自己文件引用 ----------------------------
'作用同JS的escape一样
Function escape(Str)
    Dim I, S, C, A 
    S = "" 
    For I = 1 To Len(Str)
        C = Mid(Str, I, 1) 
        A = AscW(C) 
        If(A >= 48 And A <= 57) Or(A >= 65 And A <= 90) Or(A >= 97 And A <= 122) Then
            S = S & C 
        ElseIf InStr("@*_-./", C) > 0 Then
            S = S & C 
        ElseIf A > 0 And A < 16 Then
            S = S & "%0" & Hex(A) 
        ElseIf A >= 16 And A < 256 Then
            S = S & "%" & Hex(A) 
        Else
            S = S & "%u" & Hex(A) 
        End If 
    Next 
    escape = S 
End Function
 

'作用同JS的unescape一样
Function unescape(Str)
    Dim I, NewStr, C, HaveChar, LastChar, next_1_c, next_1_num 
    HaveChar = False 
    LastChar = "" 
    NewStr = "" 
    For I = 1 To Len(Str)
        C = Mid(Str, I, 1) 
        If C = "+" Then
            NewStr = NewStr & " " 
        ElseIf Mid(Str, I, 2) = "%u" And I <= Len(Str) - 5 Then
            If IsNumeric("&H" & Mid(Str, I + 2, 4)) Then
                NewStr = NewStr & ChrW(CInt("&H" & Mid(Str, I + 2, 4))) 
                I = I + 5 
            Else
                NewStr = NewStr & C 
            End If 
            'ElseIf c="%" and i<=Len(str)-2 Then
            'If IsNumeric("&H" & Mid(str,i+1,2)) Then
            'newstr = newstr & ChrW(CInt("&H" & Mid(str,i+1,2)))
            'i = i+2
            'Else
            'newstr = newstr & c
        'End If
        ElseIf C = "%" And I <= Len(Str) - 2 Then
            next_1_c = Mid(Str, I + 1, 2) 
            If IsNumeric("&H" & next_1_c) Then
                next_1_num = CInt("&H" & next_1_c) 
                If HaveChar = True Then
                    HaveChar = False 
                    NewStr = NewStr & Chr(CInt("&H" & LastChar & next_1_c)) 
                Else
                    If Abs(next_1_num) <= 127 Then
                        NewStr = NewStr & Chr(next_1_num) 
                    Else
                        HaveChar = True 
                        LastChar = next_1_c 
                    End If 
                End If 
                I = I + 2 
            Else
                NewStr = NewStr & C 
            End If 
        Else
            NewStr = NewStr & C 
        End If 
    Next 
    unescape = NewStr 
End Function
 
'汉字乱码如：（中国，&#x4E2D;&#x56FD;）
Function ChineseToUnicode(Str)
    Dim I, S, C 
    For I = 1 To Len(Str)
        S = Mid(Str, I, 1) 
        S = "&#x" & Hex(AscW(S)) & ";" 
        C = C & S 
    Next 
    ChineseToUnicode = C 
End Function
 
'解汉字乱码
Function UnicodeToChinese(Content)
    Dim SplStr, S, C 
    SplStr = Split(Content, ";") 
    For Each S In SplStr
        If InStr(S, "&#x") > 0 Then
            S = Right(S, Len(S) - 3) 
            C = C & ChrW(CInt("&H" & S)) 
        End If 
    Next 
    UnicodeToChinese = C 
End Function
 

 
'Url加码
Function URLEncode(ByRef StrUrl)
    Dim I 
    Dim TempStr 
    For I = 1 To Len(StrUrl)
        If Asc(Mid(StrUrl, I, 1)) < 0 Then
            TempStr = "%" & Right(CStr(Hex(Asc(Mid(StrUrl, I, 1)))), 2) 
            TempStr = "%" & Left(CStr(Hex(Asc(Mid(StrUrl, I, 1)))), Len(CStr(Hex(Asc(Mid(StrUrl, I, 1))))) - 2) & TempStr 
            URLEncode = URLEncode & TempStr 
        ElseIf(Asc(Mid(StrUrl, I, 1)) >= 65 And Asc(Mid(StrUrl, I, 1)) <= 90) Or(Asc(Mid(StrUrl, I, 1)) >= 97 And Asc(Mid(StrUrl, I, 1)) <= 122) Then
            URLEncode = URLEncode & Mid(StrUrl, I, 1) 
        Else
            URLEncode = URLEncode & "%" & Hex(Asc(Mid(StrUrl, I, 1))) 
        End If 
    Next 
End Function
 
'URL解码   这个好用
Function URLDecode(ByRef StrUrl)
    Dim I 
    If InStr(StrUrl, "%") = 0 Then URLDecode = StrUrl : Exit Function 
    For I = 1 To Len(StrUrl)
        If Mid(StrUrl, I, 1) = "%" Then
            If Val("&H" & Mid(StrUrl, I + 1, 2)) > 127 Then
                URLDecode = URLDecode & Chr(Val("&H" & Mid(StrUrl, I + 1, 2) & Mid(StrUrl, I + 4, 2))) 
                I = I + 5 
            Else
                URLDecode = URLDecode & Chr(Val("&H" & Mid(StrUrl, I + 1, 2))) 
                I = I + 2 
            End If 
        Else
            URLDecode = URLDecode & Mid(StrUrl, I, 1) 
        End If 
    Next 
End Function
 
'GB2312到UTF-8
Function GBtoUTF8(szInput)
    Dim wch, uch, szRet 
    Dim X 
    Dim nAsc 
    '如果输入参数为空，则退出函数
    If szInput = "" Then
        GBtoUTF8 = szInput 
        Exit Function 
    End If 
    '开始转换
    For X = 1 To Len(szInput)
        wch = Mid(szInput, X, 1) 
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
Function EncDecURL(ByVal Content, ByVal SType)
    Dim Ystr, TStr, Temp, I, S, C 
    Ystr = "abcdefghijklmnopqrstuvwxyz0123456789:/.=& ()%" 
    TStr = "9ab1cdefghij234klmnopqrst678uvwxyz:/.05*-$[]@" 
    If SType = "解密" Or SType = "0" Then
        Temp = Ystr 
        Ystr = TStr 
        TStr = Temp 
    End If 
    For I = 1 To Len(Content)
        S = Mid(Content, I, 1) 
        If InStr(Ystr, S) > 0 Then
            S = Mid(TStr, InStr(Ystr, S), 1) 
        End If 
        C = C & S 
    Next 
    EncDecURL = C 
End Function
 
'加密URL (辅助)
Function EncURL(Content)
    EncURL = EncDecURL(Content, "加密") 
End Function
 
'解密URL (辅助)
Function DecURL(Content)
    DecURL = EncDecURL(Content, "解密") 
End Function
 
%> 
