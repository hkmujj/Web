<% 
'加密解密(2014)


'特殊Html上传加密解密 20150121 specialHtmlUploadEncryptionDecrypt(Content,"Decrypt")
Function specialHtmlUploadEncryptionDecrypt(content, sType)
    Dim splStr, splxx, c, s 
    c = "·|[*-24156*]" & vbCrLf 
    splStr = Split(c, vbCrLf) 
    For Each s In splStr
        If InStr(s, "|") > 0 Then
            splxx = Split(s, "|") 
            If sType = "1" Or sType = "解密" Or sType = "Decrypt" Then
                content = Replace(content, splxx(1), splxx(0)) 
            Else
                content = Replace(content, splxx(0), splxx(1)) 
            End If 
        End If 
    Next 
    specialHtmlUploadEncryptionDecrypt = content 
End Function 

'加密ASP代码内容
Function encAspContent(ByVal content)
    Dim splStr, c, s, THStr 
    c = "Str=Str&""|Str=Str & |If | Then|End If|&vbCrlf|Temp |Rs(|Rs.|.AddNew|(""Title"")|(""Content"")|=False|ElseIf|" 
    c = c & "Conn.Execute(""| Exit For|[Product]|.Open|.Close|Exit For|Exit Function|MoveNext:Next:|Str " 
    splStr = Split(c, "|") 
    For Each s In splStr
        If s <> "" Then
            THStr = upperCaseORLowerCase(s) 
            content = Replace(content, Chr(9), "") 'Chr(9) = Tab
            content = Replace(content, s, THStr) 
        End If 
    Next 
    encAspContent = content 
End Function 
'让大小写乱掉
Function upperCaseORLowerCase(ByVal content)
    Dim i, s, c, nRnd 
    c = "" 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        Randomize 
        nRnd = CInt(Rnd() * 1) 
        If nRnd = 0 Then
            c = c & LCase(s) 
        Else
            c = c & UCase(s) 
        End If 
    Next 
    upperCaseORLowerCase = c 
End Function 
'加密  Encryption
Function encCode(ByVal content)
    Dim i, c 
    c = "" 
    For i = 1 To Len(content)
        c = c & "%" & Hex(Asc(Mid(content, i, 1))) 
    Next 
    encCode = c 
End Function 
'解密  Decrypt
Function decCode(ByVal content)
    Dim i, c, splStr 
    c = "" 
    splStr = Split(content, "%") 
    For i = 1 To UBound(splStr)
        If splStr(i) <> "" Then
            c = c & Chr(CInt("&H" & splStr(i))) 
        End If 
    Next 
    decCode = c 
End Function 
'将汉字等转换为&#开头的unicode字符串形式
Public Function toUnicode(str)
    Dim i, j, c, p 
    toUnicode = "" 
    c = "" 
    p = "" 
    For i = 1 To Len(str)
        c = Mid(str, i, 1) 
        j = AscW(c) 
        If j < 0 Then
            j = j + 65536 
        End If 
        If j >= 0 And j <= 128 Then
            If p = "c" Then
                toUnicode = " " & toUnicode 
                p = "e" 
            End If 
            toUnicode = toUnicode & c 
        Else
            If p = "e" Then
                toUnicode = toUnicode & " " 
                p = "c" 
            End If 
            toUnicode = toUnicode & "&#" & j & ";" 
        End If 
    Next 
End Function 
'日文26字母编码
Function japan(ByVal iStr, sType)
    If IsNull(iStr) Or IsEmpty(iStr) Then
        japan = "" 
        Exit Function 
    End If 
    Dim f, i, e 
    If sType = "" Then sType = 0 
    'F=array("ゴ","ガ","ギ","グ","ゲ","ザ","ジ","ズ","ヅ","デ",_
    '"ド","ポ","ベ","プ","ビ","パ","ヴ","ボ","ペ","ブ","ピ","バ",_
    '"ヂ","ダ","ゾ","ゼ")
    'E = Array("Jn0;", "Jn1;", "Jn2;", "Jn3;", "Jn4;", "Jn5;", "Jn6;", "Jn7;", "Jn8;", "Jn9;", "Jn10;", "Jn11;", "Jn12;", "Jn13;", "Jn14;", "Jn15;", "Jn16;", "Jn17;", "Jn18;", "Jn19;", "Jn20;", "Jn21;", "Jn22;", "Jn23;", "Jn24;", "Jn25;")
    e = Split("Jn0;,Jn1;,Jn2;,Jn3;,Jn4;,Jn5;,Jn6;,Jn7;,Jn8;,Jn9;,Jn10;,Jn11;,Jn12;,Jn13;,Jn14;,Jn15;,Jn16;,Jn17;,Jn18;,Jn19;,Jn20;,Jn21;,Jn22;,Jn23;,Jn24;,Jn25;", ",") 

    'F = Array(Chr( -23116), Chr( -23124), Chr( -23122), Chr( -23120),    Chr(-23118), Chr( -23114), Chr( -23112), Chr( -23110),     Chr(-23099), Chr( -23097), Chr( -23095), Chr( -23075),   Chr(-23079), Chr( -23081), Chr( -23085), Chr( -23087),  Chr(-23052), Chr( -23076), Chr( -23078), Chr( -23082),  Chr(-23084), Chr( -23088), Chr( -23102), Chr( -23104), Chr(-23106), Chr( -23108))
    f = Split(Chr( -23116) & "," & Chr( -23124) & "," & Chr( -23122) & "," & Chr( -23120) & "," & Chr( -23118) & "," & Chr( -23114) & "," & Chr( -23112) & "," & Chr( -23110) & "," & Chr( -23099) & "," & Chr( -23097) & "," & Chr( -23095) & "," & Chr( -23075) & "," & Chr( -23079) & "," & Chr( -23081) & "," & Chr( -23085) & "," & Chr( -23087) & "," & Chr( -23052) & "," & Chr( -23076) & "," & Chr( -23078) & "," & Chr( -23082) & "," & Chr( -23084) & "," & Chr( -23088) & "," & Chr( -23102) & "," & Chr( -23104) & "," & Chr( -23106) & "," & Chr( -23108), ",") 
    japan = iStr 
    For i = 0 To 25
        If sType = 0 Then
            japan = Replace(japan, f(i), e(i)) 
        Else
            japan = Replace(japan, e(i), f(i)) 
        End If 
    Next 
End Function 
'日文26字母 加密
Function japan26(iStr)
    japan26 = japan(iStr, 0) 
End Function 
'日文26字母 解密
Function unJapan26(iStr)
    unJapan26 = japan(iStr, 1) 
End Function 
'处理内容让它成为纯HTML代码
Function handleHTML(ByVal content)
    'Content = Replace(Content, "&", "&amp;")
    content = Replace(content, "<", "&lt;") 
    handleHTML = content 
End Function 
'解开 处理内容让它成为纯HTML代码
Function unHandleHTML(ByVal content)
    'Content = Replace(Content, "&amp;", "&")
    content = Replace(content, "&lt;", "<") 
    unHandleHTML = content 
End Function 
'小写加密   [可扩展为大写与数字]
Function lcaseEnc(str)
    Dim i, n, s, c 
    c = "" 
    For i = 1 To Len(str)
        s = Mid(str, i, 1) 
        n = AscW(s) 
        If n >= 97 And n <= 122 Then
            c = c & ChrW(n + 1) 
        Else
            c = c & s 
        End If 
    Next 
    c = Replace(c, vbCrLf, "＠") 
    lcaseEnc = c 
End Function 
'小写解密
Function lcaseDec(str)
    Dim i, n, s, c 
    c = "" 
    For i = 1 To Len(str)
        s = Mid(str, i, 1) 
        n = AscW(s) 
        If n >= 97 And n <= 123 Then
            c = c & ChrW(n - 1) 
        Else
            c = c & s 
        End If 
    Next 
    c = Replace(c, "＠", vbCrLf) 
    lcaseDec = c 
End Function 

'html转换成js
Function htmlToJs(ByVal c)
    c = Replace("" & c, "\", "\\") 
    c = Replace(c, "/", "\/") 
    c = Replace(c, "'", "\'") 
    c = Replace(c, """", "\""") 
    c = Join(Split(c, vbCrLf), """);" & vbCrLf & "document.write(""") 
    c = "document.write(""" & c & """);" 
    htmlToJs = c 
End Function 
'js转换成html
Function jsToHtml(ByVal c)
    c = Replace(c, "document.write(""", "") 
    c = Replace(c, """);", "") 
    c = Replace(c, "\""", """") 
    c = Replace(c, "\'", "'") 
    c = Replace(c, "\/", "/") 
    c = Replace(c, "\\", "\") 
    jsToHtml = c 
End Function 
'html转换成Asp
Function htmlToAsp(ByVal c)
    c = Replace(c, """", """""") 
    c = Join(Split(c, vbCrLf), """)" & vbCrLf & "Response.Write(""") 
    c = "Response.Write(""" & c & """)" 
    htmlToAsp = c 
End Function 
'Html转Asp变量存储
Function htmlToAspDim(ByVal c)
    c = Replace(c, """", """""") 
    c = Join(Split(c, vbCrLf), """" & vbCrLf & "C=C & """) 
    c = "C=C & """ & c & """" 
    htmlToAspDim = c 
End Function 
'Asp转换成html
Function aspToHtml(ByVal c)
    c = Replace(c, "Response.Write(""", "") 
    c = Replace(c, """""", """") 
    aspToHtml = c 
End Function 
'文件命名规则
Function setFileName(ByVal fileName)
    Dim i, s,tStr,splstr,splReplace
	splstr=array("\","/",":","*","?","""","<",">","|",".",",")					'换这种方法是为了与PHP版通用 20160511
	splReplace=array("撇","揦","：","星","？","“","左","右","横","。","，")
	for i =0 to ubound(splstr)
		s=splstr(i)
		tStr=splReplace(i)
		fileName = Replace(fileName, s, tStr) 
	next 
    fileName = Replace(fileName, "&nbsp;", " ") 
    fileName = Replace(fileName, "&quot;", "双") 
    fileName = Replace(fileName, vbCrLf, "") 
    setFileName = fileName 
End Function 
'文件命名规则解开
Function unSetFileName(ByVal fileName)
    Dim i, s,tStr,splstr,splReplace

	splstr=array("\","/",":","*","?","""","<",">","|",".",",")
	splReplace=array("撇","揦","：","星","？","“","左","右","横","。","，")
	for i =0 to ubound(splstr)
		s=splstr(i)
		tStr=splReplace(i)
		fileName = Replace(fileName, tStr,s) 
	next 
	
    unSetFileName = fileName 
End Function 

'把Html转成ASP，并且字符转成Chr(*)样式
Function HTMLToAscChr(title)
    Dim i, s, c 
    c = "" 
    For i = 1 To Len(title)
        s = Mid(title, i, 1) 
        c = c & "Chr(" & Asc(s) & ")&" 
    Next 
    If c <> "" Then c = Left(c, Len(c) - 1) 
    HTMLToAscChr = c 
'HTMLToAscChr = "<" & "%=" & C & "%" & ">"
End Function 
'解密AscChr
Function unHTMLToAscChr(content)
    Dim i, s, c, splStr, temp 
    c = content : temp = "" 
    c = Replace(c, "Chr(", "") 
    c = Replace(c, ")&", " ") 
    c = Replace(c, ")", " ") 
    splStr = Split(c, " ") 
    For i = 0 To UBound(splStr) - 1
        s = splStr(i) 
        temp = temp & Chr(s) 
    Next 
    unHTMLToAscChr = temp 
End Function 

'变量移位
Function variableDisplacement(content, nPass)
    Dim c, i, s, LetterGroup, DigitalGroup, nLetterGroup, nDigitalGroup, nLetterLen, nDigitalLen, nX 
    '字母组
    'LetterGroup="abcdefghijklmnopqrstuvwxyz"
    LetterGroup = "yzoehijklmfgqrstuvpabnwxcd" 
    '字母长
    nLetterGroup = Len(LetterGroup) 
    '数字组
    'DigitalGroup="0123456789"
    DigitalGroup = "4539671820" 
    '数字长
    nDigitalGroup = Len(DigitalGroup) 
    c = "" 
    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        nLetterLen = InStr(LetterGroup, s) 
        nDigitalLen = InStr(DigitalGroup, s) 
        '字母处理
        If nLetterLen > 0 Then
            nX = nLetterLen + nPass 
            If nX > nLetterGroup Then
                nX = nX - nLetterGroup 
            ElseIf nX <= 0 Then
                'Call Echo("nX",nX & "," & (nLetterGroup - nX) & "/" & nLetterGroup)
                nX = nX + nLetterGroup 
            End If 
            s = Mid(LetterGroup, nX, 1) 
        '数字处理
        ElseIf nDigitalLen > 0 Then
            nX = nDigitalLen + nPass 
            If nX > nDigitalGroup Then
                nX = nX - nDigitalGroup 
            ElseIf nX <= 0 Then
                'Call Echo("nX",nX & "," & (nLetterGroup - nX) & "/" & nLetterGroup)
                nX = nX + nDigitalGroup 
            End If 
            s = Mid(DigitalGroup, nX, 1) 


        End If 
        c = c & s 
    Next 
    variableDisplacement = c 
End Function 

'【删除此行start】

'Xor加密
Function xorEnc(code, n)
    Dim c, s1, s2, s3, i 
    c = code 
    s1 = Len(c) : s3 = "" 
    For i = 0 To s1 - 1
        s2 = AscW(Right(c, s1 - i)) Xor n 
        s3 = s3 & ChrW(Int(s2)) 
    Next 
    'Chr(34) 就是等于(") 防止出错 因为"在ASP里出错
    s3 = Replace(s3, ChrW(34), "ㄨ") 
    xorEnc = s3 
End Function 
'Xor解密
Function xorDec(code, n)
    Dim c, s1, s2, s3, i 
    c = code : s3 = "" 
    c = Replace(c, "ㄨ", ChrW(34)) 
    s1 = Len(c) 
    For i = 0 To s1 - 1
        s2 = AscW(Right(c, s1 - i)) Xor n 
        s3 = s3 & ChrW(Int(s2)) 
    Next 
    s3 = Replace(s3, "＠", vbCrLf) 
    xorDec = s3 
End Function 


'处理Eval加密
Function handleEvalAddDec(ByVal content, AddStr)
    Dim c 
    c = AddStr & "Execute unDec_log(""" & encCode(content) & """)" & AddStr 
    c = c & "Function unDec_log(ByVal Content)" & AddStr 
    c = c & "    Dim I, C, SplStr " & AddStr 
    c = c & "    SplStr = Split(Content, ""%"") " & AddStr 
    c = c & "    For I = 1 To UBound(SplStr)" & AddStr 
    c = c & "        C = C & Chr(CInt(""&H"" & SplStr(I))) " & AddStr 
    c = c & "    Next " & AddStr 
    c = c & "    unDec_log = C " & AddStr 
    c = c & "End Function " & AddStr 
    handleEvalAddDec = c 
End Function 
'Eval加密
Function evalAddDec(ByVal content)
    evalAddDec = handleEvalAddDec(content, vbCrLf) 
End Function 
'Eval加密
Function VBEvalAddDec(ByVal content)
    VBEvalAddDec = escape(handleEvalAddDec(content, ":")) 
End Function 

'Xor解密    这个暂时留着
Function XDec(code)
    Dim c, s1, s2, s3, i 
    c = code 
    c = Replace(c, "ㄨ", Chr(34)) 
    s1 = Len(c) : s2 = "" : s3 = "" 
    For i = 0 To s1 - 1
        s2 = Asc(Right(c, s1 - i)) Xor 20 
        s3 = s3 & Chr(Int(s2)) 
    Next 
    s3 = Replace(s3, "＠", vbCrLf) 
    XDec = s3 
End Function 


'RGB转十六进制
Public Function toHexColor(scrColor)
    Dim HexColor 
    toHexColor = "" 
    HexColor = CStr(Hex(scrColor)) 
    Select Case Len(HexColor)
        Case 1 : toHexColor = "&H00000" & HexColor
        Case 2 : toHexColor = "&H0000" & HexColor
        Case 3 : toHexColor = "&H000" & HexColor
        Case 4 : toHexColor = "&H00" & HexColor
        Case 5 : toHexColor = "&H0" & HexColor
        Case 6 : toHexColor = "&H" & HexColor
    End Select
End Function 
'十六进制与RGB
Public Function colorRGB(color)
    Dim ByN, ByN2 
    ByN = "" : ByN2 = "" 
    Dim arrayStr(2) 
    arrayStr(0) =(color Mod ByN) 
    arrayStr(1) =((color Mod ByN2) \ ByN) 
    arrayStr(2) =(color \ ByN2) 
    colorRGB = arrayStr 
End Function 

'简单加密
Function simpleEnc(c)
    c = Replace(c, "x", "輕") 
    c = Replace(c, "f", "蚛") 
    c = Replace(c, "s", "鋋") 
    c = Replace(c, "h", "孿") 
    c = Replace(c, "t", "謣") 
    c = Replace(c, "o", "薺") 
    c = Replace(c, "<", "讞") 
    c = Replace(c, ">", "旻") 
    c = Replace(c, "a", "艓") 
    c = Replace(c, "c", "襋") 
    c = Replace(c, "n", "贁") 
    c = Replace(c, "m", "曁") 
    c = Replace(c, vbCrLf, "樞") 
    simpleEnc = c 
End Function 
'简单解密
Function simpleDec(c)
    c = Replace(c, "輕", "x") 
    c = Replace(c, "蚛", "f") 
    c = Replace(c, "鋋", "s") 
    c = Replace(c, "孿", "h") 
    c = Replace(c, "謣", "t") 
    c = Replace(c, "薺", "o") 
    c = Replace(c, "讞", "<") 
    c = Replace(c, "旻", ">") 
    c = Replace(c, "艓", "a") 
    c = Replace(c, "襋", "c") 
    c = Replace(c, "贁", "n") 
    c = Replace(c, "曁", "m") 
    c = Replace(c, "樞", vbCrLf) 
    simpleDec = c 
End Function 
'加密ASP内容成Eval来运行
Function encAspEvalRun(ByVal content)
    Dim splStr, i, c 
    '去除<% % >
    content = Replace(content, "<" & "%", "") 
    content = Replace(content, "%" & ">", "") 
    content = Replace(content, """", """""") 
    content = Replace(content, "x", "輕") 
    content = Replace(content, "f", "蚛") 
    content = Replace(content, "s", "鋋") 
    content = Replace(content, "o", "薺") 
    content = Replace(content, "e", "療") 
    splStr = Split(content, vbCrLf) 
    c = "" 
    For i = 0 To UBound(splStr)
        If InStr(splStr(i), "'") = False And Len(splStr(i)) > 1 Then
            c = c & splStr(i) & "⊙" 
        End If 
    Next 
    c = Left(c, Len(c) - 1) 
    c = "Execute MyEval(""" & c & """)" & vbCrLf 
    c = c & "Function MyEval(HtmlStr)" & vbCrLf 
    c = c & vbTab & "MyEval=Replace(HtmlStr,""⊙"",vbCrlf)" & vbCrLf 
    c = c & vbTab & "MyEval=Replace(MyEval,""輕"",""x"")" & vbCrLf 
    c = c & vbTab & "MyEval=Replace(MyEval,""蚛"",""f"")" & vbCrLf 
    c = c & vbTab & "MyEval=Replace(MyEval,""鋋"",""s"")" & vbCrLf 
    c = c & vbTab & "MyEval=Replace(MyEval,""薺"",""o"")" & vbCrLf 
    c = c & vbTab & "MyEval=Replace(MyEval,""療"",""e"")" & vbCrLf 
    c = c & "End Function" & vbCrLf 
    encAspEvalRun = "<" & "%" & vbCrLf & c & "%" & ">" 
End Function 

'【删除此行end】
%>    

