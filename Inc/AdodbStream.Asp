<% 
'AdodbStream Adodb.Stream组件操作大全 (2013,9,27)

'检测文件编码
Function checkCode(ByVal fileName)
	checkCode=getFileCharset(fileName)
End Function
'获得文件编码
Function getFileCharset(ByVal fileName)
    On Error Resume Next 
    Dim ObjStream, bintou 
    Call HandlePath(fileName)                                                       '获得完整路径
    Set ObjStream = CreateObject("ADODB.Stream")
        ObjStream.Type = 1 
        ObjStream.Mode = 3 
        ObjStream.Open 
        ObjStream.Position = 0 
        Call ObjStream.loadFromFile(fileName) 
        bintou = ObjStream.read(2)                                                      '以二进制方式读出二个字符出来(不太太)
        If TypeName(bintou) = "Null" Then getFileCharset = "gb2312" : Exit Function 
        If AscB(MidB(bintou, 1, 1)) = &HEF And AscB(MidB(bintou, 2, 1)) = &HBB Then
            getFileCharset = "utf-8" 
        ElseIf AscB(MidB(bintou, 1, 1)) = &HFF And AscB(MidB(bintou, 2, 1)) = &HFE Then
            getFileCharset = "unicode" 
        Else
            getFileCharset = "gb2312" 
        End If 
        ObjStream.Close 
    Set ObjStream = Nothing 
    If Err Then doError Err.Description, "getFileCharset，获得文件编码 函数出错，FileName=" & fileName 

End Function
'获得文件编码Senior
Function seniorCheckCode(ByVal fileName)
    '暂时不这样用，因为慢，但是还要测试20150119
    'CheckCode = CheckFileCode(FileName)            '获得文件编码(这种慢 但精准)
    'Exit Function
    'On Error Resume Next
    Dim ObjStream, bintou 
    Call HandlePath(fileName)                                                       '获得完整路径
    Set ObjStream = CreateObject("ADODB.Stream")
        ObjStream.Type = 1 
        ObjStream.Mode = 3 
        ObjStream.Open 
        ObjStream.Position = 0 
        Call ObjStream.loadFromFile(fileName) 
        bintou = ObjStream.read(2)                                                      '以二进制方式读出二个字符出来(不太太)
        If TypeName(bintou) = "Null" Then seniorCheckCode = CheckFileCode(fileName) : Exit Function '改进2015019
        'N1 = AscB(MidB(bintou, 1, 1)) : N2 = AscB(MidB(bintou, 2, 1))
        'N = N1 & N2
        If AscB(MidB(bintou, 1, 1)) = &HEF And AscB(MidB(bintou, 2, 1)) = &HBB Then
            seniorCheckCode = "UTF-8" 
        ElseIf AscB(MidB(bintou, 1, 1)) = &HFF And AscB(MidB(bintou, 2, 1)) = &HFE Then
            seniorCheckCode = "UNICODE" 
        Else
            seniorCheckCode = CheckFileCode(fileName)                                       '另一种超强获得文件编码方法
        End If 
        'CAll Rw(FileName & " >> Char_Set=" & CheckCode & " Code=" & N1  & "["& Chr(N1) &"]" &  "," & N2 & "["& Chr(N2) &"]" &"(utf-8(239,187),unicode(255,254)<br>")
        ObjStream.Close 
    Set ObjStream = Nothing 
    If Err Then doError Err.Description, "CheckCode，获得文件编码 函数出错，FileName=" & fileName 
End Function 
'数据流读出内容
Function readFile(ByVal fileName, ByVal Char_Set)
    On Error Resume Next 
    If Char_Set = "1" Or UCase(Char_Set) = "GB2312" Then
        Char_Set = "GB2312" 
    ElseIf Char_Set = "0" Or UCase(Char_Set) = "UTF-8" Then
        Char_Set = "UTF-8" 
    ElseIf Char_Set = "2" Or UCase(Char_Set) = "UNICODE" Then
        Char_Set = "UNICODE" 
    Else
        Char_Set = checkCode(fileName) 
        If Char_Set = "" Then Exit Function 
    End If 
    Dim str, stm, f, fso 
    Call HandlePath(fileName)                                                       '获得完整路径
	if checkfile(fileName)=false then 
		readFile=""
		exit function
	end if
    Set stm = CreateObject("ADODB.Stream")
        stm.Type = 2                                                                    '以本模式读取
        stm.Mode = 3 
        stm.Charset = Char_Set 
        stm.Open 
        Set fso = CreateObject("Scripting.FileSystemObject")
            Set f = fso.GetFile(fileName)
                If f.Size > 0 Then
                    Call stm.loadFromFile(fileName) 
                End If 
                str = stm.ReadText 
                stm.Close 
            Set stm = Nothing 
            readFile = str 
            If Err Then doError Err.Description, "ReadFile，数据流读出内容 函数出错，FileName=" & fileName 
        Set fso = Nothing 
    Set stm = Nothing 
End Function 
'数据流读出内容  (辅助)
Function getStext(fileName)
    getStext = readFile(fileName, "")
End Function
'以UTF-8方法读文件内容
function getFTextUTF(fileName)
	 getFTextUTF=readFile(fileName,"utf-8")
end function
'写入内容
Function writeToFile(ByVal fileName, ByVal content, ByVal Char_Set)
    On Error Resume Next 
    If Char_Set = "1" Or UCase(Char_Set) = "GB2312" Then
        Char_Set = "GB2312" 
    ElseIf Char_Set = "0" Or UCase(Char_Set) = "UTF-8" Then
        Char_Set = "UTF-8" 
    ElseIf Char_Set = "2" Or UCase(Char_Set) = "UNICODE" Then
        Char_Set = "UNICODE" 
    Else
        Char_Set = "GB2312" 
    End If 
    'Call Echo("Char_Set",Char_Set)
    Dim stm 
    Call HandlePath(fileName)                                                       '获得完整路径
    Set stm = CreateObject("ADODB.Stream")
        stm.Type = 2                                                                    '以本模式读取
        stm.Mode = 3 
        stm.Charset = Char_Set 
        stm.Open 
        Call stm.WriteText(content) 
        Call stm.SaveToFile(fileName, 2) 
        stm.Flush 
        stm.Close 
        writeToFile = fileName & "写入成功" 
    Set stm = Nothing 
    If Err Then doError Err.Description, "WriteToFile，数据流写入内容 函数出错，FileName=" & fileName & "，Content字符" & Len(content) 
End Function
'创建gb2312文件 
function createFileGBK(byval fileName, byval content)
	call writeToFile(fileName,content,"gb2312")
end function
'创建UTF-8文件（20151201）
function createFileUTF(byval fileName, byval content)
	call writeToFile(fileName,content,"utf-8")
end function
'创建累加文件2
function createAddFile2(filePath,addToStr)
	dim content
	content=readFile(filePath,"gb2312")
	content = content & addToStr
	call writeToFile(filePath,content,"gb2312")
end function
'是UTF或GB2312
'自动获得编码并且判断Html内容是否与读出文件编码一致，一致则退出内容
Function getUGText(ByVal path)
    Dim Char_Set, FanCharSet, TempCharSet 
    Dim IsContent 
    Char_Set = checkCode(path) 
    IsContent = False 
    '二次处理，防读取错误
    If Char_Set = "utf-8" Then
        FanCharSet = "gb2312" 
    ElseIf Char_Set = "gb2312" Then
        FanCharSet = "utf-8" 
    End If 
    TempCharSet = FanCharSet 
    If FanCharSet <> "" Then
        If CheckStr(LCase(getUGText), "<meta [^>]*charset *= *" & FanCharSet) = True Then
            getUGText = readFile(path, TempCharSet) 
            IsContent = True 
        End If 
    End If 
    '没有读出内容则用文件自身编码读出内容
    If IsContent = False Then
        getUGText = readFile(path, Char_Set) 
    End If 
End Function 
'检测Html内容是否设置了编码(当前页没有调用过20150119)
Function checkContentCharset(ByVal content, ByVal Char_Set)
    content = LCase(content) 
    Char_Set = LCase(Char_Set) 
    checkContentCharset = True 
    If InStr(content, "<meta ") > 0 Then
        Dim i, s, TempS, LalType, LalStr 
        For i = 1 To Len(content)
            s = Mid(content, i, 1) 
            If s = "<" Then
                TempS = LCase(Mid(content, i)) 
                LalType = LCase(Mid(TempS, 1, InStr(TempS, " "))) 
                If LalType = "<meta " Then
                    LalStr = Mid(TempS, 1, InStr(TempS, ">")) 
                    If InStr(LalStr, "charset") > 0 Then
                        If InStr(LalStr, Char_Set) = 0 Then checkContentCharset = False 
                    End If 
                End If 
            End If 
        Next 
    End If 
End Function 
'超级文件编码获取
Function checkFileCode(filePath)
    Dim splStr, c, content, ToPath, YuanStr, s, nLen 
    YuanStr = BinaryReadFile(filePath) 

    'If Len(YuanStr) > 500 Then
    'YuanStr = Right(YuanStr, Len(YuanStr) - 30)
    'End If
    '只提取2500个字符
    If Len(YuanStr) > 2500 Then
        YuanStr = Right(YuanStr, 2500) 
    End If 
    nLen = Len(YuanStr) 
    ToPath = filePath & "_Temp" 
    splStr = Split("GB2312|UTF-8|UNICODE", "|") 
    For Each s In splStr
        content = readFile(filePath, s) 
        Call WriteToFile(ToPath, content, s) 
        c = BinaryReadFile(ToPath) 
        c = Right(c, nLen) 
        If YuanStr = c Then
            checkFileCode = s 
            Call deleteFile(ToPath)                                                         '删除这个临时文件
            Exit For 
        End If 
    Next 
    Call deleteFile(ToPath)                                                         '删除这个临时文件
End Function 
'获得文件二进制内容 20150312
Function getFileBinaryContent(filePath)
    Dim stream 
    Call HandlePath(filePath)                                                       '获得完整路径
    Set stream = CreateObject("ADODB.Stream")
        stream.Mode = 3 
        stream.Type = 1 
        stream.Open 
        Call stream.loadFromFile(filePath) 
        getFileBinaryContent = stream.read 
        stream.Close 
    Set stream = Nothing 
'Response.BinaryWrite Stream.Read            '显示二进制文件内容
End Function 
'读取二进制文件处理成数字方式
Function binaryReadFile(filePath)
    Dim i, nR, c, stm 
    Call HandlePath(filePath)                                                       '获得完整路径
    Set stm = CreateObject("ADODB.Stream")
        stm.Mode = 3 
        stm.Type = 1 
        stm.Open 
        Call stm.loadFromFile(filePath) 
        c = "" : i = 0 : nR = 1 
        While i < stm.Size
            c = c & AscB(stm.read(nR)) 
            i = i + nR 
            DoEvents 
        Wend 
    Set stm = Nothing 
    binaryReadFile = c 
End Function 
'读取二进制文件并加密
Function readBinary(filePath, pass)
    Dim i, nR, c, stm 
    If pass = "" Then pass = 0                                                      '如果为空则为零20150119
    Call HandlePath(filePath)                                                       '获得完整路径
    Set stm = CreateObject("ADODB.Stream")
        stm.Mode = 3 
        stm.Type = 1 
        stm.Open 
        Call stm.loadFromFile(filePath) 
        c = "" : i = 0 : nR = 1                                                         '每次读取字符数
        While i < stm.Size
            c = c &(AscB(stm.read(nR)) + pass) & "|" 
            i = i + nR 
        Wend 
        If c <> "" Then c = Left(c, Len(c) - 1) 
    Set stm = Nothing 
    readBinary = c 
End Function 
'解密保存二进制文件
Function decSaveBinary(filePath, content, pass)
    Dim splStr, buf(), i 
    If pass = "" Then pass = 0                                                      '如果为空则为零20150119
    splStr = Split(content, "|") 
    ReDim buf(UBound(splStr)) 
    For i = 0 To UBound(splStr)
        buf(i) = CByte(CInt(splStr(i)) - pass) 
    Next 
    Call WriteBinary(filePath, buf) 
End Function 
'保存二进制文件   call writeBinary(filePath,Request.BinaryRead (Request.TotalBytes))
Sub writeBinary(filePath, Buf)
    Dim i, aBuf, size, bStream 
    Call HandlePath(filePath)                                                       '获得完整路径
    '不是二进制处理成二进制(20151211)
	if typeName(Buf)<>"Byte()" then
		size = UBound(Buf) : ReDim aBuf(size \ 2) 
		For i = 0 To size - 1 Step 2
			aBuf(i \ 2) = ChrW(Buf(i + 1) * 256 + Buf(i))
		Next 
		If i = size Then aBuf(i \ 2) = ChrW(Buf(i)) 
		aBuf = Join(aBuf, "")
	else
		aBuf=Buf
	end if
	
    Set bStream = CreateObject("ADODB.Stream")
        bStream.Type = 1 
        bStream.Open 
        With CreateObject("ADODB.Stream")
            .Type = 2 
            .Open 
            Call.WriteText(aBuf) 
            .Position = 2 
            Call.CopyTo(bStream) 
            .Close 
        End With 
        Call bStream.SaveToFile(filePath, 2) 
        bStream.Close 
    Set bStream = Nothing 
End Sub 



'---------------------------- 待定 ----------------------------
'获得文件编码 超级快速版 (从ASPBOX引用过来)  不要用
Public Function getFileCharset2(ByVal filePath)
    Dim strFileHead 
    'On Error Resume Next
    If Err Then
        Err.Clear 
    End If 
    Call HandlePath(filePath)                                                       '获得完整路径
    Dim objStream, LoadBytes 
    Set objStream = CreateObject("ADODB.Stream")
        With objStream
            .Type = 1 
            .Mode = 3 
            .Open 
            .loadFromFile filePath 
            LoadBytes =.read(3) 
            .Close 
        End With 
    Set objStream = Nothing 
    If Err.Number <> 0 Then
        Err.Clear : getFileCharset2 = "" : Exit Function 
    End If 

    strFileHead = BinToHex(LoadBytes) 
    If strFileHead = "EFBBBF" Then
        getFileCharset2 = "UTF-8" 
    Else
        strFileHead = Left(strFileHead, 4) 
        If strFileHead = "FEFF" Then
            getFileCharset2 = "UNICODE BIG" 
        ElseIf strFileHead = "FFFE" Then
            getFileCharset2 = "UNICODE" 
        ElseIf strFileHead = "3C25" Or strFileHead = "3C21" Then
            getFileCharset2 = "GB2312" 
        Else
            getFileCharset2 = "GB2312" 
        End If 
    End If 
    If Err.Number <> 0 Then
        Err.Clear : getFileCharset2 = "" : Exit Function 
    End If 
End Function 
Public Function binToHex(ByVal vStream)
    Dim reVal, i 
    reVal = 0 
    For i = 1 To LenB(vStream)
        reVal = reVal * 256 + AscB(MidB(vStream, i, 1)) 
    Next 
    binToHex = Hex(reVal) 
End Function 



'可以把字符或二进制内容转成字符，牛 从别人哪里复制过来20160112
function Base64Encode(strData)
    dim objAds,objXd
    set objAds=createobject("adodb.stream")
    objAds.Type=2
    objAds.charset="unicode"
    objAds.mode=3
    call objAds.open()
    call objAds.writeText(strData)
    objAds.Position=0
    objAds.Type=1
    'objAds.Position=2 
    set objXd=createobject("msxml.domdocument")
    call objXd.loadXml("<root/>")
    objXd.DocumentElement.DataType="bin.base64"
    objXd.DocumentElement.NodeTypedValue=objAds.read()
    Base64Encode=objXd.DocumentElement.text
end function
'转成二进制了Byte()
function Base64Decode(strData)
    dim objXd
    set objXd=createobject("msxml.domdocument")
    call objXd.loadXml("<root/>")
    objXd.DocumentElement.DataType="bin.base64"
    objXd.DocumentElement.text=strData
    Base64Decode=objXd.DocumentElement.NodeTypedValue
end function

%> 
