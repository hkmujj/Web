<% 
'���뻥�� GB2312��UTF-8ת��

'UTFתGB23
Function UTF2GB(ByVal UTFStr)
    Dim Dig, GBStr 
    For Dig = 1 To Len(UTFStr)
        '���UTF8����������%��ͷ�����ת��
        If Mid(UTFStr, Dig, 1) = "%" Then
            'UTF8�������ִ���8��ת��Ϊ����
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
 
'����Ƿ���԰�UTFתGB2312
Function CheckUTFToGB2312(ByVal UTFStr)
    Dim Dig, GBStr 
    CheckUTFToGB2312 = True 
    For Dig = 1 To Len(UTFStr)
        '���UTF8����������%��ͷ�����ת��
        If Mid(UTFStr, Dig, 1) = "%" Then
            'UTF8�������ִ���8��ת��Ϊ����
            If Len(UTFStr) >= Dig + 8 Then
                If ConvChinese(Mid(UTFStr, Dig, 9)) = "[����Error]" Then
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
 
'UTF8�������ֽ�ת��Ϊ���� �������
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
        If Err <> 0 Then ConvChinese = "[����Error]" : Exit Function '�������˳�
    Next 
End Function
 

 
'תUTF8���룬�ǳ���׼�����ã�(2014 12 11)
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
 
'GBתUTF8--��GB��������ת��ΪUTF8��������  �������׼�����Ժ���ϳ���
Function toUTF8(szInput)
    Dim wch, uch, szRet 
    Dim X 
    Dim nAsc 
    '����������Ϊ�գ����˳�����
    If szInput = "" Then
        toUTF8 = szInput 
        Exit Function 
    End If 
    '��ʼת��
    For X = 1 To Len(szInput)
        '����mid�����ֲ�GB��������
        wch = Mid(szInput, X, 1) 
        '����ascW��������ÿһ��GB�������ֵ�Unicode�ַ�����
        'ע��asc�������ص���ANSI �ַ����룬ע������
        nAsc = AscW(wch) 
        If nAsc < 0 Then nAsc = nAsc + 65536 

        If(nAsc And &HFF80) = 0 Then
            szRet = szRet & wch 
        Else
            If(nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc / 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80) 
                szRet = szRet & uch 
            Else
                'GB�������ֵ�Unicode�ַ�������0800 - FFFF֮��������ֽ�ģ��
                uch = "%" & Hex((nAsc / 2 ^ 12) Or &HE0) & "%" & _ 
                Hex((nAsc / 2 ^ 6) And &H3F Or &H80) & "%" & _ 
                Hex(nAsc And &H3F Or &H80) 
                szRet = szRet & uch 
            End If 
        End If 
    Next 

    toUTF8 = szRet 
End Function
 
'GBתunicode---��GB��������ת��Ϊunicode��������
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
 

'�ж��Ƿ�Ϊ��Ч��ʮ�����ƴ���
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
 
'---------------------- �Լ��ļ����� ----------------------------
'����ͬJS��escapeһ��
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
 

'����ͬJS��unescapeһ��
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
 
'���������磺���й���&#x4E2D;&#x56FD;��
Function ChineseToUnicode(Str)
    Dim I, S, C 
    For I = 1 To Len(Str)
        S = Mid(Str, I, 1) 
        S = "&#x" & Hex(AscW(S)) & ";" 
        C = C & S 
    Next 
    ChineseToUnicode = C 
End Function
 
'�⺺������
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
 

 
'Url����
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
 
'URL����   �������
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
 
'GB2312��UTF-8
Function GBtoUTF8(szInput)
    Dim wch, uch, szRet 
    Dim X 
    Dim nAsc 
    '����������Ϊ�գ����˳�����
    If szInput = "" Then
        GBtoUTF8 = szInput 
        Exit Function 
    End If 
    '��ʼת��
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
 
'���ܽ���URL
Function EncDecURL(ByVal Content, ByVal SType)
    Dim Ystr, TStr, Temp, I, S, C 
    Ystr = "abcdefghijklmnopqrstuvwxyz0123456789:/.=& ()%" 
    TStr = "9ab1cdefghij234klmnopqrst678uvwxyz:/.05*-$[]@" 
    If SType = "����" Or SType = "0" Then
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
 
'����URL (����)
Function EncURL(Content)
    EncURL = EncDecURL(Content, "����") 
End Function
 
'����URL (����)
Function DecURL(Content)
    DecURL = EncDecURL(Content, "����") 
End Function
 
%> 