<!--#Include virtual = "/Inc/_Config.Asp"-->  

<% 
'���ߣ��ƶ� QQ313801120  http://sharembweb.com/
'ѭ����
Class whileclass
    Function myfun(nNumb)
        If nNumb = 1 Then
            Response.Write("hello world<hr>") 
        Else
            Response.Write("no numb<hr>") 
        End If 
    End Function 
    Sub nfor(n)
        Dim i 
        For i = 1 To n
            Response.Write(i & "for��<hr>") 
        Next 
    End Sub 
    Sub nwhile(n)
        While n > 1
            n = n - 1 
            Response.Write(n & "while��<hr>") 
        Wend 
    End Sub 
    Sub ndoloop(n)
        Do While n > 1
            n = n - 1 
            Response.Write(n & "doloop��<hr>") 
        Loop 
    End Sub 
    Sub nforeach()
        Dim splStr, s 
        splStr = Array("aa", "bb", "cc") 
        For Each s In splStr
            Response.Write("s=" & s & "<hr>") 
        Next 
    End Sub 
End Class 

'�ж���
Class ifclass
    Function testif(n)
        If n > 10 Then
            Response.Write("n����10<br>") 
        ElseIf n > 5 Then
            Response.Write("n����5<br>") 
        Else
            Response.Write("nΪĬ��<br>" & n) 
        End If 
    End Function 
    Function testif2(a)
        Response.Write("testif2<hr>") 
    End Function 


End Class 

'�ֵ���
Class zdclass
    Sub testzd()
        Dim aspD, title, items, i 
        Dim aA, bB : Set aspD = Server.CreateObject("Scripting.Dictionary")
            aspD.add "Abs", "�������ֵľ���ֵ11111111" 
            aspD.add "Sqr", "������ֵ���ʽ��ƽ����aaaaaaaaaaaaaaaaaaaaaaaa" 
            aspD.add "Sgn", "���ر�ʾ���ַ��ŵ�����22222222" 
            aspD.add "Rnd", "����һ��������ɵ�����33333333333333" 
            aspD.add "Log", "����ָ����ֵ����Ȼ����ssssssssssssssss" 


            Response.Write("Abs=" & aspD("Abs") & "<hr>") 
            Response.Write("Rnd=" & aspD("Rnd") & "<hr>") 
    End Sub
End Class 

'����ѭ��
Sub testwhile()
    Dim obj : Set obj = new whileclass
        Call obj.myfun(1) 
        Response.Write("<br>33333333<br>") 
        Call obj.myfun(2) 
        Call obj.nfor(6) 
        Call obj.nwhile(6) 
        Call obj.ndoloop(6) 
        Call obj.nforeach() 

End Sub
'�����ж�
Sub testif()
    Dim obj : Set obj = new ifclass
        Call obj.testif(11) 
        Call obj.testif(6) 
        Call obj.testif(3) 
        obj.testif2 3 : obj.testif2 3 
End Sub
'�����ֵ�
Sub testzd()
    Dim obj : Set obj = new zdclass
        Call obj.testzd() 

End Sub



'��ȡ�ַ��� ����20160114
'c=[A]sharembweb.com[/A]
'0=sharembweb.com
'1=[A]sharembweb.com[/A]
'3=[A]sharembweb.com
'4=sharembweb.com[/A]
Function strCutTest(ByVal content, ByVal startStr, ByVal endStr, ByVal cutType)
    'On Error Resume Next
    Dim s1, s1Str, s2, s3, c 
    If InStr(content, startStr) = False Or InStr(content, endStr) = False Then
        c = "" 
        Exit Function 
    End If 
    Select Case cutType
        '������20150923
        Case 1
            s1 = InStr(content, startStr) 
            s1Str = Mid(content, s1 + Len(startStr)) 
            s2 = s1 + InStr(s1Str, endStr) + Len(startStr) + Len(endStr) - 1 'ΪʲôҪ��1

        Case Else
            s1 = InStr(content, startStr) + Len(startStr) 
            s1Str = Mid(content, s1) 
            'S2 = InStr(S1, Content, EndStr)
            s2 = s1 + InStr(s1Str, endStr) - 1 
        'call echo("s2",s2)
    End Select
    s3 = s2 - s1 
    If s3 >= 0 Then
        c = Mid(content, s1, s3) 
    Else
        c = "" 
    End If 
    If cutType = 3 Then
        c = startStr & c 
    End If 
    If cutType = 4 Then
        c = c & endStr 
    End If 
    strCutTest = c 
    'If Err.Number <> 0 Then Call eerr(startStr, content)
'doError Err.Description, "strCutTest ��ȡ�ַ��� ��������StartStr=" & EchoHTML(StartStr) & "<hr>EndStr=" & EchoHTML(EndStr)
End Function
 
'����ʵ��
sub testcase()

	Dim c 
	c = "[A]sharembweb.com[/A]" 
	
	Response.Write("c=" & c & "<br>") 
	
	Response.Write("0=" & strCutTest(c, "[A]", "[/A]", 0) & "<br>" & vbCrLf) 
	Response.Write("1=" & strCutTest(c, "[A]", "[/A]", 1) & "<br>" & vbCrLf) 
	'response.Write("2=" & strCutTest(c,"[A]","[/A]",2) & "<br>" & vbcrlf)
	Response.Write("3=" & strCutTest(c, "[A]", "[/A]", 3) & "<br>" & vbCrLf) 
	Response.Write("4=" & strCutTest(c, "[A]", "[/A]", 4) & "<br>" & vbCrLf) 

end sub


'ѡ��
Select Case Request("act")
    Case "testwhile" : testwhile()                                        '����ѭ��
    Case "testif" : testif()                                              '�����ж�
    Case "testzd" : testzd()                                              '�����ֵ�
    Case "testcase" : testcase()                                              '����ʵ��
	

    Case Else : displayDefault()                                          '��ʾĬ��
End Select

'��ʾĬ��
Sub displayDefault()
    Response.Write("<a href='?act=testwhile'>����ѭ��</a> <br>") 
    Response.Write("<a href='?act=testif'>�����ж�</a> <br>") 
    Response.Write("<a href='?act=testzd'>�����ֵ�</a> <br>") 
    Response.Write("<a href='?act=testcase'>����ʵ��</a> <br>") 
End Sub  
%>  

