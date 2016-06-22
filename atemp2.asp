<!--#Include virtual = "/Inc/_Config.Asp"-->  

<% 
'作者：云端 QQ313801120  http://sharembweb.com/
'循环类
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
            Response.Write(i & "for、<hr>") 
        Next 
    End Sub 
    Sub nwhile(n)
        While n > 1
            n = n - 1 
            Response.Write(n & "while、<hr>") 
        Wend 
    End Sub 
    Sub ndoloop(n)
        Do While n > 1
            n = n - 1 
            Response.Write(n & "doloop、<hr>") 
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

'判断类
Class ifclass
    Function testif(n)
        If n > 10 Then
            Response.Write("n大于10<br>") 
        ElseIf n > 5 Then
            Response.Write("n大于5<br>") 
        Else
            Response.Write("n为默认<br>" & n) 
        End If 
    End Function 
    Function testif2(a)
        Response.Write("testif2<hr>") 
    End Function 


End Class 

'字典类
Class zdclass
    Sub testzd()
        Dim aspD, title, items, i 
        Dim aA, bB : Set aspD = Server.CreateObject("Scripting.Dictionary")
            aspD.add "Abs", "返回数字的绝对值11111111" 
            aspD.add "Sqr", "返回数值表达式的平方根aaaaaaaaaaaaaaaaaaaaaaaa" 
            aspD.add "Sgn", "返回表示数字符号的整数22222222" 
            aspD.add "Rnd", "返回一个随机生成的数字33333333333333" 
            aspD.add "Log", "返回指定数值的自然对数ssssssssssssssss" 


            Response.Write("Abs=" & aspD("Abs") & "<hr>") 
            Response.Write("Rnd=" & aspD("Rnd") & "<hr>") 
    End Sub
End Class 

'测试循环
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
'测试判断
Sub testif()
    Dim obj : Set obj = new ifclass
        Call obj.testif(11) 
        Call obj.testif(6) 
        Call obj.testif(3) 
        obj.testif2 3 : obj.testif2 3 
End Sub
'测试字典
Sub testzd()
    Dim obj : Set obj = new zdclass
        Call obj.testzd() 

End Sub



'截取字符串 更新20160114
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
        '完善于20150923
        Case 1
            s1 = InStr(content, startStr) 
            s1Str = Mid(content, s1 + Len(startStr)) 
            s2 = s1 + InStr(s1Str, endStr) + Len(startStr) + Len(endStr) - 1 '为什么要减1

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
'doError Err.Description, "strCutTest 截取字符串 函数出错，StartStr=" & EchoHTML(StartStr) & "<hr>EndStr=" & EchoHTML(EndStr)
End Function
 
'测试实例
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


'选择
Select Case Request("act")
    Case "testwhile" : testwhile()                                        '测试循环
    Case "testif" : testif()                                              '测试判断
    Case "testzd" : testzd()                                              '测试字典
    Case "testcase" : testcase()                                              '测试实例
	

    Case Else : displayDefault()                                          '显示默认
End Select

'显示默认
Sub displayDefault()
    Response.Write("<a href='?act=testwhile'>测试循环</a> <br>") 
    Response.Write("<a href='?act=testif'>测试判断</a> <br>") 
    Response.Write("<a href='?act=testzd'>测试字典</a> <br>") 
    Response.Write("<a href='?act=testcase'>测试实例</a> <br>") 
End Sub  
%>  

