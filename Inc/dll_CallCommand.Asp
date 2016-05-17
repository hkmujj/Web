<!--#Include file="config.asp"-->
<%
'dll外部调用哦的  ASP小程序 处理动作，如计算数字值，回复聊天内容，基础框架 可扩展简单
Class dllCommand
'帮助信息
'regsvr32 /u E:\E盘\WEB网站\至前网站\DataDir\VB模块\dll\mydll.dll   删除
'regsvr32 E:\E盘\WEB网站\至前网站\DataDir\VB模块\dll\mydll.dll
'Dim myOBJ
'Set myOBJ = CreateObject("mydd3yun.myd2")          'set xxx = createobject("工程名.类名")
'Call MsgBox(myOBJ.testB())

    Private FrmName                     '窗体名称
    
    '初始化
    Private Sub Class_Initialize()
        
    End Sub
    '析构函数 类终止
    Private Sub Class_Terminate()
    
    End Sub
    '设置窗体名称
    Public Property Let setFrmName(Str)
         FrmName = Str
    End Property
    '获得窗体名称
    Public Property Get getFrmName()
        getFrmName = FrmName
    End Property
    
    '处理Action
    Function handAction(ByVal commandStr)
        Dim Action, Content
        commandStr = PHPTrim(commandStr)
        Content = commandStr
        '动作与内容分离
        If InStr(Content, " ") > 0 Then
            Action = Mid(Content, 1, InStr(Content, " ") - 1)
            Content = Mid(Content, InStr(Content, " ") + 1)
        End If
        '问答
        If Action = "ask" Then
             handAction = ask(Content)
        '破解Access密码
        ElseIf InStr("|access|accesspass|", "|" & LCase(Action) & "|") > 0 Then
            Dim myOBJ: Set myOBJ = New clsAccessPass
            Dim sVersion 
            handAction = HandlePath(Content) & " 密码为" & vbCrLf & myOBJ.INNER_GetAccessPwd(HandlePath(Content), sVersion, True)
        '为+-*/运算
        ElseIf Right(Content, 1) = "=" Then
            Content = Mid(commandStr, 1, Len(commandStr) - 1)
            handAction = commandStr & "" & getVBScript("function Run(a):Run=eval(a):end function ", Content)
        '测试
        ElseIf commandStr = "test" Then
            handAction = commandStr & "=" & getVBScript("function Run(a):Run=eval(a):end function ", "1+2")
        '默认问答
        Else
            handAction = ask(commandStr)
        End If
        Call CreateFile("1.txt", handAction)
    End Function
    '问答QA
    Function ask(Action)
        ask = ""
        If InStr("|35915100|", "|" & Action & "|") > 0 Then
            ask = "35915100 仿站与ASP程序维护群"
        ElseIf InStr("|313801120|qq313801120|", "|" & Action & "|") > 0 Then
            ask = "313801120为作者QQ号" & vbCrLf & "多年WEB开发程序员。"
            
        ElseIf InStr("|小云|小孙|孙晓云|", "|" & Action & "|") > 0 Then
            'Debug.Print ("3333333333")      '无用
            ask = "孙晓云是这款软件的开发人员，" & vbCrLf & "可联系QQ313801120。11111"
        End If
    End Function


    Function Eerr(Word, Str)
        Response.Write ("<font color=red>" & Word & "</font>：" & Str)
    End Function

    Function Rw(Str)
        Response.Write Str
    End Function
    '引用函数
    Function PHPTrim(Str)
        PHPTrim = TrimVbCrlf(Str)
    End Function
    Function ECHO(Word, Str)
        Call echoPrompt(Word, Str)
    End Function
    Function TrimVbCrlf(Str)
        TrimVbCrlf = PHPRTrim(PHPLTrim(Str))
    End Function
    Function echoPrompt(Word, Str)
        Response.Write "<font color=Green>" & Word & "</font>：" & Str & "<br>"
    End Function
    Function PHPRTrim(Str)
        Dim Pos, isBlankChar
        Pos = Len(Str)
        isBlankChar = True
        While isBlankChar And Pos >= 2
            If Mid(Str, Pos, 1) = " " & Mid(Str, Pos, 1) = vbTab Then '把vbTab也去掉
                Pos = Pos - 1
            ElseIf Mid(Str, Pos - 1, 2) = vbCrLf Then
                Pos = Pos - 2
            Else
                isBlankChar = False
            End If
        Wend
        PHPRTrim = RTrim(Left(Str, Pos))
    End Function
    Function PHPLTrim(Str)
        Dim Pos, isBlankChar
        Pos = 1
        isBlankChar = True
        While isBlankChar
            If Mid(Str, Pos, 1) = " " Or Mid(Str, Pos, 1) = vbTab Then '把vbTab也去掉
                Pos = Pos + 1
            ElseIf Mid(Str, Pos, 2) = vbCrLf Then
                Pos = Pos + 2
            Else
                isBlankChar = False
            End If
        Wend
        PHPLTrim = Right(Str, Len(Str) - Pos + 1)
    End Function
	 
	'Vb调用外部脚本
	Function getVBScript(ByVal Content, ByVal Action)
		Dim Sc
		Set Sc = CreateObject("ScriptControl")
		Sc.Language = "VBScript"
		Sc.AddCode Content
		Sc.AddCode "Function mytest(): mytest = ""【vb内容】"":End Function"
		 
		getVBScript = Sc.Run("Run", Action)
	End Function
			
End Class


Dim ddlCD : Set ddlCD = New dllCommand	'加载类

'call ddlCD.handAction("1+2*6+2=")
'call ddlCD.rw(ddlCD.handAction("ask 35915100"))

select case request("act")
	case "handle" : call ddlCD.rw(ddlCD.handAction(request("command")))
	case else:Default()
end select

'默认
sub Default()
%>
<form id="form1" name="form1" method="post" action="?act=handle">
  发送内容
  <input name="command" type="text" id="command" value="3+3*5=" />
  <input type="submit" name="button" id="button" value="提交" />
</form>
<form id="form2" name="form1" method="post" action="?act=handle">
  发送内容
  <input name="command" type="text" id="command" value="ask 35915100" />
  <input type="submit" name="button" id="button" value="提交" />
</form>
<%end sub%>