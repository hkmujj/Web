<%
'dll外部调用哦的  ASP小程序 处理动作，如计算数字值，回复聊天内容，基础框架 可扩展简单
Class dllSoftSafe

'帮助信息 

	private FrmName						'窗体名称
	
	'初始化
	Private Sub Class_Initialize()
		
	End Sub
	'析构函数 类终止
	Private Sub Class_Terminate()
	
	End Sub
	'设置窗体名称
	Public Property Let setFrmName(Str)
		 FrmName=Str
	End Property
	'获得窗体名称
	Public Property Get getFrmName()
		getFrmName = FrmName
	End Property
	
	'是否为数据库密码
	function isAccessPass(byval MDBPath,AccessPass)
		isAccessPass=false
        If CheckFile(MDBPath) = False Then
            Call MBInfo("出错", "没有数据库"): Exit function
        End If
        '数据库密码
        AccessPass = FileCreateDate(MDBPath)
        AccessPass = Format_Time(AccessPass, 4) & "313801120"
        If (CheckAccessPass(MDBPath, AccessPass)) = False Then
            Call MBInfo("出错", "数据库密码错误" & AccessPass): Exit function
        End If
		isAccessPass=true
	end function 


	Function CheckFile(ByVal FileName)
		On Error Resume Next
		Dim Fso
		Call HandlePath(FileName)'获得完整路径
		Set Fso = CreateObject("Scripting.FileSystemObject")
			CheckFile = Fso.FileExists(FileName)
		Set Fso = Nothing
		If Err Then doError Err.Description, "CheckFile 检查文件 函数出错，FileName=" & FileName
	End Function
	Function FileCreateDate(FileName)
		FileCreateDate = GetFileCreateDate(FileName)
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
	Function CheckAccessPass(MDBPath, AccessPass)
		On Error Resume Next
		Dim Conn
		CheckAccessPass = True
		Set Conn = CreateObject("Adodb.Connection")
			Conn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database PassWord = '" & AccessPass & "';Data Source = " & MDBPath
			Conn.Close
			If Err <> 0 Then
				Err.Clsoe
				CheckAccessPass = False
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
	Function GetFileCreateDate(ByVal FileName)
		On Error Resume Next
		Dim Fso, OpenFile
		Call HandlePath(FileName)'获得完整路径
		Set Fso = CreateObject("Scripting.FileSystemObject")
			If Fso.FileExists(FileName) = True Then
				Set OpenFile = Fso.GetFile(FileName)
					GetFileCreateDate = OpenFile.DateCreated
				Set OpenFile = Nothing
			End If
		Set Fso = Nothing
		If Err Then doError Err.Description, "GetFileCreateDate 获得文件创建时间 函数出错，FileName=" & FileName
	End Function 
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
	Function Rw(Str)
		Response.Write(Str)
	End Function
	Function MsgBox(ByVal Content)
		Content = Replace(Replace(Content, Chr(10), "\n"), Chr(13), "\n")
		Response.Write("<script>alert('" & Content & "');</script>")
	End Function
	function MBInfo(title, content)
		MBInfo = MsgBox(title)
	End function
End Class


Dim testOBJ : Set testOBJ = New dllSoftSafe	'加载类
dim AccessPass
'call ddlCD.handAction("1+2*6+2=")
'call ddlCD.rw(ddlCD.handAction("ask 35915100"))

select case request("act")
	case "handle" : call testOBJ.rw(testOBJ.isAccessPass(request("MDBPath"),AccessPass))
	case else:Default()
end select

'默认
sub Default()
%>
<form id="form1" name="form1" method="post" action="?act=handle">
  发送内容
  <input name="MDBPath" type="text" id="MDBPath" value="<%="\VB工程\Data\Access.Mdb"%>" />
  <input type="submit" name="button" id="button" value="提交" />
</form>
<%end sub%>