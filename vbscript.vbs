'Msgbox (Inputbox("请输入你的名字:")) 

Dim url, dataStr
url = "http://www.baidu.com/"
dataStr = "fileName=1.jpg&PhotoUrlID=2"
    Dim http
    Set http = CreateObject("Microsoft.XMLHTTP")
        http.Open "Post", url, False
        http.setRequestHeader "cache-control", "no-cache"
        http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"         
        http.send ()        
        'MsgBox (http.Status)
	MsgBox( BytesToBstr(http.responseBody, "utf-8"))
    Set http = Nothing

Function BytesToBstr(body,Cset) 
	dim objstream 
	set objstream = CreateObject("adodb.stream") 
	objstream.Type = 1 
	objstream.Mode =3 
	objstream.Open 
	objstream.Write body 
	objstream.Position = 0 
	objstream.Type = 2 
	objstream.Charset = Cset 
	BytesToBstr = objstream.ReadText 
	objstream.Close 
	set objstream = nothing 
End Function 
call CreateFile("E:\E盘\WEB网站\至前网站\VB工程\1.txt", "111111111111111132" & now())
msgbox(getfiletext("E:\E盘\WEB网站\至前网站\VB工程\1.txt"))


'读文件内容 (2013,9,27 
Function GetFileText(ByVal FileName)
    On Error Resume Next
    Dim Fso, FText, OpenFile 
    'GetFileText = ""   '它默认返回的就是空， 这个是多此一举 (2013,9,30)
    Call HandlePath(FileName)    '获得完整路径 
    Set Fso = CreateObject("Scripting.FileSystemObject")
        If Fso.FileExists(FileName) = True Then
            Set FText = Fso.OpenTextFile(FileName, 1)
                '加强 读空文件出错
                Set OpenFile = Fso.GetFile(FileName)
                    If OpenFile.Size = 0 Then Exit Function    '文件为空则退出 
                Set OpenFile = Nothing 
                GetFileText = FText.ReadAll 
            Set FText = Nothing 
        End If 
    Set Fso = Nothing 
    If Err Then doError Err.Description, "GetFileText 读取文件内容 函数出错，FileName=" & FileName 
End Function
'旧版,读取文件内容  (辅助)
Function GetFText(FileName)
    GetFText = GetFileText(FileName)
End Function 
'创建文件   这种方式创建文件时，会多一行，用WriteToFile创建时没有多一行
Function CreateFile(ByVal FileName, ByVal Content)
    On Error Resume Next 
    Dim FText, Fso 
    Call HandlePath(FileName)    '获得完整路径 
    Set Fso = CreateObject("Scripting.FileSystemObject")
        If ExistsZhiDuFile(FileName) = True Then    '判断是否为只读文件
            Call EditFileAttribute(FileName, 32)    '把只读属性改成存档属性 
        End If 
        Set FText = Fso.CreateTextFile(FileName, True)
            FText.WriteLine(Content)
            CreateFile = True 
        Set FText = Nothing 
    Set Fso = Nothing 
    If Err Then CreateFile = False : doError Err.Description, "CreateFile 创建文件 函数出错，FileName=" & FileName
End Function 


'处理成完成路径 (2013,9,27 
Function HandlePath(Path)    'Path前面不加ByVal 重定义，这样是为了让前面函数里可以使用这个路径完整调用
    Path = Replace(Path, "/", "\") 
    Path = Replace(Path, "\\", "\") 
    Path = Replace(Path, "\\", "\") 
	dim isDir   '为目录
	isDir=false
	if right(Path,1)="\" then
    	isDir=true
	end if 
	If InStr(Path, ":") = 0 Then 
		If Left(Path,1)="\" Then
			Path = Server.MapPath("\") & "\" & Path
		Else
			Path = Server.MapPath(".\") & "\" & Path
		End If
	End If
    Path = Replace(Path, "/", "\") 
    Path = Replace(Path, "\\", "\") 
    Path = Replace(Path, "\\", "\")  
	path = FullPath(Path)
	if isDir=true then
		Path=Path & "\"
	end if
    HandlePath = Path
End Function
'真正的路径  PHP里函数 为假返回空 
Function RealPath(ByVal Path)
	If CheckFile(Path) Then 
		RealPath = Path
		Exit Function
	End If
	If CheckFolder(Path) Then 
		RealPath = Path
		Exit Function
	End If
End Function
'完整路径
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