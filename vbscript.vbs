'Msgbox (Inputbox("�������������:")) 

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
call CreateFile("E:\E��\WEB��վ\��ǰ��վ\VB����\1.txt", "111111111111111132" & now())
msgbox(getfiletext("E:\E��\WEB��վ\��ǰ��վ\VB����\1.txt"))


'���ļ����� (2013,9,27 
Function GetFileText(ByVal FileName)
    On Error Resume Next
    Dim Fso, FText, OpenFile 
    'GetFileText = ""   '��Ĭ�Ϸ��صľ��ǿգ� ����Ƕ��һ�� (2013,9,30)
    Call HandlePath(FileName)    '�������·�� 
    Set Fso = CreateObject("Scripting.FileSystemObject")
        If Fso.FileExists(FileName) = True Then
            Set FText = Fso.OpenTextFile(FileName, 1)
                '��ǿ �����ļ�����
                Set OpenFile = Fso.GetFile(FileName)
                    If OpenFile.Size = 0 Then Exit Function    '�ļ�Ϊ�����˳� 
                Set OpenFile = Nothing 
                GetFileText = FText.ReadAll 
            Set FText = Nothing 
        End If 
    Set Fso = Nothing 
    If Err Then doError Err.Description, "GetFileText ��ȡ�ļ����� ��������FileName=" & FileName 
End Function
'�ɰ�,��ȡ�ļ�����  (����)
Function GetFText(FileName)
    GetFText = GetFileText(FileName)
End Function 
'�����ļ�   ���ַ�ʽ�����ļ�ʱ�����һ�У���WriteToFile����ʱû�ж�һ��
Function CreateFile(ByVal FileName, ByVal Content)
    On Error Resume Next 
    Dim FText, Fso 
    Call HandlePath(FileName)    '�������·�� 
    Set Fso = CreateObject("Scripting.FileSystemObject")
        If ExistsZhiDuFile(FileName) = True Then    '�ж��Ƿ�Ϊֻ���ļ�
            Call EditFileAttribute(FileName, 32)    '��ֻ�����Ըĳɴ浵���� 
        End If 
        Set FText = Fso.CreateTextFile(FileName, True)
            FText.WriteLine(Content)
            CreateFile = True 
        Set FText = Nothing 
    Set Fso = Nothing 
    If Err Then CreateFile = False : doError Err.Description, "CreateFile �����ļ� ��������FileName=" & FileName
End Function 


'��������·�� (2013,9,27 
Function HandlePath(Path)    'Pathǰ�治��ByVal �ض��壬������Ϊ����ǰ�溯�������ʹ�����·����������
    Path = Replace(Path, "/", "\") 
    Path = Replace(Path, "\\", "\") 
    Path = Replace(Path, "\\", "\") 
	dim isDir   'ΪĿ¼
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
'������·��  PHP�ﺯ�� Ϊ�ٷ��ؿ� 
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
'����·��
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