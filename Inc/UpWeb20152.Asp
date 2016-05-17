<!--#Include File="Config.Asp"-->
<%
Call UpWeb() 
Sub UpWeb()
    Dim FilePath, Content, FileType, FileStr, SplStr, Splxx, S, C, fCharSet, fEditTime, sEditTime, UpCharSet 

	'处理版本号
    If Request("act") = "getVersionNumber" Then
        Call Die("V1.1") 
    End If 

	'判断本地
    If GetIP = "127.0.0.1" Or InStr(GetIP, "192.168.") > 0 Then
        Call Die("为本地，不可以上传文件") 
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
    Else
        Content = SpecialHtmlUploadEncryptionDecrypt(Content, "Decrypt")                '解密Html内容特殊字符
        Call WriteToFile(FilePath, Content, UpCharSet) 
        Call Rw(FilePath & "|普通保存") 
    End If 
End Sub 
%> 


