<%
'Div����վ����
Function DivNavigation(AddSql)
	Dim I, C, Url, FileName, BigFolder, Did, ShowDid, Target 
	Rs.Open "Select * From [NavBigClass] " & AddSql, Conn, 1, 1 
	C = C & "<li class=left></li>" & vbCrlf
	For I = 1 To Rs.RecordCount
		FileName = Rs("FileName") 
		If FileName <> "" And Right(FileName,1)<>"/" Then FileName = FileName & ".html"  
	
		BigFolder = "/" & Rs("FolderName") & "/" 
		Did = Rs("BigClassName") 
		ShowDid = FontColorFontB(Did, Rs("FontB"), Rs("FontColor")) 
		Target = Rs("Target")
		
		Url = Rs("FileName")
		If CheckRemoteUrl(Url)=False Then
			Url = BigFolder & FileName 
			Url = Replace(Replace(Url, "//", "/"), "//", "/")
			Url = FullHttpUrl(Host(),Url)		
			Call WebDebug(Url,"act=Nav&NavDid=" & Did)'����
		End If
		
		C = C & "<li><a " & AHref(Url, "", Target) & ">" & ShowDid & "</a></li>" &vbCrlf
		C = C & "<li class=line></li>" & vbCrlf
	Rs.MoveNext : Next : Rs.Close 
	C = C & "<li class=right></li>" & vbCrlf
	DivNavigation = C 
End Function

'��վ������ַ���� (20140408�Ľ�)
function WebDebug(url, DebugUrl)
    if request("Debug") <> "" and getIP = "127.0.0.1" then
        url = "/WebAdmin/ShowWeb.Asp?Debug=1&" & DebugUrl 
        if request("MackHtml") = "False" then
            url = url & "&MackHtml=False" 
        end if 
        url = url & "&UserId=" & UserID 
    else
        url = lCase(url)                                      '��URL����Сдת��
    end if 
    WebDebug = url 
end function




%>