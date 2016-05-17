<% 
'系统信息  (2014,05,27)



'操作系统版本 
Function OperationSystem()
	Dim httpAgent,SystemVer
    httpAgent = Request.ServerVariables("HTTP_USER_AGENT") 
    If InStr(httpAgent, "NT 5.2") > 0 Then
        SystemVer = "Windows Server 2003" 
    ElseIf InStr(httpAgent, "NT 5.1") > 0 Then
        SystemVer = "Windows XP" 
    ElseIf InStr(httpAgent, "NT 5") > 0 Then
        SystemVer = "Windows 2000" 
    ElseIf InStr(httpAgent, "NT 4") > 0 Then
        SystemVer = "Windows NT4" 
    ElseIf InStr(httpAgent, "4.9") > 0 Then
        SystemVer = "Windows ME" 
    ElseIf InStr(httpAgent, "98") > 0 Then
        SystemVer = "Windows 98" 
    ElseIf InStr(httpAgent, "95") > 0 Then
        SystemVer = "Windows 95" 
    Else
        SystemVer = httpAgent 
    End If 
    OperationSystem = httpAgent
End Function 
'检测是否为手机
Function CheckMobile()
	CheckMobile=False
	If Request.ServerVariables("HTTP_X_WAP_PROFILE") <> "" Then
		CheckMobile=True
	End If
End Function 

'获得IIS版本号
function getIISVersion()
	getIISVersion=request.ServerVariables("SERVER_SOFTWARE")
end function

%>