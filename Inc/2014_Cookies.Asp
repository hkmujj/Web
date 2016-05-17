<% 
Sub setApp(AppName, AppData)
    Application.Lock 
    Application.Contents.Item(AppName) = AppData 
    Application.Unlock 
End Sub
Function getApp(AppName)
    If IsN(AppName) Then getApp = "" : Exit Function 
    getApp = Application.Contents.Item(AppName) 
End Function 
Sub removeApp(AppName)
    Application.Lock 
    Application.Contents.Remove(AppName) 
    Application.Unlock
End Sub 
'设置cookie(20151120)
Function setCookie(cookieName, cookieValue, cookieTime)
    Response.Cookies(cookieName) = cookieValue 
    Response.Cookies(cookieName).Expires = now+1 
End Function 
'获得cookie值
Function getCookie(ByVal cookieName)
    getCookie=Request.Cookies(cookieName)
End Function
'移除cookie
function removeCookie(ByVal cookieName)
	Response.Cookies(cookieName) = Empty 
	Response.Cookies(cookieName).Expires = Now() 
End function
'清除cookie
function clearCookie(ByVal cookieName)
	removeCookie(cookieName)
end function
%>

