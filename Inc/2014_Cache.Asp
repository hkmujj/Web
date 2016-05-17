<%
'无缓冲
Sub noCache()
	Response.Buffer = True
	Response.Expires = 0
	Response.ExpiresAbsolute = Now() - 1
	Response.CacheControl = "no-cache"
	Response.AddHeader "Expires",Date()
	Response.AddHeader "Pragma","no-cache"
	Response.AddHeader "Cache-Control","private, no-cache, must-revalidate"
End Sub

'清除缓冲(20150907)
function XY_clearCache(Action)
	on error resume next			'暂时这个说，应该在后台里写是否缓冲
	dim operate
	operate = LCase(RParam(Action,"operate"))
	if operate="true" and HandleisCache=false then
		call noCache()
		HandleisCache=false	
	end if
end function
%>