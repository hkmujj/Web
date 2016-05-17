<%
'json´¦Àí

'×Öµä×ªJson
function dictionaryToJson(obj)
	dim s,c,key
	For Each s In aspD
		key=s'phpno
		s=aspD(key)'phpno
		'call echo(key, s)
		c=c & key & ":"""& JsEncode__(s) &""""
		if c<>"" then
			c=c & ","
		end if
	next
	c="[{"& c &"}]"
	dictionaryToJson=c
end function

%>