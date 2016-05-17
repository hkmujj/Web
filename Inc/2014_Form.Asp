<% 
'Form

'表单提交
Sub FormSubmit()
	Dim TableName,SplStr,S,FieldName,FieldContent,FieldList,YZM
	Call OpenConn()
	SplStr=Split(Request.Form(),"&")
	TableName = Rf("TableName")
	YZM = Trim(Rf("YZM"))
	
	If YZM<>"" Then
		If Session("YZM") <>YZM Then
			Call Javascript("返回", "验证码不正确", "")
			Response.End()
		End If
	End If
	
	FieldList = LCase(GetFieldList(TableName))
	'Call Echo("FieldList",FieldList)
	'Call Echo("TableName", TableName)
	Rs.Open"Select * From ["& TableName &"]",Conn,1,3
	Rs.AddNew
		For Each S In SplStr
			FieldName = LCase(Mid(S,1,InStr(S,"=")-1))
			'FieldContent = Mid(S,InStr(S,"=")+1) 
			FieldContent = Rf(FieldName)
			If InStr(","& FieldList &",", ","& FieldName &",")>0 Then
				Rs(FieldName)=FieldContent
			End If
			'Call Echo(FieldName,FieldContent & "," & unescape(FieldContent))
		Next
	Rs.Update:Rs.Close
'	Call Echo("DialogTitle",Rf("DialogTitle"))
	'Call Die("留言内容")
	
	Call Javascript("返回", "提交"& Rf("DialogTitle") &"成功", "")
End Sub

'获得POST字段名称列表 20160226
function getFormFieldList()
	dim s,c,splstr,fieldName
	splstr=split(Request.Form,"&")
	for each s in splstr
		fieldName = LCase(Mid(s, 1, InStr(s, "=") - 1))
		if c<>"" then c=c & "|"
		c=c & fieldName
	next
	getFormFieldList=c
end function 
 
%>







