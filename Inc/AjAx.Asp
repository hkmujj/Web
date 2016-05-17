<!--#Include File = "Config.Asp"-->  
<%  
'选择动作列表
Select Case Request("act") 
	Case "SaveFile" : SaveFile( Request("FilePath"))			'保存文件
	Case "GetFunctionCode" : GetFunctionCode(Request("Lable"))	'获得函数代码
	Case "SaveQQDialogue" : Call SaveQQDialogue()				'保存点击QQ对话信息
	Case "UserLogin" : Call UserLogin()						 	'会员登录
	Case "UserReg" : Call UserReg()								'会员注册
	Case "MobileArticleList" : Call MobileArticleList()			'手机文件列表
	Case "ShowArticleNumb" : Call ShowArticleNumb()				'显示文章浏览数 累加
	Case "SaveAJAXEdit" : Call SaveAJAXEdit()					'Ajax保存内容
	Case "SaveAJAXTEXTEdit" : Call SaveAJAXTEXTEdit()					'Ajax保存TXT文件内容
	
	Case "DeleteAction" : Call DeleteAction()					'Ajax删除内容
	
	Case "showArticleList" : Call showArticleList()					'Ajax删除内容
End Select

'显示文章列表 http://127.0.0.1/inc/ajax.asp?act=showArticleList&did=阅读&numb=4
function showArticleList()
	call openconn()
	dim did,sql,i,nNumb
	did=Rq("did")
	nNumb=HandleNumber(Rq("Numb")) 	
	if nNumb<>"" then
		sql = "select top "& nNumb &" * from [product]"
	else	
		sql = "select * from [product]"
	end if
	if did<>"" then
		sql=sql & " where BigClassName='"& did &"'"
	end if
	rs.open sql,conn,1,1
	Dim json : Set json = New Cls_AB_JSON        '加载JSON类
	'call rw(json.toJSON(dataArr))
	call rw(json.toJSON(rs))
	'call rw(json.fieldToJson(rs,"|title|"))
	rs.close 
end function

function showArticleList2()
	call openconn()
	dim did,sql,i,nNumb
	did=Rq("did")
	nNumb=HandleNumber(Rq("Numb")) 	
	if nNumb<>"" then
		sql = "select top "& nNumb &" * from [product]"
	else	
		sql = "select * from [product]"
	end if
	if did<>"" then
		sql=sql & " where BigClassName='"& did &"'"
	end if
	rs.open sql,conn,1,1
	Dim json : Set json = New Cls_AB_JSON        '加载JSON类
	'call rw(json.toJSON(dataArr))
	call rw(json.toJSON(rs))
	rs.close
end function

'Ajax删除内容
Function DeleteAction() 
	Dim C	
	C=C & "MDBPath=" & Rq("MDBPath") & vbCrlf
	C=C & "TableName=" & Rq("TableName") & vbCrlf
	C=C & "FieldName=" & Rq("FieldName") & vbCrlf 
	C=C & "ID=" & Rq("ID") & vbCrlf
	'Call Rw("<pre>"& C &"</pre>"):Doevents : Response.end
	MdbPath=Request("MDBPath")
	Call OpenConn()
	Dim Sql
	Sql = "Select * From ["& Request("TableName") &"] Where Id=" & Request("Id")
	'Call Rw(MDBPath & "<hr>" & Sql)
	Rs.Open Sql,Conn,1,3
	If Not Rs.Eof Then
		Rs.Delete
		Rs.Update
		Call Rw("删除ID【"& Request("Id") &"】成功")
	Else		
		Call Rw("ID【"& Request("Id") &"】不存在，删除失败")
	End If:Rs.Close
End Function
'Ajax保存内容
Function SaveAJAXEdit() 
	Dim C
	
	C=C & "MDBPath=" & Rq("MDBPath") & "("& IsNul(Rq("MDBPath")) &")" & vbCrlf
	C=C & "MDBPath=" & HandlePath(Rq("MDBPath")) & vbCrlf
	C=C & "TableName=" & Rq("TableName") & vbCrlf
	C=C & "FieldName=" & Rq("FieldName") & vbCrlf
	C=C & "Content=" & Rq("Content") & vbCrlf
	C=C & "ID=" & Rq("ID") & vbCrlf
	Call Rw("<pre>"& C &"</pre>"):Doevents
	'MdbPath=Request("MDBPath")						' 这个不需要了，因为在Conn.Asp里已经有处理这段代码了20150311
	Call OpenConn()
	Dim Sql
	Sql = "Select * From ["& Request("TableName") &"] Where Id=" & Request("Id")
	'Call Rw(MDBPath & "<hr>" & Sql)
	Rs.Open Sql,Conn,1,3
	If Not Rs.Eof Then
		Rs(Rq("FieldName")) = Request("Content")
		Rs.Update
	End If:Rs.Close
End Function
'Ajax保存文件内容
Function SaveAJAXTEXTEdit() 
	Dim C	
	C=C & "FilePath=" & Rq("FilePath") & vbCrlf 
	
	
	C=C & "content=" & Request.Form("content") & vbCrlf 
	C=C & "count=" & Request.Form("count") & vbCrlf 
	Call Rw("<pre>"& C &"</pre>"):Doevents
	Call CreateFile(Request("FilePath"),Request("content"))
End Function

'显示文章浏览数 累加
Sub ShowArticleNumb()
	Dim Id
	Id = Request("Id")
	Call OpenConn()
	Rs.Open"Select * From [Product] Where Id=" & Id,Conn,1,3
	If Not Rs.Eof Then
		Rs("Hits") = Rs("Hits")+1
		Rs.Update
		Call Rw(Rs("Hits"))
	End If:Rs.Close
End Sub

'手机文件列表
Sub MobileArticleList()
	'Call Rw(GetFText("/2.xml"))
	Dim C,Sql,nCount,I
	Call OpenConn()
	If Session("ShowMobileArticleNumb") = "" Then Session("ShowMobileArticleNumb")=12
	nCount = 0
    C=C & "<?xml version=""1.0"" encoding=""gb2312"" ?> " & vbCrlf
    C=C & "<clists> " & vbCrlf
	Sql = "Select * From [Product] Where BigClassName='信息中心' Order By Id Desc"
	Rs.Open Sql,Conn,1,1
	For I =1 To Rs.RecordCount
		If nCount>10 Then Exit For
		If I > Session("ShowMobileArticleNumb") then
			nCount=nCount+1
		    C=C & "<clist id="""& Rs("ID") &""" title="""& Rs("Title") &""" image=""0"" time="""& Rs("AddDateTime") &"""/>" & vbCrlf
		End If
	Rs.MoveNext:Next:Rs.Close
	Session("ShowMobileArticleNumb")=Session("ShowMobileArticleNumb")+nCount 
    C=C & "</clists>" & vbCrlf
	Call Rw(C)
End Sub
'会员注册
Sub UserReg()
	Dim UserName,PassWord,Email
	UserName=Replace(Request("UserName"),"'","")
	PassWord=Replace(Request("PassWord"),"'","")
	If Request("P")="" Then
		PassWord = MD5(PassWord,2) & MD5(PassWord,4)
	End If
	Email=Replace(Request("Email"),"'","")	 
	Response.Cookies("UserName")=UserName
	Response.Cookies("UserName").Expires=Date+10	
	Call OpenConn()
	Rs.Open"Select * From [Member] Where UserName='"& UserName &"'",Conn,1,3
	If Not Rs.Eof Then
		Call Rw("用户名已被注册")
	Else
		Rs.Close
		Rs.Open"Select * From [Member] Where Email='"& Email &"'",Conn,1,3
		If Not Rs.Eof Then
			Call Rw("邮箱已被注册")
		Else
			Rs.AddNew
			Rs("UserName")=UserName
			Rs("Pwd")=PassWord
			Rs("Email")=Email
			Rs("MyMoney") = 100000
			Rs.Update
			Response.Cookies("PassWord")=PassWord
			Response.Cookies("PassWord").Expires=date+10
			Session("Email")=Rs("Email")
			Session("MyMoney")=CLng(Rs("MyMoney"))
			Call Rw("OK")
		End If
		Response.End()
	End If:Rs.Close
End Sub
'会员登录
Sub UserLogin()
	Dim UserName,PassWord,autologin
	UserName=Replace(Request("UserName"),"'","")
	PassWord=Replace(Request("PassWord"),"'","")
	If Request("P")="" Then
		PassWord = MD5(PassWord,2) & MD5(PassWord,4)
	End If
	autologin=Replace(Request("autologin"),"'","")
	 
	Response.Cookies("UserName")=UserName
	Response.Cookies("UserName").Expires=date+10
	
	Call OpenConn()
'	Call Rw("UserName=" & UserName & "|PassWord=" & PassWord & "|autologin=" & autologin)
	Rs.Open"Select * From [Member] Where UserName='"& UserName &"' Or Email='"& UserName &"'",Conn,1,3
	If Rs.Eof Then
		Call Rw("用户名不存在")
	Else
		If Rs("Pwd")<>PassWord Then
			Call Rw("密码不正确")
		Else
			Response.Cookies("PassWord")=PassWord
			Response.Cookies("PassWord").Expires=date+10
			Session("Email")=Rs("Email")
			Session("MyMoney")=CLng(Rs("MyMoney"))
			Call Rw("OK")
		End If
	End If:Rs.Close
End Sub

'保存点击QQ对话信息
Sub SaveQQDialogue()
	Dim QQ,ClickStyle
	QQ = Request("QQ")
	ClickStyle = Request("ClickStyle")
	Call CreateAddFile("/QQ/QQ.txt","聊天QQ=" & QQ & "，样式=" & ClickStyle & "，时间=" & Format_Time(Now(),8))
	Call Rw("保存QQ对话成功")
End Sub 

'获得函数代码(显示帮助信息调用)
Sub GetFunctionCode(Lable)
	Dim Content
	Content = GetFText("FSO.Asp")
	Content = StrCut(Content,"Function " & Lable,"End Function",1)
	
	Call Rw(ShowAspCodeColor(Content))
End Sub
%>