<!--#Include virtual = "/Inc/_Config.Asp"--> 
<%
'ѡ������
select case request("act") 
	case "loginuser" : loginUser()		'���QQ��Ա    aa/qqajax.asp?act=loginuser&openid=2222 
end select

'���QQ��Ա
function loginUser()
	dim openid,accesstoken,nickname,qqphoto,sex,useryear
	dim sql
	openid=request("openid")
	accesstoken=request("accesstoken")
	nickname=unescape(request("nickname"))
	qqphoto=unescape(request("qqphoto"))
	sex=unescape(request("sex"))
	useryear=unescape(request("year"))
	if useryear="" then
		useryear=0
	end if
	call echo("openid",openid)
	call echo("accesstoken",accesstoken)
	call echo("nickname",nickname)
	call echo("qqphoto",qqphoto)
	call echo("sex",sex)
	call echo("useryear",useryear)
	call openconn()
	sql="select * from "& db_PREFIX &"Member where openid='"& openid &"'"
	rs.open sql,conn,1,1
	if rs.eof then
		sql="insert into "
		  conn.Execute("insert into " & db_PREFIX & "Member (openid,accesstoken,nickname,qqphoto,sex,useryear,regip) values('" & openid & "','" & accesstoken & "','" & nickname & "','" & qqphoto & "','" & sex & "'," & useryear & ",'" & getip() & "')") 
	else		
        conn.Execute("update " & db_PREFIX & "Member  set loginip='" & getip() & "',loginCount=loginCount+1 where openid='"& openid &"'") 
	end if:rs.close
end function
%>