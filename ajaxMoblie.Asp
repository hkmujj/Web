<!--#Include file = "inc/config.Asp"-->
<%
call openconn()
dim sql
sql="select * from admin"
rs.open sql,conn,1,1
Dim json : Set json = New Cls_AB_JSON        '¼ÓÔØJSONÀà
'call eerr("",TypeName(rs))
call rw(json.toJSON(rs))



%>





