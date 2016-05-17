<!--#Include virtual = "/Inc/_Config.Asp"--> 
<%


call openconn()
conn.execute("delete from xy_weburlscan")
dim url
url="http://www.maiside.net/"
conn.execute("insert into xy_weburlscan(httpurl,isThrough) values('"& url &"',true)")
call echo("提示","恢复数据完成")
%>