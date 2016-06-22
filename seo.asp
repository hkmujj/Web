<!--#Include virtual="/Inc/_Config.Asp"-->
<%
dim httpurl,setcode,content,web_title,web_keyword,web_description,web_body
httpurl="http://bbs.moonseo.cn/article-115-1.html"
setcode="gb2312"
if request("act")="getseo" then
	httpurl=request("httpurl")
	setcode=request("setcode")
	content=gethttpurl(httpurl,setcode)
	web_title= RegExpGetStr("<TITLE>(.*)</TITLE>", content, 1)   

web_description = RegExpGetStr("<meta[^<>]*description[^<>]*[\/]?>", Content, 0)
web_description = RegExpGetStr("content=[""|']?([^""' ]*)([""|']?).*[\/]?>", web_description, 1)
web_keyword = RegExpGetStr("<meta[^<>]*keywords[^<>]*[\/]?>", Content, 0)
web_keyword = RegExpGetStr("content=[""|']?([^""' ]*)([""|']?).*[\/]?>", web_keyword, 1)
web_body=getstrcut(content,"<div class=""seoul_li"">","<div class=""d"">",2)
web_body=replace(web_body,"<","&lt;")

	call echo("title",web_title)
	call echo("keyword",web_keyword)
	call echo("description",web_description)
	call echo("body",web_body)
	call createfile("1.txt",content)
end if
%>
<form name="form1" method="post" action="?act=getseo">
  <input name="httpurl" type="text" id="httpurl" size="60" value="<%=httpurl%>">
  <select name="setcode" id="setcode">
    <option value="gb2312">gb2312</option>
    <option value="utf-8"<%if setcode="utf-8" then call rw(" selected")%>>utf-8</option>
  </select>
  <input type="submit" name="button" id="button" value="Ìá½»">
</form>