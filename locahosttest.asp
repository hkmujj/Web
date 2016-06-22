<%

'2016五一活动
if request("act")="20160501huodong" then
	response.Redirect("http://10.10.10.57/aspweb.asp?templatedir=E%3A%5CE%u76D8%5CWEB%u7F51%u7AD9%5C%u81F3%u524D%u7F51%u7AD9%5CTemplates2015%5C2016%u4E94%u4E00%u6D3B%u52A8/&templateName=Index_Model.html&nRnd=p1B2v8n4Y8o")
else
'response.Redirect("http://10.10.10.57/aspweb.asp?templatedir=E%3A%5CE%u76D8%5CWEB%u7F51%u7AD9%5C%u81F3%u524D%u7F51%u7AD9%5CTemplates2015%5C%u6D3B%u52A820160411/&templateName=Index_Model.html&nRnd=l7X2d2f4G4h")
response.Write(request("act"))
end if
%>