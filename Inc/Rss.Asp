<%@ CodePage = 65001 %> 
<!--#Include File = "Config.Asp"--> 
<% 
''时间转换函数 
Function return_RFC822_Date(ByVal myDate, ByVal TimeZone)
    Dim myDay, myDays, myMonth, myYear 
    Dim myHours, myMinutes, mySeconds 
    myDate = CDate(myDate) 
    myDay = EnWeekDayName(myDate) 
    myDays = Right("00" & Day(myDate), 2) 
    myMonth = EnMonthName(myDate) 
    myYear = Year(myDate) 
    myHours = Right("00" & Hour(myDate), 2) 
    myMinutes = Right("00" & Minute(myDate), 2) 
    mySeconds = Right("00" & Second(myDate), 2) 
    return_RFC822_Date = myDay & ", " & _ 
    myDays & " " & _ 
    myMonth & " " & _ 
    myYear & " " & _ 
    myHours & ":" & _ 
    myMinutes & ":" & _ 
    mySeconds & " " & _ 
     " " & TimeZone 
End Function 
''星期转换函数 
Function EnWeekDayName(InputDate)
    Dim Result 
    Select Case Weekday(InputDate, 1)
        Case 1 : Result = "Sun"
        Case 2 : Result = "Mon"
        Case 3 : Result = "Tue"
        Case 4 : Result = "Wed"
        Case 5 : Result = "Thu"
        Case 6 : Result = "Fri"
        Case 7 : Result = "Sat"
    End Select
    EnWeekDayName = Result 
End Function 
''月份转换函数 
Function EnMonthName(InputDate)
    Dim Result 
    Select Case Month(InputDate)
        Case 1 : Result = "Jan"
        Case 2 : Result = "Feb"
        Case 3 : Result = "Mar"
        Case 4 : Result = "Apr"
        Case 5 : Result = "May"
        Case 6 : Result = "Jun"
        Case 7 : Result = "Jul"
        Case 8 : Result = "Aug"
        Case 9 : Result = "Sep"
        Case 10 : Result = "Oct"
        Case 11 : Result = "Nov"
        Case 12 : Result = "Dec"
    End Select
    EnMonthName = Result 
End Function 
	Dim Sql,XMLContent,rsc,Web_Name,Web_Url,CategoryName,Url,ToPath
	Web_Name = "南京麦思德餐饮有限公司"
	Web_Url =  Host()
	CategoryName = "餐饮培训"
	Call OpenConn()
	''获取当前域名,其中rss.asp为你的文件名 
	Sql = "Select Top 50 * From [Product] Where OnHtml=True Order By [UpDateTime] Desc" 
	''根据自己实际修改 
	Set Rs = CreateObject("Adodb.RecordSet")
    Rs.Open Sql, Conn, 1, 1 
  '  Response.ContentType = "text/xml" 
    XMLContent = XMLContent & "<?xml version=""1.0"" encoding=""utf-8"" ?>" & vbCrlf 
    XMLContent = XMLContent & "<?xml-stylesheet type=""text/xsl"" href=""/rss/rss.xslt"" version=""1.0"" ?>" & vbCrlf 
    XMLContent = XMLContent & "<rss version=""2.0"">" & vbCrlf 
    XMLContent = XMLContent & "<channel>" & vbCrlf 
    XMLContent = XMLContent & "<title>" & Web_Name & " 的最新动态</title>" & vbCrlf 
    XMLContent = XMLContent & "<link>" & Web_Url & "</link>" & vbCrlf 
    XMLContent = XMLContent & "<language>zh-cn</language>" & vbCrlf 
    XMLContent = XMLContent & "<copyright>RSS Feed By www.hitux.com</copyright>" & vbCrlf 
    While Not Rs.Eof        
        XMLContent = XMLContent & "<item>" & vbCrlf 
        XMLContent = XMLContent & "<title><![CDATA[" & Rs("title") & "]]></title>" & vbCrlf 
		Url = Web_Url & GetHandleUrl(Rs("FileName"))	'处理URL
        XMLContent = XMLContent & "<link>" & Url & "</link>" & vbCrlf 
        XMLContent = XMLContent & "<description><![CDATA[" & Rs("WebDescription") & "]]></description>" & vbCrlf 
        XMLContent = XMLContent & "<pubDate>" & return_RFC822_Date(Rs("UpDateTime"), "08:00") & "</pubDate>" & vbCrlf 
        XMLContent = XMLContent & "<author>" & Web_Name & "</author>" & vbCrlf 
        XMLContent = XMLContent & "<category>" & CategoryName & "</category>" & vbCrlf 
        XMLContent = XMLContent & "</item>" & vbCrlf 
        Rs.MoveNext 
    Wend 
    XMLContent = XMLContent & "</channel>" & vbCrlf 
    XMLContent = XMLContent & "</rss>" & vbCrlf 
    Rs.Close 
Set Rs = Nothing 
Call CreateFolder("../Rss")
ToPath = HandlePath("../Rss/Feed.Xml")
Call WriteToFile(ToPath, XMLContent, "utf-8")
Url = "<a "& AHref(Web_Url & "/Rss/Feed.Xml","点击打开","_blank") &">"& Web_Url & "Rss/Feed.Xml" &"</a>"

Call Echo("Create Rss/Feed.Xml OK", Url)
'Response.Write XMLContent 
%> 

