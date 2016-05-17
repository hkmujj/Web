<!--#Include File = "Config.Asp"--> 
<%
'管理 
Dim Temp,Add,Sql,X,I
Dim ThisWebSite
ThisWebSite = Host			'当前域名 

Session("Administrator") = "admin"

Call OpenConn()
'选择处理
Select Case Request("act")
    Case "nav" :  Call Rw(DivNavigation("Where NavTop=True  Order By Sort Asc"))
    Case "foot" :  Call Rw(DivWebBottom())
	Case "article" : Call Rw(DivArticle)
	Case "JSShowWeb" : Call JSShowWeb()				'以JS方法显示网站一些信息，变动的
	Case "articleList","ShowProduct" 
		Session("BigClassName")=""
		Session("SmallClassName")=""
		Session("ThreeClassName")=""
		If Rq("Bid")<>"" Then ShowBidId(Rq("Bid"))
		If Rq("Sid")<>"" Then ShowSidId(Rq("Sid"))
		If Rq("Tid")<>"" Then ShowTidId(Rq("Tid"))
		ShowProduct():ArticleList() 
End Select
'JS显示网页各模块内容
Sub JSShowWeb()
	Dim S,C
	'导航
	S = DivNavigation("Where NavTop=True  Order By Sort Asc")
	S=Replace(S,vbCrlf,"\n")
	S=Replace(S,"/","\/")
	C=C & "document.all.nav.innerHTML="""& S &"""" & vbCrlf 
	'底部
	S=Replace(DivWebBottom,vbCrlf,"\n")
	S=Replace(Replace(S,"/","\/"),"""","\""")
	C=C & "document.all.foot.innerHTML="""& S &"""" & vbCrlf
	
	Call Rw(C)
End Sub
'Div版网站导航
Function DivNavigation(AddSql)
	Dim I, C, Url, FileName, BigFolder, Did, ShowDid, Target 
	Rs.Open "Select * From [NavBigClass] " & AddSql, Conn, 1, 1 
	C = C & "<li class=left></li>" & vbCrlf
	For I = 1 To Rs.RecordCount
		FileName = Rs("FileName") 
		If FileName <> "" And Right(FileName,1)<>"/" Then FileName = FileName & ".html"  
	
		BigFolder = "/" & Rs("FolderName") & "/" 	
		Did = Rs("BigClassName") 
		ShowDid = FontColorFontB(Did, Rs("FontB"), Rs("FontColor")) 
		Target = Rs("Target") 
		
		Url = Rs("FileName")
		If CheckRemoteUrl(Url)=False Then		
			Url = BigFolder & FileName 
			Url = Replace(Replace(Url, "//", "/"), "//", "/")
			Url = FullHttpUrl(ThisWebSite,Url)	
			Call WebDebug(Url,"act=Nav&NavDid=" & Did)'调试
		End If
		
		C = C & "<li><a " & AHref(Url, "", Target) & ">" & ShowDid & "</a></li>" &vbCrlf
		C = C & "<li class=line></li>" & vbCrlf
	Rs.MoveNext : Next : Rs.Close 
	C = C & "<li class=right></li>" & vbCrlf
	DivNavigation = C 
End Function
'Div版网站底部
Function DivWebBottom()
	Rs.Open"Select * From [WebSite]",Conn,1,1
	If Not Rs.Eof Then
		DivWebBottom=Rs("WebSiteBottom")
	End If:Rs.Close
End Function
'Div版显示文章
Function DivArticle()
    Dim C,Id
	Id=Trim(Request("Id"))
	Rs.Open"Select * From [Product] Where Id=" & Id,Conn,1,1
	If Not Rs.Eof Then
		C=C & "<div class=""your_position"">首页 ><Span>信息中心</Span></div>" & vbCrlf
		C=C & "<div class=""allpagenewstitle"">"& Rs("Title") &"</div>" & vbCrlf
		C=C & "<div  class=""allpagenewsinfo"">作者："& Rs("Author") &"　　发布时间："& Rs("AddDateTime") &"</div>" & vbCrlf
		C=C & "<div class=""allpagenewscontent"">" & vbCrlf
		C=C & Rs("Content") & vbCrlf
		C=C & "</div>" & vbCrlf
	Else
		C = "没有文章"
	End If:Rs.Close
	DivArticle = C
End Function
Function ShowTidId(Id)
	Rs.Open"Select * From [ThreeClass] Where Id="&Id,Conn,1,1
	If Not Rs.Eof Then
		Session("BigClassName")=Rs("BigClassName")
		Session("SmallClassName")=Rs("SmallClassName")
		Session("ThreeClassName")=Rs("ThreeClassName")
	End If:Rs.Close
End Function
'ShowSigId子类信息
Function ShowSidId(Id)
	Rs.Open"Select * From [SmallClass] Where Id="&Id,Conn,1,1
	If Not Rs.Eof Then
		Session("BigClassName")=Rs("BigClassName")
		Session("SmallClassName")=Rs("SmallClassName")
		Session("ThreeClassName")=""
	End If:Rs.Close
End Function
'ShowBigId大类信息
Function ShowBidId(Id)
	Rs.Open"Select * From [BigClass] Where Id="&Id,Conn,1,1
	If Not Rs.Eof Then
		Session("BigClassName")=Rs("BigClassName")
		Session("ThreeClassName")=""
		Session("SmallClassName")=""
	End If:Rs.Close
End Function
'ShowProduct
Function ShowProduct()
	If Session("ThreeClassName")<>"" Then
		Add=" Where BigClassName='"&Session("BigClassName")&"' And SmallClassName='"&Session("SmallClassName")&"' And ThreeClassName='"&Session("ThreeClassName")&"'"
	ElseIf Session("SmallClassName")<>"" Then
		Add=" Where BigClassName='"&Session("BigClassName")&"' And SmallClassName='"&Session("SmallClassName")&"'"
	ElseIf Session("BigClassName")<>"" Then
		Add=" Where BigClassName='"&Session("BigClassName")&"'"
	End If
End Function

'文章展示
Function ArticleList()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<title>文章列表</title>
</head>
<body>
<script language="javascript" src=""& adminDir &"Images/Function.Js"></script> 
<link href=""& adminDir &"Css/Main.Css" rel="stylesheet"  type="text/css">
<link href=""& adminDir &"Css/BodyCss.Css" rel="stylesheet"  type="text/css">
<div class="ClickYes">
<div class="Div800Center">
<%
Rw("&nbsp;<a href=""?act=ShowProduct"">产品展示 >> </a>")
Rs.Open"Select BigClassName,Id From [BigClass] Order By Sort",Conn,1,1
While Not Rs.Eof
If Session("BigClassName")=Rs(0) Then Temp="<font color=red>"&Rs(0)&"</font>" Else Temp=Rs(0)
Rw("<a href=?act=ShowProduct&Bid="&Rs("Id")&">"&Temp&"</a> | ")
Rs.MoveNext:Wend:Rs.Close
%>
</div>
<div class="div5"></div>
<%If Session("BigClassName")<>"" Then
	Rs.Open"Select BigClassName,SmallClassName,Id From [SmallClass] Where BigClassName='"&Session("BigClassName")&"' Order By Sort",Conn,1,1
	If Not Rs.Eof Then
	Rw("<div class=""Div800Center"" style=""background:#C1CDE3;"">")
	While Not Rs.Eof
	If Session("SmallClassName")=Rs(1) Then Temp="<font color=red>"&Rs(1)&"</font>" Else Temp=Rs(1)
	Rw("<a href=?act=ShowProduct&Sid="&Rs("Id")&">"&Temp&"</a> | ")
	Rs.MoveNext:Wend
	Rw("</div><div class=""div5""></div>")
	End If:Rs.Close
End If
%>
<%If Session("SmallClassName")<>"" Then
	Rs.Open"Select BigClassName,SmallClassName,ThreeClassName,Id From [ThreeClass] Where BigClassName='"&Session("BigClassName")&"' And SmallClassName='"&Session("SmallClassName")&"' Order By Sort",Conn,1,1
	If Not Rs.Eof Then
	Rw("<div class=""Div800Center"" style=""background:#C1CDE3;"">")
	While Not Rs.Eof
	If Session("ThreeClassName")=Rs(2) Then Temp="<font color=red>"&Rs(2)&"</font>" Else Temp=Rs(2)
	Rw("<a href=?act=ShowProduct&Tid="&Rs("Id")&">"&Temp&"</a> | ")
	Rs.MoveNext:Wend
	Rw("</div><div class=""div5""></div>")
	End If:Rs.Close

End If
%>
<table width="98%" border="0" cellpadding="0" cellspacing="1" bgcolor="#000000" class="border" align="center">
<form id="formPro" name="formPro" method="post" action="?act=aa">
<tr bgcolor="A4B6D7" class="title">
<td width="30" height="25" align="center" bgcolor="A4B6D7"><b>选中</b></td>
<td width="100" height="25" align="center"><span style="font-weight: bold">信息大类</span></td>
<td width="78" align="center" bgcolor="A4B6D7"><span style="font-weight: bold">信息小类</span></td>
<td width="78" align="center"><span style="font-weight: bold">信息子类</span></td>
<td align="center"><span style="font-weight: bold">信息名称</span></td>
 
<td width="40" align="center"><b>审核</b></td>
<td width="81" align="center"><span style="font-weight: bold">发布时间</span></td>
<td width="92" align="center"><b>操作选项</b></td>
</tr>
<%
If Session("BigClassName")<>"" Or Session("SmallClassName")<>"" Or Session("ThreeClassName")<>"" Then
Add=" Where "
	If Session("BigClassName")<>"" Then Add=Add&" BigClassName='"&Session("BigClassName")&"'"	
	If Session("SmallClassName")<>"" Then Add=Add&" And SmallClassName='"&Session("SmallClassName")&"'"
	If Session("ThreeClassName")<>"" Then Add=Add&" And ThreeClassName='"&Session("ThreeClassName")&"'"
ElseIf Rf("Search")<>"" Then
	Add=" Where Title like '%"&Rf("Search")&"%'"
ElseIf Rf("Sql")<>"" Then
	Add=" "&Rf("Sql")
End If
If Rq("page") <> "" Then Temp = "&page=" & Rq("page") Else Temp = "" 
Sql="Select * From [Product]"&Add&" Order By Sort"
Rs.Open Sql,Conn,1,1
If Rs.Eof Then
'Rw "暂无信息……"
Else
Rs.PageSize = 11 '每页记录条数
nCount=Rs.RecordCount '记录总数
nPageSize=Rs.PageSize
MaxPage=Rs.PageCount 
Page=request("page")
If Not IsNumeric(Page) Or Page="" Then
	Page=1
Else
	Page=Cint(Page)
End If
If Page < 1 Then
	Page = 1
ElseIf  Page > MaxPage Then
	Page = MaxPage
End If
Rs.AbsolutePage = Page
If Page = Maxpage Then
	X = nCount-(MaxPage-1)*nPageSize
Else
	X = nPageSize
End If
For I=1 To x
%>
<tr bgcolor="#E8FFED" id="TR83" class="tdbg" onMouseOver="onColor(this);" onMouseOut="offColor(this);">
<td height="22" align="center" >
<input name="selectPid" type="checkbox" id="selectPid" value="<%=Rs(0)%>" onClick="ShowColor(this,1)" />
<input name="Id" type="hidden" id="Id" value="<%=Rs(0)%>"></td>
<td width="100" align="center"><%=Rs("BigClassName")%>&nbsp;</td>
<td width="100" align="center"><%=Rs("SmallClassName")%>&nbsp;</td>
<td width="100" align="center"><%=Rs("ThreeClassName")%>&nbsp;</td>
<td align="left"><a href="#" title="文章标题" name="<%=Rs("ID")%>"><%=Rs("Title")%></a></td>
 
<td align="center"><%
 
	Call Rw(ShowOnOffImg(Rs(0), "Product", "Through", Rs("Through"), "/Inc/CallWeb.Asp?act=ShowProduct"))
 
%></td>
<td align="center"><%=Rs("UpDateTime")%>&nbsp;</td>
<td align="center" bgcolor="#E8FFED"><a href="?act=ShowEditProduct&amp;Id=<%=Rs(0)&Temp%>">修改</a> | <a href="?act=DelId&Id=<%=Rs(0)&Temp%>" onClick="return CheckDel();">删除</a></td>
</tr>
<%Rs.MoveNext:Next:End If:Rs.Close%>
<tr bgcolor="#A4B6D7" class="tdbg">
  <td height="22" colspan="9" align="left" ><a href="javascript:Checkmm('全选');">全选</a> <a href="javascript:Checkmm('反选');">反选</a><span style="text-align:center;">
  
    <select name="Sel" id="Sel" onchange="javascript:PROonclick();">
      <option>请选择操作</option>
      <%	  
	  Rw("<option value=""?act=AllSave"&Temp&""">更新信息</option>")
	  Rw("<option value=""?act=AllDel"&Temp&""">删除选择</option>")
	  %>
    </select>
  </span></td>
</tr></form>
</table>

<div style="text-align:center;">
  <%call Rw(PageControl(nCount,maxpage,page,"border=0 align=center","<p align=center>"))%></div>
</div>
</body></html>
<%End Function%>