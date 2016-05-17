<%
'引用别人

'@ *****************************************************************************
'@ 过程名:  Util.UBB.closeHtml(s)
'@ 返  回:  String (字符串)
'@ 作  用:  自动闭合html
'@ 			目前支持闭合的标签有:
'@ 			 p,DIV,span,table,ul,font,b,u,i,h1,h2,h3,h4,h5,h6
'==Param========================================================================
'@ s 	: 字符串 [String]
'==DEMO=========================================================================
'@ AB.use "Mvc" : Util.use "UBB" : AB.C.Print Util.UBB.closeHtml("<span><a href='#'>我</a>")
'@ *****************************************************************************
'call rw(closeHtml("<b><b></b>"))  这种会报错20150718
'闭合标签，但对累加处理不行
Function closeHtml(Byval s)
	on error resume next
	Dim str, Tags, arrTags, i, OpenPos, ClosePos, re, strMatchs, j, Match : str = s
	Set re = CreateObject("VBscript.RegExp")
	re.IgnoreCase = True
	re.Global = True
	Tags = "p,div,span,table,ul,font,b,u,i,h1,h2,h3,h4,h5,h6" '设置需要闭合的标签
	arrTags = Split(Tags,",")
	For i = 0 To UBound(arrTags)
		OpenPos = 0 '标签开始标记个数
		ClosePos = 0 '标签结束标记个数
		re.Pattern = "<(" + trim(arrTags(i)) + ")((\s+[^<>])+|)>"
		Set strMatchs = re.Execute(str)
		For Each Match in strMatchs
			OpenPos = OpenPos + 1
			str = Replace(str, Match.Value, "<"& LCase(Match.SubMatches(0)) & Match.SubMatches(1) &">") '转为小写
		Next
		re.Pattern = "<\/" + arrTags(i) + ">"
		Set strMatchs = re.Execute(str)
		For Each Match in strMatchs
			ClosePos = ClosePos + 1
			str = Replace(str, Match.Value, "</"& LCase(Match.SubMatches(0)) &">") '转为小写
		Next
		For j = 1 To OpenPos - ClosePos '当开始与结束标记数量不一致时，闭合当前标签
			str = str + "</" + arrTags(i) + ">"
		Next
	Next
	closeHtml = str
	Set re = Nothing
End Function
%>