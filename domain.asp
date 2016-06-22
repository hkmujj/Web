<!--#Include virtual = "/Inc/_Config.Asp"--> 
<% 

'26字母相互组合获得676条变量列表
Dim d, s, i, c, j, j3,s3,j4,s4,s2, nCount, doMain, content, domainState, regDomainList
dim domainFile,nOK,nBigCount
domainFile="domain.txt"
content = phptrim(getftext(domainFile)) 
'call rwend(content)
nCount = 0 : nOK=0 : nBigCount=50
d = "abcdefghijklmnopqrstuvwxyz"
d = "abcdefghijklmnopqrstuvwxyz"

call batchCheckDomain("aaoo","#left",1)
call echo("可注册域名列表",regDomainList)


'批量检测域名
function batchCheckDomain(domainPrefix,align,sType)
	For i = 1 To Len(d)
		s=Mid(d, i, 1) 
		s=iif(align="left",s & domainPrefix, domainPrefix & s)
		nCount = nCount + 1 
		If InStr("," & content & "|", "," & s & "|") = False Then
			domainState = IIF(getDomainState(s), 1, 0) 
			content = content & s & "|" & domainState & "," 
			Call echo("域名", s & "(" & domainState & ")")
			Call createfile(domainFile, content)
			doevents()
			nOK=nOK+1
			if nOK>nBigCount then
				exit function
			end if
		End If 
		If InStr("," & content & "|", "," & s & "|0,") > 0 Then
			regDomainList = regDomainList & s & "|" 
		End If
		if sType=2 then
			For j = 1 To Len(d)
				s2 = Mid(d, j, 1) 
				nCount = nCount + 1 				
				doMain=iif(align="left",s2 & s , s & s2 )
				If InStr("," & content & "|", "," & doMain & "|") = False Then
					domainState = IIF(getDomainState(doMain), 1, 0) 
					content = content & doMain & "|" & domainState & ","
					Call echo("域名", doMain & "(" & domainState & ")") 
					Call createfile(domainFile, content) 
					doevents() 
					nOK=nOK+1
					if nOK>nBigCount then
						exit function
					end if
				End If 
				If InStr("," & content & "|", "," & doMain & "|0,") > 0 Then
					regDomainList = regDomainList & doMain & "|" 
				End If 
			next
		end if
	next
end function
'获得域名状态
Function getDomainState(doMain)
    getDomainState = getDomainState2(doMain)
End Function
'获得域名状态
Function getDomainState1(doMain)
    Dim url, s 
    url = "http://checkdomain.xinnet.com/domainCheck?callbackparam=jQuery172044711112370714545_1459417142627&searchRandom=6&prefix=" & doMain & "&suffix=.com&_=1459417428579"	
    s = gethttpurl(url, "") 
    If InStr(s, """:[],""no""") > 0 Then
        getDomainState1 = True 
    Else
        getDomainState1 = False 
    End If 
End Function
function getDomainState2(doMain)
    Dim url, s 
    url = "http://checkdomain.xinnet.com/domainCheck?callbackparam=jQuery172044711112370714545_1459417142627&searchRandom=6&prefix=" & doMain & "&suffix=.com&_=1459417428579"	
    s = gethttpurl(url, "") 
	if instr(lcase(s),"registered")=false then
        getDomainState2 = True 
    Else
        getDomainState2 = False 
    End If  
end function

'国外域名查询网站 http://www.checkdomain.com/cgi-bin/checkdomain.pl?domain=ab.com 


sub test1()
For i = 1 To Len(d)
    s = Mid(d, i, 1) 
    nCount = nCount + 1 
    If InStr("," & content & "|", "," & s & "|") = False Then
        domainState = IIF(getDomainState(s), 1, 0) 
        content = content & s & "|" & domainState & "," 
        Call echo("域名", s & "(" & domainState & ")") 
        Call createfile(domainFile, content) 
        doevents() 
    End If 

    If InStr("," & content & "|", "," & s & "|0,") > 0 Then
        regDomainList = regDomainList & s & "|" 
    End If 
	
    For j = 1 To Len(d)
        s2 = Mid(d, j, 1) 
        nCount = nCount + 1 
        doMain = s & s2 
        If InStr("," & content & "|", "," & doMain & "|") = False Then
            domainState = IIF(getDomainState(doMain), 1, 0) 
            content = content & doMain & "|" & domainState & ","
            Call echo("域名", doMain & "(" & domainState & ")") 
            Call createfile(domainFile, content) 
            doevents() 
        End If 
        If InStr("," & content & "|", "," & s & "|0,") > 0 Then
            regDomainList = regDomainList & s & "|" 
        End If 
		
		
		For j3 = 1 To Len(d)
			s3 = Mid(d, j, 1) 
			nCount = nCount + 1 
			doMain = s & s2  & s3
			If InStr("," & content & "|", "," & doMain & "|") = False Then
				domainState = IIF(getDomainState(doMain), 1, 0) 
				content = content & doMain & "|" & domainState & ","
				Call echo("域名", doMain & "(" & domainState & ")") 
				Call createfile(domainFile, content) 
				doevents() 
			End If 
			If InStr("," & content & "|", "," & s & "|0,") > 0 Then
				regDomainList = regDomainList & s & "|" 
			End If 
			For j4 = 1 To Len(d)
				s4 = Mid(d, j, 1) 
				nCount = nCount + 1 
				doMain = s & s2  & s3 & s4
				If InStr("," & content & "|", "," & doMain & "|") = False Then
					domainState = IIF(getDomainState(doMain), 1, 0) 
					content = content & doMain & "|" & domainState & ","
					Call echo("域名", doMain & "(" & domainState & ")") 
					Call createfile(domainFile, content) 
					doevents() 
				End If 
				If InStr("," & content & "|", "," & s & "|0,") > 0 Then
					regDomainList = regDomainList & s & "|" 
				End If 
			Next
		Next
    Next 
Next 
call echo("总数",nCount)
Call echo("可注册域名", regDomainList)
end sub
%> 
