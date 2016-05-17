<% 
'对ASP处理



'批量获得函数 20150402  有点慢，不过正确的，以后慢慢改吧
Function BatchFindFunction(ConfigPath,VariableList)
	Dim ConfigContent,LCaseConfigContent,SplStr,VariableName,C
	Dim ObjCode,Content
	If CheckFile(ConfigPath)=False Then
		BatchFindFunction=""
		Call Echo("配置文件不存在", ConfigPath)
		Exit Function
	End If
	ConfigContent=GetFText(ConfigPath)
	LCaseConfigContent = LCase(ConfigContent)
	LCaseConfigContent=Replace(Replace(LCaseConfigContent, ",","|"), vbCrlf,"|")
	SplStr=Split(VariableList,"|")
	For Each VariableName In SplStr
		If VariableName<>"" And InStr(LCaseConfigContent,"|"& VariableName &"|") Then
			C=C & GetFindFileFunction(ConfigPath,VariableName) & vbCrlf 
		End If
	Next
	BatchFindFunction = C
End Function
'获得指定函数在那个文件里代码 20150402
Function GetFindFileFunction(ByVal ConfigPath,ByVal FunctionName)
	Dim FilePath,ConfigContent,SplStr,Splxx,S,S2
	Dim ObjCode,Content
	FunctionName=LCase(PHPTrim(FunctionName))
	If CheckFile(ConfigPath)=False Then
		GetFindFileFunction=""
		Call Echo("配置文件不存在", ConfigPath)
		Exit Function
	End If
	ConfigContent=GetFText(ConfigPath)
	SplStr=Split(ConfigContent,vbCrlf)
	For Each S In SplStr
		If InStr(S,",")>0 Then
			FilePath=Mid(S,1,InStr(S,",")-1)
			If InStr(FilePath,"(")>0 Then
				FilePath=Mid(FilePath,1,InStr(FilePath,"(")-1)
			End If
			S2=LCase(Mid(S,InStrRev(S,",")+1))
			If InStr("|"& S2 &"|","|"& FunctionName &"|")>0 Then
				'Call Echo(S,S2)
				Set ObjCode = New ASPCode 
				ObjCode.SetCountRow = 9999
				ObjCode.SetAspSetType = "||"		
				ObjCode.SetFindFunctionList=FunctionName
				Content=GetFText(FilePath)
				Content = ObjCode.HandleAspCode(Content)				'获得内容 
				'Call Eerr(FunctionName,ObjCode.FindFunctionContent)
				GetFindFileFunction = ObjCode.getFindFunctionContent
				Set ObjCode=Nothing
				Exit Function
			End If
		End If
	Next
End Function

'获得ASP代码
Function GetASPCode(ByVal Content)
    GetASPCode = HandleGetAspCode(Content, "") 
End Function
 
'处理获得Asp代码 C = HandleGetAspCode(C,"|第一段|去ASP开始结束标签|去注释部分|(每行不为空|去空行)|(全部为ASP|默认ASP自检打开)|处理成Execute代码|")
'以后再加强，可截取只定段数代码，如第二段，最后一段 20150329
Function HandleGetAspCode(ByVal Content, SType)
    Dim SplStr, Sx, S, C, I, YesASP, SYHCount, Wc, AspCode, TempAspCode, vbCrlfYes 
    YesASP = False                                                               '是ASP 默认为假
    SYHCount = 0                                                                 '双引号默认为0
    SplStr = Split(Content, vbCrLf)                                              '分割行
	
	'20150328
	If InStr("|" & SType & "|", "|全部为ASP|") > 0 Or InStr("|" & SType & "|", "|默认ASP自检打开|") > 0 Then
		YesASP = True                                                               '是ASP 默认为真 
	End If
    '循环分行
    For Each S In SplStr
        '循环每个字符
        AspCode = "" 
        For I = 1 To Len(S)
            Sx = Mid(S, I, 1) 
            'Asp开始
            If Sx = "<" And YesASP = False And Wc = "" Then
                If Mid(S, I + 1, 1) = "%" Then
                    YesASP = True                                                      'ASP为真
                    Sx = "<" & "%" 
                    I = I + 1 
                    If InStr("|" & SType & "|", "|去ASP开始结束标签|") = False Then
                        AspCode = AspCode & Sx 
                    End If 
                End If 
            'ASP结束
            ElseIf Sx = "%" And Mid(S, I + 1, 1) = ">" And YesASP = True And Wc = "" Then
                YesASP = False                                                                  'ASP为假
                Sx = "%" & ">" 
                I = I + 1 
                If InStr("|" & SType & "|", "|去ASP开始结束标签|") = False Then
                    C = C & Sx & vbCrLf 
                End If 
                If InStr("|" & SType & "|", "|第一段|") > 0 Or InStr("|" & SType & "|", "|第一段代码|") > 0 Or InStr("|" & SType & "|", "|获取第一段|") > 0 Or InStr("|" & SType & "|", "|1|") > 0 Then
                    HandleGetAspCode = C & AspCode 
                    Exit Function 
                End If 
            'ASP为真
            ElseIf YesASP = True Then
                '输入文本
                If(Sx = """" Or Wc <> "") Then
                    '双引号累加
                    If Sx = """" Then SYHCount = SYHCount + 1 
                    Wc = Wc & Sx
				    '判断是否"在最后
                    If SYHCount Mod 2 = 0 Then
                        If Mid(S, I + 1, 1) <> """" Then 
                            AspCode = AspCode & Wc                                          '函数代码累加
                            SYHCount = 0 : Wc = ""                                  '清除
               			End If
                    End If 
                ElseIf Sx = "'" Then                                                            '注释则退出
                    If InStr("|" & SType & "|", "|去注释部分|") > 0 Then
                        Exit For 
                    End If 
                    AspCode = AspCode & Mid(S, I) 
                    Exit For 
                Else
                    AspCode = AspCode & Sx 
                End If 

            End If 
        Next 
        If YesASP = True Then			'
            vbCrlfYes = True 
            If TrimVbCrlf(AspCode) = "" Then
                If InStr("|" & SType & "|", "|每行不为空|") > 0 Or InStr("|" & SType & "|", "|去空行|") > 0 Then
                    vbCrlfYes = False 
                End If 
            End If 
            If vbCrlfYes = True Then
                If InStr("|" & SType & "|", "|处理成Execute代码|") > 0 Then
                    AspCode = TrimVbCrlf(AspCode) : TempAspCode = LCase(AspCode) 
                    C = C & AspCode & " : " 
                    If Left(TempAspCode, 3) = "if " And Right(TempAspCode, 5) <> " then" Then
                        C = C & " End If : " 
                    End If 
                Else
					
        			AspCode = RTrim(AspCode)            '去除右边空格
                    C = C & AspCode & vbCrlf 
                End If 
            End If 
        End If 
    Next
	If InStr("  " & C,vbCrlf)>0 Then C = Left(C,Len(C)-2) 
    HandleGetAspCode = C 
End Function
 
'删除ASP里注释 (精)
Function DelAspNote(ByVal Content)
    Dim SplStr, I, YesASP, YesWord, Sx, S, Wc, AspCode, SYHCount 
    Dim YesFunction                                                 '函数是否为真
    YesASP = False                                                  '是ASP 默认为假
    YesFunction = False                                             '是函数 默认为假
    YesWord = False                                                 '是单词 默认为假
    SYHCount = 0                                                    '双引号默认为0
    SplStr = Split(Content, vbCrLf)                                 '分割行
    '循环分行
    For Each S In SplStr
        '循环每个字符
        For I = 1 To Len(S)
            Sx = Mid(S, I, 1) 
            'Asp开始
            If Sx = "<" And Wc = "" Then                            '输出文本必需为空 Wc为输出内容 如"<%" 排除 修改于20140412
                If Mid(S, I + 1, 1) = "%" Then
                    YesASP = True                             'ASP为真
                    Sx = "<" & "%" 
                    I = I + 1                                 '加1而不能加2，要不然<%function Test() 就截取不到
                End If 
            'ASP结束
            ElseIf Sx = "%" And Mid(S, I + 1, 1) = ">" And Wc = "" Then'Wc为输出内容
                YesASP = False                'ASP为假
                Sx = "%" & ">" 
                I = I + 1                     '不能加2，只能加1，因为这里定义ASP为假，它会在下一次显示上面的 'ASP运行为假
            End If 
            If YesASP = True Then
                '输入文本
                If(Sx = """" Or Wc <> "") Then
                    '双引号累加
                    If Sx = """" Then SYHCount = SYHCount + 1 
                    Wc = Wc & Sx 
                    '判断是否"在最后
                    If SYHCount Mod 2 = 0 Then
                        If Mid(S, I + 1, 1) <> """" Then
                            AspCode = AspCode & Wc '函数代码累加
                            SYHCount = 0 : Wc = "" '清除
                        End If 
                    End If 
                ElseIf Sx = "'" Then          '注释则退出
                    Exit For 
                Else
                    AspCode = AspCode & Sx 
                End If 
            Else
                AspCode = AspCode & Sx 
            End If 
            DoEvents 
        Next 
        AspCode = RTrim(AspCode)   '去除右边空格
        AspCode = AspCode & vbCrLf 
        DoEvents 
    Next 
    If AspCode <> "" Then AspCode = Left(AspCode, Len(AspCode) - 2) 
    DelAspNote = AspCode 
End Function

'寻找ASP函数列表 20140811
'nSub与nFunction 是为了统计用的，默认为空就好  
'Call FindAspFunction(Content,0,0,"(全部为ASP|默认ASP自检打开)|变量处理|加密变量")
Function FindAspFunction(Content, nSub, nFunction, SType)
    Dim SplStr, I, YesASP, YesWord, Sx, S, Wc, Zc, S1, AspCode, SYHCount
	Dim UpWord																	 '上一个词 
    Dim TempS, DimList 
    Dim YesFunction                                                              '函数是否为真 
    YesASP = False                                                               '是ASP 默认为假
    YesFunction = False                                                          '是函数 默认为假
    YesWord = False                                                              '是单词 默认为假
    SYHCount = 0                                                                 '双引号默认为0
    SplStr = Split(Content, vbCrLf)                                              '分割行
    If nSub = "" Then nSub = 0                                                   '为空则默认为零
    If nFunction = "" Then nFunction = 0                                         '为空则默认为零 
	
	'20150328
	If InStr("|" & SType & "|", "|全部为ASP|") > 0 Or InStr("|" & SType & "|", "|默认ASP自检打开|") > 0 Then
		YesASP = True                                                               '是ASP 默认为真 
	End If
    '循环分行
    For Each S In SplStr
        '循环每个字符
        For I = 1 To Len(S)
            Sx = Mid(S, I, 1) 
            'Asp开始
            If Sx = "<" And Wc = "" Then                            '输出文本必需为空 Wc为输出内容 如"<%" 排除 修改于20140412
                If Mid(S, I + 1, 1) = "%" Then
                    YesASP = True                                                                   'ASP为真
                    I = I + 1                                                                       '加1而不能加2，要不然<%function Test() 就截取不到
                End If 
            'ASP结束
            ElseIf Sx = "%" And Mid(S, I + 1, 1) = ">" And Wc = "" Then'Wc为输出内容
                YesASP = False                                                       'ASP为假
                I = I + 1                                                            '不能加2，只能加1，因为这里定义ASP为假，它会在下一次显示上面的 'ASP运行为假
            End If 
            If YesASP = True Then
                '输入文本
                If(Sx = """" Or Wc <> "") Then
                    '双引号累加
                    If Sx = """" Then SYHCount = SYHCount + 1 
                    Wc = Wc & Sx 
                    '判断是否"在最后
                    If SYHCount Mod 2 = 0 Then
                        If Mid(S, I + 1, 1) <> """" Then
                            S1 = Right(Replace(Mid(S, 1, I - Len(Wc)), " ", ""), 1) '必需放在这里，要不会出错
                            If YesFunction = True Then AspCode = AspCode & Wc       '函数代码累加
                            SYHCount = 0 : Wc = ""                                  '清除
                        End If 
                    End If 
                ElseIf Sx = "'" Then                                                 '注释则退出
                    If YesFunction = True Then AspCode = AspCode & Mid(S, I) 
                    Exit For 
                '字母
                ElseIf CheckABC(Sx) = True Or(Sx = "_" And Zc <> "") Or Zc <> "" Then
                    Zc = Zc & Sx 
                    S1 = LCase(Mid(S & " ", I + 1, 1)) 
                    If InStr("abcdefghijklmnopqrstuvwxyz0123456789", S1) = 0 And Not(S1 = "_" And Zc <> "") Then'最简单判断
                        TempS = Mid(S, I + 1) 
                        If InStr("|function|sub|", "|" & LCase(Zc) & "|") Then
                            '函数开始
                            If YesFunction = False And LCase(UpWord) <> "end" Then
                                YesFunction = True 
                                If LCase(Zc) = "sub" Then nSub = nSub + 1 Else nFunction = nFunction + 1 
                                DimList = GetFunDimName(TempS) 
                            ElseIf YesFunction = True And LCase(UpWord) = "end" Then'获得上一个单词
                                AspCode = AspCode & Zc & vbCrLf 
                                If InStr("|" & SType & "|", "|加密变量|") > 0 Then AspCode = AspCode & "'[NOHandle]" & EncCode(DimList) & vbCrLf & vbCrLf 
                                YesFunction = False 
                            End If 
                        ElseIf YesFunction = True And LCase(Zc) = "dim" Then
                            DimList = DimList & "," & GetDimName(TempS) 
                        ElseIf YesFunction = True And InStr("|" & SType & "|", "|变量处理|") > 0 Then 
                            Zc = ReplaceDim2(DimList, Zc)
                        End If 
                        UpWord = Zc                                           '记住当前单词
                        If YesFunction = True Then AspCode = AspCode & Zc 
                        Zc = "" 
                    End If 
                ElseIf YesFunction = True Then
                    AspCode = AspCode & Sx 
                End If 
            End If 
            DoEvents 
        Next 
        AspCode = RTrim(AspCode)  '去除右边空格
        If YesFunction = True Then
            AspCode = AspCode & vbCrLf                                 'Asp换行
        End If 
        DoEvents 
    Next 
    FindAspFunction = AspCode 
End Function


'ASP代码首字母大写 (待完善)
Function FirstLetterUppercase(Content)
    Dim SplStr, I, YesASP, YesWord, Sx, S, Wc, Zc, S1, AspCode, SYHCount 
    Dim YesFunction                                                      '函数是否为真

    YesASP = False                                                       '是ASP 默认为假
    YesFunction = False                                                  '是函数 默认为假
    YesWord = False                                                      '是单词 默认为假
    SYHCount = 0                                                         '双引号默认为0
    SplStr = Split(Content, vbCrlf)                                      '分割行
    '循环分行
    For Each S In SplStr
        '循环每个字符
        For I = 1 To Len(S)
            Sx = Mid(S, I, 1) 
            'Asp开始
            If Sx = "<" And Wc = "" Then                            '输出文本必需为空 Wc为输出内容 如"<%" 排除 修改于20140412
                If Mid(S, I + 1, 1) = "%" Then
                    YesASP = True                                                                   'ASP为真
                    I = I + 1                                                                       '加1而不能加2，要不然<%function Test() 就截取不到
                    Sx = "<%" 
                End If 
            'ASP结束
            ElseIf Sx = "%" And Mid(S, I + 1, 1) = ">" And Wc = "" Then'Wc为输出内容
                YesASP = False                                                       'ASP为假
                I = I + 1                                                            '不能加2，只能加1，因为这里定义ASP为假，它会在下一次显示上面的 'ASP运行为假
                AspCode = AspCode & "%" & ">" 
            End If 
            If YesASP = True Then
                '输入文本
                If(Sx = """" Or Wc <> "") Then
                    '双引号累加
                    If Sx = """" Then SYHCount = SYHCount + 1 
					Wc = Wc & Sx
                    '判断是否"在最后
                    If SYHCount Mod 2 = 0 Then
                        If Mid(S, I + 1, 1) <> """" Then 
                            S1 = Right(Replace(Mid(S, 1, I - Len(Wc)), " ", ""), 1) '必需放在这里，要不会出错
                            If YesFunction = True Then AspCode = AspCode & Wc       '函数代码累加
                            SYHCount = 0 : Wc = ""                                  '清除
						End If
                    End If 
                ElseIf Sx = "'" Then                                                 '注释则退出
                    If YesFunction = True Then AspCode = AspCode & Mid(S, I) 
                    Exit For 
                '字母
                ElseIf CheckABC(Sx) = True Or(Sx = "_" And Zc <> "") Or Zc <> "" Then
                    Zc = Zc & Sx 
                    S1 = LCase(Mid(S & " ", I + 1, 1)) 
                    If InStr("abcdefghijklmnopqrstuvwxyz0123456789", S1) = 0 And Not(S1 = "_" And Zc <> "") Then'最简单判断
                        AspCode = AspCode & UCase(Left(Zc, 1)) & Mid(Zc, 2) 
                        Zc = "" 
                    End If 
                Else
                    AspCode = AspCode & Sx 
                End If 
            End If 
            DoEvents 
        Next 
        AspCode = Rtrim(AspCode)   '去除右边空格
        If YesFunction = True Then
            AspCode = AspCode & vbCrlf                                 'Asp换行
        End If 
        AspCode = AspCode & vbCrlf 
        DoEvents 
    Next 
    FirstLetterUppercase = AspCode 
End Function 

'字符大小写随机显示 20150329
Function ContentLUShow(Content)
	Dim I,S,C
    For I = 1 To Len(Content)
        S = Mid(Content, I, 1) 
		S = IIF(PHPRnd(0,1)="0",LCase(S),UCase(S))
		C=C & S
	Next
	ContentLUShow = C
End Function
'ASP转VB内容 20150727
function ASPToVB(content)
	'判断是否移除 '开始'结束 中间内容 
	content= deleteStrCut(content,vbCrlf & "'开始",vbCrlf & "'结束",1,0)
	'判断是否移除 '开始'结束 中间内容 
	content= deleteStrCut(content,vbCrlf & "'VB不引用 Start", vbCrlf & "'VB不引用 End",1,0)
	'Content = getStrCut(Content,"<" & "%","%" & ">",2)
	Content = Replace(Content,"Response.End()","Response.End")
	Content = Replace(Content,"Server."&"MapPath(Path)","App"&".Path & ""\"" & Path")
	Content = Replace(Content,"Server."&"MapPath(""/"")","App."&"Path & ""\""")
	Content = Replace(Content,"Server."&"MapPath(""\"")","App."&"Path & ""\""")
	Content = Replace(Content,"Server."&"MapPath("".\"")","App."&"Path & "".\""")
	Content = Replace(Content,"Server."&"MapPath(""./"")","App."&"Path & "".\""")
	'20150614
	Content = Replace(Content,""" & Request."&"ServerVariables(""HTTP_HOST"") & ""","127.0.0.1")
	Content = Replace(Content,vbTab,"    ")			'替换Tab			 
	Content = Replace(Content,"request(""Debug"")"," 1=2 and ""11""")			'屏蔽管理			
	Content = Replace(Content,"Request(""gl"")"," 1=2 and ""11""")			'屏蔽管理 		
	Content = Replace(Content,"Request."&"QueryString(""Debug"")"," 1=2 and ""11""")			'屏蔽管理 
	ASPToVB = content
end function
'asp转VB.net内容
function ASPToVBNET(content)
	dim splstr,i,s,noSpaceS,s2,c,LCaseS,functionName
	dim leftS,noteS,centerS,tempCenterS
	splstr=split(content,vbcrlf)
	for each s in splstr
		LCaseS=phptrim(lcase(s)):noSpaceS=phptrim(s)
			's=centerS
		if instr(Lcase(s)," doerror ")>0 and instr(Lcase(s),"call ")=false then
			s=RegExp_Replace(s, " doError ", " call doError(") & ")"
			'call echo("s",s)
		end if
		
		if left(LCaseS,4)="sub " then
			s=functionParamtoNet(s)						'处理函数里参数
		elseif left(LCaseS,9)="function " then
			s=functionParamtoNet(s)						'处理函数里参数
			s2=mid(s,instr(Lcase(s),"function ") + 9)
			functionName=phptrim( mid(s2,1,instr(s2,"(")-1))
			s = s & vbcrlf & functionName & "="""""
		elseif left(LCaseS,4)="set " then
			s=mid(s, instr(lcase(s),"set ")+3)
		elseif left(LCaseS,4)="dim " then
			s=dimAddDefaultValue(s)
			
		elseif left(LCaseS,1)="." then
			leftS=mid(s,1,instr(s,".")-1)
			noteS=""
			if instr(s,"'")>0 then
				noteS=mid(s,instr(s,"'"))
			end if
			centerS=mid(s,instr(s,"."))
			if instr(s,"'")>0 then
				centerS=mid(centerS,1,instr(centerS,"'")-1)
			end if
			tempCenterS=centerS
			centerS=leftS & "call " & mid(centerS,1,instr(centerS," ")) & "(" & mid(centerS,instr(centerS," ")) & ")" & noteS
			if instr(tempCenterS,"=")=false and instr(tempCenterS,"(")=false  then
				s=centerS
				'call echo(s,centerS)
			end if
		
		
		elseif left(LCaseS,2)="if" then
			leftS=mid(s,1,instr(lcase(s),"if")-1)
			noteS=""
			if instr(s,"'")>0 then
				noteS=mid(s,instr(s,"'"))
			end if
			centerS=mid(s,instr(lcase(s),"if"))
			if instr(s,"'")>0 then
				centerS=mid(centerS,1,instr(centerS,"'")-1)
			end if
			if right(RTrimVBTab(lcase(centerS)),5)<>" then" then
				centerS=centerS & " :end if  "
			end if
			centerS = leftS & centerS & noteS
			
		elseif left(LCaseS,4)="wend" then
			s= "end While"
			
		end if 
		
		c=c & s & vbcrlf
	next
	ASPToVBNET=c
end function
'定义变量加默认值，指定
function dimAddDefaultValue(dimList)  
	dim splstr,s,c,startStr,dimValue,tempDimList,noteS
	tempDimList=dimList
	startStr=mid(dimList,1,instr(Lcase(dimList),"dim ")+3)
	dimList=mid(dimList,len(startStr)) 
	splstr=split(dimList,",")
	for each s in splstr			
		noteS=""
		if instr(s,":")>0 then
			noteS=mid(s,instr(s,":"))
			s=mid(s,1,instr(s,":")-1)
			'call echo("s",s)
		elseif instr(s,"'")>0 then
			noteS=mid(s,instr(s,"'"))
			s=mid(s,1,instr(s,"'")-1)
			'call echo("s",s)
		end if
		
		if instr(lcase("|c|content|html|htmlstr|strhtml|cssstr|strcss|StyleName|parentUrl|parentParentUrl|parentParentParentUrl|szRet|Url|UrlFileName|urlList|UrlParam|WebSite|FanCharSet|s1|s2|s3|c1|c2|c3|c4|c5|En1|En2|En3|En4|s|temp|temp1|temp2|temp3|yinghaoLabel|parentName|valueStr|valueStr1|valueStr2|valueStr3|noDanYinShuangYinStr|PubAHrefList|PubATitleList|yinghaoLabel|parentName|valueStr|nLen|ParamList|ParamNameList|ParamList|Str_unicode|valueLabel|domainName|labelName|"), "|" & phptrim(lcase(s)) & "|")>0 or right(phptrim(lcase(s)),3)="str" then
			dimValue= dimValue & s & "=""""" & vbCrlf
		end if
		if c<>"" then c=c & ","
		c=c & s & noteS
	next
	dimAddDefaultValue=tempDimList & vbcrlf & dimValue
end function
'函数参数转Net里能用(20151021)
function functionParamtoNet(content)
	dim leftStr,rightStr,valueStr,splstr,s,c,LCaseS
	leftStr=mid(content,1,instr(content,"("))
	rightStr=mid(content,instr(content,")"))
	valueStr=mid(content,instr(content,"(")+1)
	valueStr=mid(valueStr,1,instr(valueStr,")")-1)
	
	splstr=split(valueStr,",")
	for each s in splstr
		s=phptrim(s)
		LCaseS=lcase(s)
		if instr(LCaseS,"byval")=false and instr(LCaseS,"byref")=false then
			s=" byref " & s
		end if
		if c<>"" then c=c & ","
		c=c & s
	next
	'call echo("leftStr",leftStr)
	'call echo("rightStr",rightStr)
	'call echo("valueStr",valueStr)
	functionParamtoNet = leftStr & c & rightStr
end function

'call echo("单词",GetBeforeNStr("aa bb cc", "单词", 2)) 
'call echo("输入",GetBeforeNStr("""a a"" ""b b"" ""c  c""", "输入", 2))
'call echo("符号",GetBeforeNStr(" + - * / \ ", "符号", 2)) 
'call echo("数字",GetBeforeNStr(" 1 222.22 3 44", "数字", 2)) 
'call echo("数字",GetBeforeNStr(" aa - 333 ""a b c""", "全部", 2))
'获得下N个单词或其它等 (20050810)
function GetBeforeNStr(content,sType,nNumb)
    dim splStr, sx, s, c, i, sYHCount, s1, s2, tempZc, nI 
    dim wc                                                        '输入文本存储内容
	dim ArrayWC(999)											  '输入文本数组内容
	dim ArrayWCCount											  '输入文本数组内容总数
    dim zc                                                        '单词存储内容
	dim ArrayZC(999)											  '单词存储内容 数组
	dim ArrayZCCount											  '单词存储内容 数组总数
    dim nc                                                        '为数字累加
	dim ArrayNC(999)											  '为数字累加 数组
	dim ArrayNCCount											  '为数字累加 数组总数
    dim fc                                                        '符号
	dim ArrayFC(999)											  '符号 数组
	dim ArrayFCCount											  '符号 数组总数
	
	dim ArrayCountStr(999)											  '数组总内容
	dim nArrayCount											  	  '数组总数

	
    dim yesOK                                                     '是否OK
    dim tempS 
    dim upWord                                                    '上一个变量
    dim upWord2                                                   '上上一个变量
    dim endCode                                                   '当前行字符往后面代码

    dim beforeStr                                                 '前一个字符
    dim afterStr                                                  '后一个字符
    dim rowC                                                      '每行代码
    dim tempRow                                                   '每行完整代码
    dim downRow                                                   '下一行代码
     
    dim noteCode 
	
    sYHCount = 0                                                  '双引号默认为0
	
	
    ArrayWCCount=0											  '输入文本数组内容总数
    ArrayZCCount=0											  '单词存储内容 数组总数
    ArrayNCCount=0											  '为数字累加 数组总数
    ArrayFCCount=0											  '符号 数组总数	
	nArrayCount=0											  '数组总数
	
	
	

    splStr = split(content, vbCrLf)                               '分割行
    '循环分行
    for nI = 0 to uBound(splStr)
        s = splStr(nI) 
        tempS = s 
        rowC = "" : tempRow = ""         '清空每行ASP代码和暂存完整行代码
        noteCode = ""                    '默认注释代码为空
        downRow = ""                     '下一行代码
        if(nI + 1) <= uBound(splStr) then
            downRow = splStr(nI + 1)  
        end if 
        nc = ""                          '数字为空

        '循环每个字符
        for i = 1 to len(s)
            sx = mid(s, i, 1) 
            beforeStr = right(replace(mid(s, 1, i - 1), " ", ""), 1)                       '上一个字符
            afterStr = left(replace(mid(s, i + 1), " ", ""), 1)                            '下一个字符
            endCode = mid(s, i + 1)                                                        '当前字符往后面代码 一行

            '输入文本
            if(sx = """" or wc <> "") then
                '双引号累加
                if sx = """" then sYHCount = sYHCount + 1 
                '判断是否"在最后
                if sYHCount mod 2 = 0 then
                    if mid(s, i + 1, 1) <> """" then
                        wc = wc & sx 
                        s1 = right(replace(mid(s, 1, i - len(wc)), " ", ""), 1) '必需放在这里，要不会出错

                        '处理输入内容
                        wc = right(wc, len(wc) - 1) 
                        wc = left(wc, len(wc) - 1) 
                        wc = replace(wc, """""", "\""") 
                        wc = """" & wc & """" 
						
						
    ArrayWCCount=ArrayWCCount+1											  '输入文本数组内容总数						
	ArrayWC(ArrayWCCount)=wc											  '输入文本数组内容		
	nArrayCount=nArrayCount+1											  '数组总数		
	ArrayCountStr(nArrayCount)=wc										  '数组总内容

                        rowC = rowC & wc                                        '函数代码累加
                        sYHCount = 0 : wc = ""                                  '清除
                    else
                        wc = wc & sx 
                    end if 
                else
                    wc = wc & sx 
                end if 

            '+-*\=&   字符特殊处理
            elseIf inStr(".&=,+-*/:()><", sx) > 0 and nc = "" then
                if zc <> "" then
                    tempZc = zc 
                    upWord2 = upWord                '记录上上一个变量
                    upWord = lCase(tempZc)          '记录上一个变量名称   为小写
                    rowC = rowC & zc 
                    zc = ""                         '清空字母内容
                end if  
                upWord2 = upWord                                         '记录上上一个变量
                upWord = sx 
				
    ArrayFCCount=ArrayFCCount+1											  '符号 数组总数	
	ArrayFC(ArrayFCCount)=sx											  '符号 数组
	nArrayCount=nArrayCount+1											  '数组总数
	ArrayCountStr(nArrayCount)=sx										  '数组总内容 
			
            elseIf sx = "'" then
                noteCode = "//" & mid(s, i + 1) 
                exit for 
            '字母
            elseIf checkABC(sx) = true or sx = "_" or zc <> "" then
                zc = zc & sx 
                yesOK = true 
                s1 = mid(s & " ", i + 1, 1) 
                s2 = mid(zc, 1, 1) 
                if checkABC(s1) <> true and s1 <> "_" then
                    yesOK = false 
                end if 
                '允许变量后面是数字
                if checkNumber(s1) = true and checkABC(s2) = true then
                    yesOK = true 
                end if 
                if yesOK = false then
                    'Rem注释
                    if lCase(zc) = "rem" then
                        sx = mid(s, i + 1) : i = i + len(sx) 
                        zc = zc & sx 
                    end if 
 

    ArrayZCCount=ArrayZCCount+1											  '单词存储内容 数组总数
	ArrayZC(ArrayZCCount)=zc											  '单词存储内容 数组
	nArrayCount=nArrayCount+1											  '数组总数
	ArrayCountStr(nArrayCount)=zc										  '数组总内容
	
            '注释则退出
                    tempZc = zc 

                    upWord2 = upWord                '记录上上一个变量
                    upWord = lCase(tempZc)          '记录上一个变量名称   为小写
                    rowC = rowC & zc 
                    zc = ""                         '清空字母内容
                end if 
            '为数字
            elseIf checkNumber(sx) = true or nc <> "" then
                nc = nc & sx 
                if afterStr <> "." and checkNumber(afterStr) = false then
                    rowC = rowC & nc 
					

    ArrayNCCount=ArrayNCCount+1											  '为数字累加 数组总数
	ArrayNC(ArrayNCCount)=nc											  '为数字累加 数组
	nArrayCount=nArrayCount+1											  '数组总数
	ArrayCountStr(nArrayCount)=nc										  '数组总内容
					
					
                    nc = "" 
                end if 
            else
                rowC = rowC & sx 
            end if 
            tempRow = tempRow & sx 
        next 
        c = c & rowC 
    next 

	
	if sType="单词" then
		if ubound(ArrayCountStr)>=nNumb then
			GetBeforeNStr=ArrayZC(nNumb)
		end if	
	elseif sType="输入" then
		if ubound(ArrayCountStr)>=nNumb then
			GetBeforeNStr=ArrayWC(nNumb)
		end if	
	elseif sType="符号" then
		if ubound(ArrayCountStr)>=nNumb then
			GetBeforeNStr=ArrayFC(nNumb)
		end if	
	elseif sType="数字" then
		if ubound(ArrayCountStr)>=nNumb then
			GetBeforeNStr=ArrayNC(nNumb)
		end if	
	elseif sType="全部" then
		if ubound(ArrayCountStr)>=nNumb then
			GetBeforeNStr=ArrayCountStr(nNumb)	
		end if	
	else
		GetBeforeNStr=c
	end if
end function


'获得字符串里的第一次出现的Dim变量组   call rw(getStrDim("  dim a, b,        c,d   "  & vbcrlf &" ss    'sdfasdf"))
function getStrDim(byval content)
	content=replace(replace(content,vbcrlf,":"),vbtab," ")
	if instr(content,":")>0 then
		content=mid(content,1,instr(content,":")-1)
	end if
	if instr(content,"'")>0 then
		content=mid(content,1,instr(content,"'")-1)
	end if
	content=phptrim(content)
	if left(content,3)="dim" then
		content=mid(content,4)
	end if
	content=replace(content," ","")
	dim splstr,s,c
	if instr(content,"(")>0 then
		splstr=split(content,",")
		for each s in splstr
			if s <>"" then
				if c<>"" then c=c & ","
				if instr(s,"(")>0 then
					s=mid(s,1,instr(s,"(")-1)
				end if
				c=c & s
			end if
		next
		content=c
	end if 
	getStrDim=content
end function

'获得函数字符串里的变量组 call rw(getFunctionStrDim("aaa(ByVal Content,a,b,c,d)"))
function getFunctionStrDim(byval content)
	dim splstr,s,c,lCaseS
	if instr(content,")")>0 then
		content=mid(content,1,instr(content,")")-1)
	end if
	if instr(content,"(")>0 then
		content=mid(content,instr(content,"(")+1)
	end if
	content=replace(content,","," ")
	splstr=split(content," ")
	for each s in splstr
		lCaseS=lcase(s)
		if lCaseS<>"" and lCaseS<> "byval" and lCaseS<>"byref" then
			if c<>"" then c=c & ","
			c=c & s
		end if
	next
	getFunctionStrDim=c
end function

%>