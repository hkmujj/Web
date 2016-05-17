<% 
'Dictionary 字典 (2013,9,27)
 
'开始 
Dim AspD, AspHelpTitle, AspHelpContent 
Set AspD = CreateObject("Scripting.Dictionary")
AspD.Add "Abs", "返回数字的绝对值" 
AspD.Add "Sqr", "返回数值表达式的平方根" 
AspD.Add "Sgn", "返回表示数字符号的整数" 
AspD.Add "Rnd", "返回一个随机生成的数字" 
AspD.Add "Log", "返回指定数值的自然对数" 
AspD.Add "Exp", "返回e的指定次幂" 
AspD.Add "Sin", "返回指定弧度的正弦值" 
AspD.Add "Cos", "返回指定弧度的余弦值" 
AspD.Add "Tan", "返回指定弧度的正切值" 
AspD.Add "Atn", "返回指定数值的反正切值" 
AspD.Add "Int", "返回指定数字的整数部分" 
AspD.Add "Round", "对参数按指定的位数进行四舍五入" 
AspD.Add "Fix", "返回指定数值的整数部分" 
AspD.Add "Now", "返回当前的日期时间" 
AspD.Add "Date", "返回当前的系统日期" 
AspD.Add "Time", "返回当前的系统时间" 
AspD.Add "Year", "从指定的表达式中提取年份" 
AspD.Add "Month", "从指定的表达式中提取月份" 
AspD.Add "Day", "从指定表达式中提取日期" 
AspD.Add "Hour", "从指定的表达式中提取整点值" 
AspD.Add "Minute", "从指定的表达式中提取分钟数" 
AspD.Add "Second", "从指定的表达式中提取秒数" 
AspD.Add "DateAdd", "将指定日期添加上指定时间间隔并返回新值" 
AspD.Add "DateDiff", "计算两个日期之间的时间间隔" 
AspD.Add "DatePart", "返回给定日期中的某一部分" 
AspD.Add "DateSerial", "根据提供的年月日的数值返回一个Date子类型的数据" 
AspD.Add "DateValue", "将指定的字符串转化为Date子类型的数据" 
AspD.Add "MonthName", "根据指定的数字返回代表该月份的字符串" 
AspD.Add "Weekday", "返回代表一周中某一天的整数" 
AspD.Add "WeekdayName", "根据周次返回代表该周次的字符串" 
AspD.Add "TimeSerial", "根据提供的时分秒的数值返回一个Date子类型的数据" 
AspD.Add "TimeValue", "将指定的字符串转化为Date子类型的数据" 
AspD.Add "CBool", "将表达式转化为Boolean子类型" 
AspD.Add "CByte", "将表达式转化为Byte子类型" 
AspD.Add "CDbl", "将表达式转化为Double子类型" 
AspD.Add "Chr", "将ANSI字符代码（～）转化为对应的字符" 
AspD.Add "CInt", "将表达式转化为Integer数值子类型" 
AspD.Add "CLng", "将表达式转化为Long数值子类型" 
AspD.Add "Cstr", "将表达式转化为String子类型" 
AspD.Add "CCur", "将表达式转化为Currency子类型" 
AspD.Add "CDate", "将日期表达式转化为Date子类型" 
AspD.Add "CSng", "将表达式转化为Single数值子类型" 
AspD.Add "Asc", "返回与字符串首字母相对应的ANSI字符代码（～）" 
AspD.Add "Oct", "返回八进制的字符串" 
AspD.Add "Hex", "返回十六进制的字符串" 
AspD.Add "FormatNumber", "格式化指定的数值表达式" 
AspD.Add "FormatDateTime", "格式化指定的日期时间表达式" 
AspD.Add "FormatCurrency", "格式化指定的货币表达式" 
AspD.Add "FormatPercent", "格式化某一数值为百分数" 
AspD.Add "IsArray", "判断某一变量是否为数组变量" 
AspD.Add "IsDate", "判断某一变量是否为或者可以转化为日期表达式" 
AspD.Add "IsEmpty", "判断变量是否为空" 
AspD.Add "IsNull", "判断表达式中是否包含有效数据" 
AspD.Add "IsNumeric", "判断表达式的值是否为数字" 
AspD.Add "Is_object_", "判断表达式的值是否为某一对象" 
AspD.Add "VarType", "返回指定变量的子类型" 
AspD.Add "TypeName", "返回指定变量子类型的信息" 
AspD.Add "MsgBox", "弹出含有提示信息的对话框" 
AspD.Add "InputBox", "弹出可输入数据的对话框" 
AspD.Add "ScriptEngine", "返回一个代表当前所使用的脚本语言的字符串" 
AspD.Add "ScriptEngineBuildVersion", "返回所使用的脚本引擎的编译版本号" 
AspD.Add "ScriptEngineMajorVersion", "返回所使用的脚本引擎的主版本号" 
AspD.Add "ScriptEngineMinorVersion", "返回所使用脚本引擎的次版本号" 
AspD.Add "Array", "利用所给的参数返回一个数组" 
AspD.Add "UBound", "返回数组的最大可用下标" 
AspD.Add "LBound", "返回数组的最小可用下标" 
AspD.Add "Filter", "返回一个数组其中包含了满足条件的字符串" 
AspD.Add "InStr", "返回某字符串在指定字符串中第一次出现的位置" 
AspD.Add "InStrRev", "返回某字符串在指定字符串中从结尾算起的出现位置" 
AspD.Add "Join", "将数组中的字符串连接为一个字符串" 
AspD.Add "LCase", "将字符串中的字母全部转换为小写形式" 
AspD.Add "Left", "从字符串的左端取指定数目的字符" 
AspD.Add "Len", "返回字符串的长度或者存储某一变量所需要的字节数" 
AspD.Add "Ltrim", "返回不带前导空格的字符串" 
AspD.Add "Replace", "将字符串中指定字符串替换为其他内容" 
AspD.Add "Right", "从字符串的右侧获取指定数目的字符" 
AspD.Add "Rtrim", "删除字符串的后续空格" 
AspD.Add "Mid", "从字符串中指定位置获取指定长度的字符串" 
AspD.Add "Split", "返回包含指定数目子字符串的一维数组" 
AspD.Add "StrComp", "对两个指定的字符串进行比较" 
AspD.Add "StrReverse", "将指定字符串倒序排列并返回重新排列后的字符串" 
AspD.Add "String", "返回由指定数目的某一字符组成的字符串" 
AspD.Add "Trim", "删除字符串的前导和后续空格" 
AspD.Add "Space", "返回由指定数目的空格组成的字符串" 
AspD.Add "UCase", "将字符串中的字母全部转换为大写形式" 
AspD.Add "LoadPicture", "返回一个图片对象" 
AspD.Add "Create_object_", "创建并返回一个对象" 
AspD.Add "Get_object_", "从指定文件中获得一个Automation对象" 
AspD.Add "RGB", "返回代表某一颜色的整数" 
AspHelpTitle = AspD.keys 
AspHelpContent = AspD.Items 
'获得ASP内置函数帮助说明
Function ASPHelpTip(ByVal FunName, ByVal Content)
	Dim I, S, C 
	For I = 0 To AspD.Count - 1
		If LCase(AspHelpTitle(I)) = LCase(FunName) Then
			If LCase(FunName) = "chr" Then
				S = Left(Content, InStr(Content, ")") - 1) 
				S = Replace(S, "(", "") 
				C = vbCrlf & "Chr(" & S & ")=" 
				If CheckNumber(S) = True Then    '判断Chr 里面是否为数字
					If S = "10" Or S = "13" Then
						C = C & "(回车键)" 
					ElseIf S = "32" Then
						C = C & "(空格)" 
					ElseIf S = "0" Then
						C = C & "(中断程序,停止执行)" 
					Else
						C = C & "(" & Chr(S) & ")" 
					End If 
					C = Replace(C, "'", "&#x27;") 
				Else
					C = C & S 
				End If 
			End If 
			ASPHelpTip = AspHelpContent(I) & C 
			Exit For 
		End If 
	Next 
End Function 
'处理用户输入的值
Function HandleWrite(Jc)
	Dim S 
	S = Left(Jc, Len(Jc) - 1) 
	S = Right(S, Len(S) - 1) 
	S = LCase(Trim(S)) 
	If S = "adodb.connection" Then
		Jc = """Adodb.Connection""" 
	ElseIf S = "adodb.recordset" Then
		Jc = """Adodb.RecordSet""" 
	End If 
	HandleWrite = Jc 
End Function 
'结束     
'简体与繁体转换 用法 JianTiTabFanTi("罢摆败颁办绊帮绑镑谤",2)2为繁体
Function JianTiTabFanTi(Path, Content, SType)
	Dim JianStr, FanStr, JianS, FanS, I, S, C 
	C = GetFText(Path) & vbCrlf 
	SType = Cstr(SType)    '字符类型 
	JianStr = StrCut(C, "简体" & vbCrlf, vbCrlf, 2) 
	FanStr = StrCut(C, "繁体" & vbCrlf, vbCrlf, 2)  
	If 1 = 2 Then
		If Len(Content) * 2 < Len(JianStr) Then
			For I = 1 To Len(Content)
				S = Mid(Content, I, 1) 
				'繁体转简体
				If SType = "2" And InStr(JianStr, S) > 0 Then
					JianS = Mid(JianStr, InStr(JianStr, S), 1) 
					FanS = Mid(FanStr, InStr(JianStr, S), 1) 
					Content = Replace(Content, JianS, FanS) 
					Call Echo("", "") 
				ElseIf InStr(FanStr, S) > 0 Then
					JianS = Mid(JianStr, InStr(FanStr, S), 1) 
					FanS = Mid(FanStr, InStr(FanStr, S), 1) 
					Content = Replace(Content, FanS, JianS) 
				End If 
			Next 
			JianTiTabFanTi = Content 
			Exit Function 
		End If 
	End If 
	For I = 1 To Len(JianStr)
		JianS = Mid(JianStr, I, 1) 
		FanS = Mid(FanStr, I, 1) 
		'繁体转简体
		If SType = "2" Then
			Content = Replace(Content, JianS, FanS) 
		Else
			Content = Replace(Content, FanS, JianS) 
		End If 
	Next 
	JianTiTabFanTi = Content 
End Function 
'近意词替换
Function JiYiCiReplace(Path, Content)
	Dim SplStr, Splxx, I, S, C 
	SplStr = Split(GetFText(Path), vbCrlf) 
	For Each S In SplStr
		If InStr(S, ",") And Left(Trim(S), 1) <> "#" Then
			Splxx = Split(S, ",") 
			If UBound(Splxx) >= 1 Then
				Content = Replace(Content, Splxx(0), Splxx(1)) 
			End If 
		End If 
	Next 
	JiYiCiReplace = Content 
End Function
'替换字符
Function CiReplace(THContent, Content)
    Dim SplStr, Splxx, I, S, C
    SplStr = Split(THContent, vbCrLf)
    For Each S In SplStr
        If InStr(S, ",") And Left(Trim(S), 1) <> "#" Then
            Splxx = Split(S, ",")
            If UBound(Splxx) >= 1 Then
                Content = Replace(Content, Splxx(0), Splxx(1))
            End If
        End If
    Next
    CiReplace = Content
End Function

%> 

