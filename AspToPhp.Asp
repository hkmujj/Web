<!--#Include virtual = "/Inc/Config.Asp"-->
<% 
'在ASP里 reDim是重新定义数组的。嗯，明白
'atemp.php 打开加载配置在 PHP2/ImageWaterMark/PHPAuxiliary.php  修改

Dim requestact                               '返回act值，在VB里面没用，是在web里用的
dim tempPHPContent							'暂存php内容
Dim c, msgStr, sType 
requestact = Request("act") 
sType = 1 
If sType = 2 Or Request("act") = "inc" Then
    c = createIncDirList() 
    c = c & getAspToPHPContent("atemp.asp", msgStr) 
    Call writeToFile("/PHP2/Web/" & "index.php", c,"gb2312") 
ElseIf Request("act") = "web1asp" Then
    Call getAspToPHPContent("web/1.asp", msgStr) 
    Call rw(msgStr) 
ElseIf Request("act") = "webfunction" Then
    Call getAspToPHPContent("web/function.asp", msgStr) 
    Call rw(msgStr) 
ElseIf Request("act") = "webfunction2" Then
    Call getAspToPHPContent("web/function2.asp", msgStr) 
    Call rw(msgStr) 
	
ElseIf Request("act") = "webtemplate" Then
    Call getAspToPHPContent("web/template.asp", msgStr) 
    Call rw(msgStr) 
ElseIf Request("act") = "setAccess" Then
    Call getAspToPHPContent("web/setAccess.asp", msgStr) 
    Call rw(msgStr)
ElseIf Request("act") = "aspweb" Then
    Call getAspToPHPContent("aspweb.asp", msgStr) 
    Call rw(msgStr)
ElseIf Request("act") = "config" Then 
    Call getAspToPHPContent("inc/_config.asp", msgStr) 
    Call rw(msgStr)
	
	
elseif request("act")="handlecontent" then
	call writeToFile("atemp.asp",request("content"),"gb2312")
	tempPHPContent=getAspToPHPContent("atemp.asp", msgStr)	
	tempPHPContent= "<pre>" & replace(tempPHPContent,"<","&lt;")	
	if getip()<>"127.0.0.1" then
		call deletefile("./atemp.asp")
		call deletefile("./atemp.php")
	end if
	'收录20160126
	dim folderPath
	folderPath="/AspToPhpCode/" & getip() & "/"
	call createDirFolder(folderPath)
	call writeToFile(folderPath & "/" & format_Time(now(),11) & ".txt",request("content"),"gb2312")	
	
	
	
	if getIP()<>"127.0.0.1" then
		tempPHPContent=tempPHPContent& "<script src=""http://sharembweb.com/Jquery/jquery.Min.js""></script><script src=""http://sharembweb.com/Jquery/qq.js""></script>"
	end if
	
	call rwend(tempPHPContent)
	
Else
    Call getAspToPHPContent("atemp.asp", msgStr)
    Call rw(msgStr)
End If 
call deletefile("./1.txt")
Call rw(VBEchoTimer()) 

'获得ASP转PHP内容
Function getAspToPHPContent(filePath, msgStr)
    '(20150805不完善架构不好)
    Dim content, c 
    Dim obj : Set obj = new aspToPhpClass
		'开启导入函数功能20160115
		if request("isOpenImportFunction")<>"" then
			obj.setIsOpenImportFunction=true
		end if
		
        '设置为函数列表
        c = "|CLng|ubound|LBound|Chr|replace|Write|Cookies|QueryString|TypeName|Redirect|FormatNumber|Cstr|Asc|Join|Len|PHPEcho|ASPEcho|LCase|UCase|IsNumeric|IsNull|Int|Execute|ServerVariables|CreateObject|Hex|AscW|time|IsObject|isArray|"
        'php里用到的
        c = c & "|header|echo|" 
        c = c & "|case|Timer|DateDiff|Fix|now|getYMDHMS|IsDate|Year|Month|Day|Hour|Minute|Second|array|AspTrim|WebPageControl|HandleAction|XY_getLableValue|ReplaceValueParam|LTrim|RTrim|newGetArray|AspLTrim|AspRTrim|DateAdd|ASPDate|" 
        c = c & "handleMakeHtmlFileEditTime|XY_admin|XY_test|ManyTimesHandleTemplateAction|XY_executeSQL|XY_NavDidExists|NewRParam|newReplaceValueParam|XY_PHP_GeneralInfoList|XY_PHP_GeneralList||XY_PHP_GetFieldValue|XY_PHP_SinglePage|aspMD5|session_set_cookie_params|getMsg1|myMD5|inputText|inputText2|ADSql|showSelectcolumn|setCookie|columnList|columnList2|showColumnList|getEditorStr|handleInputHiddenTextArea|displayUploadDialog|ADSqlRf|handleInputCheckBox|inputCheckBox3|arrayToString|resetAccessData|XY_PHP_NavList|XY_PHP_DetailList|getColumnName|batchDeleteTempStr|createFileUTF|handleRemoteJsCssParam|handleIsWebSite|handleGetContentUrlList|replaceGlobleVariable|handleHttpUrlArray|MainNav|dispalyManage|addEditDisplay|saveAddEdit|del|sortHandle|displayDefaultLayout|loadWebConfig|adminIndex|handleRGV|makeAllHtml|getRsUrl|handleUpDownArticle|upArticle|downArticle|getColumnUrl|handleReplaceValueParam|" 
		c=c & "|displayLayout|saveRobots|saveSiteMap|isOpenTemplate|updateWebsiteStat|contentTranscoding|stat2016|displayTemplatesList|XY_DisplayWrap|XY_GetColumnUrl|XY_GetOnePageUrl|copyHtmlToWeb|makeHtmlWebToZip|getRandArticleId|handleWebUrl|handleWebUrl|getOnePageUrl|XY_PHP_GetColumnContent|XY_PHP_CommentList|displayEditor|specialStrReplace|unSpecialStrReplace|"
		c=c & "|getBrType|templateFileList|delTemplateFile|addEditFile|folderSearch|getDirJsList|getDirCssList|getDateilTemplate|websiteDetail|getFieldConfigList|handlePower|handlePower|checkPower|getFormFieldName|handlePower|getHandleFieldList|getHandleDate|strtolower|strtoupper|runScanWebUrl|PHPStrLen|Right|reaFile|scanCheckDomain|callFunction2|"
		c = c & "|ord|intval|is_array|gettype|"
		
        Dim addC 
        addC = GetFText("\Config\网站函数总杂205.txt")  & GetScanFunctionNameList(getftext(filePath))
        addC = Replace(Replace(addC, ",", "|"), vbCrLf, "|") 
		
		
        obj.setFunctionNameList = obj.getWordList() & c & addC
        '设置不显示函数名称列表 删除函数
        obj.setNotShowFunctionList = "PHPRand|PHPRnd|CheckErr|IIF|PageControl|DelHtml|GetFText|GetFileText|HandlePath|GetArrayList|GetHttpPage|GetURLContent|GetHttpUrl|CheckFunValue|PHPTrim|HttpHost|GetHost|PHPPrint|getRsPageNumber|ADSqlRf|Rr|getArray|is_numeric|regReplace|easp_Replace|getHandleDate|PHPLTrim|PHPRTrim|getHttpUrlState|checkDomainName|phpDirName|getHttpUrlServerName|" 
        '设置留空函数名称列表 清空函数内容
        c = "IsMail|MBInfo|PrintFormInfo|BytesToBstr|Agent|isNul|AsaiLinkAdd|AsaiLinkDel|Iswww|XY_RelatedInformationList|HandleDispalyArticleList|CheckAccessPass|" 
        c = c & "|CheckSql|CreateMdb|EditTable|GetFieldAlt|GetFieldList|GetDifferentTableFieldList|CompactDB|HandleSqlServer|GetDataTableList|CheckTable|CheckField|GetTableList_Conn|EditAccessPassWord|checkDomainName|" 
        c = c & "|ExportToExcel|ImportToDatabase|getFieldValue|getConnFieldValue|setFieldValue|SetTableField|FileBinaryAddAccess|ShowAccessStream|DeleteAction|AddDateViewWebEffect|" 
        c = c & "|CopyDateToWebDate|CheckEval|PHPExec|handleGetHttpPage|HandleContentCode|XY_PHP_CustomNavList_1||" 
        obj.setSpaceFunctionList = c 
		
		'设置默认类名称列表
		obj.setClassNameList="myclass"
		'设置全局变量列表
		c="|@$_REQUEST|@$_POST|@$_GET|@$_COOKIE|@$_SESSION|"
		c=c & "|WEBCOLUMNTYPE|WEB_VIEWURL|WEB_ADMINURL|EDITORTYPE|ROOT_PATH|"
		obj.setGlobalDimList=c
		 
		'设置转换成语言为js20160321
		if request("language")="js" then
			obj.setToLanguage="js"
		end if
		
        '设置显示HTML部分
        obj.setIsDisplayHtml = True 

        content = getFText(filePath) 
        content = deleteAspTempCode(content)                                            '删除里暂存代码

        content = obj.aSPToPHP(content) 
        '测试用到
        msgStr = obj.getTempActionList() 
    Set obj = Nothing 
    content = phpTrim(content)
    getAspToPHPContent = content 
End Function
 
'删除里暂存代码
Function deleteAspTempCode(content)
    Dim startStr, endStr, i, s 
    startStr = "'【删除此行start】" : endStr = "'【删除此行end】" 
    For i = 1 To 9
        If InStr(content, startStr) = False Then
            Exit For 
        End If 
        s = getStrCut(content, startStr, endStr, 1) 
        content = Replace(content, s, "") 
    Next 
    deleteAspTempCode = content 
End Function
 

'创建Inc目录里asp文件列表
Function createIncDirList()
    Dim splStr, fileList, fileName, filePath, content, pHPDir, pHP_IncDir, phpFileName, phpFilePath, includeStr,nCount
    pHPDir = "/PHP2/Web/" 
    'call deleteFolder(pHPDir)
    Call createFolder(pHPDir) 
    pHP_IncDir = pHPDir & "Inc/" 
    Call createFolder(pHP_IncDir)
    fileList = "|Common.Asp|StringNumber.Asp|Print.Asp|Time.Asp|2014_Array.asp|Cai.Asp|Check.asp|Incpage.Asp|2015_Param.Asp|2014_Author.Asp|2014_Css.Asp|2015_Editor.Asp|" 
    fileList = fileList & "2014_Js.Asp|2015_Color.Asp|2015_Formatting.Asp|URL.Asp|2015_APGeneral.Asp|2015_ToMyPHP.Asp|FunHTML.Asp|2015_NewWebFunction.Asp|ASPPHPAccess.asp|EncDec.Asp|IE.asp|2016_WebControl.asp|2016_SaveData.Asp|Html.asp|2016_Log.Asp|2014_Template.Asp|SystemInfo.Asp|PinYin.Asp|2014_Form.Asp|" 
    'fileList = fileList & "|2014_Action.Asp|2014_MainInfo.Asp||2014_News.Asp|2014_Search.Asp|2014_SiteMap.Asp|2014_Template.Asp|Access.Asp|AjAx.Asp|AutoAdd.Asp|Create_Html.Asp|2014_Class.Asp|2015_ToPhpCms.Asp||"
    splStr = Split(fileList, "|")  : nCount=0
    For Each fileName In splStr
        If fileName <> "" And InStr(fileName, "#") = False Then
            filePath = "/Inc/" & fileName 
            phpFileName = handleFilePathArray(fileName)(3) & ".php" 
            phpFilePath = pHP_IncDir & phpFileName 
            If checkFile(phpFilePath) = False Then
                content = getAspToPHPContent(filePath, "") 
                If LCase(fileName) = LCase("Create_Html.Asp") Then
                    content = "<?PHP require_once './Config.php';?>" & content 
                End If
                Call writeToFile(phpFilePath, content,"gb2312") 
				nCount=nCount+1
                Call echo(nCount & "、" & fileName, filePath) 
            Else
                Call echo("已经存在", phpFilePath) 
            End If 
            doevents 
            includeStr = includeStr & "require_once './Inc/" & phpFileName & "';" & vbCrLf 
        End If 
    Next 
    '不为空
    If includeStr <> "" Then
        Dim addToIncludeStr 
        addToIncludeStr = addToIncludeStr & "require_once './../ImageWaterMark/Include/ASP.php';" & vbCrLf 
        addToIncludeStr = addToIncludeStr & "require_once './../ImageWaterMark/Include/Conn.php';" & vbCrLf 
        addToIncludeStr = addToIncludeStr & "require_once './../ImageWaterMark/Include/MySqlClass.php';" & vbCrLf 
        includeStr = addToIncludeStr & includeStr 
        includeStr = "<?PHP" & vbCrLf & includeStr & vbCrLf & "?>" & vbCrLf 
    End If 
    createIncDirList = includeStr 


	
End Function
 

'nDimAraayLeftKuoHao(nThisIndex) 改成 nDimAraayLeftKuoHao(5)     要不然会不出，这个行，待研究

'ASP转PHP 201504
Class aspToPhpClass
	Dim toLanguage																	'转换成什么语言 
	dim toLanguageLinkCha															'转换语言变量与变量之间链接字符
	dim toLanguageVarNameCha															'转换语言变量前字符
	
    Dim ifJudge(99)                                                                 'IF判断
    Dim ifJudgeRow(99)                                                              'IF判断 是否为一行
    Dim ifJudgeRowStr(99)                                                           'IF判断 为一行内容
    Dim ifRsObjHandStr                                                              'IF判断Rs句柄内容
    Dim forJudge                                                                    'For判断
    Dim trueJudge                                                                   '真假判断
    Dim forFirstVarName                                                             'For第一个变量名
    Dim forEachName                                                                 'For的Each名称
    Dim isForStep                                                                   '是否为For循环步数
    Dim forStep                                                                     'For的Step步骤名称
    Dim functionName                                                                '记录函数名称
    Dim isNotShowFunctionName                                                       '记录不显示函数名称
    Dim isFunctionDimName                                                           '是否为函数定义名称
    Dim isSpaceFunctionName                                                         '记录留空函数名称
    Dim isSelectCase                                                                '是Select Case判断
    Dim isSelectMoreCheck                                                           '为select判断的多个判断如 case 1,2,""  这个在PHP不能用，因为php只能用一个
    Dim isCase                                                                      'Case是否开启 只让第一次显示后台:为:
    Dim nCaseCount                                                                  'Case总数累加

    Dim isDimAraay(99)                                                              '是否为定义数组
    Dim nDimAraayLeftKuoHao(99)                                                     '定义数组左括号数
    Dim nDimAraayRightKuoHao(99)                                                    '定义数组右括号数
    Dim dimAraayList                                                                '定义数组列表
    Dim dimAraayStart                                                               '定义数组列表开始  与结束 判断
    Dim nDimArrayQianTao                                                            '定义数据嵌套数
    Dim functionNameList                                                            '函数名称列表
    Dim notShowFunctionList                                                         '不显示函数名称列表
    Dim spaceFunctionList                                                           '留空函数名称列表
	dim globalDimList																'全局变量列表
    Dim noteCode                                                                    '注释代码
    Dim isToSetCookies                                                              '为设置Cookies
    Dim nMidDuanLuo                                                                 'mid分段数  为了处理最后无参数的情况
    Dim isWhile                                                                     'while是否为真
    Dim ismultidimensionalArray                                                     '是多维数组
    Dim isOpenTable                                                                 '打开数据库表
    Dim isLoop                                                                      '为循环字符


    Dim nVariableSetValue(99)                                                       '变量设置值累加数


    Dim upWord                                                                      '上一个变量
    Dim upWord2,upWord3,upWord4,upWord5,upWord6                                                                     '上上一个变量
    Dim upWordYuan, upWordYuan2,upWordYuan3,upWordYuan4,upWordYuan5,upWordYuan6                                                     '上一个源变量与上上一个源变量

    Dim endCode                                                                     '当前行字符往后面代码

    Dim beforeStr                                                                   '前一个字符
    Dim afterStr                                                                    '后一个字符
    Dim rowC                                                                        '每行代码
    Dim tempRow                                                                     '每行完整代码
    Dim nc                                                                          '为数字累加
    Dim arrayVar(20)                                                                '数据变量
    Dim arrayVarLeftKuoHao(99)                                                      '数据变量 左括号数
    Dim arrayVarRightKuoHao(99)                                                     '数据变量 右括号数

    Dim tempActionList                                                              '暂存动作列表
    Dim nShowFunction                                                               '函数显示数
    Dim functionReturn                                                              '函数返回，给return 后台=用
    Dim isFunctionReturn                                                            '判断函数是否有返回
    Dim systemFunctonList 
    Dim downRow                                                                     '下一行代码
    Dim downWord                                                                    '下一行单词
    Dim isDisplayHtml                                                               '是否显示HTML代码
    Dim tempWord                                                                    '暂存单词

    Dim functionInDimList                                                           '函数里定义的Dim变量
    Dim isSearverAddKuoHu                                                           'Server里有括弧
    Dim isHeader                                                                    '是header函数
    Dim dimRsStr                                                                    '定义RS内容 给 If Rs.Eof Then Exit For   用
    Dim isOpenRsStr                                                                 '是打开数据记录
    Dim onDimRsStr                                                                  '开户定义Rs字符
    Dim nRsStrCount                                                                 'Rs定义总数
    Dim nThisIndex                                                                  '当前索引
	dim importFunctionList															'需要导入系统函数列表（20160115
	dim isOpenImportFunction														'是否打开导入函数
	dim isClass,className,classNameList,dimClassNameList												'是否为类，类名称，类名称列表
	dim rowEndStr							'行最后字符
	dim isZD,ZDName,ZDNameList														'是否字典，字典名称，字典名称列表
	dim pubRowS				'一行内容

    '初始化
    Private Sub class_Initialize()
        systemFunctonList = "|class|if|for|while|select|iif|function|sub|dim|response|call|left|right|rnd|mid|cint|instr|instrrev|vbtab|vbcrlf|split|mod|server|" 
        systemFunctonList = LCase(systemFunctonList) 
        isDisplayHtml = False                                                           '设置默认是否显示HTML内容
		isOpenImportFunction=false														'设置默认是否打开导入函数
		isClass=false:className="":classNameList="|":dimClassNameList="|"										'类初始化
		isZD=false:ZDName="":ZDNameList="|"											'字典初始化
    End Sub
	
 
    '获得 暂存动作列表
    Public Property Get getGlobalDimList()
        getGlobalDimList = globalDimList 
    End Property 
    '设置函数名称列表
    Public Property Let setGlobalDimList(str)
        globalDimList = "|" & str & "|"
    End Property
    '获得 暂存动作列表
    Public Property Get getTempActionList()
        getTempActionList = tempActionList 
    End Property 
    '设置函数名称列表
    Public Property Let setFunctionNameList(str)
        functionNameList = LCase("|" & str & "|") 
    End Property 
    '设置不显示函数名称列表
    Public Property Let setNotShowFunctionList(str)
        notShowFunctionList = LCase("|" & str & "|") 
    End Property 
    '设置留空函数名称列表
    Public Property Let setSpaceFunctionList(str)
        spaceFunctionList = LCase("|" & str & "|") 
    End Property 
    '设置是否显示HTML部分
    Public Property Let setIsDisplayHtml(str)
        isDisplayHtml = str 
    End Property
    '设置是否导入函数功能
    Public Property Let setIsOpenImportFunction(str)
        isOpenImportFunction = str 
    End Property
    '设置类名称列表
    Public Property Let setClassNameList(str)
        classNameList = "|" & str & "|"
    End Property 
	'设置转换成什么语言	
    Public Property Let setToLanguage(str)
        toLanguage=str
    End Property 
	
 
	'处理PHP;分号  20160405
	function handlePHPSemicolon(byval code)
		dim i,s,c,leftStr,rightStr
		if right(code,1)=";" then
			for i=2 to len(code)
				s=mid(code,len(code)-i+1,1)
				if (s=" " or s=vbtab) and  leftStr="" then
					rightStr=rightStr & s
				else
					leftStr= s & leftStr
				end if
			next
			code=leftStr & ";" & rightStr
		end if
		handlePHPSemicolon = code
	end function
    '处理获得Asp代码
    Function aSPToPHP(ByVal content)
        Dim splStr, sx, tempSx, s, ganJingS, c, i, yesASP, sYHCount, s1, s2, tempZc, nI 
        Dim noHandleStr                                             '不处理字符与长度
        Dim wc                                                                          '输入文本存储内容
        Dim zc                                                                          '字母文件存储内容
        Dim yesOK                                                                       '是否OK
        Dim tempS, tempS2, lCaseS 
        Dim splxx, tempI 
        yesASP = False                                                                  '是ASP 默认为假
        sYHCount = 0                                                                    '双引号默认为0
        onDimRsStr = False                                                              '开户定义Rs字符
        nRsStrCount = 0                                                                 'Rs定义字符内容默认为0

		'不同语言用不能链接符，与变量前缀
		if toLanguage="js" then
			toLanguageLinkCha="+"
		else
			toLanguageLinkCha="."
			toLanguageVarNameCha="$"
		end if
		
        noHandleStr = "[#不处理#]" 
        
        splStr = Split(content, vbCrLf)                                                 '分割行
        '循环分行
        For nI = 0 To UBound(splStr)		
            s = splStr(nI) : pubRowS=s			'为一行内容
			
            s = Replace(Replace(s, Chr(10), ""), Chr(13), "") 				'奇怪为什么 s里会有 chr(10)与chr(13) 呢？  所以就要把 chr(10)与chr(13) 删除
            ganJingS = PHPTrim(s) 					'干净s
            lCaseS = LCase(ganJingS)                '小写干净的S内容
            tempS = s 								'暂存s内容
            rowC = "" : tempRow = ""                                                        '清空每行ASP代码和暂存完整行代码
            noteCode = ""                                                                   '默认注释代码为空
            downRow = ""                                                                    '下一行代码
            downWord = ""                                                                   '下一行单词
            If(nI + 1) <= UBound(splStr) Then
                downRow = splStr(nI + 1) 					'下一行内容
                downWord = getBeforeWord(downRow) 			'下一行单词
				'call echo("【下一行内容与单词】" & downRow,downWord)
            End If 
            nc = ""                                                                         '数字为空
            nDimArrayQianTao = 0                                                            '定义数据嵌套数
            'call echo("s",replace(lCaseS,"<","&lt;"))
            '不处理打头为 on 这种的
            If Left(pHPTrim(LCase(s)), 3) = "on " Then
                s = ""
			'不处理 
            ElseIf InStr(lCaseS, "if err then ") > 0 Or InStr(lCaseS, "err.description") > 0 Or InStr(lCaseS, "err.number") > 0 Or InStr(lCaseS, "if err<>") > 0 Then
                s = ""
			'把 )(0) 替换成 )[0]
            ElseIf InStr(lCaseS, "conn.execute(") > 0 And InStr(lCaseS, ")(0)") > 0 Then
                s = Replace(s, ")(0)", ")[0]") 
                'call eerr("s",s)
			'删除数字库记录集
            ElseIf InStr(lCaseS, "createobject(""adodb.recordset"")") > 0 Then
                s = "" 
			'为不处理字符，直接把 这行里的'去掉
            ElseIf Left(ganJingS, len(noHandleStr)) = noHandleStr Then
                c = c & Mid(ganJingS, len(noHandleStr) + 1) 
                s = ""
			'session失效
            ElseIf InStr(lCaseS, "session.timeout") > 0 Then
                s = "session_set_cookie_params(" & phptrim(Mid(s, InStr(s, "=") + 1)) & ")"
			'PHP里不处理
            ElseIf InStr(lCaseS, "'phpno") > 0 Then
                s = ""				
				
				
			'这应该是测试
			elseif ganJingS="'rsPHP" then
				call echo("isOpenRsStr",isOpenRsStr)
			
            End If

            isLoop = True                                                                   '循环字符为真
            '循环每个字符
            For i = 1 To Len(s)
                sx = Mid(s, i, 1) : tempSx = sx 
                beforeStr = Right(Replace(Mid(s, 1, i - 1), " ", ""), 1)                        '上一个字符
                afterStr = Left(Replace(Mid(s, i + 1), " ", ""), 1)                             '下一个字符
                endCode = Mid(s, i + 1)                                                         '当前字符往后面代码 一行

                'Asp开始
                If sx = "<" And yesASP = False And wc = "" Then
                    If Mid(s, i + 1, 1) = "%" Then
                        yesASP = True                                                                   'ASP为真
						
						'为转换成js语言
						if toLanguage="js" then
							sx="<script type=""text/javascript"">"
						'默认为转PHP
						else
							sx = "<" & "?" 				'前面换行删除了，因为在<input value=''>会有问题 20160315
							'如果ASP不是在一行则打头加上PHP标识，但是这种方法也不精准 20160307
							if instr(endCode,"%" & ">")=false then
								sx=sx & "PHP"
							end if	
						end if					
						upword="<" & "%"
						
                        i = i + 1 
                        rowC = rowC & sx 
                    ElseIf isDisplayHtml = True Then
                        rowC = rowC & sx                                                                '给不是ASP代码使用
                    End If 

                'ASP结束
                ElseIf sx = "%" And Mid(s, i + 1, 1) = ">" And yesASP = True And wc = "" Then
                    yesASP = False                                                                  'ASP为假		
					'为转换成js语言
					if toLanguage="js" then
						sx="</script>"
					else
	                    sx = "?" & ">"
					end if
                    i = i + 1 
                    rowC = rowC & sx
                'ASP为真
                ElseIf yesASP = True Then
                    '输入文本
                    If(sx = """" Or wc <> "") Then
                        '双引号累加
                        If sx = """" Then sYHCount = sYHCount + 1 
                        '判断是否"在最后
                        If sYHCount Mod 2 = 0 Then
                            If Mid(s, i + 1, 1) <> """" Then
                                wc = wc & sx 
                                s1 = Right(Replace(Mid(s, 1, i - Len(wc)), " ", ""), 1) '必需放在这里，要不会出错

                                '处理输入内容
                                wc = Right(wc, Len(wc) - 1) 
                                wc = Left(wc, Len(wc) - 1)                                

                                wc = aSPVarToPHPVar(wc, "输入")                         'ASP变量转PHP变量

                        
								if wc="'[为字典不输出内容]'" and isZD=true then
									wc=""
								end if

                                rowC = rowC & wc                                        '函数代码累加
                                sYHCount = 0 : wc = ""                                  '清除
                            Else
                                wc = wc & sx 
                            End If 
                        Else
                            wc = wc & sx 
                        End If 

                    '+-*\=&   字符特殊处理
                    ElseIf InStr(".&=,+-*/:()><[]", sx) > 0 And yesASP = True And nc = "" Then
                        If zc <> "" Then
                            tempZc = zc 
                            zc = aSPVarToPHPVar(zc, "变量")              'ASP变量转PHP变量

							upWordYuan6 = upWordYuan5                    '第6源单词 等于 源第5单词
							upWordYuan5 = upWordYuan4                    '第5源单词 等于 源第4单词
							upWordYuan4 = upWordYuan3                    '第4源单词 等于 源第3单词
							upWordYuan3 = upWordYuan2                    '第3源单词 等于 源第2单词
							upWordYuan2 = upWordYuan                     '第2源单词 等于 源第1单词
                            upWordYuan = tempZc                          '记录上一变量
							upWord6=upWord5
							upWord5=upWord4
							upWord4=upWord3
							upWord3=upWord2
                            upWord2 = upWord                             '记录上上一个变量
                            upWord = LCase(tempZc)                       '记录上一个变量名称   为小写
                            rowC = rowC & zc 
                            zc = ""                                      '清空字母内容
                        End If 
						if sx="=" then 
							rowC=rtrim(rowC)				 
						end if
                        'call echo("Sx",Sx)
                        rowC = rowC & aSPVarToPHPVar(sx, "") 
						upWordYuan6 = upWordYuan5                    '第6源单词 等于 源第5单词
						upWordYuan5 = upWordYuan4                    '第5源单词 等于 源第4单词
						upWordYuan4 = upWordYuan3                    '第4源单词 等于 源第3单词
						upWordYuan3 = upWordYuan2                    '第3源单词 等于 源第2单词
						upWordYuan2 = upWordYuan                     '第2源单词 等于 源第1单词
						upWord6=upWord5
						upWord5=upWord4
						upWord4=upWord3
						upWord3=upWord2
                        upWord2 = upWord                                                                '记录上上一个变量
                        upWord = sx 
                    '注释则退出
                    ElseIf sx = "'" Then
                        noteCode = "//" & Mid(s, i + 1)
						if left(noteCode,7)="//【PHP】" then
							noteCode=mid(noteCode,8)
						end if
                        If isSpaceFunctionName = "留空函数" Or isNotShowFunctionName = "不显示函数" Then
                            noteCode = "" 
                        End If 
                        Exit For 
                    '字母
                    ElseIf checkABC(sx) = True Or sx = "_" Or zc <> "" Then
                        zc = zc & sx 
                        yesOK = True 
                        s1 = Mid(s & " ", i + 1, 1) 
                        s2 = Mid(zc, 1, 1) 
                        If checkABC(s1) <> True And s1 <> "_" Then
                            yesOK = False 
                        End If 
                        '允许变量后面是数字
                        If checkNumber(s1) = True And checkABC(s2) = True Then
                            yesOK = True 
                        End If 
                        If yesOK = False Then
                            'Rem注释
                            If LCase(zc) = "rem" Then
                                sx = Mid(s, i + 1) : i = i + Len(sx) 
                                zc = zc & sx 
                            End If 

                            tempZc = zc 
                            zc = aSPVarToPHPVar(zc, "变量")              'ASP变量转PHP变量

							upWordYuan6 = upWordYuan5                    '第6源单词 等于 源第5单词
							upWordYuan5 = upWordYuan4                    '第5源单词 等于 源第4单词
							upWordYuan4 = upWordYuan3                    '第4源单词 等于 源第3单词
							upWordYuan3 = upWordYuan2                    '第3源单词 等于 源第2单词
							upWordYuan2 = upWordYuan                     '第2源单词 等于 源第1单词
                            upWordYuan = tempZc                          '记录上一变量
							upWord6=upWord5
							upWord5=upWord4
							upWord4=upWord3
							upWord3=upWord2
                            upWord2 = upWord                             '记录上上一个变量
                            upWord = LCase(tempZc)                       '记录上一个变量名称   为小写
                            rowC = rowC & zc 
                            zc = ""                                      '清空字母内容
                        End If 
                    '为数字
                    ElseIf checkNumber(sx) = True Or nc <> "" Then
                        nc = nc & sx 
                        If afterStr <> "." And checkNumber(afterStr) = False Then
                            nc = aSPVarToPHPVar(nc, "数字")              'ASP变量转PHP变量
                            rowC = rowC & nc 
                            nc = "" 
                        End If 
                    Else
                        rowC = rowC & sx 

                    End If 
                    tempRow = tempRow & sx 

                ElseIf isDisplayHtml = True Then
                    rowC = rowC & sx 
                End If 

                If isLoop = False Then Exit For 
            Next 
			
            If trimVbCrlf(rowC) <> "}" And yesASP = True Then                               '当ASP为真时
                rowC = rowC & aSPVarToPHPVar("", "换行")
            End If 
			'一行就;分号  则为空
            If pHPTrim(rowC) = ";" Then
                rowC = "" 
            End If 
            '为HTML加换行
            If yesASP = False And rowC <> "" And isDisplayHtml = True Then
				'删除ASP加载部分
                If Left(Trim(rowC), 4) = "<!--" And Right(Trim(rowC), 3) = "-->" Then
                    rowC = "" 
                Else
                    rowC = rowC & vbCrLf 
                End If 
            End If
            '为真，则开始定义的RS(20151120)      【重点】
            If onDimRsStr = True Then
                onDimRsStr = False 
                tempS = getStrCut(c, "<rs start>" & nRsStrCount, "<rs end>", 2) 
                tempS2 = LCase(splStr(nI - 1)) 
                If InStr(tempS2, ".recordcount") > 0 Or InStr(tempS2, ".eof") > 0 Then
                    c = c & tempS & vbCrLf 
                Else
                    c = Replace(c, "<rs start>" & nRsStrCount & tempS & "<rs end>", vbCrLf & tempS) 
                End If 
            End If 

			'删除当前字典名称20160308   【重点】
			if isZD=true then
				rowC=replace(rowC,"$"& zdName &"=array();","$"& zdName &"=array(")
			end if
			'call echo("notecode", noteCode)
			'rowC=rowC & noteCode                                                                '加上注释代码
			rowC=handlePHPSemicolon(rowC) & noteCode				'处理php;分号
            c = c & rowC
            c = c & isOpenRsStr 		'加上打开RS字符  
            isOpenRsStr = "" 
			'对不显示函数 与留宿函数 处理  不要换行20151223
			if isNotShowFunctionName <> "不显示函数" and isSpaceFunctionName <> "留空函数" then	
	            If yesASP = True Then c = c & vbCrLf                                            'Css为真则
			end if
        Next 
        'Call PrintArray(ArrayVar)

        '最后删除rs定义暂存
        If 1 = 1 Then
            For tempI = 0 To 1000
                If InStr(c, "<rs start>") = False Then Exit For 
                tempS = getStrCut(c, "<rs start>", "<rs end>", 1) 
                c = Replace(c, tempS, "") 
            Next 
        End If
        '最后处理PHP正确
        c = handlePHPCorrect(c) 
		aSPToPHP =  saveHandleContent(c)

    End Function
    'ASP变量转PHP变量   Variable(变量)
    Function aSPVarToPHPVar(varName, sType)
        Dim s, action, tempVarName, temp1
		'加载php文件里函数列表   不优化，会让程序跑慢 暂时留着20160306
		 if instr("|"& importFunctionList &"|","|"& varname & "|")=false then
			importFunctionList = importFunctionList & varname & "|"
		end if 

        '让Sub转function
        If LCase(varName) = "sub" Then
            varName = "function" 
        ElseIf sType = "输入" Then
            varName = Replace(varName, """""", """")                                        '把两个""替换成一个"
            varName = Replace(Replace(varName, "\", "\\"), "'", "\'")                       ',"""","\""")  这个不需要了
			
			
			if trim(lcase(varName))="scripting.dictionary"and upword2="createobject" then
				varName="[为字典不输出内容]"
				ZDName=upword4
				if lcase(ZDName)="server" then
					ZDName=upWordYuan6
				end if
				ZDNameList=ZDNameList & ZDName & "|"
				rowC=mid(rowC,1,instr(rowC, "$" & ZDName)+len(ZDName)) & "=array(" 
				isZD=true
			end if
			
            '对分割处理值  替换
            If varName = "{\\$" Then
                If upWord2 = "startstr" And upWord = "=" Then
                    varName = "{$" 
                End If 
            ElseIf varName = "\\$}" Then
                If upWord2 = "endstr" And upWord = "=" Then
                    varName = "$}" 
                End If 
            '为读数据库字符，则转小写
            ElseIf InStr("|rs|rsx|rss|rst|rsd|temprs|temprs2|rstemp|", "|" & upWord2 & "|") > 0 Then
                varName = LCase(varName) 
            End If 

            aSPVarToPHPVar = "'" & varName & "'" 
            If isSpaceFunctionName = "留空函数" Or isNotShowFunctionName = "不显示函数" Then
                aSPVarToPHPVar = "" 
            End If 
            If aSPVarToPHPVar = "'""'" Then
                aSPVarToPHPVar = "'[#晕这是一个双引号，显示就出错#]'" 
            End If 
            Exit Function 
        End If 

        tempVarName = varName                                                           '暂时变量名称
        s = LCase(varName) 

        '清除不需要显示的变量或函数
        If InStr(LCase("|add|randomize|byval|byref|exit|each|option|explicit|expires|not|do|set|redim|recordcount|eof|bof|AddNew|update|movenext|Err|Clear|"), "|" & s & "|") > 0 And s <> "" Then
            aSPVarToPHPVar = "" 
            Exit Function 
        End If 

        If s = "request" And phptrim(Mid(endCode, 1, 1)) = "." Then
            aSPVarToPHPVar = "" : Exit Function 
        End If 

        '特殊词转换
        If LCase(varName) = "print_r" Then
            varName = "PHPPrint_R" 
        ElseIf LCase(varName) = "die" Then
            varName = "PHPDie" 
        ElseIf LCase(varName) = "print" Then
            varName = "PHPPrint" 
        ElseIf LCase(varName) = "exec" Then
            varName = "PHPExec" 
        ElseIf LCase(varName) = "str_repeat" Then
            varName = "PHPStr_Repeat" 
        ElseIf LCase(varName) = "strlen" Then
            varName = "PHPStrLen" 
        ElseIf LCase(varName) = "str_shuffle" Then
            varName = "PHPStr_Shuffle" 
        ElseIf LCase(varName) = "str_split" Then
            varName = "PHPStr_Split" 
        ElseIf LCase(varName) = "max" Then
            varName = "PHPMax" 
        ElseIf LCase(varName) = "min" Then
            varName = "PHPMin" 

        ElseIf LCase(varName) = "echo" Then
            varName = "ASPEcho" 
        ElseIf LCase(varName) = "trace" Then
            varName = "FlashTrace" 
        ElseIf LCase(varName) = "redirect" Then
            varName = "AspRedirect" 
        ElseIf LCase(varName) = "trim" Then
            varName = "AspTrim" 
        ElseIf LCase(varName) = "ltrim" Then
            varName = "AspLTrim" 
        ElseIf LCase(varName) = "rtrim" Then
            varName = "AspRTrim" 
            'elseIf lCase(varName) = "savefile" then
            'varName = "AspSavefile"
        ElseIf LCase(varName) = "date" Then
            varName = "ASPDate" 
		'20160215
        ElseIf LCase(varName) = "chrw" Then
            varName = "chr" 


        '函数替换  不要这种，对汉字处理不好2016043
        'ElseIf LCase(varName) = "len" Then
        '    varName = "strlen" 

        ElseIf LCase(varName) = "openconn" Then
            varName = "conn=OpenConn" 
        ElseIf s = "default" Then
            varName = "defaultStr" 
        ElseIf s = "md5" Then
            varName = "aspMD5" 


        ElseIf s = "request" Then
            varName = "@$_REQUEST" 
            dimAraayList = dimAraayList & s & "|"                                           '加入数组列表里
        ElseIf s = "form" Then
            varName = "@$_POST" 
            dimAraayList = dimAraayList & s & "|"                                           '加入数组列表里
        ElseIf s = "querystring" Then
            varName = "@$_GET" 
            dimAraayList = dimAraayList & s & "|"                                           '加入数组列表里
        ElseIf s = "cookies" Then
            varName = "@$_COOKIE" 
            dimAraayList = dimAraayList & s & "|"                                           '加入数组列表里
        ElseIf s = "session" Then
            varName = "@$_SESSION" 
            dimAraayList = dimAraayList & s & "|"                                           '加入数组列表里

        ElseIf s = "public" Or s = "private" Then
            'call echo(s,s)
            varName = "dim" 
            s = "dim" 
            '下一个单词是function 或 sub时 就为空退出
            If InStr("function|sub", LCase(getBeforeWord(endCode))) > 0 Then
                aSPVarToPHPVar = "" 
                Exit Function 
            End If 

        ElseIf s = "isnumeric" Then
            aSPVarToPHPVar = "is_numeric" 
			if isNotShowFunctionName = "不显示函数" or isSpaceFunctionName = "留空函数" then
				aSPVarToPHPVar=""
			end if
            Exit Function 
		
		'ASP与PHP同步时，下面就不要
        ElseIf s = "lcase" Then			
            varName = "strtolower" 
        ElseIf s = "ucase" Then			
            varName = "strtoupper" 			
        ElseIf s = "asc" Then
            varName = "ord"

        ElseIf s = "clng" Then
            varName = "intval"
        ElseIf s = "isarray" Then
            varName = "is_array"
        ElseIf s = "typename" Then
            varName = "gettype"
		'替换这个函数，因为这个函数在php已经是系统内置函数了20160511
        ElseIf s = "readfile" Then
            varName = "reaFile"
			
			
			
            
			
        End If 
        '[]特殊字符处理
        If varName = "[" Or varName = "]" Then
            If isSpaceFunctionName = "留空函数" Then
                varName = "" 
            End If 
            aSPVarToPHPVar = varName 
            Exit Function 
        End If 

        '数组定义 让数字字符小写
        If LCase(varName) = "array" Then
            varName = "array" 
        End If 


        '处理不显示函数
        If isNotShowFunctionName = "不显示函数" Then
            If InStr("|function|sub|", "|" & getArrayVar() & "|") = False Then
                aSPVarToPHPVar = "" 
                Exit Function 
            Else
                If InStr("|function|sub|", "|" & LCase(varName) & "|") = False Then
                    aSPVarToPHPVar = "" 
                    Exit Function 
                End If 
            End If 
        ElseIf isSpaceFunctionName = "留空函数" Then
            If InStr("|function|sub|", "|" & getArrayVar() & "|") = False Then
                aSPVarToPHPVar = "" 
                Exit Function 
			'真假判断  为假
            ElseIf trueJudge = False And s <> "function" Then
                aSPVarToPHPVar = "" 
                Exit Function 
            End If 
        End If 


        '是否为IF一行判断代码
        If ifJudgeRow(getArrayIndex()) = 2 Then
            ifJudgeRowStr(getArrayIndex()) = ifJudgeRowStr(getArrayIndex()) & varName
        End If 

        '判断是否为系统函数 真为追加变量
        If checkAspSystemFunction(s) = True Then
            Call AddArrayVar(s)                                                             '追加数组变量
            arrayVarLeftKuoHao(getArrayIndex()) = 0                                         '数据变量 左括号数
            arrayVarRightKuoHao(getArrayIndex()) = 0                                        '数据变量 右括号数
        End If 

        action = getArrayVar()                                                          '获得当前动作
        nThisIndex = getArrayIndex()                                                    '当前索引
        
		if action="class" then
			If LCase(upWord) <> "end" And s = "class" Then
				varName="class"
				isClass=true
				className=""
            ElseIf sType = "变量" Then
				if className="" then
					className=varname
					classNameList=classNameList & varName & "|"
					varName=varName &"{"
                elseIf upWord = "end" And s = "class" Then
                    varName = "}" 
					'删除二次是因为  class类处理二次20160307
                	Call DelArrayVar()                                                              '删除数组变量 
                	Call DelArrayVar()                                                              '删除数组变量 
					isClass=false:className=""
                Else
                    varName = handleVarName(s, varName, sType)                                      '获得变量前面是否加
                End If
			'Class里也有换行处理
            ElseIf sType = "换行" Then				 
                varName = handleBR(action, s, varName)                                          '处理换行 
			ElseIf sType <> "数字" Then
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
 				 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
			 
            End If  
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显 
        elseIf action = "if" Then
            If LCase(upWord) <> "end" And s = "if" Then
                varName = "if(" 
                ifJudge(getArrayIndex()) = True                                                 'IF为真
                ifJudgeRow(getArrayIndex()) = 1                                                 '为一行内容
                ifJudgeRowStr(getArrayIndex()) = ""                                             '清空一行内容
                ifRsObjHandStr = ""                                                             '清空Rs句柄内容
            ElseIf s = "then" Then
                varName = "){" & ifRsObjHandStr                                                 '加上Rs句柄内容
                ifJudge(getArrayIndex()) = False 
                ifJudgeRow(getArrayIndex()) = 2 
            ElseIf s = "else" Then
                varName = "}else{" 
                If ifJudgeRow(getArrayIndex()) = 2 Then
                    varName = ";" & varName 
                End If 
                ifJudgeRow(getArrayIndex()) = 3 
            ElseIf s = "elseif" Then
                varName = "}else if(" 
                If ifJudgeRow(getArrayIndex()) = 2 Then
                    varName = ";" & varName 
                End If 
                ifJudgeRow(getArrayIndex()) = 3                                                 '为3
                ifJudge(getArrayIndex()) = True 
            ElseIf s = "end" Then
                varName = "" 
            ElseIf s = "if" And upWord = "end" Then                                         'lCase(getBeforeWord(tempRow)) 这是什么东东
                varName = "}" 
                Call DelArrayVar()                                                              '删除数组变量
                Call DelArrayVar()                                                              '删除数组变量
                ifJudgeRow(getArrayIndex()) = 0 
                ifJudgeRowStr(getArrayIndex()) = ""                                             'IF结束


            ElseIf sType = "变量" Then
                If upWord = "end" And s = "if" Then
                    varName = "" 
                    ifJudgeRow(getArrayIndex()) = 0 
                    ifJudgeRowStr(getArrayIndex()) = ""                                             'IF结束
                Else
                    varName = handleVarName(s, varName, sType)                                      '获得变量前面是否加
                End If 
            ElseIf sType = "换行" Then
                ifJudgeRow(getArrayIndex()) = ifJudgeRow(getArrayIndex()) + 1                   '累加判断一行内容
                varName = handleBR(action, s, varName)                                          '处理换行 
                'elseIf s = "=" and ifJudge(getArrayIndex()) = true and inStr("<>", beforeStr) = false then
                'tempWord= getBeforeNStr(endCode, "全部", 2)
                'call echo("tempWord",tempWord)
                'if tempWord<>"&" and tempWord<>"and" and tempWord<>"or" then
                'varName = "=="
                '
            'end if
            ElseIf s = "=" And ifJudge(GetArrayIndex()) = True And InStr("<>", beforeStr) = False Then
                varName = "==" 

            ElseIf sType <> "数字" Then
                varName = handleVarName(s, varName, sType)                                      '处理数组类型

            End If 

            Call addMsg(action, tempVarName, sType, varName)                                '追加回显


        ElseIf action = "iif" Then
            If s = "iif" Then
                varName = "IIF" 

            ElseIf s = "=" Then
                If arrayVarLeftKuoHao(getArrayIndex()) - 1 = arrayVarRightKuoHao(getArrayIndex()) Then
                    varName = "==" 
                End If 

            ElseIf sType = "换行" Then
                ifJudgeRow(getArrayIndex()) = ifJudgeRow(getArrayIndex()) + 1                   '累加判断一行内容
                varName = handleBR(action, s, varName)                                          '处理换行 

            ElseIf sType <> "数字" Then
                varName = handleVarName(s, varName, sType)                                      '处理数组类型

            End If 

            Call addMsg(action, tempVarName, sType, varName)                                '追加回显
        ElseIf action = "for" Then
            If s = "for" Then
                If upWord = "exit" Then
                    varName = "break;" 
                    Call DelArrayVar()                                                              '删除数组变量
                Else
                    If LCase(getBeforeWord(endCode)) = "each" Then
                        varName = "foreach(" 
                    Else
                        varName = "for(" 
                    End If 

                    forJudge = True 
                    forFirstVarName = "" 
                    forEachName = "" 
                    forStep = "++" 
                    isForStep = False 
                End If 
            ElseIf s = "in" Then
                varName = "" 
            ElseIf s = "to" Then
                '为步数则大于
                If InStr(LCase(endCode), " step ") > 0 Then
                    varName = "; " & forFirstVarName & ">=" 
                Else
                    varName = "; " & forFirstVarName & "<=" 
                End If 
            ElseIf s = "exit" Then
                varName = "" 
            ElseIf s = "next" Then
                varName = "}" 
                Call DelArrayVar()                                                              '删除数组变量
            ElseIf s = "step" Then
                forStep = "--" 
                isForStep = True 
                varName = "" 
            ElseIf sType = "变量" Then
                If isForStep = True Then
                    varName = "" 
                Else
                    varName = handleVarName(s, varName, sType)                                      '获得变量前面是否加
                    If forFirstVarName = "" Then
                        forFirstVarName = varName                                                       '记录For第一个变量名
                        If upWord = "each" Then
                            forEachName = varName                        '记录循环each变量名
                            varName = "" 
                        End If 
                    End If 
                End If 
            ElseIf sType = "换行" Then
                isForStep = False                                                               '换行了，循环频数为假
                varName = handleBR(action, s, varName)                                          '处理换行 
            ElseIf sType = "数字" Then
                If isForStep = True Then
                    'forStep=forStep & varName
                    varName = "" 
                End If 
            Else
                If isForStep = True Then
                    varName = "" 
                Else
                    varName = handleVarName(s, varName, sType)                                      '处理数组类型
                End If 
            End If 


            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        '设置传值
        ElseIf action = "select" Then
            If LCase(upWord) <> "end" And s = "select" Then
                varName = "" 
                isSelectCase = True 
                isCase = False 
                nCaseCount = 0                                                                  'Case总数为0
            ElseIf s = "end" Then
                varName = "" 
            ElseIf s = "select" And upWord = "end" Then                                     'lCase(getBeforeWord(tempRow)) 这是什么东东
                varName = "}"  
				'这里面要删除二次，要不出错，为什么(20160115)
                Call DelArrayVar()                                                              '删除数组变量
                Call DelArrayVar()                                                              '删除数组变量
				
            ElseIf s = ":" Then
                varName = ";" 
            ElseIf s = ":" And isCase = True Then
                isCase = False 

            ElseIf s = "," And isSelectMoreCheck = "多个判断" Then
                varName = ";" 
            ElseIf s = "else" Then
                varName = "default" 

            ElseIf s = "case" Then
                varName = LCase(varName)                                                        '必需为小写，在PHP里
                If isSelectCase = True Then
                    varName = "switch (" 
                '如果下一个单词为else 则这个case为空
                ElseIf LCase(getBeforeWord(endCode)) = "else" Then
                    varName = "" 
                End If 
                isSelectMoreCheck = "多个判断" 
                isCase = True 
                nCaseCount = nCaseCount + 1                                                     'Case总数累加
            ElseIf sType = "换行" Then
                If isSelectCase = True Then
                    varName = varName & " ){" 
                    isSelectCase = False 
                Else
                    varName = handleBR(action, s, varName)                                          '处理换行
                    If LCase(downWord) = "case" Then
                        If nCaseCount > 1 Then
                            varName = varName & "break;" 
                        End If 
                    End If 
                End If 
                isSelectMoreCheck = "清除多个判断" 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 
            '代码分割了，是第二段代码了
            If s = ":" Then
                isSelectMoreCheck = "清除多个判断" 
            End If 

            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        ElseIf action = "dim" Then
            '为多维数组则为空
            If ismultidimensionalArray = "是多维数组" Then
				if toLanguage="js" then
					varName = "var" 
				else
                	varName = "" 
				end if
            End If 
            If s = "dim" Then
				if toLanguage="js" then
					varName = "var" 
				else
                	varName = "" 
				end if
                isDimAraay(getArrayIndex()) = "定义变量" 
                functionInDimList = functionInDimList & Replace(getStrDim(endCode), ",", "|") & "|" '函数里定义的变量累加
            ElseIf s = "," Then
                If isDimAraay(getArrayIndex()) = "定义变量" Then
                    varName = "='';" 
                '数组完成又成变量了,因为，代码后面还有变量呀
                Else
                    isDimAraay(getArrayIndex()) = "定义变量" 
                    varName = "" 
                End If 
            ElseIf s = ":" Then
                If isDimAraay(getArrayIndex()) = "定义变量" Then
                    varName = "='';" 
				'追加于20160311
				else
					varName=""
                End If
                isDimAraay(getArrayIndex()) = ""                                                '清除定义变量
                Call DelArrayVar()                                                              '删除数组变量
            ElseIf s = "(" Then

                dimAraayList = dimAraayList & LCase(upWord) & "|" 

                varName = "=array(" 
                isDimAraay(getArrayIndex()) = ""                                                '清空为定义数组
                nDimAraayLeftKuoHao(5) = 1                                        '定义数组左括号数 初始化为1
                nDimAraayRightKuoHao(5) = 0                                       '定义数组右括号数 初始化为0

                temp1 = Mid(endCode, 1, InStr(endCode, ")") - 1) 
                ismultidimensionalArray = "" 
                If InStr(temp1, ",") > 0 Then
                    ismultidimensionalArray = "是多维数组" 
                End If 

            ElseIf s = ")" Then
                ismultidimensionalArray = ""                                                    '多维数组 为空
                varName = ");" 

            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行
				Call DelArrayVar()                                                              '删除数组变量 20160408  换行则删除dim定义的
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 


            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        '设置传值
        ElseIf action = "response" Then
            If s = "response" Then
                varName = "" 
                isHeader = False 
            ElseIf s = "write" Then
				if toLanguage="js" then
					varName = "document.write" 
				else
	                varName = "echo" 
				end if

                'elseIf s = "cookies" or s = "querystring" then
                'varName = "@$_COOKIE"
                'isToSetCookies = "设置Cookies开始"                     '注意这个标识可以拿掉了(20151123)

            ElseIf s = "flush" Then
                varName = "PHPFlush" 
            ElseIf s = "end" Then
                varName = "die" 

            ElseIf s = "," And isHeader = True Then
                varName = toLanguageLinkCha & " ': ' " & toLanguageLinkCha

            ElseIf s = "addheader" Then
                varName = "header(" 
                isHeader = True 

            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
                If isHeader = True Then
                    varName = ");" 
                End If 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        '显示传值
        ElseIf action = "request" Then
            If s = "request" Then
                If getBeforeStr(endCode) = "." Then
                    varName = "" 
                End If 
            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        ElseIf action = "function" Then
            If s = "function" Then
                If upWord = "end" Then

                    If isNotShowFunctionName = "不显示函数" Then
                        varName = "" 
                    '可用可无 (不对有用20151119)
                    ElseIf isFunctionReturn = True Then
                        varName = "return @$" & functionName & ";" & vbCrLf & "}" 
                    Else
                        varName = "}" 
                    End If 
                    If isSpaceFunctionName = "留空函数" Then
                        varName = "//留空函数" & vbCrLf & varName 
                    End If 

                    functionName = ""                                                               '清空函数名称
                    isNotShowFunctionName = ""                                                      '清除不显示函数
                    isSpaceFunctionName = ""                                                        '清空留空
                    Call DelArrayVar()                                                              '删除数组变量
                    Call DelArrayVar()                                                              '删除数组变量
                    functionInDimList = ""                                                          '清空函数里定义变量
                '退出 并删除一次数组
                ElseIf upWord = "exit" Then
                    'call echo(varName & " " & functionNameList & " " & functionname,isFunctionReturn)
                    '有返回值
                    If isFunctionReturn = True Then
                        varName = "return @$" & functionName & ";" 
						
                    Else
                        varName = "return '';" 														'函数在退出之前没有返回值，则返回个空
                    End If 
                    Call DelArrayVar()                                                              '删除数组变量
                Else
                    varName = "function " 
                    trueJudge = True                                                                '判断为真
                    functionName = "" 
                    isFunctionDimName = "函数变量名称" 
                    isFunctionReturn = False 

                    'call echo("getBeforeWord",lCase(getBeforeWord(endCode)))
                    If InStr("|" & notShowFunctionList & "|", "|" & LCase(getBeforeWord(endCode)) & "|") > 0 Then
						
                        isNotShowFunctionName = "不显示函数" 
                        varName = "" 
                    ElseIf InStr("|" & spaceFunctionList & "|", "|" & LCase(getBeforeWord(endCode)) & "|") > 0 Then
                        isSpaceFunctionName = "留空函数" 
                    'call echo("留空函数",spaceFunctionList)
                    End If 
                End If 
            ElseIf s = ")" And trueJudge = True Then
                varName = "){" 
                trueJudge = False 
                isFunctionDimName = "关闭函数变量名称" 

            ElseIf sType = "变量" Then
                If functionName = "" Then
                    functionName = varName 
                    functionNameList = functionNameList & LCase(varName) & "|"                      '函数名称累加
                    'call echo(varName,functionNameList)
                    nShowFunction = 1 

                    functionInDimList = Replace(getFunctionStrDim(endCode), ",", "|") & "|"         '函数里定义的变量 这里不要累加
                ElseIf s = "end" Then
                    varName = "" 
                Else
                    '函数定义参数变量给定义数组列表
                    If trueJudge = True Then
                        dimAraayList = dimAraayList & LCase(varName) & "|" 
                    End If 
                    varName = handleVarName(s, varName, sType)                                      '获得变量前面是否加
                End If 

            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
            Else
                varName = handleVarName(s, varName, sType)                                      '获得变量前面是否加
            End If 

            If isNotShowFunctionName = "不显示函数" Then
                varName = "" 
            End If 

            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        ElseIf action = "call" Then
            If s = "call" Then
                varName = "" 
            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显


        ElseIf action = "left" Then
            If s = "left" Then
                varName = "substr" 
            ElseIf s = "," Then
                varName = ", 0 ," 
            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

		'屏蔽
        ElseIf action = "right" and 1=2 Then
            If s = "right" Then
                varName = "substr" 
            ElseIf s = "," Then
                varName = ", -" 
            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        ElseIf action = "rnd" Then
            If s = "rnd" Then
                varName = "rnd" 
                '下一个字符不等于(
                If afterStr <> "(" Then
                    varName = varName & "()" 
                    Call DelArrayVar()                                                              '删除数组
                End If 
            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        ElseIf action = "cint" Then
            If s = "cint" Then
                varName = "intval" 
                Call DelArrayVar() 
            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        ElseIf action = "fix" Then                                                      '不启用
            If s = "fix" Then
                varName = "floor" 
                Call DelArrayVar() 
            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显



        ElseIf action = "instr" Then
            If s = "instr" Then
                varName = "instr"                                                               '注意可以改成这个  strpos
            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        ElseIf action = "instrrev" Then
            If s = "instrrev" Then
                varName = "strrpos" 
            ElseIf s = "&" Then
                varName = toLanguageLinkCha
            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        ElseIf action = "split" Then
            If s = "split" Then
                varName = "aspSplit" 
                Call DelArrayVar() 
                '必需为等于号
                If upWord = "=" Then
                    dimAraayList = dimAraayList & LCase(upWord2) & "|"                              '加入数组列表里
                End If 
            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型

            End If 
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        ElseIf action = "mod" Then
            If s = "mod" Then
                varName = "%" 
                Call DelArrayVar() 
            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        ElseIf action = "vbtab" Then
            If s = "vbtab" Then
                varName = """\t""" 
                Call DelArrayVar() 
            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        ElseIf action = "vbcrlf" Then
            If s = "vbcrlf" Then
                varName = "vbCrlf()" '"""\n""" 
                Call DelArrayVar() 
            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        ElseIf action = "mid" Then
            If s = "mid" Then
                varName = "mid" 
                nMidDuanLuo = 0
            ElseIf s = "," Then
                nMidDuanLuo = nMidDuanLuo + 1 
            ElseIf sType = "换行" Then
                varName = handleBR(action, s, varName)                                          '处理换行 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型
            End If 

            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        '追加于20150923
        ElseIf action = "server" Then
            If s = "server" Then
                varName = "" 
                isSearverAddKuoHu = False                                                       'Server里有括弧
            ElseIf s = "." Then
                varName = "" 
            ElseIf s = "=" Then
                varName = "" 
            ElseIf s = "scripttimeout" Then
                varName = "set_time_limit(" 
                isSearverAddKuoHu = True                                                        'Server里有括弧
            ElseIf sType = "换行" Then
                'Server里有括弧
                If isSearverAddKuoHu = True Then
                    varName = varName & ")" 
                End If 
                varName = varName & ";" 
                Call DelArrayVar()                                                              '删除数组变量
            End If 
            'call echo("s",s)
            Call addMsg(action, tempVarName, sType, varName)                                '追加回显

        ElseIf action = "while" Then
            If s = "while" Then
                varName = "while(" 
                isWhile = True 
            ElseIf s = "loop" Or s = "wend" Then
                varName = "}" 
                Call DelArrayVar()                                                              '删除数组变量
            ElseIf sType = "换行" Then
                If isWhile = True Then
                    isWhile = False 
                    varName = varName & "){" 
                Else
                    varName = handleBR(action, s, varName)                                          '处理换行
                End If 
            Else
                varName = handleVarName(s, varName, sType)                                      '处理数组类型

            End If 

            Call addMsg(action, tempVarName, sType, varName)                                '追加回显
        ElseIf sType = "换行" Then
            varName = handleBR(action, s, varName)                                          '处理换行
            Call addMsg("换行", tempVarName, sType, varName)                                '追加回显


        ElseIf sType = "变量" Then 
			'全部大写为php全局变量 20160228			
            If instr("|"& globalDimList &"|","|"& varName &"|")=false Then
            	varName = handleVarName(s, varName, sType)
			end if 
            Call addMsg("变量", tempVarName, sType, varName)                                '追加回显


        '后一个字符为空 相当于在此行最后
        ElseIf afterStr = "" Then                                                       '最后一个字符
            varName = handleVarName(s, varName, sType)                                      '处理数组类型
            If getBeforeStr(rowC) <> "" And getAfterStr(rowC) <> "{" And getAfterStr(rowC) <> "}" And s <> "" Then
                varName = varName & ";" 
            End If 
            Call addMsg("最后一个字符为空", tempVarName, sType, varName)                    '追加回显

        Else
            varName = handleVarName(s, varName, sType)                                      '处理数组类型
            Call addMsg("空", tempVarName, sType, varName)                                  '追加回显

        End If 
		'call echo("varName",varName)
        aSPVarToPHPVar = varName 
    End Function
	
    '获得变量前面是否加$号 VarName=handleVarName(S,VarName)  这里面有数组的处理
    Function handleVarName(s, varName, sType) 
        Dim tempVarName 
        tempVarName = varName  
        '函数里可以调用自身函数（20151120）
        If functionName = varName Then
			'<>& 是处理这种   mytestfun=mytestfun & "aaa"  20160307  
            If getBeforeNStr(endCode, "全部", 1) <> "=" and getBeforeNStr(endCode, "全部", 1) <> "&"  Then
                handleVarName = varName 
                Exit Function 
            End If 
        End If 
		
		'为字典 aa.add "title","value"   20160308
		if instr(ZDNameList, "|"& varName &"|")>0 and getBeforeNStr(endCode, "全部", 1)="."then
			handleVarName = "" 
			Exit Function 
		end if
		if isZD=true and varName="," then
			handleVarName = "=>" 
			Exit Function 		
		elseif isZD=true and varName=":" then
			handleVarName = "," 
			Exit Function 		
		end if

        '函数有返回值，返回函数本身名称
        If LCase(varName) = LCase(functionName) Then
            isFunctionReturn = True 
        End If 

        If isNotShowFunctionName = "不显示函数" Then
            handleVarName = "" 
            Exit Function 
        End If 
        '为空退出
        If varName = "" Then
            handleVarName = varName 
            Exit Function 
        '为数字退出
        ElseIf InStr("0123456789", Left(varName, 1)) > 0 Or Left(varName, 1) = """" Then
            handleVarName = varName 
            Exit Function 
        End If 

        If s = "and" Then
            varName = "&&" 
        ElseIf s = "or" Then
            varName = "||" 
        ElseIf s = "true" Then
            varName = "true" 
        ElseIf s = "false" Then
            varName = "false" 
        '判断函数
        ElseIf InStr("|" & functionNameList & "|", "|" & LCase(varName) & "|") > 0 Then
		
			'call echo(varName,getBeforeNStr(endCode, "全部", 1))
            If LCase(varName) = LCase(functionName) Then                                    '暂时不启用
                varName = toLanguageVarNameCha & varName 
			
			'系统函数，临时用下20160307
			elseif getBeforeNStr(endCode, "全部", 1)<>"(" and instr(lcase("|CLng|ubound|Chr|replace|Write|Cookies|QueryString|TypeName|Redirect|FormatNumber|Cstr|Asc|Join|Len|PHPEcho|ASPEcho|LCase|UCase|IsNumeric|IsNull|Int|Execute|ServerVariables|CreateObject|Hex|AscW|time||case|Timer|DateDiff|Fix|now|getYMDHMS|IsDate|Year|Month|Day|Hour|Minute|Second|array|AspTrim|WebPageControl|HandleAction|XY_getLableValue|ReplaceValueParam|LTrim|RTrim|newGetArray|AspLTrim|AspRTrim|DateAdd|ASPDate|"),lcase(varName))=false then
				varname=varName & "("
				rowEndStr=")"
			
            End If 
        ElseIf functionReturn = "函数返回" And varName = "=" Then
            varName = "" 
            functionReturn = "函数返回结束" 
        '判断数组
        ElseIf InStr("|" & dimAraayList & "|", "|" & varName & "|") > 0 And getBeforeNStr(endCode, "全部", 1) <> "." Then
            varName = toLanguageVarNameCha & varName 
        ElseIf s = "(" Then
            'call echo(upword,dimAraayList)
            If isFunctionDimName = "函数变量名称" Then
                isFunctionDimName = "启用函数变量名称" 
            End If 

            '系统函数)数累加
            arrayVarLeftKuoHao(nThisIndex) = arrayVarLeftKuoHao(nThisIndex) + 1 
            'call echo(arrayVarLeftKuoHao(nThisIndex),nThisIndex)
            '为定义数组则累加(数  nDimArrayQianTao为深度
            If nDimArrayQianTao > 0 Then
                nDimAraayLeftKuoHao(5) = nDimAraayLeftKuoHao(5) + 1 
            'call echo(nDimArrayQianTao,"煺")
            End If 
            If isToSetCookies = "设置Cookies开始" Then
                varName = "[" 
            '此变量为数组
            ElseIf InStr("|" & dimAraayList & "|", "|" & LCase(upWord) & "|") > 0 Or(varName = "(" And upWord = ")") Then'让函数myfun()(1)  后面(1)为数组
                varName = "[" 
                dimAraayStart = "开始" 
                isDimAraay(nThisIndex) = ""                                                '为定义数组
                If nDimArrayQianTao = 0 Then
                    nDimAraayLeftKuoHao(5) = 1                                        '定义数组左括号数 初始化为1
                    nDimAraayRightKuoHao(5) = 0                                       '定义数组右括号数 初始化为0 
                End If 
                nDimArrayQianTao = nDimArrayQianTao + 1                                         '定义数据嵌套数累加
                'Call echo(varName, nThisIndex & "," & nThisIndex & ", nDimArrayQianTao=" & nDimArrayQianTao) 
            End If 

		'注意：Response.QueryString("aaa" & Response.QueryString("aaa") )  对这样处理还是有问题    nThisIndex  替换获得数据，解决 Response.QueryString("aaa")
        ElseIf s = ")" Then


            If isFunctionDimName = "启用函数变量名称" Then
                isFunctionDimName = "关闭函数变量名称" 
            End If 
            '系统函数)数累加
            arrayVarRightKuoHao(nThisIndex) = arrayVarRightKuoHao(nThisIndex) + 1 
            'Call echo("nThisIndex", nThisIndex) 

            'call echo(arrayVarLeftKuoHao(nThisIndex),arrayVarRightKuoHao(nThisIndex))
            'call echo(getArrayVar(),nMidDuanLuo)
            '()号相等
            If arrayVarLeftKuoHao(nThisIndex) = arrayVarRightKuoHao(nThisIndex) Then
                If InStr("|mid|", "|" & getArrayVar() & "|") > 0 Then
                    If nMidDuanLuo = 1 Then
                        varName = ",-1)" 
                    End If 
                    Call DelArrayVar()                                                              '删除数组变量
                ElseIf InStr("|if|for|while|select|", "|" & getArrayVar() & "|") = False Then
                    'call echo("del getArrayVar()",getArrayVar())
                    Call DelArrayVar()                                                              '删除数组变量
                End If 
            End If 

            '为定义数组则累加)数 nDimArrayQianTao为深度
            If nDimArrayQianTao > 0 Then
                nDimAraayRightKuoHao(5) = nDimAraayRightKuoHao(5) + 1 
            End If 
            'call echo(nDimAraayLeftKuoHao(nThisIndex),nDimAraayRightKuoHao(nThisIndex))
            If isToSetCookies = "设置Cookies开始" Then
                varName = "]" 
                isToSetCookies = "设置Cookies结束" 
            '数组
            ElseIf dimAraayStart = "开始" Then
                'call echo(varName & " " & nThisIndex & "=" & nThisIndex, nDimAraayLeftKuoHao(nThisIndex)-nDimArrayQianTao+1 & " = "&  nDimAraayRightKuoHao(nThisIndex))
                'call echo(nDimAraayLeftKuoHao(nThisIndex)-nDimArrayQianTao+1 & "=" &  nDimAraayRightKuoHao(nThisIndex),nThisIndex)
                'call echo(nDimAraayLeftKuoHao(nThisIndex)-nDimArrayQianTao+1, nDimAraayRightKuoHao(nThisIndex)   & "    nDimArrayQianTao=" & nDimArrayQianTao) 
				'call echo(nDimAraayLeftKuoHao(nThisIndex) - nDimArrayQianTao + 1 , nDimAraayRightKuoHao(nThisIndex) )
                If nDimAraayLeftKuoHao(5) - nDimArrayQianTao + 1 = nDimAraayRightKuoHao(5) Then
                    'call echo(isDimAraay(nThisIndex),"1111111111111")
                    varName = "]" 
                    nDimArrayQianTao = nDimArrayQianTao - 1 
                    If nDimArrayQianTao = 0 Then
                        dimAraayStart = "结束" 
                    'call echo("","结束")
                    End If 
                End If 
            End If 


        '二组数组
        ElseIf dimAraayStart = "开始" And varName = "," Then
            'call eerr("varname",varName)  这里面没有+1，那是因为它不是)符号 ，而是,这个符号
            If nDimAraayLeftKuoHao(5) - nDimArrayQianTao = nDimAraayRightKuoHao(5) Then
                'call echo(isDimAraay(nThisIndex),"1111111111111")
                varName = "][" 

            End If 

        '为单词
        ElseIf InStr(".&=,+-*/:()<>", varName) = False Then                      '加$前排除特殊字符

            '此单词为数组
            If afterStr = "(" Then
                If InStr("|" & dimAraayList & "|", "|" & LCase(varName) & "|") = False Then     '屏蔽重复
                    dimAraayList = dimAraayList & LCase(varName) & "|" 
                End If 
            End If 
			'类名称列表
			if instr(classNameList,"|"& varName &"|")>0 then
				'call echo(varname,classNameList)
				varName=varName & "();"
				'call echo(upword,upword3)
				dimClassNameList=dimClassNameList & upword3 & "|"
            elseIf sType = "变量" And functionName <> "" Then
			
                If InStr("|" & LCase(functionInDimList) & "|db|conn|rs|rsx|rss|rst|rsd|temprs|temprs2|rstmep|", "|" & LCase(Trim(varName)) & "|") = False Then'函数里有的定义的变量列表
                    If InStr(varName, "=") > 0 Then
                        varName = "$GLOBALS['" & Mid(varName, 1, InStr(varName, "=")) & "']" & Mid(varName, InStr(varName, "=")) 
                    Else
                        If instr("|"& globalDimList &"|","|"& varName &"|")=false Then
                            varName = "$GLOBALS['" & varName & "']" 
                        End If 
                    'call echo(functionInDimList,varname)
                    End If 
					
                Else
                    varName = toLanguageVarNameCha & varName 
                End If 
            Else
                varName = toLanguageVarNameCha & varName 
            End If 




            '为打开数据库表处理   处理rs.open
            If LCase(getBeforeNStr(endCode, "全部", 2)) = "open" Then
                varName = "" 
            '处理rs.recordcount
            ElseIf LCase(getBeforeNStr(endCode, "全部", 2)) = "recordcount" Then 
			    'varName = " $GLOBALS['conn']->num_rows($" & tempVarName & "Obj)" 
			    varName = " @mysql_num_rows($" & tempVarName & "Obj)" 
            '处理rs.close
            ElseIf LCase(getBeforeNStr(endCode, "全部", 2)) = "close" Or LCase(getBeforeNStr(endCode, "全部", 2)) = "addnew" Or LCase(getBeforeNStr(endCode, "全部", 2)) = "update" Or LCase(getBeforeNStr(endCode, "全部", 2)) = "movenext" Then
                varName = "" 


            ElseIf LCase(getBeforeNStr(endCode, "全部", 2)) = "eof" Then
                If LCase(upWord) = "not" And upWord2 = "while" Then
                    varName = "$" & tempVarName & "= $GLOBALS['conn']->fetch_array($" & tempVarName & "Obj)" 
                ElseIf LCase(upWord) = "not" Then
				
                    'varName = "$GLOBALS['conn']->num_rows($" & tempVarName & "Obj)!=0"
                    varName = "@mysql_num_rows($" & tempVarName & "Obj)!=0"
                Else
                    'varName = "$GLOBALS['conn']->num_rows($" & tempVarName & "Obj)==0" 
                    varName = "@mysql_num_rows($" & tempVarName & "Obj)==0" 
                End If 
                If getArrayVar() = "if" Then
                    '暂时不用，因为rs.conn已经使用
                    If InStr(LCase(" " & endCode & " "), "exit for") > 0 Then
                        ifRsObjHandStr = ""  
                    Else 
                    End If 
                    onDimRsStr = True 
                End If
            '为执行SQL语句
            ElseIf LCase(getBeforeNStr(endCode, "全部", 2)) = "execute" Then
                varName = "conn" 
            'call echo(varName,lcase(GetBeforeNStr(endcode, "全部", 2)))
            ElseIf s = "open" Then
                If InStr(endCode, " conn") > 0 Then
                    nRsStrCount = nRsStrCount + 1 
                    isOpenRsStr = "<rs start>" & nRsStrCount & copyStr(" ", 4) & "$" & upWord2 & "=mysql_fetch_array($" & upWord2 & "Obj);" & "<rs end>" 
                End If 
                varName = "$" & upWordYuan2 & "Obj=$GLOBALS['conn']->query(" 
                isOpenTable = "打开数据库表" 
            ElseIf s = "end" Or s = "close" Then
                varName = "" 
            End If 
        ElseIf s = ":" Then		 
            varName = rowEndStr &";" 
			rowEndStr=""
            nVariableSetValue(getArrayIndex()) = 0                                          '为换行，让变量赋值判断清零
        ElseIf s = "&" Then
            varName = toLanguageLinkCha
        ElseIf InStr(".", varName) > 0 Then
			'call echo(upword,dimClassNameList)
			'为类的定义变量名称
			if instr(dimClassNameList,"|"& upword &"|")>0 then
				varName="->"
			else
            	varName = ""
			end if 
        ElseIf s = "," And isOpenTable = "打开数据库表" Then
            If LCase(getBeforeNStr(endCode, "全部", 1)) = "conn" Then
                varName = ")" 
                isLoop = False                                                                  '不循环字符了
            End If 

        ElseIf s = "=" Then
			'call echo(11,nVariableSetValue(getArrayIndex()))
            '变量有赋值处理
            If nVariableSetValue(getArrayIndex()) >= 1 Then
                'call echo(varname,nVariableSetValue(getArrayIndex()))
                nVariableSetValue(getArrayIndex()) = nVariableSetValue(getArrayIndex()) + 1 
                If nVariableSetValue(getArrayIndex()) > 2 Then
					'在if里可以这个用==  20160321
					if getArrayVar()="if" then
	                    varName = "==" 
					end if
                'call echo(nVariableSetValue(getArrayIndex()),getArrayIndex())
                End If 
            End If
        End If 

        If sType = "变量" Then
            If getBeforeNStr(endCode, "全部", 1) = "=" Then
                If nVariableSetValue(getArrayIndex()) = 0 Then
                    'call echo(varname,"varname")
                    nVariableSetValue(getArrayIndex()) = 1 
                End If 
            End If 
        End If 


        If isFunctionDimName = "启用函数变量名称" Then
            'if left(varname,1)="$" and lcase(upword)<>"byref" then            '当设置为byrel了才加入 变量可以通过引用传递 因为它是必需是变量才可以用的，php比较严谨(20150828)
            If Left(varName, 1) = "$" And LCase(upWord) = "byref" Then                      '为引用变量 这样做的目录是为了使用函数不出错(20150923)
                varName = "&" & varName 
            'call echo(upword,varname)
            End If 
        End If 
        '让前面有@号的不要再在前面加$号了(20151123)
        If Mid(varName, 1, 2) = "$@" Then
            varName = Mid(varName, 2) 
        End If 
        handleVarName = varName 
    End Function
 
    '处理换行 varName = handleBR(action,s,varName)                '处理换行
    Function handleBR(action, s, varName) 
		dim isRowIf
		'判断是否为一行，也排除 else if 这种
		isRowIf=ifJudgeRowStr(getArrayIndex()) <> "" And action = "if" And s = "" and left(phptrim(lcase(pubRowS)),3)="if "
		' 
        '变量有赋值处理为零
        nVariableSetValue(getArrayIndex()) = 0				
		'为字典换行
		if isZD=true and upword6="add" and instr(rowc,"=>")>0 then
			'call echo(upword6,rowc)
			handleBR = "," 
			Exit Function
		elseif isZD=true and instr(rowc,"=>")=false then 		'要想字典结束完整，就必需在字段定义完最后空一行
			handleBR = ");" 
			isZD=false
			Exit Function 
		elseif isZD=true then
			handleBR = "," 
			Exit Function
		end if
		
		
		'放在字典下面，要不然就会字典会出错
		'排除一行if			终于解决了 if a=a then exit for  这种ASP  20160405
		if instr("{};",right(phptrim(rowc),1))>0 and isRowIf=false then
			'call echo(right(phptrim(rowc),1),"怪怪马")
			handleBR=""
			exit function
		end if
		
        'IF为真
        If ifJudge(getArrayIndex()) = True Then
            Call DelArrayVar()                                                              '删除数组变量
            'call echo("if为真" & ifJudgeRowStr(GetArrayIndex()), getArrayIndex())

        '判断IF是否为一行内容    '改进于20160405   对if为一行进制加强判断
        ElseIf isRowIf Then
            'call echo(ifJudgeRow(getArrayIndex()),ifJudgeRowStr(getArrayIndex()))
            ifJudgeRowStr(getArrayIndex()) = "" 
            ifJudgeRow(getArrayIndex()) = 0 
			'call echo(pubRowS,upword)
            varName = ";}" & dimRsStr 
            dimRsStr = ""                                                                   '清除定义RS内容
            Call DelArrayVar()                                                              '删除数组变量
        'For
        ElseIf forJudge = True Then
            If forEachName = "" Then
                varName = "; " & forFirstVarName & "" & forStep & "){"                          'forStep默认为++
            Else
                varName = varName & " as $key=>" & forEachName & "){" 
            End If 
            forJudge = False 

        'dim
        ElseIf isDimAraay(getArrayIndex()) = "定义变量" And getAfterStr(rowC) <> ";" Then
            varName = "='';" 
        'call eerr(GetAfterStr(rowC),rowC)
        ElseIf getAfterStr(rowC) <> ";" Then
            varName = ";" 
        End If 
        '移除内部定义函数
        If InStr("|dim|vbtab|vbcrlf|mod|response|request|call|left|right|rnd|mid|cint|instr|instrrev|split|", "|" & action & "|") > 0 Then
            isDimAraay(getArrayIndex()) = ""                                                '清除定义变量
            Call DelArrayVar()                                                              '删除数组变量
        'call echo("换行删除系统函数", getArrayIndex())
        End If 
		
		varName = rowEndStr & varName
		rowEndStr=""
				
        handleBR = varName 
        isOpenTable = False                                                             '打开数据库表为假


        'IF一行清零 必需放在最后，因为前面要处理
        ifJudgeRow(getArrayIndex()) = 0 
    End Function
 

    '处理PHP完整
    Function handlePHPCorrect(content)
        Dim splStr, i, s, c, s2, s3 
        splStr = Split(content, vbCrLf) 
        For Each s In splStr
            s = PHPRTrim(s) : s2 = trimVbCrlf(s) : s3 = phpTrim(Replace(s, " ", "")) 

            If s2 = ";" Then
                s = "" 
            ElseIf s3 = "<?PHP;" Then
                s = "<?PHP" 
            ElseIf s3 = "?><?PHP;" Then
                s = "?>" & vbCrLf & "<?PHP" 
			elseif instr(s,"<?PHP =")>0 and instr(s,"?") then
				s=replace(s,"<?PHP =","<?= ")
            Else
                '去掉{;}; 后台的;
                If Right(s, 1) = ";" And(InStr(s, "{") > 0 Or InStr(s, "}") > 0) Then
                    s2 = PHPRTrim(Left(s, Len(s) - 1)) 
                    If Right(s2, 1) = "{" Or Right(s2, 1) = "}" Then
                        s = Left(s, Len(s) - 1) 
                    End If 
                End If 
                '对 ;}  这种处理  把前面 ;  删除掉 (20151109)
                If Left(s3, 1) = ";" Then
                    s = Mid(s, InStr(s, ";") + 1) 
                End If 
            End If 
            c = c & s & vbCrLf 
        Next 
        handlePHPCorrect = c 
    End Function
 


    '检测是否为ASP系统函数  这个非常重要 对系统变量特殊处理
    Function checkAspSystemFunction(ByVal functionName)
        If InStr(systemFunctonList, "|" & LCase(functionName) & "|") > 0 Then
            checkAspSystemFunction = True 
        Else
            checkAspSystemFunction = False 
        End If 
    End Function
 


    '获得前一个字符 GetBeforeStr(EndCode)
    Function getBeforeStr(ByVal content)
        getBeforeStr = Left(trimVbCrlf(content), 1) 
    End Function
 

    '获得后一个字符
    Function getAfterStr(ByVal content)
        getAfterStr = Right(trimVbCrlf(content), 1) 
    End Function
 

    '获得前一个单词
    Function getBeforeWord(ByVal content)
        content = trimVbCrlf(content) 
        content = Replace(Replace(Replace(Replace(content, "(", " ( "), ",", "  , "), ")", " ) "), ".", " . ") 
        If InStr(content, " ") Then
            content = Mid(content, 1, InStr(content, " ") - 1) 
        End If 
        getBeforeWord = content 
    End Function
 
    '获得前前一个单词
    Function getBefore2Word(ByVal content)
        content = trimVbCrlf(content) 
        content = Replace(Replace(Replace(Replace(content, "(", " ( "), ",", "  , "), ")", " ) "), ".", " . ") 
        If InStr(content, " ") Then
            content = Mid(content, InStr(content, " ") + 1) 
        End If 
        If InStr(content, " ") Then
            content = Mid(content, 1, InStr(content, " ") - 1) 
        End If 
        getBefore2Word = content 
    End Function
 

    '获得后一个单词
    Function getAfterWord(ByVal content)
        content = trimVbCrlf(content) 
        If InStr(content, " ") Then
            content = Mid(content, InStr(content, " ") + 1) 
        End If 
        getAfterWord = content 
    End Function
 

    '追加数组变量
    Function addArrayVar(varName)
        Dim i 
        For i = 0 To UBound(arrayVar)
            If arrayVar(i) = "" Then
                arrayVar(i) = varName 
                Exit For 
            End If 
        Next 
    End Function
 

    '获得数组变量  从最后往前  (应该是从前向后)
    Function getArrayVar()
        Dim i 
        For i = 0 To UBound(arrayVar)
            If arrayVar(i) <> "" Then
                getArrayVar = arrayVar(i) 
            Else
                Exit For 
            End If 
        Next 
    End Function
 
    '获得数组索引
    Function getArrayIndex()
        Dim i 
        getArrayIndex = 0 
        For i = 0 To UBound(arrayVar)
            If arrayVar(i) <> "" Then
                getArrayIndex = i 
            Else
                Exit For 
            End If 
        Next 
    End Function
 

    '删除数组变量 从最后往前  (应该是从前向后)
    Function delArrayVar()
        Dim i 
        For i = 0 To UBound(arrayVar)
            If arrayVar(i) = "" Then
                If i > 0 Then
                    arrayVar(i - 1) = "" 
                End If 
                Exit For 
            End If 
        Next 
    End Function
 



    '追加回显 call addMsg("if",tempVarName,sType,varName)
    Sub addMsg(lableName, tempVarName, sType, varName)
        'getArrayIndex()替换成功 nThisIndex 显示精准
        tempActionList = tempActionList & nThisIndex & "、" & lableName & "(" & tempVarName & ")(" & sType & ")(" & varName & ")<br>" & vbCrLf 
    End Sub
 


    '内容红色
    Function textRed(content)
        textRed = "<font color=red>" & content & "</font>" 
    End Function
 
    '获得单词列表，从文本里提取
    Function getWordList()
        Dim c, splStr, i, s, splxx 
        c = getFText("\VB工程\Config\网站函数总杂205.txt") 
        splStr = Split(c, vbCrLf) 
        c = "|" 
        For Each s In splStr
            If InStr(s, ",") > 0 Then
                splxx = Split(s, ",") 
                c = c & splxx(1) & "|" 
            End If 
        Next 
        getWordList = c 
    End Function
	
	
	
	
	
	
	
	
	
	function saveHandleContent(c) 
		dim content,splstr,s
        Call createfile("1.txt", c) 
        c = FormattingJSPHP(c)                                                          '格式化PHP20150824
        c = Replace(c, "'[#晕这是一个双引号，显示就出错#]'", "'""'")               '还愿，真晕了20151112  程序写得太差了            人啊到原因了，是上面的那个格式化jsphp有问题
		
		'输入PHP系统函数
		if isOpenImportFunction=true then
			'20160115
			dim importFunctionTemplate,importFunctionContent,funStr
			'导入函数列表20160115
			if importFunctionList<>"" then
				importFunctionTemplate=getFText("\PHP2\ImageWaterMark\Include\ASP.php") 
				importFunctionTemplate=importFunctionTemplate & getFText("\PHP2\ImageWaterMark\Include\sys_Cai.Php") 
				importFunctionTemplate=importFunctionTemplate & getFText("\PHP2\ImageWaterMark\Include\sys_FSO.php") 
				importFunctionTemplate=importFunctionTemplate & getFText("\PHP2\ImageWaterMark\Include\sys_Print.php") 
				importFunctionTemplate=importFunctionTemplate & getFText("\PHP2\ImageWaterMark\Include\sys_System.php") 
				importFunctionTemplate=importFunctionTemplate & getFText("\PHP2\ImageWaterMark\Include\sys_Url.php") 
				splstr=split(importFunctionList,"|")
				for each s in splstr
					if s <>"" then
						'call echo( "function "& s &"(",s)
						funStr=getStrCut(importFunctionTemplate, "function "& s &"(", vbcrlf & "}" & vbcrlf, "1|lu")		'不分大小写
						if funStr<>"" then
						
						end if
						importFunctionContent=importFunctionContent & funStr
					end if
				next
			end if
			'加上gb2312头
			importFunctionContent="header(""Content-Type: text/html; charset=gb2312"");" & vbcrlf & importFunctionContent

			
			if request("language")="js" and isOpenImportFunction=true then
				c = "<script type=""text/javascript"" src=""http://sharembweb.com/asptophp/asp.js""></script>" & vbcrlf & c
			elseif importFunctionContent<>"" then
				importFunctionContent = vbcrlf & "<?PHP" & vbcrlf & "//系统函数列表" & vbcrlf & importFunctionContent & vbcrlf & "?>"				
				c=c & importFunctionContent
			end if
		end if
        saveHandleContent = c                                                                    '返回内容


        Dim tempC, tempContent                                                          '(20151119)
        If requestact = "web1asp" Then
            tempContent = getftext("/web/1.php") 
            tempContent = Mid(tempContent, 1, InStr(tempContent, "//=========") - 1) & Mid(c, InStr(c, "//=========")) 
            Call writeToFile("/web/1.php", tempContent,"gb2312")                                  '(20151119)
        elseIf requestact = "webfunction" Then
            Call writeToFile("/web/function.php", phptrim(c),"gb2312")                              '(20151119)
        elseIf requestact = "webfunction2" Then
            Call writeToFile("/web/function2.php", phptrim(c),"gb2312")                              '(20151119)
        elseIf requestact = "webtemplate" Then
            tempContent = getftext("/web/template.php")
			c=replace(c,"../aspweb.asp","../phpweb.php")
			tempContent = Mid(tempContent, 1, InStr(tempContent, "?>") + 1) & c
			 Call writeToFile("/web/template.php", tempContent,"gb2312")  
			
        ElseIf requestact = "aspweb" Then
            tempContent = getftext("/phpweb.php") 
            tempContent = Mid(tempContent, 1, InStr(tempContent, "//=========") - 1) & Mid(c, InStr(c, "//=========")) 
            Call writeToFile("/phpweb.php", tempContent,"gb2312") 
			
        ElseIf requestact = "config" Then
            tempContent = getftext("/web/config.php")  
            tempContent = Mid(tempContent, 1, InStr(tempContent, "//=========") - 1) & Mid(c, InStr(c, "//=========")) 
            Call writeToFile("/web/config.php", phptrim(tempContent),"gb2312") 
			

        ElseIf requestact = "setAccess" Then
            Call writeToFile("/web/setAccess.php", phptrim(c),"gb2312")
			
		ElseIf requestact = "js" Then
			c= "<script type=""text/javascript"" src=""ASP.js""></script>" & c
            Call writeToFile("/atemp.html", phptrim(c),"gb2312")
        End If

		
		content=getftext("atemp.php")
		if instr(content,"//=========")>0 then
			c=mid(content,1,instr(content,"//=========")+len("//=========")+1) & "?>" & c
		else			
			c = getFText("PHP2/ImageWaterMark/PHPAuxiliary.php") & c 
			c = Replace(c, "//PHP函数文件列表", getPHPFunctionList("PHP2/Web/Inc/")) 
			c = trimVbCrlf(c)                                                               '去除两边空格
			c = Replace(c, "//Php ", "")                                                    '替换里面的php设定好的配置20150914
		end if
		

		if checkfolder("./PHP2/")=false and checkfolder("./Public/")=true then
			c="<?PHP" & vbcrlf & replace(authorInfo(""),"'","//") & vbcrlf & "require_once './Public/ASP.php';" & vbcrlf & "require_once './Public/Conn.php';" & vbcrlf & "require_once './Public/MySqlClass.php';" & vbcrlf & "?>" & vbcrlf & c
		end if


        Call writeToFile("atemp.php", c,"gb2312")
	end function
	
 
    '获得php生成后的函数文件列表
    Function getPHPFunctionList(folderPath)
        Dim splStr, content, fileName, c 
        content = GetFileFolderList(folderPath, True, "php", "名称", "", "", "") 
        splStr = Split(content, vbCrLf) 
        For Each fileName In splStr
            If fileName <> "" And LCase(fileName) <> "config.php" And LCase(fileName) <> "create_html.php" And Left(fileName, 1) <> "#" Then
                c = c & "require_once 'PHP2/Web/Inc/" & fileName & "';" & vbCrLf 
            End If 
        Next 
        getPHPFunctionList = c 
    End Function
 
End Class 
%>    
