<% 
'Version1.0 20150902




'格式化JS或PHP程序
Function formattingJSPHP(ByVal content)
    Dim splStr, sx, tempSx, s, clearS, c, i, j, yesASP, nSYHCount, s1, s2, tempZc, rowC, tempRow, noteStr, downRow, downWord, nc, isLoop, beforeStr, afterStr, endCode,tempEndCode 
    Dim UpWord, UpWord2, yesOK 
    Dim wc                                                                          '输入文本存储内容
    Dim wcType                                                                      '输入文本类型，如 " 或 '
    Dim zc                                                                          '字母文件存储内容
    Dim tempS 
    Dim RSpaceStr, RSpaceNumb, addToRSpaceNumb                                      '行空格与空格数,追加行数
    Dim MultiLineComment                                                            '多行注释
    MultiLineComment = ""                                                           '多行注释默认为空
    RSpaceStr = "    " : RSpaceNumb = 0                                             '行空格值 与 空格数 设置初值
    nSYHCount = 0                                                                   '双引号默认为0
    Dim isAddToSYH                                                                  '是否累加双引号

    splStr = Split(content, vbCrLf)                                                 '分割行
    '循环分行
    For j = 0 To UBound(splStr)
        s = splStr(j) 
        s = phptrim(s) 
        s = Replace(Replace(s, Chr(10), ""), Chr(13), "") '奇怪为什么 s里会有 chr(10)与chr(13) 呢？
        clearS = PHPTrim(s)                                                             '清除两边空格与退格
        tempS = s                                                                       '暂存S
        rowC = "" : tempRow = ""                                                        '清空每行ASP代码和暂存完整行代码
        noteStr = ""                                                                    '默认注释代码为空
        downRow = ""                                                                    '下一行代码
        downWord = ""                                                                   '下一行单词
        addToRSpaceNumb = 0                                                             '清空追加行数
        If(j + 1) <= UBound(splStr) Then
            downRow = splStr(j + 1) 
            downWord = GetBeforeNStr(downRow, "全部", 1) 
        End If 
        nc = ""                                                                         '数字为空
        isLoop = True                                                                   '循环字符为真
        '循环每个字符
        For i = 1 To Len(s)
            sx = Mid(s, i, 1) : tempSx = sx 
            beforeStr = Right(Replace(Mid(s, 1, i - 1), " ", ""), 1)                        '上一个字符
            afterStr = Left(Replace(Mid(s, i + 1), " ", ""), 1)                             '下一个字符
            endCode = Mid(s, i + 1)                                                         '当前字符往后面代码 一行
            '输入文本
            If(sx = """" Or sx = "'" And wcType = "") Or sx = wcType Or wc <> "" Then
                isAddToSYH = True 
                '这是一种简单的方法，等完善(20150914)
                If isAddToSYH = True And beforeStr = "\" Then
					
                    If Len(wc) >=1 Then
                        If isStrTransferred(wc)=true Then		'为转义字符为真
							'call echo(wc,isStrTransferred(wc))
                            isAddToSYH = False 
                        End If 
                    Else
                        isAddToSYH = False 
                    End If 
					'call echo(wc,isAddToSYH)
                End If 
                If wc = "" Then
                    wcType = sx 
                End If 

                '双引号累加
                If sx = wcType And isAddToSYH = True Then nSYHCount = nSYHCount + 1             '排除上一个字符为\这个转义字符(20150914)


                '判断是否"在最后
                If nSYHCount Mod 2 = 0 And beforeStr <> "\" Then
                    If Mid(s, i + 1, 1) <> wcType Then
                        wc = wc & sx 
                        rowC = rowC & wc                      '行代码累加
						'call echo("wc",wc)
                        nSYHCount = 0 : wc = ""               '清除
                        wcType = "" 
                    Else
                        wc = wc & sx 
                    End If 
                Else
                    wc = wc & sx 
                End If

            '多行注释
            ElseIf(sx = "/" And afterStr = "*") Or MultiLineComment <> "" Then
                noteStr = Mid(s, i) 
                'call echo("多行注释",noteStr)
                If MultiLineComment <> "" Then MultiLineComment = MultiLineComment & vbCrLf 
                MultiLineComment = MultiLineComment & noteStr 
                If InStr(noteStr, "*/") > 0 Then
                    rowC = rowC & MultiLineComment 
                    MultiLineComment = "" 
                End If 
                Exit For 
            '注释则退出 单选注释
            ElseIf sx = "/" And afterStr = "/" Then
                noteStr = Mid(s, i) 
                rowC = rowC & noteStr 
                Exit For 

            '+-*\=&   字符特殊处理
            ElseIf InStr(".&=,+-*/:()><;", sx) > 0 And nc = "" Then
                If zc <> "" Then
                    tempZc = zc 
                    UpWord2 = UpWord                                                         '记录上上一个变量
                    UpWord = LCase(tempZc)                                                   '记录上一个变量名称   为小写
                    rowC = rowC & zc 
                    zc = ""                                                                  '清空字母内容
                End If
				'清除如if(1=1){;;;;;;;;}   20151224
				if sx=";" and instr("{;",right(trim(rowC),1))>0 then
					'call eerr(endcode,rowc)
					sx=""
				end if
                rowC = rowC & sx 
                UpWord2 = UpWord                                                                '记录上上一个变量
                UpWord = sx 
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

                    UpWord2 = UpWord                                                         '记录上上一个变量
                    UpWord = LCase(tempZc)                                                   '记录上一个变量名称   为小写
                    rowC = rowC & zc 
                    zc = ""                                                                  '清空字母内容
                End If 
            '为数字
            ElseIf checkNumber(sx) = True Or nc <> "" Then
                nc = nc & sx 
                If afterStr <> "." And checkNumber(afterStr) = False Then
                    rowC = rowC & nc 
                    nc = "" 
                End If 
            Else
                'JS或PHP程序{}处理
                If sx = "{" Then 		    	
					If phptrim(rowC) <> "" Then
						addToRSpaceNumb = 1
					Else
						RSpaceNumb = RSpaceNumb + 1 
					End If  
                ElseIf sx = "}" Then
                    If RSpaceNumb > 0 Then
                        If phptrim(rowC) <> "" Then							
							'删除后台注释20151224
							tempEndCode =endCode
							if instr(tempEndCode,"//")>0 then
								tempEndCode=mid(tempEndCode,1,instr(tempEndCode,"//")-1) 
							end if					
                            '修复if(a==b){}  方法		
                            If phptrim(tempEndCode) <> "" Then
                                addToRSpaceNumb = -1 
                            Else
                                addToRSpaceNumb = 0 
                            End If 
                        Else
                            RSpaceNumb = RSpaceNumb - 1 
                        End If 
                    End If  
                End If 

                yesOK = True 
                If sx = " " And i > 1 Then
                    If Mid(s, i - 1, 1) = " " Then
                        yesOK = False 
                    End If 
                End If 
                If yesOK = True Then
                    rowC = rowC & sx 
                'call echo("sx",sx)
                End If 
            End If 
            '暂存每行内容
            tempRow = tempRow & sx 
            If isLoop = False Then Exit For 
        Next 
        If rowC = ";}" Then
            rowC = "}" 
        End If 
        If rowC <> "" And RSpaceNumb > 0 Then
            rowC = CopyStrNumb(RSpaceStr, RSpaceNumb) & Replace(rowC, vbCrLf, vbCrLf & CopyStrNumb(RSpaceStr, RSpaceNumb)) '修进20150902
        End If 
        RSpaceNumb = RSpaceNumb + addToRSpaceNumb 
        If MultiLineComment = "" Then                                                   '修进20150902
            c = c & rowC & vbCrLf 
        End If 
    Next 
    formattingJSPHP = c 
End Function 





'生成 Install.php 文件
function printAccessToPHPInstallFile()
    Dim content, splTableName, tableName, splFieldName, fieldName, fieldType, c ,tempTableName
	dim tableNameList			'表名称列表
    Call openconn() 
    content = getTableList() 
    'call echo("content",content)
    splTableName = Split(content, vbCrLf) 
    For Each tableName In splTableName
		tempTableName=tableName
        tableName = LCase(tableName)                                                    '转成小写
		if instr(tableName,"_")>0 then
			tableName=mid(tableName,instr(tableName,"_")+1) 
		end if
		'排除重复表名称
		if instr("|"& tableNameList &"|","|"& tableName &"|")=false then
			content = getFieldList(tempTableName) 
			Call echo(tableName, content) 
			splFieldName = Split(content, ",") 
			c = c & """{$DB_PREFIX}" & tableName & """ => ""CREATE TABLE `{$DB_PREFIX}" & tableName & "` (" & vbCrLf 
			c = c & "`id` int(5) unsigned NOT NULL auto_increment," & vbCrLf 
			For Each fieldName In splFieldName
				If fieldName <> "" Then 
					fieldName = LCase(fieldName)                                                    '转成小写
					fieldType = getFieldAlt(tempTableName, fieldName) 
					'call echo(copyStr("&nbsp;",8) & "fieldName",fieldName   & "("& fieldType &")")        'mediumtext
					
					If fieldType = "VarChar" Then
						c = c & "`" & fieldName & "` varchar(255) NOT NULL default ''," & vbCrLf 
					ElseIf fieldType = "Int" Then
						c = c & "`" & fieldName & "` int(8) NOT NULL default '0'," & vbCrLf 
					ElseIf fieldType = "Float" Then
						c = c & "`" & fieldName & "` float(8) NOT NULL default '0'," & vbCrLf  
					ElseIf fieldType = "YesNo" Then
						c = c & "`" & fieldName & "` tinyint(1) NOT NULL default '0'," & vbCrLf  						
						
					'  InStr("|aboutcontent|bodycontent|content|template|", "|" & fieldName & "|") > 0 Or InStr(fieldName, "module") > 0
					ElseIf fieldType="Text" Then
						c = c & "`" & fieldName & "` mediumtext," & vbCrLf 
	
					Else
						c = c & "`" & fieldName & "` varchar(250) NOT NULL default ''," & vbCrLf 
					End If 
				End If 
			Next 
			c = c & "PRIMARY KEY  (`Id`)" & vbCrLf 
			c = c & "){$char};""," & vbCrLf 
		end if
    Next 
    Dim installFilePath, startStr, endStr, searchStr, replaceStr 
    startStr = "//access start" 
    endStr = "//access end" 
    installFilePath = "\PHP2\ImageWaterMark\Include\Install.php" 
    content = getftext(installFilePath) 
    searchStr = getStrCut(content, startStr, endStr, 1) 
    replaceStr = startStr & vbCrLf & c & endStr 
    content = Replace(content, searchStr, replaceStr) 
    'call eerr("",content)
    Call createfile(installFilePath, content) 
    Call echo("操作完成，打印数据库到php能生成文件", "<a href=""http://127.0.0.1/PHP2/ImageWaterMark/Include/Install.php?db_PREFIX=xy_"">http://127.0.0.1/PHP2/ImageWaterMark/Include/Install.php?db_PREFIX=xy_</a>") 
End function 
%> 
