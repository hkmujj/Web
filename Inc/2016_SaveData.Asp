<% 
'保存数据   ?act=saveData
Function saveData(sType)
	
	if session("yzm")="" then
		call eerr("提示","验证码失效")
	end if
	
	'if instr("|"& getFormFieldList() &"|","|yzm|") then		 
	if session("yzm")<>request.Form("yzm") then
		call eerr("提示","验证码错误")
	end if
	 

    '保存文章评论
    If sType = "articlecomment" Then 
		if request.Form("content")="" then
			call eerr("提示","评论为空")
		end if
        Call saveArticleComment(Request("itemid"), ADSql(Request("content")))
    ElseIf sType = "feedback" Then
		if request.Form("guestname")="" then
			call eerr("提示","姓名为空")
		end if
		call autoSavePostData("", "feedback", "isthrough|numb|0,ip||"& getip() &",columnid||" & request.QueryString("columnid"))
		Call echo("提示", "反馈提交成功，等待管理员审核") 
    ElseIf sType = "guestbook" Then
		
		if request.Form("guestname")="" then
			call eerr("提示","姓名为空")
		end if
	
		call autoSavePostData("", "guestbook", "isthrough|numb|0,ip||"& getip() &",columnid||" & request.QueryString("columnid"))
		Call echo("提示", "留言提交成功，等待管理员审核")
	
	 ElseIf sType = "articledetail" Then	
		call autoSavePostData("", "articledetail", "title|bodycontent,ip||"& getip())
		Call echo("提示", "文章提交成功")		
    End If 
	response.End()
End Function 
'保存文章评论 20160129
Function saveArticleComment(itemid, bodycontent) 
    conn.Execute("insert into " & db_PREFIX & "tablecomment (tableName,itemid,ip,bodycontent,adddatetime) values('ArticleDetail'," & itemid & ",'" & getIP() & "','" & bodycontent & "','" & Now() & "')") 
    Call echo("提示", "评论提交成功，等待管理员审核") 
End Function
'自动保存POST数据到表
function autoSavePostData(id, tableName, fieldNameList)
	dim sql 
	sql=getPostSql(id, tableName, fieldNameList)
    '检测SQL
    If checksql(sql) = False Then
        Call errorLog("出错提示：<hr>sql=" & sql & "<br>") 
        Exit Function 
    End If
	'conn.execute(sql)			'checksql这一步就已经执行了不需要再执行了20160410
end function
'获得Post发送表单处理SQL语句 20160309
Function getPostSql(id, tableName, fieldNameList)
    Dim valueStr, editValueStr, sql
    Dim splStr, splxx, s, fieldList 
    Dim fieldName, defaultFieldValue                                                '字段名称
    Dim fieldSetType                                                                '字段设置类型
    Dim fieldValue                                                                  '字段值

    Dim systemFieldList                                                             '表字段列表
    systemFieldList = getHandleFieldList(db_PREFIX & tableName, "字段配置列表") 

    Dim postFieldList                                                               'post字段列表
    Dim splPost, fieldContent, fieldConfig 
    postFieldList = getFormFieldList() 
    splPost = Split(postFieldList, "|") 
    For Each fieldName In splPost
        fieldContent = Request.Form(fieldName)  
        If InStr(systemFieldList, "," & fieldName & "|") > 0 And InStr("," & fieldList & ",", "," & fieldName & ",") = False Then
            '为自定义的
            If InStr(fieldNameList, "," & fieldName & "|") > 0 Then
                fieldConfig = Mid(fieldNameList, InStr(fieldNameList, "," & fieldName & "|") + 1) 
            Else
                fieldConfig = Mid(systemFieldList, InStr(systemFieldList, "," & fieldName & "|") + 1) 
            End If 
            fieldConfig = Mid(fieldConfig, 1, InStr(fieldConfig, ",") - 1) 
            'call echo("config",fieldConfig)
            'call echo(fieldName,fieldContent)

            splxx = Split(fieldConfig & "|||", "|") 
            fieldName = splxx(0)                                                            '字段名称
            fieldSetType = splxx(1)                                                         '字段设置类型
            defaultFieldValue = splxx(2)                                                    '默认字段值
            fieldValue = ADSqlRf(fieldName)                                                 '代替上面，因为它处理了'符号
            'md5加密
            If fieldSetType = "md5" Then
                fieldValue = myMD5(fieldValue) 
            End If 

            If fieldSetType = "yesno" Then
                If fieldValue = "" Then
                    fieldValue = defaultFieldValue 
                End If 
            '不为数字类型加单引号
            ElseIf fieldSetType = "numb" Then
                If fieldValue = "" Then
                    fieldValue = defaultFieldValue 
                End If 

            ElseIf fieldName = "flags" Then
                'PHP里用法
                If EDITORTYPE = "php" Then
                    If fieldValue <> "" Then
                        fieldValue = "|" & arrayToString(fieldValue, "|") 
                    End If 
                Else
                    fieldValue = "|" & arrayToString(Split(fieldValue, ", "), "|") 
                End If 


                fieldValue = "'" & fieldValue & "'" 

            '为时期
            ElseIf fieldSetType = "date" Then
                If fieldValue = "" Then
                    fieldValue = Date() 
                End If 

            Else
                fieldValue = "'" & fieldValue & "'" 
            End If 
			
			fieldValue=unescape(fieldValue)			'解码20160418
		
            If fieldList <> "" Then
                fieldList = fieldList & "," 
                valueStr = valueStr & "," 
                editValueStr = editValueStr & "," 
            End If 
            fieldList = fieldList & fieldName 
            valueStr = valueStr & fieldValue 
            editValueStr = editValueStr & fieldName & "=" & fieldValue 
        End If 
    Next 
    '自定义字段是否需要写入默认值
    splStr = Split(fieldNameList, ",") 
    For Each s In splStr
        If InStr(s, "|") > 0 Then
            splxx = Split(s & "|||", "|") 
            fieldName = splxx(0)                                                            '字段名称
            fieldSetType = splxx(1)                                                         '字段设置类型
            fieldValue = splxx(2)                                                           '默认字段值

            If InStr(systemFieldList, "," & fieldName & "|") > 0 And InStr("," & fieldList & ",", "," & fieldName & ",") = False Then
                If fieldSetType <> "yesno" And fieldSetType <> "numb" Then
                    fieldValue = "'" & fieldValue & "'" 
                End If 
                If fieldList <> "" Then
                    fieldList = fieldList & "," 
                    valueStr = valueStr & "," 
                    editValueStr = editValueStr & "," 
                End If 
                fieldList = fieldList & fieldName 
                valueStr = valueStr & fieldValue 
                editValueStr = editValueStr & fieldName & "=" & fieldValue 
            'call echo(fieldName,fieldSetType)
            End If 
        End If 
    Next 

    If id = "" Then
        sql = "insert into " & db_PREFIX & "" & tableName & " (" & fieldList & ",updatetime) values(" & valueStr & ",'" & Now() & "')" 
    Else
        sql = "update " & db_PREFIX & "" & tableName & " set " & editValueStr & ",updatetime='" & Now() & "' where id=" & id 
    End If 
    getPostSql = sql 
End Function 
%> 

