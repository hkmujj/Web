<%
'######################################################################
'���ñ���20150724  ����������Ȼ�죬��������Ժ��Լ���д
'######################################################################
'Dim obj : Set obj = New Cls_AB_JSON
'call rw(obj.toJSON)
Class Cls_AB_JSON
	Public Collection, Count, QuotedVars, Kind, StrEncode
	Private s_charset

	Private Sub Class_Initialize
		 '��ʱ��ɾ����һЩ����
	End Sub

	Private Sub Class_Terminate
		Set Collection = Nothing
	End Sub

	Public Property Let CharSet(ByVal s)
		s_charset = Ucase(s)
	End Property
	Public Property Get CharSet()
		CharSet = s_charset
	End Property

	'@ *****************************************************************************************
	'@ ��  ��:  AB.Json.Kind ���� (�ɶ�/��д)
	'@ ��  ��:  Integer (����) ����0����ʾ��Json���󣻷���1����ʾ�����顣
	'@ ��  ��:  �����½�Json��������ͣ������Կɶ���д
	'@ 			�����Կ������úͻ�ȡ�½�Json��������ͣ�0��ʾ��Json����1��ʾ�����顣
	'@ 			�����ʹ�� AB.Json.New ����������AspBox Json����������ָ�������ԡ�
	'==DESC=====================================================================================
	'@ ����: ��
	'==DEMO=====================================================================================
	'@ ��
	'@ *****************************************************************************************

	'@ *****************************************************************************************
	'@ ��  ��:  AB.Json.QuotedVars ���� (�ɶ�/��д)
	'@ ��  ��:  AB.Json.QuotedVars [ = boolean]
	'@ ��  ��:  Boolean (����ֵ) ��ǰ���õ����JSON�ַ����Ƿ�ʹ�����Ű�����������
	'@ ��  ��:  �����Ƿ��������JSON�ַ�����������ʹ�����ţ������Կɶ���д
	'@ 			��������˴�����Ϊ��(True)���������JSON�ַ����е����ƽ��������ţ����򲻰������š�������Ĭ��Ϊ��(True)��
	'==DESC=====================================================================================
	'@ ���� boolean: Ϊ True ���ʾ���JSON�ַ��������������Ű�����Ϊ False ���ʾ�������š�Ĭ��Ϊ True
	'==DEMO=====================================================================================
	'@ �ڲ����ô����Ի������ô�����Ϊ��(True)������£������JSON�ַ���Ϊ���¸�ʽ��
	'@ {"name":"Ray", "country":"China", "city":"Leshan"}
	'@ ��������˴�����Ϊ��(False)���������JSON�ַ���Ϊ���¸�ʽ��
	'@ {name:"Ray", country:"China", city:"Leshan"}
	'@ *****************************************************************************************

	'@ *****************************************************************************************
	'@ ��  ��:  AB.Json.StrEncode ���� (�ɶ�/��д)
	'@ ��  ��:  AB.Json.StrEncode [ = boolean]
	'@ ��  ��:  Boolean (����ֵ) ��ǰ�Ƿ����������Json�ַ���Ҫ��������
	'@ ��  ��:  ���������Json�ַ����Ƿ�������ģ������Կɶ���д
	'@ 			�����Կ������������Json�ַ����Ƿ�������ĵȶ��ֽ��ַ���
	'@ 			�������javascript��Ϊ�˱���������룬ͨ����Ҫ����ģ�
	'@ 			�������������ĳ��ϣ��������ڴ���Flash�е�ActionScript������Ҫ�����������ַ���
	'@ 			������Ĭ��Ϊ��(True)����Ҫ������ֽ��ַ���
	'==DESC=====================================================================================
	'@ ����: ��
	'==DEMO=====================================================================================
	'@ ��
	'@ *****************************************************************************************

	Public Function [New](ByVal k)
		Set [New] = New Cls_AB_JSON
		Select Case LCase(k)
			Case "0", "object" [New].Kind = 0
			Case "1", "array"  [New].Kind = 1
		End Select
	End Function

	Private Property Get Counter
		Counter = Count
		Count = Count + 1
	End Property

	'@ *****************************************************************************************
	'@ ��  ��:  AB.Json.Pair ���� (�ɶ�/��д)
	'@ ��  ��:  AB.Json.Pair(name)[ = value] ���� AB.Json(name) [ = value]
	'@ ��  ��:  Object (ASP����) �� String (�ַ���) ������Ӧ��ֵ�д����ֵ
	'@ ��  ��:  ����Json�����еĶ�ֵ�������Կɶ���д
	'@ 			��������Json�����Ҫ���������Ƕ�Json���ݸ�ֵ�����ԣ�
	'@ 			��ֵʱ��ֵ�������ַ�����Ҳ��������һ��AspBox Json����
	'@ 			��������ʵ�ʱ�д�Ĺ����п���ʡ��Ϊ AB.Json(name) ��
	'==DESC=====================================================================================
	'@ ����: ��
	'==DEMO=====================================================================================
	'@ Dim jso : Set jso = ab.json.jsObject()
	'@ jso("a") = "x" : jso("b") = "y" : jso("c") = "z"
	'@ PrintCn jso.JsString '��������{"a":"x","b":"y","c":"z"}
	'@ PrintCn jso.Pair("b") '��������y
	'@ jso.Pair("b") = "-"
	'@ PrintCn jso.JsString '��������{"a":"x","b":"-","c":"z"}
	'@ *****************************************************************************************

	Public Property Let Pair(Byval p, Byval v)
		If IsNull(p) Then p = Counter
		If vartype(v) = 9 Then
			If TypeName(v) = "Cls_AB_JSON" Then
				Set Collection(p) = v
			Else
				Collection(p) = v
			End If
		Else
			Collection(p) = v
		End If
	End Property

	Public Default Property Get Pair(Byval p)
		If IsNull(p) Then p = Count - 1
		If IsObject(Collection(p)) Then
			Set Pair = Collection(p)
		Else
			Pair = Collection(p)
		End If
	End Property

	'@ *****************************************************************************************
	'@ ������:  AB.Json.Clean ����
	'@ ��  ��:  �޷���ֵ
	'@ ��  ��:  ɾ��Json�����ȫ����ֵ�����ô˷�������ɾ��AspBox Json�����е�ȫ����ֵ��
	'==DESC=====================================================================================
	'@ ����: ��
	'==DEMO=====================================================================================
	'@ Dim jso : Set jso = ab.json.jsObject()
	'@ jso("a") = "x" : jso("b") = "y" : jso("c") = "z"
	'@ PrintCn jso.JsString '��������{"a":"x","b":"y","c":"z"}
	'@ jso.Clean
	'@ PrintCn jso.JsString '��������{}
	'@ *****************************************************************************************

	Public Sub Clean
		Collection.RemoveAll
	End Sub

	'@ *****************************************************************************************
	'@ ������:  AB.Json.Remove name
	'@ ��  ��:  �޷���ֵ
	'@ ��  ��:  ɾ��Json��ֵ�е�һ����ô˷�������ɾ��AspBox Json�����е�ĳһ���ֵ��
	'==DESC=====================================================================================
	'@ ���� name : ɾ��Json�����е�ĳһ���ֵ
	'==DEMO=====================================================================================
	'@ Dim jso : Set jso = ab.json.jsObject()
	'@ jso("a") = "x" : jso("b") = "y" : jso("c") = "z"
	'@ PrintCn jso.JsString '��������{"a":"x","b":"y","c":"z"}
	'@ jso.Remove "b"
	'@ PrintCn jso.JsString '��������{"a":"x","c":"z"}
	'@ *****************************************************************************************

	Public Sub Remove(Byval name)
		If Collection.Exists(name) Then Collection.Remove name
	End Sub

	'@ *****************************************************************************************
	'@ ������:  AB.Json.toJson Object
	'@ ��  ��:  AB.Json.Stringify Object
	'@ ��  ��:  String (�ַ���) ��ֱ������Json��ֵ�е�ֵ���ַ���
	'@ ��  ��:  ��Ŀ��ת��ΪJson��ֵ�е�ֵ��
	'@ 			���ô˷������԰Ѹ������͵�����ת��Ϊ��ֱ������Json��ֵ�е�ֵ���ַ�����
	'@ 			Ŀ��������ַ��������顢�ֵ����Rs��¼���ȡ�
	'==DESC=====================================================================================
	'@ ���� o : �������ַ��������飬����
	'==DEMO=====================================================================================
	'@ Dim a(1,1) '����
	'@ a(0,0) = "zero - zero"
	'@ a(0,1) = "zero - one"
	'@ a(1,0) = "one - zero"
	'@ a(1,1) = "one - one"
	'@ PrintCn AB.Json.toJSON(a) '��������[["zero - zero","zero - one"],["one - zero","one - one"]]
	'@ Dim o_json : Set o_json = AB.Json.New(0) '�ɲ���Json����
	'@ o_json("name") = "Booty Music" : o_json("artist") = "Deepside"
	'@ PrintCn AB.Json.toJSON(o_json) '��������{"name":"Booty Music","artist":"Deepside"}
	'@ 'Rem �߼��﷨:
	'@ PrintCn AB.Json.toJSON(AB.Dict) '���Dictionary�ֵ����
	'@ PrintCn AB.Json.toJSON(AB.db.ExeC("SELECT id,name FROM [Table] WHERE id<=2")) '���Rs��¼������
	'@ -----------------------
	'@ Dim temp
	'@ temp = NULL '## NUll : null
	'@ temp = Date '## Date : "2012-2-4"
	'@ temp = "id=6" '## String : "id=6"
	'@ temp = Empty '## Empty : ""
	'@ temp = "" '## String : ""
	'@ temp = cdbl(1.2888) '## Double,Integer... : 1.2888
	'@ temp = array(1,2,"a",array("x","y")) '## Array :
	'@ temp = True  '## Boolean : true/false
	'@ Rem '## ==Object==
	'@ Set temp = Nothing '## Nothing : {}
	'@ Set temp = AB.Dict '## Dictionary : {"name":"Lajox","msg":"Hello World"}
	'@ Set temp = AB.db.ExeC("SELECT id,name FROM [LB_C_Media] WHERE id<=2") '## RecordSet : [{"id":1,"name":"The Dawn"},{"id":2,"name":"lubov"}]
	'@ Rem '## --
	'@ Dim dict_1 : Set dict_1 = new(ab.dictName)
	'@ Dim dict_2 : Set dict_2 = new(ab.dictName)
	'@ dict_1("id") = 1 : dict_1("name") = "The Dawn"
	'@ dict_2("id") = 2 : dict_2("name") = "lubov"
	'@ temp = Array(dict_1, dict_2) '## �ۺ��� : [t,{"id":1,"name":"The Dawn"},{"id":2,"name":"lubov"}]
	'@ PrintCn AB.Json.toJSON(temp)
	'@ '---------------------
	'@ ab.use "sc"
	'@ Dim sc : Set sc = ab.sc.new
	'@ sc.Lang = "js"
	'@ sc.Add "function jsobj(){ var person = {name: {a:'zhangsan'}, pass: '123', fn: function(){alert(this.name.a+':'+this.pass);} }; return person; }"
	'@ sc.Add "function jsarr(){ return [1,5,8,9]; }"
	'@ Dim x : set x = sc.eval("jsobj()")
	'@ Dim y : set y = sc.eval("jsarr()")
	'@ ab.trace(x)
	'@ ab.trace(y)
	'@ PrintCn x.pass '�����123
	'@ ab.use "json"
	'@ PrintCn AB.Json.toJSON(x) '�����{"name":{"a":"zhangsan"},"pass":"123"}
	'@ PrintCn AB.Json.toJSON(y) '�����[1,5,8,9]
	'@ Dim str : str = "{name:'alonely', age:24, email:['test1@163.com','weed2@gmail.com'], family:{parents:['����','ĸ��']}}"
	'@ Dim z : set z = AB.Json.toObject(str)
	'@ PrintCn z.email.[0] '�����test1@163.com
	'@ PrintCn z.family.parents.[0] '���������
	'@ *****************************************************************************************

	Public Function toJSON(Byval o)
		'On Error Resume Next
		Dim s,i,j : i = 0 : j = 0
		Dim bFI, bFJ, oCount, jsLib, jso
		Dim tempJsDir : tempJsDir = "/" 'Ŀ¼����AspBox/jsLib/core/��
		Dim tempJsonJs : tempJsonJs = "json2.js"			'�����û�õģ���֪���������õ�
		Dim tempJsonJsPath : tempJsonJsPath = tempJsDir&tempJsonJs '�ļ�·������AspBox/jsLib/core/json2.js��
		'call echo(ubound(a),VarType(o))
		Select Case VarType(o)
			Case 1 'Null
				toJSON = "null"
			Case 7 'Date
				toJSON = """" & JsEncode__(CStr(o)) & """"
			Case 8 'String 
				Select Case TypeName(o)
					Case "IRequest", "IRequestDictionary" 'Request.QueryString/Request.Form/Request.Cookies
						If o = Request.QueryString Or o = Request.Form Or o = Request.Cookies Then
							bFI = True
							toJSON = toJSON & "{"
							For Each i In o
								If bFI Then bFI = False Else toJSON = toJSON & ","
								toJSON = toJSON & IIF(QuotedVars, """" & IIF(StrEncode,JSEncode(i),JSEncode__(i)) & """", JsEncode__(i)) & ":" & toJSON(o(i))
							Next
							toJSON = toJSON & "}"
						End If
					Case "JScriptTypeInfo"
						If isFile(tempJsonJsPath) Then
							AB.Use "jsLib"
							Set jsLib = AB.jsLib.New
							jsLib.BasePath = tempJsDir
							jsLib.Inc(tempJsonJs)
							Set jso = jsLib.Object
							toJSON = jso.toString(o)
							Set jso = Nothing
							Set jsLib = Nothing
						Else
							toJSON = o
						End If
					Case Else
						If o="" Then
							toJSON = """"""
						Else 
							toJSON = """" & IIF(StrEncode,JSEncode(o),JSEncode__(o)) & """"
						End If
				End Select
			Case 9 'Object
				Select Case TypeName(o)
					Case "Nothing","Empty"
						toJSON = "{}"
					Case "Recordset"
						If o.State = 0 Then '�˼�¼�������ѹر�
							toJSON = "[]"
						Else
							If IsNul(o) Then '�˼�¼������Ϊ�ռ�¼����û������
								toJSON = "[]"
							Else
								On Error Resume Next
								Set o = o.Clone
								On Error Goto 0
								If o.RecordCount = 1 Then '������¼��,��¼�ֶ���,�ֶ�ֵ
									bFI = True
									toJSON = toJSON & "[{"
									For j = 0 To o.Fields.Count-1
										If bFI Then bFI = False Else toJSON = toJSON & ","
										If QuotedVars Then
											toJSON = toJSON & """" & IIF(StrEncode,JSEncode(lcase(o.Fields(j).Name)),JSEncode__(lcase(o.Fields(j).Name))) & """:" & toJSON(o.Fields(j).Value)
										Else
											toJSON = toJSON & IIF(StrEncode,JSEncode(lcase(o.Fields(j).Name)),JSEncode__(lcase(o.Fields(j).Name))) & ":" & toJSON(o.Fields(j).Value)
										End If
									Next
									toJSON = toJSON & "}]"
								Else
									bFJ = True
									toJSON = toJSON & "["
									o.MoveFirst
									Do While Not o.Eof
										If bFJ Then bFJ = False Else toJSON = toJSON & ","
										bFI = True
										toJSON = toJSON & "{"
										For j = 0 To o.Fields.Count-1
											If bFI Then bFI = False Else toJSON = toJSON & ","
											If QuotedVars Then
												toJSON = toJSON & """" & IIF(StrEncode,JSEncode(lcase(o.Fields(j).Name)),JSEncode__(lcase(o.Fields(j).Name))) & """:" & toJSON(o.Fields(j).Value)
											Else
												toJSON = toJSON & IIF(StrEncode,JSEncode(lcase(o.Fields(j).Name)),JSEncode__(lcase(o.Fields(j).Name))) & ":" & toJSON(o.Fields(j).Value)
											End If
										Next
										toJSON = toJSON & "}"
										i = i + 1
										o.MoveNext
									Loop
									toJSON = toJSON & "]"
								End If
							End If
						End If
					Case "Dictionary", "IRequest", "IRequestDictionary", "IApplicationObject", "ISessionObject"
						If TypeName(o) = "IRequest" Then
							Dim req : Set req = Server.CreateObject(AB.dictName)
							req.add "clientcertificate", o.ClientCertificate
							req.add "cookies", o.cookies
							req.add "form", o.form
							req.add "querystring", o.queryString
							req.add "servervariables", o.serverVariables
							req.add "totalbytes", o.totalBytes
							toJSON = toJSON(req)
						ElseIf TypeName(o) = "IApplicationObject" Or TypeName(o) = "ISessionObject" Then
							oCount = o.Contents.Count
							If oCount = 0 Then '�˶����ǿյģ���û���κμ�ֵ
								toJSON = "{}"
							Else 'Application����Session�����
								bFI = True
								toJSON = toJSON & "{"
								For Each i In o.Contents
									If bFI Then bFI = False Else toJSON = toJSON & ","
									toJSON = toJSON & IIF(QuotedVars, """" & IIF(StrEncode,JSEncode(i),JSEncode__(i)) & """", IIF(StrEncode,JSEncode(i),JSEncode__(i))) & ":" & toJSON(o(i))
								Next
								toJSON = toJSON & "}"
							End If
						ElseIf TypeName(o) = "Dictionary" Then
							oCount = o.Count
							If oCount = 0 Then '�˶����ǿյģ���û���κμ�ֵ
								toJSON = "{}"
							Else 'Dictionary����
								bFI = True
								toJSON = toJSON & "{"
								For Each i In o
									If bFI Then bFI = False Else toJSON = toJSON & ","
									toJSON = toJSON & IIF(QuotedVars, """" & IIF(StrEncode,JSEncode(i),JSEncode__(i)) & """", IIF(StrEncode,JSEncode(i),JSEncode__(i))) & ":" & toJSON(o(i))
								Next
								toJSON = toJSON & "}"
							End If
						Else
							bFI = True
							toJSON = toJSON & "{"
							For Each i In o
								If bFI Then bFI = False Else toJSON = toJSON & ","
								toJSON = toJSON & IIF(QuotedVars, """" & IIF(StrEncode,JSEncode(i),JSEncode__(i)) & """", IIF(StrEncode,JSEncode(i),JSEncode__(i))) & ":" & toJSON(o(i))
							Next
							toJSON = toJSON & "}"
						End If
					Case "Cls_AB_JSON"
						bFI = True
						toJSON = toJSON & IIF(o.Kind, "[", "{")
						For Each i In o.Collection
							If bFI Then bFI = False Else toJSON = toJSON & ","
							toJSON = toJSON & IfThen(o.Kind=0, IIF(QuotedVars, """" & IIF(StrEncode,JSEncode(i),JSEncode__(i)) & """", IIF(StrEncode,JSEncode(i),JSEncode__(i))) & ":") & toJSON(o(i))
						Next
						toJSON = toJSON & IIF(o.Kind, "]", "}")
					Case "JScriptTypeInfo"
						If isFile(tempJsonJsPath) Then
							AB.Use "jsLib"
							Set jsLib = AB.jsLib.New
							jsLib.BasePath = tempJsDir
							jsLib.Inc(tempJsonJs)
							Set jso = jsLib.Object
							toJSON = jso.toString(o)
							Set jso = Nothing
							Set jsLib = Nothing
						Else
							toJSON = "{}"
						End If
					Case "VBScriptTypeInfo"
						toJSON = "{}"
					Case Else 'e.g. TypeName(Server) = "IServer"
						toJSON = "{}"
				End Select
			Case 11 'Boolean
				toJSON = IIF(o, "true", "false")
			Case 12, 8192, 8204, 8209 'Array
				toJSON = RenderArray(o, 1, "")
			Case 2, 3, 17, 19  'int, long, byte
				toJSON = CLng(o)
			Case 4, 5, 6, 14  'single, double, currency
				toJSON = Replace(CDbl(o), ",", ".")
			Case Else
				toJSON = """" & IIF(StrEncode,JSEncode(Cstr(o)),JSEncode__(Cstr(o))) & """"
		End select
		'On Error Goto 0
	End Function
	
	'����תJSON�������Լ�������ֶ���ʾ 20150727
	function fieldToJson(rs,byval setFieldList)
		dim bFI, c,fieldName,fieldValue,j
		setFieldList = "|" & LCase(setFieldList) & "|"
		bFI = True
		c = c & "{"
		while not rs.eof
		For j = 0 To rs.Fields.Count-1
			fieldName = rs.Fields(j).Name
			fieldValue = rs.Fields(j).Value
			'�����ֶ������
			if instr(setFieldList,"|"& LCase(fieldName) &"|")>0 then
				If bFI Then bFI = False Else c = c & ","
				If QuotedVars Then
					c = c & """" & IIF(StrEncode,JSEncode(fieldName),JSEncode__(fieldName)) & """:" & toJSON(fieldValue)
				Else
					c = c & IIF(StrEncode,JSEncode(fieldName),JSEncode__(fieldName)) & ":" & toJSON(fieldValue)
				End If
			end if
		Next
		c = c & "}"
		rs.movenext:wend:rs.close
		c= "["& c &"]"		
		fieldToJson = c
	end function
	
	'ʹ��Ϊ���� 
	Private Function RenderArray(Byval arr, Byval depth, Byval parent)
		'call echo("arr",ubound(arr)):doevents
		Dim first : first = LBound(arr, depth)
		Dim last : last = UBound(arr, depth)
		Dim index, rendered
		Dim limiter : limiter = ","
		RenderArray = "["
		For index = first To last
			If index = last Then
				limiter = ""
			End If
			On Error Resume Next
			rendered = RenderArray(arr, depth + 1, parent & index & "," )
			If Err = 9 Then
				On Error GoTo 0
				RenderArray = RenderArray & toJSON(Eval("arr(" & parent & index & ")")) & limiter
			Else
				RenderArray = RenderArray & rendered & "" & limiter
			End If
		Next
		RenderArray = RenderArray & "]"
	End Function
	
	Public Function JsEncode(ByVal s)
		If isNul(s) Then JsEncode = "" : Exit Function
		Dim arr1, arr2, i, j, c, p, t
		arr1 = Array(&h27,&h22,&h5C,&h2F,&h08,&h0C,&h0A,&h0D,&h09)
		arr2 = Array(&h27,&h22,&h5C,&h2F,&h62,&h66,&h6E,&h72,&h749)
		For i = 1 To Len(s)
			p = True
			c = Mid(s, i, 1)
			For j = 0 To Ubound(arr1)
				If c = Chr(arr1(j)) Then
					t = t & "\" & Chr(arr2(j))
					p = False
					Exit For
				End If
			Next
			If p Then
				Dim a
				a = AscW(c)
				If a > 31 And a < 127 Then
					t = t & c
				ElseIf a > -1 Or a < 65535 Then
					t = t & "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)
				End If
			End If
		Next
		JsEncode = t
		'call echo("JsEncode s",s & "t=" & t)
	End Function
	Private Function JsEncode__(ByVal s)
		'call echo("JsEncode__",s)
		If isNul(s) Then JsEncode__ = "" : Exit Function
		Dim arr1, arr2, i, j, c, p, t
		arr1 = Array(&h27,&h22,&h5C,&h2F,&h08,&h0C,&h0A,&h0D,&h09)
		arr2 = Array(&h27,&h22,&h5C,&h2F,&h62,&h66,&h6E,&h72,&h749)
		For i = 1 To Len(s)
			p = True
			c = Mid(s, i, 1)
			For j = 0 To Ubound(arr1)
				If c = Chr(arr1(j)) Then
					t = t & "\" & Chr(arr2(j))
					p = False
					Exit For
				End If
			Next
			If p Then t = t & c
		Next
		JsEncode__ = t
	End Function
	
	Public Function Stringify(Byval o)
		Stringify = toJSON(o)
	End Function

	'@ *****************************************************************************************
	'@ ������:  AB.Json.JsString ����
	'@ ��  ��:  String (�ַ���) Json��ʽ�ַ���
	'@ ��  ��:  �õ�ǰJson��������Json�ַ���
	'@ 			���ô˷������԰ѵ�ǰAspBox Json����ֱ����Json��ʽת��Ϊ�ַ�����
	'==DESC=====================================================================================
	'@ ����: ��
	'==DEMO=====================================================================================
	'@ Dim jsa : Set jsa = ab.json.jsArray()
	'@ jsa(0) = "a" : jsa(1) = "b" : jsa(2) = "c"
	'@ PrintCn jsa.JsString '��������["a","b","c"]
	'@ PrintCn AB.Json.toStr(jsa) '��������["a","b","c"]
	'@ *****************************************************************************************

	Public Property Get jsString
		jsString = toJSON(Me)
	End Property

	'@ *****************************************************************************************
	'@ ������:  AB.Json.toObject(s) ����
	'@ ��  ��:  AB.Json.Parse(s)
	'@ ��  ��:  ASPʵ�壬�����Ƕ���Ҳ��������������
	'@ ��  ��:  ���ַ���ת��ΪASPʵ��, json�ַ�����תΪJS Object����
	'==DESC=====================================================================================
	'@ ���� s : �ַ���
	'==DEMO=====================================================================================
	'# Dim strTest
	'# 'strTest = "{name:'alonely', age:24, email:['test1@163.com','weed2@gmail.com'], family:{parents:['����','ĸ��']}, func:function(){} }"
	'# 'strTest = "{name:""alonely"", age:24, email:[""test1@163.com"",""weed2@gmail.com""], family:{parents:[""����"",""ĸ��""],toString:function(){return ""��ͥ��Ա"";}}}"
	'# strTest = "{name:'alonely', age:24, email:['test1@163.com','weed2@gmail.com'], family:{parents:['����','ĸ��']}}"
	'# Dim objTest
	'# AB.Use "json"
	'# Set objTest = AB.Json.toObject(strTest)
	'# AB.Trace objTest
	'# PrintCn objTest.name & "���ʼ���ַ��" & objTest.email.[0] '�����alonely���ʼ���ַ��ycplxl1314@163.com
	'# PrintCn "�����ʼ���ַ" & objTest.email.length & "��" '����������ʼ���ַ2��
	'# PrintCn objTest.family.parents.[0] '���������
	'# Dim father
	'# AB.Json.setJSArrayItem father, objTest.family.parents, 0
	'# PrintCn father '���������
	'# PrintCn ab.json.getJSArrayItem(objTest.family.parents, 0) '���������
	'@ *****************************************************************************************

	Public Function toObject(Byval o)
		On Error Resume Next
		If Err Then Err.Clear
		Dim i,j : i = 0 : j = 0
		Dim Sc, jsLib, jso, temp
		Dim tempJsDir : tempJsDir = AB.BasePath & "jsLib/core/" 'Ŀ¼����AspBox/jsLib/core/��
		Dim tempJsonJs : tempJsonJs = "json2.js"
		Dim tempJsonJsPath : tempJsonJsPath = tempJsDir&tempJsonJs '�ļ�·������AspBox/jsLib/core/json2.js��
		Select Case VarType(o)
			Case 9 'vbObject
				Set toObject = o
			Case 8 'vbString
				Select Case TypeName(o)
					Case "JScriptTypeInfo"
						If isFile(tempJsonJsPath) Then
							AB.Use "jsLib"
							Set jsLib = AB.jsLib.New
							jsLib.BasePath = tempJsDir
							jsLib.Inc(tempJsonJs)
							Set jso = jsLib.Object
							If IsObject(jso.toObject(o)) Then Set toObject = jso.toObject(o) Else toObject = jso.toObject(o)
							Set jso = Nothing
							Set jsLib = Nothing
						Else
							toObject = o
						End If
					Case "String"
						If LCase(o) = "null" Then 'null
							toObject = Null
						ElseIf LCase(o) = "empty" Then 'empty
							toObject = Empty
						ElseIf o = "" Or o = """""" Then 'empty
							toObject = ""
						ElseIf LCase(o) = "nothing" Then 'nothing
							Set toObject = Nothing
						ElseIf LCase(o) = "true" Or LCase(o) = "false" Then 'boolean
							toObject = CBool(o)
						ElseIf isInt(o) Then 'int, long, byte
							toObject = CLng(o)
						ElseIf Test(o,"double") Then  'single, double
							toObject = CDbl(o)
						ElseIf Test(o,"date") Then  'Date
							toObject = CDate(o)
						Else
							' If RegTest(Trim(o),"^\s*[(""']*[\{].*?[\}][""')]*\s*$") Or RegTest(Trim(o),"^\s*[(""']*[\[].*?[\]][""')]*\s*$") Then 'json format
								' o = RegReplace(Trim(o),"^\s*[(""']*([\{].*?[\}])[""')]*\s*$","$1")
								' o = RegReplace(Trim(o),"^\s*[(""']*([\[].*?[\]])[""')]*\s*$","$1")
								' If isFile(tempJsonJsPath) Then
									' AB.Use "jsLib"
									' Set jsLib = AB.jsLib.New
									' jsLib.BasePath = tempJsDir
									' jsLib.Inc(tempJsonJs)
									' Set jso = jsLib.Object
									' If IsObject(jso.toObject(o)) Then Set temp = jso.toObject(o) Else temp = jso.toObject(o)
									' Set jso = Nothing
									' Set jsLib = Nothing
								' Else
									' If IsObject(Me.jsEval(o)) Then Set temp = Me.jsEval(o) Else temp = Me.jsEval(o)
								' End If
								' If IsObject(temp) Then Set toObject = temp Else toObject = temp
							' End If
							'o = RegReplace(Trim(o),"^\s*[(""']*([\{].*?[\}])[""')]*\s*$","$1")
							'o = RegReplace(Trim(o),"^\s*[(""']*([\[].*?[\]])[""')]*\s*$","$1")
							AB.Use "Sc"
							Set Sc = AB.Sc.New
							Sc.Lang = "js"
							Sc.Add "var jsonObject = " & o & ";"
							If IsObject(Sc.CodeObject.jsonObject) Then Set toObject = Sc.CodeObject.jsonObject Else Set toObject = Nothing
						End If
					Case Else
						'toObject = o
						If IsObject(Me.jsEval(o)) Then Set temp = Me.jsEval(o) Else temp = Me.jsEval(o)
						If IsObject(temp) Then Set toObject = temp Else toObject = temp
				End Select
			Case Else
				'toObject = o
				If IsObject(Me.jsEval(o)) Then Set temp = Me.jsEval(o) Else temp = Me.jsEval(o)
				If IsObject(temp) Then Set toObject = temp Else toObject = temp
		End select
		If Err Then : toObject = o : Err.Clear : End If
		On Error Goto 0
	End Function


	Public Function Parse(Byval o)
		If IsObject(toObject(o)) Then Set Parse = toObject(o) Else Parse = toObject(o)
	End Function

	'@ *****************************************************************************************
	'@ ������:  AB.Json.setJSArrayItem(o, objJSArray, index) ����
	'@ ��  ��:  void
	'@ ��  ��:  ���ñ���oֵΪobjJSArray�����index������ֵ
	'==DESC=====================================================================================
	'@ ���� o 			[Any] 		: ����o
	'@ ���� objJSArray 	[Object] 	: JS Array ����
	'@ ���� index 		[Integer] 	: JS Array������
	'==DEMO=====================================================================================
	'# Dim strTest
	'# strTest = "{name:'alonely', age:24, email:['test1@163.com','weed2@gmail.com'], family:{parents:['����','ĸ��']}}"
	'# Dim objTest
	'# AB.Use "json"
	'# Set objTest = AB.Json.toObject(strTest)
	'# PrintCn objTest.family.parents.[0] '���������
	'# Dim father
	'# AB.Json.setJSArrayItem father, objTest.family.parents, 0
	'# PrintCn father '���������
	'# PrintCn ab.json.getJSArrayItem(objTest.family.parents, 0) '���������
	'@ *****************************************************************************************

	Public Sub setJSArrayItem(ByRef o, ByVal objJSArray, ByVal index)
		On Error Resume Next
		AB.Use "Sc"
		Dim Sc : Set Sc = AB.Sc.New
		Sc.Lang = "js"
		Sc.Add "var itemTemp=null; function setJSArray(arr, index){try { itemTemp = arr[index]; } catch(e) { itemTemp = null; }}"
		Sc.Run "setJSArray", Array(objJSArray, index)
		Set o = Sc.CodeObject.itemTemp
		If Err.number=0 Then
			Exit Sub
		End If
		o = Sc.CodeObject.itemTemp
		On Error Goto 0
	End Sub

	'@ *****************************************************************************************
	'@ ������:  AB.Json.getJSArrayItem(objJSArray, index) ����
	'@ ��  ��:  void
	'@ ��  ��:  ��ȡobjJSArray�����index������ֵ
	'==DESC=====================================================================================
	'@ ���� objJSArray 	[Object] 	: JS Array ����
	'@ ���� index 		[Integer] 	: JS Array������
	'==DEMO=====================================================================================
	'# Dim strTest
	'# strTest = "{name:'alonely', age:24, email:['test1@163.com','weed2@gmail.com'], family:{parents:['����','ĸ��']}}"
	'# Dim objTest
	'# AB.Use "json"
	'# Set objTest = AB.Json.toObject(strTest)
	'# PrintCn objTest.family.parents.[0] '���������
	'# PrintCn ab.json.getJSArrayItem(objTest.family.parents, 0) '���������
	'@ *****************************************************************************************

	Public Function getJSArrayItem(ByVal objJSArray, ByVal index)
		On Error Resume Next
		AB.Use "Sc"
		Dim Sc : Set Sc = AB.Sc.New
		Sc.Lang = "js"
		Sc.Add "function getJSArray(arr, index){var temp=null; try { temp = arr[index]; } catch(e) {} return temp;}"
		If IsObject(Sc.Run("getJSArray", Array(objJSArray, index))) Then
			Set getJSArrayItem = Sc.Run("getJSArray", Array(objJSArray, index))
		Else
			getJSArrayItem = Sc.Run("getJSArray", Array(objJSArray, index))
		End If
		On Error Goto 0
	End Function

	'@ *****************************************************************************************
	'@ ������:  AB.Json.jsEval(o) ����
	'@ ��  ��:  js����
	'@ ��  ��:  ���ַ���תΪjs����
	'==DESC=====================================================================================
	'@ ���� s : �ַ���
	'==DEMO=====================================================================================
	'# Dim str : str = "{name:'alonely', age:24, email:['test1@163.com','weed2@gmail.com'], family:{parents:['����','ĸ��']}}"
	'# Dim temp
	'# AB.Use "json"
	'# If IsObject(AB.Json.jsEval(str)) Then Set temp = AB.Json.jsEval(str) Else temp = AB.Json.jsEval(str)
	'# AB.Trace temp
	'@ *****************************************************************************************

	Public Function jsEval(Byval o)
		On Error Resume Next
		If Err Then Err.Clear
		Dim Sc, jsLib, jso, temp
		AB.Use "Sc"
		Set Sc = AB.Sc.New
		Sc.Lang = "js"
		Sc.Add "function toEval(str) { var _temp = null; try { eval('_temp = (' + str + ');'); } catch(e) { _temp = null; } return (_temp);}"
		Set jso = Sc.Object
		If IsObject(jso.toEval(o)) Then Set jsEval = jso.toEval(o) Else jsEval = jso.toEval(o)
		Set jso = Nothing
		Set Sc = Nothing
		If Err.Number <> 0 Then : Err.Clear : jsEval = o : End If
		On Error Goto 0
	End Function

	Public Function vbEval(Byval o)
		On Error Resume Next
		If Err Then Err.Clear
		If IsObject(Eval(o)) Then Set vbEval = Eval(o) Else vbEval = Eval(o)
		If Err.Number <> 0 Then : Err.Clear : vbEval = o : End If
		On Error Goto 0
	End Function

	'@ *****************************************************************************************
	'@ ������:  AB.Json.JSArrayToVBArray ����
	'@ ��  ��:  Array (VB����)
	'@ ��  ��:  ����ǰJS������� ת�� VB����
	'==DESC=====================================================================================
	'@ ���� jsArr: [Object] JS�������
	'==DEMO=====================================================================================
	'@ ab.use "sc"
	'@ Dim sc : Set sc = ab.sc.new
	'@ sc.Lang = "js"
	'@ sc.Add "function foo(){ return [1,5,8,9]; }"
	'@ Dim s,y,x
	'@ set x = sc.eval("foo();") 'js�������
	'@ 'Response.Write x.[2]
	'@ ab.use "json"
	'@ Dim temp : temp = AB.Json.JSArrayToVBArray(x) '����vb����
	'@ AB.Trace(temp)
	'@ *****************************************************************************************

	Public Function JSArrayToVBArray(ByVal jsArr)
		On Error Resume Next
		Dim temp, myVBArray
		If TypeName(jsArr) = "JScriptTypeInfo" Then
			temp = jsArr.join("{@Lx.Split@}") '�Ƚ� JSArray ת��Ϊ [�ַ���]
			myVBArray = Split(temp, "{@Lx.Split@}") '�ٽ� [�ַ���] ת��Ϊ VBArray
		End If
		JSArrayToVBArray = myVBArray
		On Error Goto 0
	End Function

	'@ *****************************************************************************************
	'@ ������:  AB.Json.VBArrayToJSArray ����
	'@ ��  ��:  Object (JS�������)
	'@ ��  ��:  ����ǰVB���� ת�� JS�������
	'==DESC=====================================================================================
	'@ ���� vbArr: [Array] VB����
	'==DEMO=====================================================================================
	'@ Dim arr
	'@ arr = Array(1,3,8,7)
	'@ AB.Use "json"
	'@ Dim temp : Set temp = AB.Json.VBArrayToJSArray(arr) '����js�������
	'@ AB.Trace(temp)
	'@ PrintCn AB.Json.toJSON(arr) '�����[1,3,8,7]
	'@ PrintCn AB.Json.toJSON(temp) '�����[1,3,8,7]
	'@ Dim obj : Set obj = AB.Json.toObject(AB.Json.toJSON(temp))
	'@ AB.Trace(obj)
	'@ PrintCn obj.[0]
	'@ *****************************************************************************************

	Public Function VBArrayToJSArray(ByVal vbArr)
		On Error Resume Next
		AB.Use "Sc"
		Dim objSC : Set objSC = AB.Sc.New
		objSC.Lang = "js"
		Dim objModule : Set objModule = objSC.Modules.Add("NewModule")
		objModule.AddCode "function jsVBArrayToJSArray(vbArr){ var temp = new VBArray(vbArr); return temp.toArray(); }"
		Dim objCodeObject : Set objCodeObject = objModule.CodeObject
		If IsObject(objCodeObject.jsVBArrayToJSArray(vbArr)) Then
			Set VBArrayToJSArray = objCodeObject.jsVBArrayToJSArray(vbArr)
		Else
			VBArrayToJSArray = objCodeObject.jsVBArrayToJSArray(vbArr)
		End If
		On Error Goto 0
	End Function

	'@ *****************************************************************************************
	'@ ������:  AB.Json.Flush ����
	'@ ��  ��:  Stream (������) Json�ļ�
	'@ ��  ��:  �õ�ǰJson��������������Json�ļ�
	'@ 			���ô˷��������õ�ǰJson�����������ֱ�����Json�ļ�������������ݵ�ContentType��"application/json"��
	'==DESC=====================================================================================
	'@ ����: ��
	'==DEMO=====================================================================================
	'@ AB.Json.Flush
	'@ *****************************************************************************************

	Public Sub Flush
		Response.Clear()
		Response.Charset = s_charset
		Response.ContentType = "application/json"
		NoCache()
		If TypeName(Response) <> "Empty" Then
			Put jsString
		ElseIf WScript <> Empty Then
			WScript.Echo(jsString)
		End If
	End Sub

	'@ *****************************************************************************************
	'@ ������:  AB.Json.Clone ����
	'@ ��  ��:  AspBox Json����
	'@ ��  ��:  ���Ƶ�ǰJson����Ϊ�µ�AspBox Json����
	'@ 			���ô˷������Ը��Ƶ�ǰJson����Ϊ�µ�AspBox Json����ԭ����Json������Ӱ�졣
	'==DESC=====================================================================================
	'@ ����: ��
	'==DEMO=====================================================================================
	'@ Dim jsa : Set jsa = ab.json.jsArray()
	'@ jsa(0) = "a" : jsa(1) = "b" : jsa(2) = "c"
	'@ Dim jsb : Set jsb = jsa.Clone
	'@ jsb(3) = "d" : jsb(4) = "e"
	'@ PrintCn AB.Json.toStr(jsa) '��������["a","b","c"]
	'@ PrintCn AB.Json.toStr(jsb) '��������["a","b","c","d","e"]
	'@ *****************************************************************************************

	Public Function Clone
		Set Clone = ColClone(Me)
	End Function

	'@ *****************************************************************************************
	'@ ������:  AB.Json.jsObject ����
	'@ ��  ��: 	Object (����)
	'@ ��  ��:  ����Json����
	'==DESC=====================================================================================
	'@ ����: ��
	'==DEMO=====================================================================================
	'@ Dim member
	'@ Set member = ab.json.jsObject() '��ͬ�ڣ� Set member = AB.Json.New(0)
	'@ member("name") = "Turul"
	'@ member("surname") = "Topuz"
	'@ member("message") = "Hello World"
	'@ member.Flush '��������{"name":"Tu\u011Frul","surname":"Topuz","message":"Hello World"}
	'@ *****************************************************************************************

	Public Function jsObject
		Set jsObject = AB.Json.New(0)
	End Function

	'@ *****************************************************************************************
	'@ ������:  AB.Json.jsArray ����
	'@ ��  ��: 	Array (����)
	'@ ��  ��:  ����Json����
	'==DESC=====================================================================================
	'@ ����: ��
	'==DEMO=====================================================================================
	'@ Dim jsa : Set jsa = ab.json.jsArray() '��ͬ�ڣ� Set jsa = AB.Json.New(1)
	'@ jsa(0) = "a" : jsa(1) = "b" : jsa(2) = "c"
	'@ jsa.Flush '��������["a","b","c"]
	'@ *****************************************************************************************

	Public Function jsArray
		Set jsArray = AB.Json.New(1)
	End Function

	'@ *****************************************************************************************
	'@ ������:  AB.Json.toStr ����
	'@ ��  ��: 	String (�ַ���)
	'@ ��  ��:  �����������ͷ����ַ��� [��(js object)��(js array)��ʽ]
	'==DESC=====================================================================================
	'@ ���� o: String/Object/Array �ַ���/����(Json�����ֵ���󡢼�¼�������)/����
	'==DEMO=====================================================================================
	'@ PrintCn AB.Json.toStr("hello world")
	'@ PrintCn AB.Json.toStr(Array("a","b","c"))
	'@ PrintCn AB.Json.toStr(AB.Dict)
	'@ PrintCn AB.Json.toStr(AB.db.ExeC("SELECT id,name FROM [Table] WHERE id<=2"))
	'@ *****************************************************************************************

	Public Function toStr(Byval o)
		toStr = AB.Json.toJSON(o)
	End Function

	'@ *****************************************************************************************
	'@ ������:  AB.Json.QueryToJSON ����
	'@ ��  ��:  AspBox Json����
	'@ ��  ��:  ��ѯ���ݿⲢ����Json����
	'==DESC=====================================================================================
	'@ ���� dbconn: ���ݿ����Ӷ���
	'@ ���� sql: Ҫִ�е�SQL���
	'==DEMO=====================================================================================
	'@ Dim dbJson : Set dbJson = AB.Json.QueryToJSON(AB.db.Conn, "SELECT id,name FROM [Table] WHERE id <= 2")
	'@ dbJson.Flush
	'@ Put AB.Json.toJSON(dbJson) '������ʽ��[{"id":1,"name":"The Dawn"},{"id":2,"name":"lubov"}]
	'@ *****************************************************************************************

	Public Function QueryToJSON(Byval dbconn, Byval sql)
		On Error Resume Next
		Dim Rs, jsa, Col, tpobj
		Set jsa = jsArray()
		If Not IsObject(dbconn) Or TypeName(dbconn)<>"Connection" Then:Err.Clear:Set QueryToJSON=jsa:Exit Function:End If
		Set Rs = dbconn.Execute(sql)
		If Err.Number<>0 Then : Err.Clear : Set QueryToJSON = jsa : Exit Function : End If
		Dim i : i = 0
		Do While Not (Rs.EOF Or Rs.BOF)
			Set tpobj = jsObject()
			jsa(i) = tpobj
			For Each Col In Rs.Fields
				jsa(i)(Col.Name) = Col.Value
			Next
			i = i + 1
			Set tpobj = Nothing
			Rs.MoveNext
		Loop
		Rs.Close : Set Rs = Nothing
		Set QueryToJSON = jsa
		On Error Goto 0
	End Function



	Private Function ColClone(Byval core)
		Dim jsc, i
		Set jsc = new Cls_AB_JSON
		jsc.Kind = core.Kind
		For Each i In core.Collection
			If IsObject(core(i)) Then
				Set jsc(i) = ColClone(core(i))
			Else
				jsc(i) = core(i)
			End If
		Next
		Set ColClone = jsc
	End Function



	Private Function Escape__(Byval s)
		Dim cDoubleQuote, cRevSolidus, cSolidus
		cDoubleQuote = &h22
		cRevSolidus = &h5C
		cSolidus = &h2F
		Dim i, cur, temp
		For i = 1 To (Len(s))
			cur = Mid(s, i, 1)
			If AscW(cur) > &h00 and AscW(cur) < &h1F Then
				cur = EscapeQuence(cur)
			ElseIf AscW(cur) >= &hC280 and AscW(cur) <= &hC2BF Then
				cur = "\u00" + Right(padLeft(hex(AscW(cur) - &hC200), 2, 0), 2)
			ElseIf AscW(cur) >= &hC380 and AscW(cur) <= &hC3BF Then
				cur = "\u00" + Right(padLeft(hex(AscW(cur) - &hC2C0), 2, 0), 2)
			Else
				Select Case AscW(cur)
					Case cDoubleQuote: cur = EscapeQuence(cur)
					Case cRevSolidus: cur = EscapeQuence(cur)
					Case cSolidus: cur = EscapeQuence(cur)
				End Select
			End If
			temp = temp & cur
		Next
		Escape__ = temp
	End Function
	Private Function EscapeQuence(Byval digit)
		EscapeQuence = "\u00" + Right(padLeft(hex(AscW(digit)), 2, 0), 2)
	End Function
	Private Function padLeft(Byval value, Byval totalLength, Byval paddingChar)
		padLeft = Right(cloneIt(paddingChar, totalLength) & value, totalLength)
	End Function
	private Function cloneIt(Byval str, Byval n)
		Dim i:For i = 1 To n : cloneIt = cloneIt & str : Next
	End Function

End Class
%>