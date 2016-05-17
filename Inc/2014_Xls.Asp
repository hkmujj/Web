<%
'xml文件操作

'批量导入xls文件，并保存 Call BatchImportXlsFile("E:\E盘\WEB网站\网站案例\", "优秀网址大全2.txt")
Function BatchImportXlsFile(FolderPath,SaveXmlFilePath)
	Dim I,S,S2,Sql,connStr,TableList,SplTable,TableName,TableFieldList,SplField,FieldName,UrlList,Url
	Dim XlsFilePath,XlsFileList,SplXls
	 
	XlsFileList =  GetFileFolderList(FolderPath, True, "xls", "", "", "", "") 
	SplXls = Split(XlsFileList,vbCrlf)
	For Each XlsFilePath In SplXls
		connStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& XlsFilePath &";Extended Properties='Excel 8.0;HDR=No;IMEX=1';" 
		Call CloseConn()
		set conn=server.createobject("adodb.connection") 
		Set Rs = CreateObject("Adodb.RecordSet")
		Set Rsx = CreateObject("Adodb.RecordSet")
		Set Rss = CreateObject("Adodb.RecordSet")
		Set Rst = CreateObject("Adodb.Recordset")
		Set Rsd = CreateObject("Adodb.Recordset")
		Set TempRs = CreateObject("Adodb.RecordSet")
		Set TempRs2 = CreateObject("Adodb.RecordSet")
		Set RsTemp = CreateObject("Adodb.RecordSet")	
		'Call Eerr("connStr",connStr)
		Conn.Open  connStr
		
		TableList = GetTableList_Conn(Conn)
		SplTable = Split(TableList,vbCrlf)
		For Each TableName IN SplTable
			If TableName<>"" Then
				TableFieldList = GetFieldList(TableName)
				SplField = Split(TableFieldList,",")
				If Ubound(SplField)<>-1 Then
					'Call Echo(Ubound(SplField),TableFieldList):Doevents
					
					Rsx.Open"Select * From ["& TableName &"]",Conn,1,1
					'Call Echo("记录",Rsx.RecordCount)
					While Not Rsx.Eof
						I=0
						For Each FieldName In SplField
							If FieldName <>"" Then
								Url = Rsx(I)
								'Call Echo(FieldName & "("& CheckUrl(Rsx(I)) &")", Rsx(I) )
								If CheckUrl(Url) = True And InStr(vbCrlf & UrlList & vbCrlf, vbCrlf & Url & vbCrlf)=False Then
									UrlList = UrlList & Url & vbCrlf
									Call Echo(TableName,Url):Doevents
								End If
								I=I+1
							End If
						Next
					Rsx.MoveNext:Wend:Rsx.Close
				End If
			End If
		Next
	Next
	UrlList=CStr(UrlList)
	Call Echo("UrlList",UrlList)
	Call WriteToFile(SaveXmlFilePath,UrlList,1)
	'Call CreateFile("优秀网址大全.txt",UrlList)
End Function
%>