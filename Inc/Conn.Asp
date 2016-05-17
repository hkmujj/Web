<% 
Public MDBPath,AccessPass					'数据库路径与密码 
Dim DatabaseType						'数据库类型，Access与SqlServer之分
Dim SqlServerHostIP  					'SQL主机IP地址（服务器名）在2008版本使用
Dim SqlServerDatabaseName 				'SQL数据库名
Dim SqlServerUsername  					'SQL数据库用户名
Dim SqlServerPassword 					'SQL数据库用户密码 
 

'打开数据库
Public Sub OpenConn() 
	Dim ConnStr  
	If Request("MDBPath")<>"" And IsNul(Request("MDBPath"))=False Then
		MDBPath=HandlePath(Request("MDBPath"))
	End If
	'call eerr("",MDBPath)
	'连接MMD数据库
	If MDBPath<>"" Then		
    	If Instr(MDBPath, ":") = 0 Then MDBPath = Server.MapPath(Replace(MDBPath, "//", "/")) 
		Set Conn = CreateObject("Adodb.Connection") 
		Conn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database PassWord = '"& AccessPass &"';Data Source = "& MDBPath
	Else	
		'SqlServerDatabaseName,SqlServerUsername,SqlServerPassword,SqlServerHostIP
		If DatabaseType = "SqlServerWebData" Then
			SqlServerDatabaseName = "WebData" 'SQL数据库名
			SqlServerUsername = "sa" 'SQL数据库用户名
			SqlServerPassword = "sa" 'SQL数据库用户密码
			SqlServerHostIP = ".\SQLEXPRESS" 'SQL主机IP地址（服务器名）在2008版本使用 
			'call eerr("","是我")
						
		ElseIf DatabaseType = "SqlServerLocalData" Then
			SqlServerDatabaseName = "LocalData" 'SQL数据库名
			SqlServerUsername = "as" 'SQL数据库用户名
			SqlServerPassword = "as" 'SQL数据库用户密码
			SqlServerHostIP = "127.0.0.1,1433" 'SQL主机IP地址（服务器名）在2008版本使用
			
		ElseIf DatabaseType = "RemoteSqlServer" Then
			'远程SqlServer数据库
			SqlServerDatabaseName = "qds0140159_db"
			SqlServerUsername = "qds0140159": SqlServerPassword = "L4dN4eRd"
			SqlServerHostIP = "qds-014.hichina.com,1433"
			
		ElseIf DatabaseType = "RemoteSqlServer20150727" Then
			'远程SqlServer数据库
			SqlServerDatabaseName = "qds157513275_db"
			SqlServerUsername = "qds157513275": SqlServerPassword = "313801120"
			SqlServerHostIP = "qds157513275.my3w.com,1433"
			call echo("","11")

			
		ElseIf SqlServerDatabaseName<>"" Then 
			SqlServerHostIP = "127.0.0.1,1433" 'SQL主机IP地址（服务器名）在2008版本使用 
			'SqlServerHostIP ="(local),1433"  '不可以 
			
		End If 
		ConnStr = " Password = " & SqlServerPassword & "; user id =" & SqlServerUsername & "; Initial Catalog =" & SqlServerDatabaseName & "; data source =" & SqlServerHostIP & ";Provider = sqloledb;"
		'Call Echo("ConnStr", ConnStr):Doevents
		Set Conn = CreateObject("Adodb.Connection")
		 
		
		Conn.Open ConnStr 

	End If
	 
	'加个判断目的是为了重复设置，调用时会出错 如GetRecordCount(Table, AddSql)   20150313
	If TypeName(Rs) <> "Recordset" Then
		Set Rs = CreateObject("Adodb.RecordSet")
		Set Rsx = CreateObject("Adodb.RecordSet")
		Set Rss = CreateObject("Adodb.RecordSet")
		Set Rst = CreateObject("Adodb.Recordset")
		Set Rsd = CreateObject("Adodb.Recordset")
		Set TempRs = CreateObject("Adodb.RecordSet")
		Set TempRs2 = CreateObject("Adodb.RecordSet")
		Set RsTemp = CreateObject("Adodb.RecordSet")	
	End If 
End Sub
'关闭数据库
Function CloseConn()
	On Error Resume Next
    Conn.Close
    Set Conn = Nothing
End Function
'获得数据库句柄  ThisConn=GetAccessConn(ThisConn,ThisMDBPath,"")
Function GetAccessConn(AccessConn,ThisMDBPath,ThisAccessPass)
	ThisMDBPath=HandlePath(Request("ThisMDBPath"))
	Set AccessConn = CreateObject("Adodb.Connection")
	AccessConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database PassWord = '"& ThisAccessPass &"';Data Source = "& ThisMDBPath
	GetAccessConn = AccessConn
End Function
%>
