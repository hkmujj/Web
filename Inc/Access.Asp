<% 
'Access (2013,10,22)   ��Ҫ�����£�̫��̫��(20151117)


'===================================== SQl�������� =====================================





'������Access���ݿ�ֵ (20151116)  �÷� conn.execute("update product set content='"& AD(c) &"'")
Function handleAccessData(ByVal valueStr)
    handleAccessData = Replace(valueStr, "'", "''") 
End Function 
'������Access���ݿ�ֵ   ��������
Function aDSql(valueStr)
    aDSql = handleAccessData(valueStr) 
End Function 
'�ж����ݿ�����
Function checkAccessPass(MDBPath, AccessPass)
    On Error Resume Next 
    Dim conn 
    checkAccessPass = True 
    Set conn = CreateObject("Adodb.Connection")
        conn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database PassWord = '" & AccessPass & "';Data Source = " & MDBPath 
        conn.Close 
        If Err <> 0 Then
            Err.Clsoe 
            checkAccessPass = False 
        End If 
End Function
'�ж�SQl�Ƿ���ȷ
Function checkSql(sql)
    On Error Resume Next 
    conn.Execute(sql) 
    checkSql = IIF(Err = 0, True, False) 
End Function 
'������
Function createTable(tableName)
    Dim sql 
    If checkSql(tableName) = False Then
        sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key)" 
        conn.Execute(sql) 
        createTable = True 
    Else
        createTable = False 
    End If 
End Function 
'�жϱ��Ƿ���ʾ ���������
Function showIsTable(tableName)
    showIsTable = ExistsTable(tableName) 
    If showIsTable = False Then
        Call rw("������[" & tableName & "]�ɹ���<br>") 
    Else
        Call rw("��[" & tableName & "]�Ѿ�����<br>") 
    End If 
End Function 
'��¼���� �ܼ�¼�����ж��Ƿ��м�¼
Function getRecordCount(tableName, addSql)
    'Call Eerr("","Select Count(*) From [" & Table & "]" & AddSql)
    Call OpenConn() 
    getRecordCount = conn.Execute("Select Count(*) From [" & tableName & "] " & addSql)(0) 
End Function 
Function tableCount(Table, addSql)
    tableCount = getRecordCount(Table, addSql) 
End Function 
'�򿪱�RsOpenTable
Function openTable(tableName, addSql)
    rs.Open "Select * From [" & tableName & "] " & addSql, conn, 1, 1 
End Function 
'�������
Function insertTable(tableName, TableFields, values)
    conn.Execute("Insert Into[" & tableName & "](" & TableFields & ") Values('" & values & "')") 
End Function 
'��ʾ���ֶ�����
Function getTableText(tableName, fieldName, addSql)
    On Error Resume Next 
    getTableText = conn.Execute("Select " & fieldName & " From [" & tableName & "]" & addSql)(0) 
    If Err.Number <> 0 Then Call eerr("������", tableName & "-" & fieldName & "-" & addSql) 
End Function 
'����Mdb
Function createMdb(accessPath)
    Dim conn 
    Set conn = CreateObject("ADOX.Catalog")
        conn.create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & accessPath) 
        Call rw("�������ݿ�[" & accessPath & "]�ɹ���<br>" & vbCrLf) 
    Set conn = Nothing 
End Function 
'ɾ����
Function deleteTable(tableName)
    tableName = Trim(tableName) 
    deleteTable = False 
    If tableName = "" Then
        Exit Function 
    End If 
    If CheckTable(tableName) Then
        deleteTable = True 
    End If 
    'Conn.Execute("Drop Table [" & Table & "]")
'End If
End Function 
'ɾ��ȫ����
Function deleteAllTable(tableList)
    Dim splStr, tableName, c 
    splStr = Split(tableList, vbCrLf) 
    For Each tableName In splStr
        If c <> "" Then c = c & vbCrLf 
        c = c & "ɾ������" & tableName & "��" & deleteTable(tableName) 
    Next 
    deleteAllTable = c 
End Function 

'�жϱ�
Function checkTable(tableName)
    On Error Resume Next 
    conn.Execute("Select * From [" & tableName & "]") 
    If Not Err.Number = 0 Then
        Err.Clear                                                                       '����ô���
        checkTable = False 
    Else
        checkTable = True 
    End If 
End Function 
'�������� (����)
Function existsTable(Table)
    existsTable = checkTable(Table) 
End Function 
'�޸ı���
Function editTable(db, tableName, NewTable)
    Dim connStr, oCat, oTbl 
    editTable = False 
    If checkTable(tableName) And checkTable(NewTable) = False Then
        connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & db 
        Set oCat = CreateObject("ADOX.Catalog")
            oCat.ActiveConnection = connStr 
            Set oTbl = CreateObject("ADOX.Table")
                Set oTbl = oCat.Tables(tableName)                                               'Ҫ�������ı�����OldTable
                    oTbl.Name = NewTable                                                            '�±���
                Set oCat = Nothing 
            Set oTbl = Nothing 
            editTable = True 
        Set oCat = Nothing 
    End If 
End Function 
'�ж��ֶ�
Function checkField(tableName, fieldName)
    checkField = False 
    If checkTable(tableName) Then
        On Error Resume Next 
        conn.Execute("Select " & fieldName & " From [" & tableName & "]") 
        If Not Err.Number = 0 Then
            Err.Clear                                                                       '����ô���
            checkField = False 
        Else
            checkField = True 
        End If 
        Err.Clear 
    End If 
End Function 
'�������� (����)
Function existsField(Table, fieldName)
    existsField = checkField(Table, fieldName) 
End Function 
'�޸��ֶ�,
Function editField(tableName, fieldName, toFieldName)
    'On Error Resume Next
    Dim FieldType 
    editField = False 
    If checkTable(tableName) Then
        If checkField(tableName, toFieldName) = False And checkField(tableName, fieldName) Then
            FieldType = GetFieldAlt(tableName, fieldName) 
            If FieldType <> "" Then
                'MsgBox (tableName & vbCrLf & FieldName & vbCrLf & ToFieldName & "����" & FieldType)
                conn.Execute("Alter Table [" & tableName & "] Add [" & toFieldName & "] " & FieldType & "") '�����ֶ�
                conn.Execute("Update [" & tableName & "] Set " & toFieldName & "=[" & fieldName & "]") '��A�ֶ����ݸ��Ƶ�B�ֶ�������
                conn.Execute("Alter Table [" & tableName & "] Drop [" & fieldName & "]") 'ɾ���ֶ�
            End If 
            If Err.Number <> 0 Then MsgBox(tableName & vbCrLf & fieldName & vbCrLf & toFieldName & "����" & FieldType) 
        End If 
    End If 
End Function 
'����ֶ�����
Function getFieldAlt(tableName, fieldName)
    Dim n, Fn, Fld, rs 
    Set rs = CreateObject("Adodb.RecordSet")
        rs.Open "Select * From [" & tableName & "]", conn, 1, 1 
        Fn = rs.Fields.Count 
        For n = 1 To Fn - 1
            Set Fld = rs.Fields.Item(n)
                If lcase(fieldName) = lcase(Fld.Name) Then
                    'MsgBox (FieldName & vbCrLf & Rs.Fields(N).Type)
                    Select Case rs.Fields(n).Type
                        Case "202"
                            getFieldAlt = "VarChar" 
                        Case "203"
                            getFieldAlt = "Text" 
                        Case "7"
                            getFieldAlt = "DateTime" 
                        Case "3"
                            getFieldAlt = "Int" 
                        Case "11"
                            getFieldAlt = "YesNo" 
                        Case "5"
                            getFieldAlt = "float" 
                        Case Else
                            MsgBox(rs.Fields(n).Type & vbCrLf & fieldName & "û������") 
                            getFieldAlt = "û����" & rs.Fields(n).Type 
                    End Select
                    Exit For 
                End If 
            Set Fld = Nothing 
        Next : rs.Close 
End Function
'���ָ�����ݿ�����б�
Function getTableList_Conn(conn)
    Dim s, c,rs
    Set rs = conn.OpenSchema(20)
        rs.MoveFirst 
        Do Until rs.EOF
            If rs("TABLE_TYPE") = "TABLE" Then
                s = rs("TABLE_NAME") 
                If c <> "" Then c = c & vbCrLf 
                c = c & s 
            End If 
            rs.MoveNext 
        Loop 
        getTableList_Conn = c 
End Function
'������ݿ���б�
Function getTableList()
    getTableList = getTableList_Conn(conn) 
End Function 
'��ñ��ֶ��б�
Function getFieldList(tableName)
    Dim Fn, c, Fld, n 
    On Error Resume Next 
    '�ֶ�
    tempRs.Open "Select * From [" & tableName & "]", conn, 1, 1 
    If tempRs.Fields.Count = 0 Then Exit Function                                   'Ϊ���˳�
    Fn = tempRs.Fields.Count 
    If Err <> 0 Then Call eerr("�ֶ�", tempRs.Fields.Count) 
    c = "," 
    For n = 1 To Fn - 1
        Set Fld = tempRs.Fields.Item(n)
            c = c & Fld.Name & "," 
        Set Fld = Nothing 
    Next : tempRs.Close 
    getFieldList = c 
End Function 
'��ñ��ֶ������б�20160226
Function getFieldConfigList(tableName)
    Dim Fn, c, Fld, n ,s
    On Error Resume Next 
    '�ֶ�
    tempRs.Open "Select * From [" & tableName & "]", conn, 1, 1 
    If tempRs.Fields.Count = 0 Then Exit Function                                   'Ϊ���˳�
    Fn = tempRs.Fields.Count 
    If Err <> 0 Then Call eerr("�ֶ�", tempRs.Fields.Count) 
    c = "," 
    For n = 1 To Fn - 1
        Set Fld = tempRs.Fields.Item(n)
			if Fld.Type=3 or Fld.Type=11 then
				s="|numb|0"
			elseif Fld.Type=7 then
				s="|time|"
			elseif Fld.Type=203 then
				s="|textarea|"				
			else
				s="||"
			end if
            c = c & Fld.Name & s & ","
			
        Set Fld = Nothing 
    Next : tempRs.Close 
    getFieldConfigList = c 
End Function
'��ò�ͬ���ݿ���ֶ��б� (.mdb Or .xls)
Function getDifferentTableFieldList(ConnObj, Table)
    Dim Fn, c, Fld, n 
    '�ظ���������ݿ�
    If TypeName(ConnObj) = "Connection" Then
        Set conn = CreateObject("Adodb.Connection")
            conn.Open(ConnObj) 
    End If
    '�ֶ�
    rs.Open "Select * From [" & Table & "]", conn, 1, 1 
    Fn = rs.Fields.Count 
    c = "," 
    For n = 1 To Fn - 1
        Set Fld = rs.Fields.Item(n)
            c = c & Fld.Name & "," 
        Set Fld = Nothing 
    Next : rs.Close 
    getDifferentTableFieldList = c 
End Function 
'Call Echo("ѹ�����ݿ�", CompactDB("D:\��������\html1\Admin\Data\Data.mdb", False))
'=====================ѹ������========================= boolIs97 �������û��ʲô��
'ѹ�����ݿ�
Function compactDB(dBPath, boolIs97)
    Dim fso, Engine, strDBPath, JET_3X
	dBPath=handlePath(dBPath)
    strDBPath = Left(dBPath, InStrRev(dBPath, "\")) 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(dBPath) Then
            fso.CopyFile dBPath, strDBPath & "temp.mdb" 
            Set Engine = CreateObject("JRO.JetEngine")
                If boolIs97 = True Then
                    Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;" & "Jet OLEDB:Engine Type=" & JET_3X 
                Else
                    Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb" 
                End If 
                fso.CopyFile strDBPath & "temp1.mdb", dBPath 
                fso.deleteFile(strDBPath & "temp.mdb") 
                fso.deleteFile(strDBPath & "temp1.mdb") 
            Set Engine = Nothing 
            compactDB = "<font color=#FF0000><b>������ݿ�, " & dBPath & ", �Ѿ�ѹ���ɹ�!</b></font>" & vbCrLf 
        Else
            compactDB = "<font color=#FF0000><b>���ݿ����ƻ�·������ȷ. ������!</b></font>" & vbCrLf 
        End If 
    Set fso = Nothing 
End Function 
'�޸����ݿ���������
Function editAccessPassWord(dBPath, oldPass, newPass)
    Dim path, miJRO 
    Set miJRO = CreateObject("JRO.JetEngine")
        Call HandlePath(dBPath)                                                         '�������·��
        path = "Test_News.Mdb" 
        Call HandlePath(path)                                                           '�������·��
        Call deleteFile(path) 
        On Error Resume Next 
        miJRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dBPath & ";Jet OLEDB:Database Password=" & oldPass, "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & path & ";Jet OLEDB:Database Password=" & newPass 
        If Err.Number <> 0 Then editAccessPassWord = "�������" : Err.Clear : Exit Function 
        Call deleteFile(dBPath)                                                         'ɾ�����ݿ�
        Call MoveFile(path, dBPath)                                                     '�ƶ������ݿ�
        editAccessPassWord = "�޸ĳɹ���������Ϊ[" & newPass & "]" 
    Set miJRO = Nothing 
End Function 
'����SqlServer�����﷨(AccessתSqlServer)
'SqlServer���Ĵ���Ҫ��Ƚϸߣ��ո�ֻ��Ϊ''��������""���о�����ֵ���ַ���������
Function handleSqlServer(ByVal content)
    Dim i, s, TempS, TempS2, c, splStr 
    splStr = Split(content, ",") 
    For Each content In splStr

        content = Replace(content, """""", "''") 
        content = Replace(content, "''''", "''") 
        content = Replace(content, "Date()", "GetDate()") 
        content = Replace(content, "Now()", "GetDate()") 
        'Content = Replace(Content,"YesNo Default Yes"," Int Default 1")
        'Content = Replace(Content,"YesNo Default No"," Int Default 0")
        'Content = Replace(Content,"YesNo,"""," Int Default 1,""")

        content = Replace(content, "YesNo Default Yes", " Bit Default 'true'") 
        content = Replace(content, "YesNo Default No", " Bit Default 'false'") 
        content = Replace(content, " YesNo", "  Bit Default 'true'") 
        If InStr(content, "YesNo") > 0 Then
            Call eerr("Content", content) 
        End If 
        For i = 1 To Len(content)
            s = Mid(content, i, 1) 
            If s = " " Then
                TempS = LCase(Mid(content, i + 1, 7)) 
                TempS2 = LCase(Mid(content, i + 1, 8)) 
                If TempS = "varchar" And TempS2 <> "varchar(" Then
                    c = c & " VarChar(255)" 
                    i = i + 7 
                Else
                    c = c & s 
                End If 
            Else
                c = c & s 
            End If 
        Next 
        c = c & "," 
    Next 
    If c <> "" Then c = Left(c, Len(c) - 1) 
    handleSqlServer = c 
End Function 
'====================================================================
'���ݿ� ���뵼������ Call ExportToExcel(GetDataTableList("Product","")) �ѱ�����xls��Ϸ���


'Excel�������� Call ExportToExcel("�����б�")
Sub exportToExcel(tableContent)
    Response.ContentType = "application/vnd.ms-Excel" 
    Response.AddHeader "Content-Disposition", "attachment;Filename=Results.xls" 
    Response.Write("<body>") & vbCrLf 
    Response.Write(tableContent) & vbCrLf 
    Response.Write("</body>") & vbCrLf 
    Response.Write("</html>") & vbCrLf 
End Sub 
'���뵽���ݿ��� ������
Sub importToDatabase()
    '��XML
    Dim xlsFile, xmlConn, xmlRs 
    xlsFile = Server.MapPath("1.xls") 
    Set xmlConn = CreateObject("Adodb.Connection")
        xmlConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" & xlsFile 
        Set xmlRs = CreateObject("Adodb.RecordSet")
            '��MDB
            Dim mdbFile, mdbConn, mdbRs 
            mdbFile = Server.MapPath(""& adminDir &"Data/Data.mdb") 
            Set mdbConn = CreateObject("Adodb.Connection")
                mdbConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database PassWord = '';Data Source = " & mdbFile 
                Set mdbRs = CreateObject("Adodb.RecordSet")

                    '���xlsȫ���ֶ�
                    Dim Fn, n, Fld, xlsFieldList 
                    xmlRs.Open "Select * From [Excel$]", xmlConn, 1, 1 
                    Fn = xmlRs.Fields.Count 
                    For n = 1 To Fn - 1
                        Set Fld = xmlRs.Fields.Item(n)
                            xlsFieldList = xlsFieldList & Fld.Name & "," 
                        Set Fld = Nothing 
                    Next : xmlRs.Close 
                    Call echo("xlsFieldList", xlsFieldList) 
                    '��xls���ݵ���mdb
                    Dim splStr, fieldName 
                    splStr = Split(xlsFieldList, ",") 
                    xmlRs.Open "Select * From [Excel$]", xmlConn, 1, 1 
                    Call echo("xls��������", xmlRs.RecordCount) 
                    'On Error Resume Next
                    While Not xmlRs.EOF
                        mdbRs.Open "Select * From [Product]", mdbConn, 1, 3 
                        mdbRs.AddNew 
                        For Each fieldName In splStr
                            fieldName = Trim(fieldName) 
                            If fieldName <> "" Then
                                Call echo("��ʾ�ֶ�", fieldName) 
                                Doevents 
                                mdbRs(fieldName) = xmlRs(fieldName) 
                            End If 
                        Next 
                        Call echo("���Ӽ�¼", "") 
                        'Call Rw("<hr>")
                        mdbRs.Update : mdbRs.Close 
                    xmlRs.MoveNext : Wend : xmlRs.Close 
End Sub
'������ݱ��б� GetDataTableList("Product")
Function getDataTableList(tableName, addSql)
    Dim DataFieldList, splStr, i, c, fieldName 
    c = c & "<table width=""100%"" border=""1"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf 
    'Call OpenConn()         '�����ݿ�
    DataFieldList = getFieldList(tableName) 
    'Call Echo(TableName & "�ַ��б�",DataFieldList)
    splStr = Split(DataFieldList, ",") 
    c = c & "  <tr>" & vbCrLf 
    For i = 1 To UBound(splStr) - 1
        c = c & "    <td>" & splStr(i) & "</td>" & vbCrLf 
    Next 
    c = c & "  </tr>" & vbCrLf 

    rs.Open "Select * From [" & tableName & "] " & addSql, conn, 1, 1 
    While Not rs.EOF
        c = c & "  <tr>" & vbCrLf 
        For Each fieldName In splStr
            If fieldName <> "" Then
                c = c & "    <td>" & rs(fieldName) & "</td>" & vbCrLf 
            End If 
        Next 
        c = c & "  </tr>" & vbCrLf 
    rs.MoveNext : Wend : rs.Close 
    c = c & "</table>" & vbCrLf 
    getDataTableList = c 
End Function 
'����׷���ֶ�   �� Call Echo("׷���ֶ�",SetTableField("Admin,aa,Int Default 1|Admin,bb,VarChar(255) Default 'str'|"))
Function setTableField(conn, content)
    Dim splStr, splxx, i, s, c, tableName, fieldName, sql 
    splStr = Split(content, "|") 
    For i = 0 To UBound(splStr)
        s = splStr(i) 
        If s <> "" Then
            splxx = Split(s, ",") 
            tableName = splxx(0) 
            fieldName = splxx(1) 
            On Error Resume Next 
            sql = "alter table " & tableName & " add column " & fieldName & " " & splxx(2) & "" 
            conn.Execute sql 
            If Err.Number = 0 Then
                c = c & "[��" & tableName & "]���ֶ�" & fieldName & "���������ֶγɹ�<hr>" & vbCrLf 
            Else
                Err.Close 
                c = c & "[��" & tableName & "]���ֶ�" & fieldName & "������ʧ��<hr>" & vbCrLf 
            End If 
        End If 
    Next 
    setTableField = c 
End Function 

'�˺����Ὣ recordset ��λ���ʹ���ת��Ϊ�׶���Ӣ��ʶ���� �������ù�������Ӧ��
Function getTypeString(typeId)
    Select Case typeId
        Case 0 : getTypeString = "Empty"
        Case 16 : getTypeString = "TinyInt"
        Case 2 : getTypeString = "SmallInt"
        Case 3 : getTypeString = "Int"
        Case 20 : getTypeString = "BigInt"
        Case 17 : getTypeString = "UnsignedTinyInt"
        Case 18 : getTypeString = "UnsignedSmallInt"
        Case 19 : getTypeString = "UnsignedInt"
        Case 21 : getTypeString = "UnsignedBigInt"
        Case 4 : getTypeString = "Single"
        Case 5 : getTypeString = "Double"
        Case 6 : getTypeString = "Currency"
        Case 14 : getTypeString = "Decimal"
        Case 131 : getTypeString = "Numeric"
        Case 11 : getTypeString = "Bit"
        Case 10 : getTypeString = "Error"
        Case 132 : getTypeString = "UserDefined"
        Case 12 : getTypeString = "Variant"
        Case 9 : getTypeString = "IDispatch"
        Case 13 : getTypeString = "IUnknown"
        Case 72 : getTypeString = "GUID"
        Case 7 : getTypeString = "Date"
        Case 133 : getTypeString = "DBDate"
        Case 134 : getTypeString = "DBTime"
        Case 135 : getTypeString = "Datetime"
        Case 8 : getTypeString = "BSTR"
        Case 129 : getTypeString = "Char"
        Case 200 : getTypeString = "VarChar"
        Case 201 : getTypeString = "LongVarChar"
        Case 130 : getTypeString = "WChar"
        Case 202 : getTypeString = "Text"
        Case 203 : getTypeString = "Memo"
        Case 128 : getTypeString = "Binary"
        Case 204 : getTypeString = "VarBinary"
        Case 205 : getTypeString = "LongVarBinary"
        Case Else : getTypeString = "Unknown Type"
    End Select
End Function 
'��ò�Ʒָ���������ID�б�
Function getProductRndIdList(rs, nItem)
    Dim nCount, i, j, RndIdList, PidList, nOK, id, splxx 
    'ִ�м���
    nCount = rs.RecordCount 
    nItem = CInt(nItem)                                                             'ת����������
    nOK = 0 
    Randomize Timer                                                                 '��ʼ�������������
    'ѭ������
    For j = 1 To 3
        For i = 1 To nCount
            id = Int(Rnd * nCount + 1) 
            If InStr("," & RndIdList & ",", "," & id & ",") = False Then
                RndIdList = RndIdList & id & "," 
                nOK = nOK + 1 
                If nOK >= nItem Then
                    'Call Eerr(RndIdList &"nOK=" & nOK,"nItem=" & nItem)
                    Exit For 
                End If 
            End If 
        Next 
        If nOK >= nItem Then Exit For 
    Next 
    '��ò�ƷID
    splxx = Split(RndIdList, ",") 
    For i = 1 To nCount
        If InStr("," & RndIdList & ",", "," & i & ",") > 0 Then
            PidList = PidList & rs(0) & "," 
        End If 
    rs.MoveNext : Next 
    If PidList <> "" Then PidList = Left(PidList, Len(PidList) - 1) 
    getProductRndIdList = PidList 
'Call Eerr(RndIdList,PidList)
End Function 
'���ĳ�ֶ����� 20150129  ����getFieldValue("WebSite","WebTitle","")     '�����վ����
Function getFieldValue(tableName, fieldName, addSql)
    Dim rs 
    Call OpenConn() 
    Set rs = CreateObject("Adodb.RecordSet")
        rs.Open "Select * From [" & tableName & "] " & addSql, conn, 1, 1 
        If Not rs.EOF Then
            getFieldValue = rs(fieldName) 
        End If : rs.Close 
End Function
'���ָ���ֶ����ݡ�ָ��Conn����
Function getConnFieldValue(objConn, tableName, fieldName, addSql)
    On Error Resume Next 
    Dim rs 
    Set rs = CreateObject("Adodb.RecordSet")
        rs.Open "Select * From [" & tableName & "] " & addSql, objConn, 1, 1 
        If Not rs.EOF Then
            'Call Echo(TableName, FieldName):Doevents
            getConnFieldValue = rs(fieldName) 
        End If : rs.Close 
        If Err <> 0 Then Call eerr(tableName, fieldName) 
End Function
'��ĳ�ֶ�д���� 20150129  ����Call setFieldValue("WebSite","WebTitle","�й���","")     '������վ����
Function setFieldValue(tableName, fieldName, writeValue, addSql)
    Dim rs 
    Call OpenConn() 
    setFieldValue = False 
    Set rs = CreateObject("Adodb.RecordSet")
        rs.Open "Select * From [" & tableName & "] " & addSql, conn, 1, 3 
        If Not rs.EOF Then
            rs(fieldName) = writeValue 
            rs.Update 
            setFieldValue = True 
        End If : rs.Close 
End Function

'�ļ�������д�����ݿ�20150312 Call FileBinaryAddAccess(Conn,"Admin","test1","","","UploadFiles\testimages2015_source.jpg")
Function fileBinaryAddAccess(conn, tableName, fieldName, id, addSql, filePath)
    Dim sql, rs 
    Set rs = CreateObject("Adodb.RecordSet")
        sql = "Select " & fieldName & " From [" & tableName & "]" 
        sql = getWhereAnd(sql, addSql) 
        If id <> "" Then
            sql = getWhereAnd(sql, "Where Id=" & id) 
        End If 
        rs.Open sql, conn, 1, 3 
        If Not rs.EOF Then
            rs(fieldName).AppendChunk GetFileBinaryContent(filePath) 
            'Call eerr(sql,filepath):Doevents
            rs.Update 
        End If : rs.Close 
End Function
'��ʾ���ݿ�ָ���ֶ�����������20150312 Call ShowAccessStream(Conn,"Admin","test1","","")
Function showAccessStream(conn, tableName, fieldName, id, addSql)
    Dim sql, rs 
    Set rs = CreateObject("Adodb.RecordSet")
        sql = "Select " & fieldName & " From [" & tableName & "]" 
        sql = getWhereAnd(sql, addSql) 
        If id <> "" Then
            sql = getWhereAnd(sql, "Where Id=" & id) 
        End If 
        rs.Open sql, conn, 1, 3 
        If Not rs.EOF Then
            'Response.ContentType = "img/*"
            Response.BinaryWrite rs(fieldName).GetChunk(7500000) 
        End If : rs.Close 
End Function
%>     




