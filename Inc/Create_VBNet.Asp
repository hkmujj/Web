<!--#Include File = "Config.Asp"--> 

<%
'  2015_ASP.Asp   dimAddDefaultValue������Ĭ�ϱ���Ϊ��

call CreateVBDialog()
Sub CreateVBDialog()
	Dim SplStr,I,S,C,Content,RootFolder,zFormFolder,DataFolder,Path,ToPath,FormName,BasList,FrmList,ToolsMenu,FrmListStr,OpenFrom,nTools
	dim TempContent,StartStr,endStr,splxx,functionName,centerStr
	c="|Print.Asp|Access.Asp|Cai.Asp|EncDec.Asp|Function.Asp|GetWord.Asp|RegExp.Asp|StringNumber.Asp|XML.Asp|Time.Asp|URL.Asp|FSO.Asp|Check.Asp|Html.Asp|AdodbStream.Asp|Dictionary.Asp|PinYin.Asp|2014_Array.Asp|2014_Css.Asp|2014_GBUTF.Asp|2015_Formatting.Asp|2015_Param.Asp|2015_ASP.Asp|2015_Color.Asp|FunHTML.Asp|2015_PHP.Asp|2015_WebSql.Asp|2015_hexSwitch.Asp"
	
	zFormFolder="D:\WindowsApplication1\zForm\Z_"
	
	
	dim searchStr,replaceStr
	'����ģ���б�
	SplStr=Split(c,"|")'|ZClassAspCode.Asp
	For Each Path In SplStr
		If Path <> "" And 1=1 Then
			FormName = Mid(Path,1,InStrRev(Path,".")-1)
			'�ж��ļ������ݴ�ͷ�����ǰ׺ ����VB���û����
			If InStr("0123456789",Left(FormName,1))>0 Then
				FormName="y_" & FormName							'��������ǰ��������
			End If
			Content = ReadFile("/inc/" & Path,"")		
			
			searchStr = getStrCut(Content,"'footer start","'footer end",2)
			if searchStr<>"" then
				replaceStr=vbcrlf & getftext("\DataDir\VBģ��\2015_Param.Asp.txt")
				Content=replace(Content,searchStr,replaceStr)
				call echo(path,"�滻footer�ɹ�")
			end if 
			'ɾ��������������   'NOVBNet start                'NOVBNet end
			for i = 1 to 3
				searchStr = getStrCut(Content,"'NOVBNet start","'NOVBNet end",2)
				if searchStr<>"" then 
					Content=replace(Content,searchStr,"") 
				else
					exit for
				end if 
			next
			'���
			If Path = "ZClassAspCode.Asp" Then
 
			'ģ��
			Else
				ToPath = zFormFolder & FormName & ".vb"
				Content = StrCut(Content,"<" & "%","%" & ">",2)
				Content = "Module "& FormName & vbCrlf & Content & vbcrlf & "End Module"
			End If
			tempContent=content
			content=replace(content,"Ʋ�X���ǣ������Һᡣ��",  "Ʋ�X���ǣ�"" & Chr(-24144) & ""���Һᡣ��") 
			
			If Path = "Time.Asp" Then			
				content=replace(content,"d = CStr(Day(timeStr))","d = Format(timeStr, ""dd"")")
			end if
			
			content=replace(content,vbtab,"    ")
			content=ASPToVB(content) 
			content=ASPToVBNET(content) 
			
			
			if Content="" then
				call eerr(FormName,tempContent)
			end if
			


			'����Ϊ�� 
			splxx=split("getThisUrl|chkPost|getUpdateFile|showErr|printFormInfo|popupDownFile|downRemoteFile|c10to2|agent|getWebHome|webDebug|getThisUrlFileName|webDoMain|checkUrlName|checkUrlName|handleContentCode|replaceContentModule|replaceContentRowModule|addModuleReplaceArray|moduleFindContent|vBGetTimer|vBEchoTimer|||||","|")
			for each functionName in splxx
				if functionName<>"" then
					startStr="Function " & functionName
					endStr="End Function"
					s=GetStrCut(content,startStr,endStr,1)
					if s<>"" then
						centerStr=""
					
					 	startStr=mid(s,1,instr(s,")"))
						if functionName="GetFileText" then					
							centerStr= vbcrlf & "Call HandlePath(filePath)    '�������·�� "& vbcrlf &"If IO.File.Exists(filePath) = True Then" & vbcrlf
							centerStr=centerStr & "    Dim TxtReader As IO.StreamReader = New IO.StreamReader(filePath, System.Text.Encoding.Default)" & vbcrlf
							centerStr=centerStr & "    MyTxtReader = TxtReader.ReadToEnd" & vbcrlf
							centerStr=centerStr & "    TxtReader.Close()" & vbcrlf
							centerStr=centerStr & "End If" & vbcrlf 
						end if
						content=replace(content,s,  startStr & vbcrlf  & functionName & "=""" & centerStr & vbcrlf & endStr)
					end if
					
					startStr="Sub " & functionName
					endStr="End Sub"
					s=GetStrCut(content,startStr,endStr,1)
					if s<>"" then
					 	startStr=mid(s,1,instr(s,")"))
						content=replace(content,s,  startStr & vbcrlf & endStr)
					end if
				end if
			next

			'ɾ������
			splxx=split("getTimerSet|calculationTimer|getTimer|VBGetTimer|VBEchoTimer|errorText|getPort|||||||||||","|")
			for each functionName in splxx
				if functionName<>"" then
					startStr="Sub " & functionName
					endStr="End Sub"
					s=GetStrCut(content,startStr,endStr,1)
					if s<>"" then
						content=replace(content,s,  "")
					end if
					
					startStr="Function " & functionName
					endStr="End Function"
					s=GetStrCut(content,startStr,endStr,1)
					if s<>"" then
						content=replace(content,s,  "")
					end if
				end if
			next
			
			'�滻
			content=RegExp_Replace(content,"If Err Then","If Err.Number Then") 
			content=RegExp_Replace(content," Err <> 0 "," Err.Number <> 0 ") 
			content=RegExp_Replace(content,"If Err Then","If Err.Numbe Then")
			content=RegExp_Replace(content,"Set xml ","xml")
			content=RegExp_Replace(content,"Set Stm ","Stm")
			
			
			
			
			
			content=replace(content,"App.Path ","Application.StartupPath() ")				'ϵͳ·��
			content=replace(content,"C2 = ""Ʋ�X���ǣ������Һᡣ��""","C2 = ""Ʋ�X���ǣ�"" &  chr(""-24144"") &  ""���Һᡣ��""")
			
			content=replace(content,"Cstr(Day(s_Time))","Cstr(Format(Now(), ""dd""))")				'ɾ��set
			
			
			
			content=replace(content," LenB("," Len(")
			content=replace(content," AscB("," Asc(")
			content=replace(content," MidB("," Mid(")
			content=replace(content," ChrB("," Chr(")	 
			
			content=replace(content,"(LenB(","(Len(")
			content=replace(content,"(AscB(","(Asc(")
			content=replace(content,"(MidB(","(Mid(")
			content=replace(content,"(ChrB(","(Chr(")	 
			
			content=replace(content,"""��""", "chr(-24143)" )	 
			content=replace(content,"""��""", "chr(-24144)" )	 
			content=replace(content,"D = Cstr(Day(timeStr))", "D = CStr(CDate(timeStr).Day)" )	 	'��vb.net������ڴ���
			
			
			
			Content = Replace(Content,vbTab,"    ")			'�滻Tab
			Call Echo(FormName,ToPath)
			Call CreateFile(ToPath,Content)
		End If
	Next
end sub
%>