<!--#Include File = "../Inc/Config.Asp"--> 
<!--#Include File = "function.asp"--> 
<!DOCTYPE html> 
<html xmlns="http://www.w3.org/1999/xhtml"> 
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" /> 
<title>模板文件管理</title> 
</head> 
<body> 
<style type="text/css"> 
<!-- 
body { 
    margin-left: 0px; 
    margin-top: 0px; 
    margin-right: 0px; 
    margin-bottom: 0px; 
} 
a:link,a:visited,a:active { 
    color: #000000; 
    text-decoration: none; 
} 
a:hover { 
    color: #666666; 
    text-decoration: none; 
} 
.tableline{ 
    border: 1px solid #999999; 
} 
body,td,th { 
    font-size: 12px; 
} 
a { 
    font-size: 12px; 
} 
--> 
</style> 
<script language="javascript"> 
function checkDel() 
{ 
    if(confirm("确认要删除吗？删除后将不可恢复！")) 
    return true; 
    else 
    return false; 
} 
</script> 
<% 

If Session("adminusername") = "" Then
    Call eerr("提示", "未登录，请先登录") 
End If 


Select Case Request("act")
    Case "templateFileList" : displayTemplateDirDialog(Request("dir")) : templateFileList(Request("dir"))'模板列表
    Case "delTemplateFile" : Call delTemplateFile(Request("dir"), Request("fileName")) : displayTemplateDirDialog(Request("dir")) : templateFileList(Request("dir"))		'删除模板文件
    Case "addEditFile" : displayTemplateDirDialog(Request("dir")) : Call addEditFile(Request("dir"), Request("fileName"))'显示添加修改文件
    Case Else : displayTemplateDirDialog(Request("dir"))                                        '显示模板目录面板
End Select

'模板文件列表
Sub templateFileList(dir)
    Dim content, splStr, fileName, s 
	if  Session("adminusername") = "ASPPHPCMS"  then
		content = getDirFileNameList(dir,"")
	else
	    content = getDirHtmlNameList(dir)
	end if
    splStr = Split(content, vbCrLf) 
    For Each fileName In splStr
		if fileName<>"" then
			s = "<a href=""../aspweb.asp?templatedir=" & escape(dir) & "&templateName=" & fileName & """ target='_blank'>预览</a> " 
			Call echo("<img src='Images/Icon/2/htm.gif'>" & fileName, s & "| <a href='?act=addEditFile&dir=" & dir & "&fileName=" & fileName & "'>修改</a> | <a href='?act=delTemplateFile&dir=" & Request("dir") & "&fileName=" & fileName & "' onclick='return checkDel()'>删除</a>") 
		end if
    Next 
End Sub
 
'删除模板文件
Sub delTemplateFile(dir, fileName)
    Dim filePath 
	
	call handlePower("删除模板文件")						'管理权限处理
	
    filePath = dir & "/" & fileName 
    Call deleteFile(filePath) 
    Call echo("删除文件", filePath) 
End Sub

'显示面板样式列表
function displayPanelList(dir)
	dim content,splstr,s,c
	content=getDirFolderNameList(dir)
	splstr=split(content,vbcrlf)
	c="<select name='selectLeftStyle'>"
	for each s in splstr
		s="<option value=''>"& s &"</option>"
		c=c & s & vbcrlf
	next
	displayPanelList = c & "</select>"	
end function
 

'添加修改文件
Function addEditFile(dir, fileName)
    Dim filePath,promptMsg
	
    If Right(LCase(fileName), 5) <> ".html" and Session("adminusername") <> "ASPPHPCMS" Then
        fileName = fileName & ".html" 
    End If 
    filePath = dir & "/" & fileName
	
	if checkFile(filePath)=false then
		call handlePower("添加模板文件")						'管理权限处理
	else
		call handlePower("修改模板文件")						'管理权限处理	
	end if
	 
    '保存内容
    If Request("issave") = "true" Then
        Call createfile(filePath, Request("content")) 
		promptMsg="保存成功"
    End If 
%> 
<form name="form1" method="post" action="?act=addEditFile&issave=true"> 
  <table width="99%" border="0" cellspacing="0" cellpadding="0" class="tableline"> 
    <tr> 
      <td height="30">目录<% =dir%><br> 
      <input name="dir" type="hidden" id="dir" value="<% =dir%>" /></td> 
    </tr> 
    <tr> 
      <td>文件名称 
      <input name="fileName" type="text" id="fileName" value="<% =fileName%>" size="40"><%=promptMsg%>
      <br> 
      <textarea name="content"  style="width:99%;height:480px;"id="content"><%call rw(getFText(filePath))%></textarea></td> 
    </tr> 
    <tr> 
      <td height="40" align="center"><input type="submit" name="button" id="button" value=" 保存 " /></td> 
    </tr> 
  </table> 
</form>
<% End Function
'文件夹搜索
Function displayTemplateDirDialog(dir)
	dim folderPath
%> 
<form name="form2" method="post" action="?act=templateFileList"> 
  <table width="99%" border="0" cellspacing="0" cellpadding="0" class="tableline"> 
    <tr> 
      <td height="30"><input name="dir" type="text" id="dir" value="<% =dir%>" size="60" /> 
        <input type="submit" name="button2" id="button2" value=" 进入 " /><%
		folderPath=dir & "/images/column/"
		if checkFolder(folderPath) then
			call rw("面板样式" & displayPanelList(folderPath))
		end if
		folderPath=dir & "/images/nav/"
		if checkFolder(folderPath) then
			call rw("导航样式" & displayPanelList(folderPath))
		end if
		%></td> 
    </tr> 
  </table> 
</form> 
<% End Function%>
