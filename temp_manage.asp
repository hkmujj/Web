<!--#Include virtual = "/Inc/_Config.Asp"-->
<%

select case request("act")
	case "delfile" : call_delfile()
	case "delfolder" : call_delfolder()
	case else:displayDefault()
end select

'删除文件
function call_delfile()
	dim filePath
	filePath=request("filePath")
	if getip()<>"127.0.0.1" then
		call deleteFile(filePath)
		call eerr("删除文件",filePath)
	else
		call eerr("提示","为本地，请自行手动删除 文件")
	end if
end function
'删除文件夹
function call_delfolder()
	dim folderPath
	folderPath=request("folderPath")
	if folderPath="" then
		call eerr("出错","文件夹为空")
	end if
	if getip()<>"127.0.0.1" then
		call deleteFolder(folderPath)
		call eerr("删除文件夹",folderPath)
	else
		call eerr("提示","为本地，请自行手动删除 文件夹")
	end if
end function
'显示默认
function displayDefault()
	dim content,splstr,fileName,s,filePath,folderPath
	content=getDirFileList("/","")
	splstr=split(content,vbcrlf)
	for each filePath in splstr
		s="<a href='?act=delfile&filePath="& filePath &"' onclick='return jsconfirm();'>删除</a>"
		call echo(filePath,s)
	next
	
	content=getDirFolderList("/")
	splstr=split(content,vbcrlf)
	for each folderPath in splstr
		s="<a href='?act=delfolder&folderPath="& folderPath &"' onclick='return jsconfirm();'>删除</a>"
		call echo(folderPath,s)
	next
end function
%>
<script>
// 确定操作
function jsconfirm(){
	if(confirm("你确定要操作吗？\n操作后将不可恢复"))
	return true;
	else
	return false;
}
</script>
