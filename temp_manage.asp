<!--#Include virtual = "/Inc/_Config.Asp"-->
<%

select case request("act")
	case "delfile" : call_delfile()
	case "delfolder" : call_delfolder()
	case else:displayDefault()
end select

'ɾ���ļ�
function call_delfile()
	dim filePath
	filePath=request("filePath")
	if getip()<>"127.0.0.1" then
		call deleteFile(filePath)
		call eerr("ɾ���ļ�",filePath)
	else
		call eerr("��ʾ","Ϊ���أ��������ֶ�ɾ�� �ļ�")
	end if
end function
'ɾ���ļ���
function call_delfolder()
	dim folderPath
	folderPath=request("folderPath")
	if folderPath="" then
		call eerr("����","�ļ���Ϊ��")
	end if
	if getip()<>"127.0.0.1" then
		call deleteFolder(folderPath)
		call eerr("ɾ���ļ���",folderPath)
	else
		call eerr("��ʾ","Ϊ���أ��������ֶ�ɾ�� �ļ���")
	end if
end function
'��ʾĬ��
function displayDefault()
	dim content,splstr,fileName,s,filePath,folderPath
	content=getDirFileList("/","")
	splstr=split(content,vbcrlf)
	for each filePath in splstr
		s="<a href='?act=delfile&filePath="& filePath &"' onclick='return jsconfirm();'>ɾ��</a>"
		call echo(filePath,s)
	next
	
	content=getDirFolderList("/")
	splstr=split(content,vbcrlf)
	for each folderPath in splstr
		s="<a href='?act=delfolder&folderPath="& folderPath &"' onclick='return jsconfirm();'>ɾ��</a>"
		call echo(folderPath,s)
	next
end function
%>
<script>
// ȷ������
function jsconfirm(){
	if(confirm("��ȷ��Ҫ������\n�����󽫲��ɻָ�"))
	return true;
	else
	return false;
}
</script>
