<!--#Include File = "Inc/_Config.Asp"-->
<%
select case request("act")
	case "makePHPInstallFile" : call printAccessToPHPInstallFile():call echo("����php�������ݿ��ļ��ɹ�","\PHP2\ImageWaterMark\Include\Install.php")
	case else : showDefault()
end select

'��ʾĬ��
sub showDefault()
	call echo("����", "<a href=""?act=makePHPInstallFile"">����Install.php�ļ�</a>")
end sub
%>
