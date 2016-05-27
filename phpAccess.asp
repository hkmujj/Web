<!--#Include File = "Inc/_Config.Asp"-->
<%
select case request("act")
	case "makePHPInstallFile" : call printAccessToPHPInstallFile():call echo("生成php创建数据库文件成功","\PHP2\ImageWaterMark\Include\Install.php")
	case else : showDefault()
end select

'显示默认
sub showDefault()
	call echo("操作", "<a href=""?act=makePHPInstallFile"">生成Install.php文件</a>")
end sub
%>
