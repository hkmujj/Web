<!--#Include virtual="/Inc/_Config.Asp"-->
<% 

dim webDir,incDir,content,splstr,filePath,fileName,incC,c
webDir="E:\EÅÌ\PHPÔ´Âë\PhpApplication\case1\"
incDir = webDir & "/inc"
call deleteFolder(incDir)
call createFolder(incDir)
call echo("incDir",incDir)
content=getDirPhpList("E:\EÅÌ\WEBÍøÕ¾\ÖÁÇ°ÍøÕ¾\PHP2\Web\Inc")
content=content & vbcrlf & getDirPhpList("E:\EÅÌ\WEBÍøÕ¾\ÖÁÇ°ÍøÕ¾\PHP2\ImageWaterMark\Include")
call echo("content",content)
splstr=split(content,vbcrlf)
for each filePath in splstr
	fileName=getStrFileName(filePath)
	call echo(fileName,filePath)
	if instr(lcase(fileName),"install")=false then
		incC=incC & "require_once './Inc/"& fileName &"';" & vbcrlf
	end if
	call copyFile(filePath, incDir & "/" & fileName)
next

c="<?php" & vbcrlf & incC & copyStr(vbcrlf,5) & "?>"
call createFile(webDir & "/index.php",c)


%>