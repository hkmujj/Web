<!--#Include virtual = "/Inc/_Config.Asp"-->
<%

call rwend(16*16)
dim i,s,c
for i = 0 to 10
	s="[0,"& i*17 &"]"
	if c<>"" then c=c & ","
	c=c & s
next

c="["& c &"]"
call rw(c)

		
%>
