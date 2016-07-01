<!--#Include File = "Inc/_Config.Asp"-->            
<!--#Include File = "web/function.Asp"-->      
<%
 
'Xor加密
Function xorEnc(code, n)
    Dim c, s1, s2, s3, i 
    c = code 
    s1 = Len(c) : s3 = "" 
    For i = 0 To s1 - 1
        s2 = AscW(Right(c, s1 - i)) Xor n 
        s3 = s3 & ChrW(Int(s2)) 
    Next 
    'Chr(34) 就是等于(") 防止出错 因为"在ASP里出错
    s3 = Replace(s3, ChrW(34), "ㄨ") 
    xorEnc = s3 
End Function 

call rw(xorEnc("313801120",2))
%>
