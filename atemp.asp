<!--#Include File = "Inc/_Config.Asp"-->            
<!--#Include File = "web/function.Asp"-->      
<%
 
'Xor����
Function xorEnc(code, n)
    Dim c, s1, s2, s3, i 
    c = code 
    s1 = Len(c) : s3 = "" 
    For i = 0 To s1 - 1
        s2 = AscW(Right(c, s1 - i)) Xor n 
        s3 = s3 & ChrW(Int(s2)) 
    Next 
    'Chr(34) ���ǵ���(") ��ֹ���� ��Ϊ"��ASP�����
    s3 = Replace(s3, ChrW(34), "��") 
    xorEnc = s3 
End Function 

call rw(xorEnc("313801120",2))
%>
