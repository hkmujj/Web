<% 


'����function2�ļ�����
Function callFunction_cai()
    Select Case Request("stype")
		case "callFunction_cai_test" : callFunction_cai_test()											'����
        Case Else : Call eerr("callFunction_caiҳ��û�ж���", Request("stype"))
    End Select
End Function

'����
function callFunction_cai_test()
   call echo("����","callFunction_cai_test")
end function 


%>                  

