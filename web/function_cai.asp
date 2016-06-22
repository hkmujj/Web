<% 


'调用function2文件函数
Function callFunction_cai()
    Select Case Request("stype")
		case "callFunction_cai_test" : callFunction_cai_test()											'测试
        Case Else : Call eerr("callFunction_cai页里没有动作", Request("stype"))
    End Select
End Function

'测试
function callFunction_cai_test()
   call echo("测试","callFunction_cai_test")
end function 


%>                  

