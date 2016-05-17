<%
'*---------- Cache缓冲 ---------*'

' AddCache array(dhIn,"",ShowHtml,"说明Index")
CONST CACHEPREFIX = "CACHE_" '缓存前缀,不能为空

'说明：
'         1,缓存的格式为Application(CACHEPREFIX & key) = array("缓存key","缓存时间","缓存内容","缓存说明","到期时间")
'           2,缓存key不区分大小写

' AddCache array(dhIn,"",SHtml,"说明Index")

'添加缓存,不检查缓存是否存在，如果存在则相当于更新缓存
'varAry :   参数，格式为:array("缓存key","缓存时间","缓存内容","缓存说明")
'           缓存key ：application()格式相同
'           缓存时间：单位秒,可以为负数,表示立即过期,可以为空，空或不是数字则默认为20分钟过期
'           缓存内容：缓存数据，不支持对对象的缓存
'           缓存说明：缓存描述说明
Function AddCache(varAry)
	Dim c,ary(4)
	If Not IsArray(varAry) Then
		Response.Write "Error:AddCache(varAry)参数错误，参数不是数组"
		Response.End()
	End If
	If UBound(varAry) <> 3 Then
		Response.Write "Error:AddCache(varAry)参数错误，数组长度错误"
		Response.End()
	End If
	If varAry(0) = "" Then
		Response.Write "Error:AddCache(varAry)错误，key不能为空"
		Response.End()
	End If
	If varAry(1) = "" or Not IsNumeric(varAry(1)) Then varAry(1) = 1200
	Application.Lock()
	Application(CACHEPREFIX & varAry(0)) = array(varAry(0),varAry(1),varAry(2),varAry(3),DateAdd("s",varAry(1),Now()))
	Application.UnLock()
End Function

' 检查某个缓存是否存在,存在返回True否则返回False
' key : 缓存key
Function CheckCache(key)
	Dim k
	For Each k In Application.Contents
		If LCase(k) = LCase(CACHEPREFIX & key) Then CheckCache = True : Exit Function
	Next
	CheckCache = False
End Function

' 清除某个缓存
Function RemoveCache(key)
	Application.Lock()
	Application.Contents.Remove(CACHEPREFIX & key)
	Application.UnLock()
End Function

'更新缓存,如果缓存不存在则出错，所以更新之前先用CheckCache(key)检查
'varAry :   参数，格式为:array("缓存key","缓存时间","缓存内容","缓存说明")
'           缓存key ：application()格式相同
'           缓存时间：单位秒,可以为负数,表示立即过期,可以为空，空或不是数字则默认为20分钟过期
'           缓存内容：缓存数据，不支持对对象的缓存
'           缓存说明：缓存描述说明
' 注意   ： 如果不更新varAry某个值则设置该值为null即可，
'           如UpdateCache(array("key",null,"内容",null)),就是不更新过期时间和说明
Function UpdateCache(varAry)
	Dim app
	app = GetCache(varAry(0))
	If Not IsNull(varAry(1)) Then app(1) = varAry(1)
	If Not IsNull(varAry(2)) Then app(2) = varAry(2)
	If Not IsNull(varAry(3)) Then app(3) = varAry(3)
	If app(1) = "" or Not IsNumeric(app(1)) Then app(1) = 1200
	Application.Lock()
	Application(CACHEPREFIX & app(0)) = array(app(0),app(1),app(2),app(3),DateAdd("s",app(1),Now()))
	Application.UnLock()
End Function

' 打印cache，做调试用
Function PrintCache(key)
    Dim App, I, T
    If CheckCache(key) Then
        App = GetCache(key)
        Response.Write "<pre>{" & Chr(10)
        Response.Write Chr(32) & "缓存名称" & Chr(32) & ":" & Chr(32) & CACHEPREFIX & App(0) & Chr(10)
        Response.Write Chr(32) & "缓存key " & Chr(32) & ":" & Chr(32) & App(0) & Chr(10)
        Response.Write Chr(32) & "缓存时间" & Chr(32) & ":" & Chr(32) & App(1) & Chr(10)
        Response.Write Chr(32) & "到期时间" & Chr(32) & ":" & Chr(32) & App(4) & Chr(10)
        Response.Write Chr(32) & "是否到期" & Chr(32) & ":" & Chr(32) & App(5) & Chr(10)
        Response.Write Chr(32) & "缓存说明" & Chr(32) & ":" & Chr(32) & App(3) & Chr(10)
        ' 内容
        Response.Write Chr(32) & "缓存内容" & Chr(32) & ":" & Chr(32)
        T = VarType(App(2))
        If InStr(",0,1,2,3,4,5,6,7,8,11,", "," & T & ",") > 0 Then
            Response.Write App(2)
        Else
            Response.Write TypeName(App(2))
        End If
        Response.Write Chr(10)
        
        Response.Write "}</pre>" & Chr(10)
    Else
        Response.Write "不存在该缓存"
    End If
 End Function
 
' 获取缓存
' 返回数组，格式为：array("缓存key","缓存时间","缓存内容","缓存说明","到期时间",是否过期True|False)
' 如果不存在则出错，所以获取之前先用CheckCache(key)检查
Function GetCache(key)
    Dim App, IsExp
    App = Application(CACHEPREFIX & key)
    IsExp = False
    If DateDiff("s", Now(), App(4)) <= 0 Then IsExp = True
    GetCache = Array(App(0), App(1), App(2), App(3), App(4), IsExp)
End Function
 
' ==========================
'   输出缓冲内容
' ==========================
Function HcShow(key)
    Dim App, I, T
    If CheckCache(key) Then
        App = GetCache(key)
        ' Response.Write chr(32) & "缓存内容" & chr(32) & ":" & chr(32)
        T = VarType(App(2))
        If InStr(",0,1,2,3,4,5,6,7,8,11,", "," & T & ",") > 0 Then
            HcShow = App(2)
        Else
            HcShow = TypeName(App(2))
        End If
    Else
        HcShow = "不存在该缓存"
    End If
End Function

' 换行再定义
Sub Br(Str)
    Response.Write Str & "<br />" & vbCrLf
End Sub

' 清除全部缓冲
Function RemoveAllCache()
    Application.Lock
    Application.Contents.RemoveAll
    Application.Unlock
End Function
%>