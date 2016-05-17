<%
'详细页Banner图 20150123
function XY_DetailBanner(action)
    dim DefaultImage, did, sid, tid, MainId, ProId, ImgUrl 
    DefaultImage = RParam(action, "defaultimage")                                   '默认图片

    did = PubNavDid 
    if did = "" then did = PubProDid 
    sid = PubNavSid 
    if sid = "" then sid = PubProSid 
    tid = PubProTid 
    if tid = "" then tid = PubProTid 

    if PubID <> "" then
        if PubActionType = "MainInfo" then
            MainId = PubID 
        elseIf PubActionType = "Article" then
            ProId = PubID 
        end if 
    end if 

        '产品大类
    if did <> "" then
        rs.open "Select * From [BigClass] Where BigClassName='" & did & "' And ImgPath<>''", conn, 1, 1 
        if not rs.eOF then
            ImgUrl = rs("ImgPath") 
        end if : rs.close 
    end if 
        '产品小类
    if sid <> "" then
        rs.open "Select * From [SmallClass] Where BigClassName='" & did & "' And SmallClassName='" & sid & "' And ImgPath<>''", conn, 1, 1 
        if not rs.eOF then
            ImgUrl = rs("ImgPath") 
        end if : rs.close 
    end if 
        '产品子类
    if tid <> "" then
        rs.open "Select * From [ThreeClass] Where BigClassName='" & did & "' And SmallClassName='" & sid & "' And ThreeClassName='" & tid & "' And ImgPath<>''", conn, 1, 1 
        if not rs.eOF then
            ImgUrl = rs("ImgPath") 
        end if : rs.close 
    end if 

        '导航大类
    if did <> "" then
        rs.open "Select * From [NavBigClass] Where BigClassName='" & did & "' And ImgPath<>''", conn, 1, 1 
        if not rs.eOF then
            ImgUrl = rs("ImgPath") 
        end if : rs.close 
    end if 
        '导航小类
    if sid <> "" then
        rs.open "Select * From [NavSmallClass] Where BigClassName='" & did & "' And SmallClassName='" & sid & "' And ImgPath<>''", conn, 1, 1 
        if not rs.eOF then
            ImgUrl = rs("ImgPath") 
        end if : rs.close 
    end if 
        '产品详细   不需要这个，因为这样会把上传的产品用到Banner详细页里，这样就不好了呀，你这个笨蛋。。。20150124
    if ProId <> "" and 1 = 2 then
        rs.open "Select * From [Product] Where Id=" & ProId & " And (BigFiles<>'' Or SmallFiles<>'')", conn, 1, 1 
        if not rs.eOF then
            ImgUrl = IIF(rs("BigFiles") <> "", rs("BigFiles"), rs("SmallFiles")) 
        end if : rs.close 
    end if 

    XY_DetailBanner = IIF(ImgUrl <> "", ImgUrl, DefaultImage) 
    XY_DetailBanner = HandleHttpUrl(XY_DetailBanner) 

    if 1 = 2 then
        call echo("PubNavDid", PubNavDid) 
        call echo("PubNavSid", PubNavSid) 
        call echo("PubProDid", PubProDid) 
        call echo("PubProSid", PubProSid) 
        call echo("PubProTid", PubProTid) 

        call echo("PubActionType", PubActionType) 
        call echo("PubID", PubID) 
        call echo("PubClassName", PubClassName) 
        call echo("PubWD", PubWD) 
        call echo("ImgUrl", ImgUrl) 
    end if 

end function
 

'Banner切换
function BannerSwitch(byVal did, byVal TopNumb)
    dim i, c, addSql, sql, url, CutStrNumb 
    CutStrNumb = 22 
    if lCase(left(did, 4)) = "sql:" then
        addSql = mid(did, 5) 
    else
        addSql = "Where BigClassName='" & did & "' And Recommend<>0 Order By Sort" 
    end if 
    c = c & "<div class=""bannerlb pr"" id=""flashs"">" & vbCrLf 
    sql = "Select * From [Product] " & addSql 
    rs.open sql, conn, 1, 1 
    for i = 1 to rs.recordCount
        if i > TopNumb then exit for 

        url = rs("FileName") 
        if CheckRemoteUrl(url) = false then
            url = GetHandleUrl(rs("FileName")) 
            call WebDebug(url, "act=CreateArticle&ID=" & rs("Id"))                      '调试
        end if 

        ShowTitle = CutStr(rs("Title"), CutStrNumb, "... ") 
        ShowTitle = FontColorFontB(ShowTitle, rs("FontB"), rs("FontColor")) 

        c = c & "    <div title=""" & ShowTitle & """ alt=""" & ShowTitle & """ onClick=""window.open('" & url & "')"" class=""bgitem"" id=""flashbg" & i - 1 & """  style=""background: url('" & rs("BigFiles") & "') no-repeat scroll center top;" & vbCrLf 
        c = c & "height: 496px;cursor:pointer; width: 100%; margin:0 auto;""></div>" & vbCrLf & vbCrLf 


    rs.moveNext : next : rs.close 

          'C=C & "<script type=""text/javascript"" src=""/Skins/WeiZhanLue/Index/flash.js""></'script>" & vbCrlf
    c = c & "</div> " & vbCrLf 

    BannerSwitch = c 
end function
 
'Banner切换2
function BannerSwitch2(byVal did, byVal TopNumb)
    dim i, c 
    rs.open "Select * From[Product] Where BigClassName='" & did & "'", conn, 1, 1 
    c = c & "<div class=""container"" id=""idTransformView2"">" & vbCrLf 
    c = c & "  <ul class=""slider slider2"" id=""idSlider2"">" & vbCrLf 
    for i = 1 to TopNumb
        if rs.eOF then TopNumb = i - 1 : exit for 
        c = c & "    <li><img src=""" & rs("BigFiles") & """/></li>" & vbCrLf 
    rs.moveNext : next : rs.close 
    c = c & "  </ul>" & vbCrLf 
    c = c & "  <ul class=""num"" id=""idNum2"">" & vbCrLf 

    for i = 1 to TopNumb
        c = c & "    <li>" & i & "</li>" & vbCrLf 
    next 
    c = c & "  </ul>" & vbCrLf 
    c = c & "</div>" & vbCrLf 
    BannerSwitch2 = c 
end function

%>
