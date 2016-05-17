<% 


'技巧：计算图片等比率缩放方法 宽=213,高=160  缩放后宽=100,高=75   100除213得到小数后两位，四舍五入
'document.write("\n<br>显示" + 100/213)
'document.write("\n<br>计算" + 160*0.47)


'用法
'Call Rw(HandleShowImg("UploadFiles/Soso.gif",100,60,""))
'Call Rw(HandleShowImg("UploadFiles/Soso.gif",100,60,"高"))
'Call Rw(HandleShowImg("UploadFiles/Soso.gif",100,60,"宽"))
'获得图片类型    getFileContentType("/aa.jpg")

'VB不引用 Start
'psd不行，等加
Class GPS
    Dim aso 
    Private Sub class_Initialize()
        Set aso = CreateObject("ADODB.Stream")
            aso.Mode = 3 
            aso.Type = 1 
            aso.Open 
    End Sub
    Private Sub class_Terminate()
        Set aso = Nothing 
    End Sub 
    Private Function bin2Str(Bin)
        Dim g, str, clow 
        For g = 1 To LenB(Bin)
            clow = MidB(Bin, g, 1) 
            If AscB(clow) < 128 Then
                str = str & Chr(AscB(clow)) 
            Else
                g = g + 1 
                If g <= LenB(Bin) Then str = str & Chr(AscW(MidB(Bin, g, 1) & clow)) 
            End If 
        Next 
        bin2Str = str 
    End Function 
    Private Function binVal(bin)
        Dim ret 
        Dim h 
        ret = 0 
        For h = LenB(bin) To 1 Step - 1
            ret = ret * 256 + AscB(MidB(bin, h, 1)) 
        Next 
        binVal = ret 
    End Function 
    Private Function binVal2(bin)
        Dim ret, g 
        ret = 0 
        For g = 1 To LenB(bin)
            ret = ret * 256 + AscB(MidB(bin, g, 1)) 
        Next 
        binVal2 = ret 
    End Function 
    '///以下是调用代码/// 获得图片文件大小
    Function getImageSize(filespec)
        Dim bFlag, p1 
        If checkFile(filespec) = False Then Exit Function                               '图片文件不存 则退出

        aso.loadFromFile(filespec) 
        'On Error Resume Next
        Dim ret(3) 
        'call echo("filespec",filespec)
        bFlag = aso.read(3) 
        'Call Echo("bFlag1",Hex(binVal(bFlag)))
        Select Case Hex(binVal(bFlag))
            Case "4E5089"
                aso.read(15) 
                ret(0) = "PNG" 
                ret(1) = binVal2(aso.read(2)) 
                aso.read(2) 
                ret(2) = binVal2(aso.read(2)) 
            Case "464947"
                aso.read(3) 
                ret(0) = "GIF" 
                ret(1) = binVal(aso.read(2)) 
                ret(2) = binVal(aso.read(2)) 
            Case "FFD8FF"
                Do
                Do : p1 = binVal(aso.read(1)) : Loop While p1 = 255 And Not aso.EOS
                If p1 > 191 And p1 < 196 Then Exit Do Else aso.read(binval2(aso.read(2)) - 2) 
                Do : p1 = binVal(aso.read(1)) : Loop While p1 < 255 And Not aso.EOS 
                Loop While True 
                aso.read(3) 
                ret(0) = "JPG" 
                ret(2) = binval2(aso.read(2)) 
                ret(1) = binval2(aso.read(2)) 
            Case Else
                If Left(bin2Str(bFlag), 2) = "BM" Then
                    aso.read(15) 
                    ret(0) = "BMP" 
                    ret(1) = binval(aso.read(4)) 
                    ret(2) = binval(aso.read(4)) 
                Else
                    ret(0) = "" 
                End If
        End Select
        ret(3) = "width=""" & ret(1) & """ height=""" & ret(2) & """" 
        getImageSize = ret 
    End Function 
End Class 
'VB不引用 End

'处理显示图片
Function handleShowImg(imagename, setWidth, setHeight, SType)
    Dim filePath 
    Dim fs 

    filePath = HandlePath(imagename) 
    '/////建立GPS类实体/////

    Dim GetPicSize 
    Set GetPicSize = new GPS
        Set fs = CreateObject("Scripting.FileSystemObject")
            '/////获取图片类型/////
            Dim PicSuffixName 
            PicSuffixName = fs.GetExtensionName(filePath) 
            Dim PD                                          '//Picture Dimension
            Dim PWidth, PHeight 
            Select Case PicSuffixName
                Case "gif", "bmp", "jpg", "png"
                    '/////调用GPS通用类中的GetImageSize函数获取图片尺寸/////

                    PD = GetPicSize.getImageSize(filePath) 
                    PWidth = PD(1)                         '//获取图片宽度
                    PHeight = PD(2)                        '//获取图片高度
            End Select
        Set fs = Nothing 
    Set GetPicSize = Nothing 
    Dim PXWidth, PXHeight 
    Dim Pp                                                                          '//长宽比

    Call echo("PWidth", PWidth) 
    Call echo("PHeight", PHeight) 

    Call echo("setWidth", setWidth) 
    Call echo("setHeight", setHeight) 

    '获得图片宽为零或为空则
    If PWidth = 0 Or PWidth = "" Then
        PXWidth = 0 
        PXHeight = 0 
    Else
        If setHeight < setWidth And setHeight > 0 Then
            Pp = FormatNumber(PHeight / PWidth, 2)          '//高宽 百分比

            PXHeight = setHeight 
            PXWidth = FormatNumber(setHeight * Pp, 0) 
        Else
            Pp = FormatNumber(PWidth / PHeight, 2)          '//长宽 百分比

            PXWidth = setWidth 
            PXHeight = FormatNumber(setWidth / Pp, 0) 
            Call echo(Pp, setWidth & "   " & setWidth / Pp) 
        End If 
    End If 

    Call echo("宽百分比", FormatNumber(PWidth / PHeight, 2)) 
    Call echo("高百分比", FormatNumber(PHeight / PWidth, 2)) 

    Call echo("PXWidth", PXWidth) 
    Call echo("PXHeight", PXHeight) 

    If SType = "宽" Then
        If setWidth < PXWidth Then PXWidth = setWidth 
        handleShowImg = PXWidth 
    ElseIf SType = "高" Then
        If setHeight < PXWidth Then PXWidth = setHeight 
        handleShowImg = PXWidth 
    Else
        handleShowImg = "<img onerror=javascript:this.src='product_pic/wutu.jpg';this.width='92';this.height='130' src='" & imagename & "' border='0' width='" & PXWidth & "' height='" & PXHeight & "'>显示图片" 
    End If 
    'response.Write "<br>"
    'response.write PXWidth&"X"&PXHeight  '当前图片显示的尺寸
    'response.Write "<br>"
    'response.write PWidth&"X"&PHeight  '图片原始尺寸
'onload='javascript:if(this.width>940) this.width=940;'                加载完后处理
End Function 
'获得图片宽高
Function getImageWidthHeight(filePath, width, height)
    Dim GetPicSize, fs 
    Call HandlePath(filePath)                                                       '获得完整路径
    Set GetPicSize = new GPS
        Set fs = CreateObject("Scripting.FileSystemObject")
            Dim PicSuffixName 
            PicSuffixName = fs.GetExtensionName(filePath) 
            Dim PD 
            Select Case PicSuffixName
                Case "gif", "bmp", "jpg", "png"
                    PD = GetPicSize.getImageSize(filePath) 
                    width = PD(1)                          '//获取图片宽度
                    height = PD(2)                         '//获取图片高度
            End Select
        Set fs = Nothing 
    Set GetPicSize = Nothing 
	getImageWidthHeight=split(width & "|" & height,"|")
End Function 
'获得图片宽值
Function getImageWidthValue(filePath)
    Dim width 
    Call GetImageWidthHeight(filePath, width, "") 
    getImageWidthValue = width 
End Function 
'获得文件内容类型
Function getFileContentType(filePath)
    Dim GetPicSize, fs 
    Call HandlePath(filePath)                                                       '获得完整路径
    Set GetPicSize = new GPS
        Set fs = CreateObject("Scripting.FileSystemObject")
            Dim PD 
            PD = GetPicSize.getImageSize(filePath) 
            getFileContentType = PD(0)                      '//获取图片类型
        Set fs = Nothing 
    Set GetPicSize = Nothing 
End Function
'获得图片类型 
function getImageType(filePath)
	getImageType = getFileContentType(filePath)
end function
'瀑布流方式显示图片列表
Function imageWaterfall(folderPath, nRow, SetWidth)
    Dim content, splStr, i, ImgPath, c, s, nProportion, width, height, thisI, nNumb 
    Dim SplC(99), SplN(99) 
    nRow = nRow - 1                                                                 '减一，循环从1开始就不显示内容，晕
    '初始化
    For i = 0 To nRow
        SplC(i) = "" 
        SplN(i) = 0 
    Next 
    'Content = GetDirFileNameList(FolderPath,"")
    content = GetFileFolderList(folderPath, True, "全部", "", "全部文件夹", "NO", "") 
    content = Replace(LCase(content), LCase(Server.MapPath("/")), "")               '强力替换路径

    splStr = Split(content, vbCrLf) : folderPath = Replace(folderPath, "\", "/") 
    For Each ImgPath In splStr
        s = "<span><img src=" & host & ImgPath & " width='" & SetWidth & "'></span>" 
        Call GetImageWidthHeight(ImgPath, width, height) 
        If width > SetWidth Then
            nProportion = FormatNumber(SetWidth / width, 2) 
            width = SetWidth 
            height = FormatNumber(height * nProportion, 0) 
        End If 
        thisI = 0 : nNumb = "" 
        For i = 0 To nRow
            If nNumb = "" Then
                nNumb = SplN(i) 
            ElseIf SplN(i) < nNumb Then
                thisI = i 
                nNumb = SplN(i) 
            End If 
        Next 
        SplN(thisI) = SplN(thisI) + height 
        SplC(thisI) = SplC(thisI) & s 
    Next 
    For i = 0 To nRow
        c = c & "<div class=""ImgDivBox"">" & SplC(i) & "</div>" & vbCrLf 
    Next 
    c = c & "<div class=clear></div>" 
    imageWaterfall = c 
End Function 

'图片等比例缩放(20151218)  自动加宽 加高  加padding-left 加padding-top    例 Call rw("<img src='"& tempFilePath &"'" & zoomImage(imgWidth, imgHeight, 35, 35, 0, 0, "") & " />")
Function zoomImage(nWrapWidth,nWrapHeight, imgWidth, imgHeight, nSetWidth, nSetHeight, nLeft, nTop, sType)
    Dim htmlStr, css, nNewWidth, nNewHeight
	
	
    If nWrapWidth = "" Then nWrapWidth = 0 
    If nWrapHeight = "" Then nWrapHeight = 0 
	nWrapWidth=cint(nWrapWidth) : nWrapHeight=cint(nWrapHeight)
	 
    If nLeft = "" Then nLeft = 0 
    If nTop = "" Then nTop = 0  
	imgWidth=cint(imgWidth) : imgHeight=cint(imgHeight) : nSetWidth=cint(nSetWidth) : nSetHeight=cint(nSetHeight) : nLeft=cint(nLeft) : nTop=cint(nTop)
	
	dim temp_nLeft, temp_nTop				'暂存
	temp_nLeft=nLeft : temp_nTop=nTop
	
    If imgWidth / imgHeight >= nSetWidth / nSetHeight Then
        If imgWidth > nSetWidth Then
            nNewWidth = nSetWidth 
            nNewHeight =(imgHeight * nSetWidth) / imgWidth 
        Else
            nNewWidth = imgWidth 
            nNewHeight = imgHeight 
        End If 
    Else
        If imgHeight > nSetHeight Then
            nNewHeight = nSetHeight 
            nNewWidth =(imgWidth * nSetHeight) / imgHeight 
        Else
            nNewWidth = imgWidth 
            nNewHeight = imgHeight 
        End If 
    End If

    nNewWidth = Int(nNewWidth) 
    nNewHeight = Int(nNewHeight) 
    nLeft = nLeft + nSetWidth - nNewWidth 
    nTop = nTop + nSetHeight - nNewHeight 
    htmlStr = " width=""" & nNewWidth & """ height=""" & nNewHeight & """" 
  
    if nWrapHeight>0 then
		'call echo("nWrapHeight",nWrapHeight)
			if instr(sType,"|imgulmiddel|")>0 then
				nTop=(nWrapHeight-nNewHeight)/2 
			end if 
	end if
	
	
	
	
	
	
	
    If nLeft > 0 and nLeft<>temp_nLeft Then
        css = css & "padding-left:" & nLeft & "px;" 
    End If 
    If nTop > 0 and nTop<>temp_nLeft Then
        css = css & "padding-top:" & nTop & "px;" 
    End If 
    'Call echo("nLeft", nLeft) 
    'Call echo("nTop", nTop) 
    If css <> "" Then
        css = " style=""" & css & """" 
    End If 
    zoomImage = htmlStr & css 
End Function

%>  
