<%

Class xmlZIP
    Dim zipPathDir, zipPathFile, zipFileExt 
    Dim starTime, endTime 
    Dim pubXmlZipDir                                                                'xml打包目录
    Dim pubIsEchoMsg                                                                '是否回显打印信息
    Dim webRootPathDir                                                              '网站根目录路径

    Sub run(thisWebRootPathDir, xmlZipDir, isEchoMsg, xmlFileName)
        pubXmlZipDir = xmlZipDir & "/"                                                  'xml目录
        pubIsEchoMsg = isEchoMsg                                                        '是否回显
        '此为默认当前文件夹
        webRootPathDir = thisWebRootPathDir                                             '等于它
        'call echo("webRootPathDir",webRootPathDir)

        zipPathDir = Server.MapPath("./") & "\" 

        '在此更改要打包文件夹的路径
        'ZipPathDir="D:\MYWEB\WEBINFO"
        '生成的xml文件
        zipPathFile = xmlFileName 
        If zipPathFile = "" Then
            zipPathFile = "update.xml" 
        End If 
        '不进行打包的文件扩展名
        zipFileExt = "db;bak" 
        If Right(zipPathDir, 1) <> "\" Then zipPathDir = zipPathDir & "\" 

        Response.Write("保存路径：" & zipPathDir & zipPathFile & "<hr>") 
        '开始打包
        CreateXml(zipPathDir & zipPathFile) 
    End Sub 
    '遍历目录内的所有文件以及文件夹
    Sub loadData(dirPath)
        Dim xmlDoc 
        Dim fso                                                                         'fso对象
        Dim objFolder                                                                   '文件夹对象
        Dim objSubFolders                                                               '子文件夹集合
        Dim objSubFolder                                                                '子文件夹对象
        Dim objFiles                                                                    '文件集合
        Dim objFile                                                                     '文件对象
        Dim objStream 
        Dim pathname, xFolder, xFPath, xFile, xPath, xStream 
        Dim pathNameStr 

        If pubIsEchoMsg = True Then
            Response.Write("==========" & dirPath & "==========<br>") 
        End If 
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
            Set objFolder = fso.GetFolder(dirPath)                                          '创建文件夹对象

                If pubIsEchoMsg = True Then
                    Response.Write dirPath 
                    Response.flush 
                End If 
                Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
                    xmlDoc.load(Server.MapPath(zipPathFile)) 
                    xmlDoc.async = False 
                    '写入每个文件夹路径
                    Set xFolder = xmlDoc.SelectSingleNode("//root").AppendChild(xmlDoc.CreateElement("folder"))
                        Set xFPath = xFolder.AppendChild(xmlDoc.CreateElement("path"))
                            xFPath.text = Replace(dirPath, webRootPathDir, "") 'zipPathDir  注意
                            Set objFiles = objFolder.files
                                For Each objFile In objFiles
                                    If LCase(dirPath & objFile.Name) <> LCase(Request.ServerVariables("PATH_TRANSLATED")) And LCase(dirPath & objFile.Name) <> LCase(dirPath & zipPathFile) Then
                                        If ext(objFile.Name) Then
                                            pathNameStr = dirPath & "" & objFile.Name 
                                            If pubIsEchoMsg = True Then
                                                Response.Write "---<br/>" 
                                                Response.Write pathNameStr & "" 
                                                Response.flush 
                                            End If 
                                            '================================================

                                            '写入文件的路径及文件内容
                                            Set xFile = xmlDoc.SelectSingleNode("//root").AppendChild(xmlDoc.CreateElement("file"))
                                                Set xPath = xFile.AppendChild(xmlDoc.CreateElement("path"))
                                                    'call echo(zipPathDir,webRootPathDir)
                                                    xPath.text = Replace(pathNameStr, webRootPathDir, "") 'zipPathDir  注意
                                                    '创建文件流读入文件内容，并写入XML文件中
                                                    Set objStream = Server.CreateObject("ADODB.Stream")
                                                        objStream.Type = 1 
                                                        objStream.Open() 
                                                        objStream.loadFromFile(pathNameStr) 
                                                        objStream.position = 0 
                                                        Set xStream = xFile.AppendChild(xmlDoc.CreateElement("stream"))
                                                            xStream.SetAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes" 
                                                            '文件内容采用二制方式存放
                                                            xStream.dataType = "bin.base64" 
                                                            xStream.nodeTypedValue = objStream.read() 
                                                        Set objStream = Nothing 
                                                    Set xPath = Nothing 
                                                Set xStream = Nothing 
                                            Set xFile = Nothing 
                                            '================================================

                                        End If 
                                    End If 
                                Next 
                                If pubIsEchoMsg = True Then
                                    Response.Write "<p>" 
                                End If 
                                xmlDoc.Save(Server.MapPath(zipPathFile)) 
                            Set xFPath = Nothing 
                        Set xFolder = Nothing 
                    Set xmlDoc = Nothing 
                    '创建的子文件夹对象
                    Set objSubFolders = objFolder.subFolders
                        '调用递归遍历子文件夹
                        For Each objSubFolder In objSubFolders
                            pathname = dirPath & objSubFolder.Name & "\" 
                            loadData(pathname) 
                        Next 
                    Set objFolder = Nothing 
                Set objSubFolders = Nothing 
            Set fso = Nothing 
    End Sub
    '创建一个空的XML文件，为写入文件作准备
    Sub createXml(filePath)
        '程序开始执行时间
        starTime = Timer() 
        Dim xmlDoc, root 
        Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
            xmlDoc.async = False 
            Set root = xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
                xmlDoc.appendChild(root) 
                xmlDoc.appendChild(xmlDoc.CreateElement("root")) 
                If InStr(filePath, ":") = False Then filePath = Server.MapPath(filePath) 
                xmlDoc.Save(filePath) 
            Set root = Nothing 
        Set xmlDoc = Nothing 
        'call eerr(ZipPathDir & "inc/",pubXmlZipDir)
        'call echo(ZipPathDir & "newweb/http：揦揦www。thinkphp。cn揦/web/",pubXmlZipDir)
        loadData(pubXmlZipDir) 
        '程序结束时间
        endTime = Timer() 
        Response.Write("页面执行时间：" & FormatNumber((endTime - starTime), 3) & "秒") 
    End Sub 
    '判断文件类型是否合法
    Function ext(fileName)
        ext = True 
        Dim temp_ext, e 
        temp_ext = Split(zipFileExt, ";") 
        For e = 0 To UBound(temp_ext)
            If Mid(fileName, InStrRev(fileName, ".") + 1) = temp_ext(e) Then ext = False 
        Next 
    End Function 
End Class 
%>