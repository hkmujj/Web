<% 
'FSO  Scripting.FileSystemObject组件操作大全 (2013,9,27)

'读文件内容 (2013,9,27
Function getFileText(ByVal filePath)
    On Error Resume Next 
    Dim fso, fText, OpenFile 
    'GetFileText = ""   '它默认返回的就是空， 这个是多此一举 (2013,9,30)
    Call handlePath(filePath)                                                       '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(filePath) = True Then
            Set fText = fso.openTextFile(filePath, 1)
                '加强 读空文件出错
                Set OpenFile = fso.GetFile(filePath)
                    If OpenFile.Size = 0 Then getFileText = "" : Exit Function                      '文件为空则退出
                Set OpenFile = Nothing 
                getFileText = fText.ReadAll 
            Set fText = Nothing 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "GetFileText 读取文件内容 函数出错，filePath=" & filePath) 
End Function 
'旧版,读取文件内容  (辅助)
Function getFText(filePath)
    getFText = getFileText(filePath) 
End Function 
'读文件类型 (2013,9,27
Function getFileType(ByVal filePath)
    On Error Resume Next 
    Dim fso, myFile, OpenFile 
    'GetFileText = ""   '它默认返回的就是空， 这个是多此一举 (2013,9,30)
    Call handlePath(filePath)                                                       '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(filePath) = True Then
            Set myFile = fso.GetFile(filePath)
                getFileType = myFile.Type 
        End If
            Set fso = Nothing
            If Err Then Call doError(Err.Description, "GetFileText 读取文件类型 函数出错，filePath=" & filePath) 
End Function

'检查文件
Function checkFile(ByVal filePath)
    On Error Resume Next 
    Dim fso 
    Call handlePath(filePath)                                                       '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        checkFile = fso.fileExists(filePath) 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "CheckFile 检查文件 函数出错，filePath=" & filePath) 
End Function 
'判断文件,相当于检查文件 (辅助)
Function existsFile(filePath)
    existsFile = checkFile(filePath) 
End Function 
'创建文件   这种方式创建文件时，会多一行，用WriteToFile创建时没有多一行
Function createFile(ByVal filePath, ByVal content)
    On Error Resume Next 
    Dim fText, fso 
    Call handlePath(filePath)                                                       '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If existsZhiDuFile(filePath) = True Then                                        '判断是否为只读文件
            Call EditFileAttribute(filePath, 32)                                            '把只读属性改成存档属性
        End If 
        'call echo("filePath",filePath)
        Set fText = fso.createTextFile(filePath, True)
            fText.WriteLine(content) 
            createFile = True 
        Set fText = Nothing 
    Set fso = Nothing 
    If Err Then createFile = False : Call doError(Err.Description, "CreateFile 创建文件 函数出错，filePath=" & filePath) 
End Function 
'保存文件 (辅助)
Function saveFile(filePath, content)
    saveFile = createFile(filePath, content) 
End Function 
'创建文件，存在累加 【2013,9,27】
Function createAddFile(ByVal filePath, ByVal content)
    On Error Resume Next 
    Dim fso, fText 
    Call handlePath(filePath)                                                       '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(filePath) = False Then                                        '文件不存在，则创建
            Set fText = fso.createTextFile(filePath, True)
                fText.WriteLine(content) 
            Set fText = Nothing 
            createAddFile = "创建文件成功" 
        Else                                                                            '文件存在，则累加
            createAddFile = addToFile(filePath, content) 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "CreateAddFile 创建文件，存在累加 函数出错，filePath=" & filePath) 
End Function 
'创建文件，旧版
Function createAddUpFile(filePath, content)
    createAddUpFile = createAddFile(filePath, content) 
End Function 
'追加文件信息
Function addToFile(ByVal filePath, ByVal content)
    On Error Resume Next 
    Dim fso, fText 
    Call handlePath(filePath)                                                       '获得完整路径
    If existsZhiDuFile(filePath) = True Then                                        '判断是否为只读文件
        Call EditFileAttribute(filePath, 32)                                            '把只读属性改成存档属性
    End If 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(filePath) = True Then
            Set fText = fso.openTextFile(filePath, 8)
                fText.WriteLine(content) 
            Set fText = Nothing 
            addToFile = "追加信息成功" 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "AddToFile 追加文件信息 函数出错，filePath=" & filePath & "<br>Content字符长" & Len(content) & " Content类型" & TypeName(content) & "【" & content & "】") 
End Function 
'保存字节文件(20151106)
Function saveByteFile(ByVal filePath, content)
    Dim ads 
    filePath = handlePath(filePath) 
    Set ads = CreateObject("ADODB.Stream")
        With ads
            .Type = 1 
            .Open 
            .Write content 
            'Response.Write("<br><b>["&LocalFileName&"]</b><br>")
            .SaveToFile filePath, 2 
            .Cancel 
            .Close 
        End With 
    Set ads = Nothing 
End Function 
Function createByteFile(filePath, content)

End Function 
'移动文件
Function moveFile(ByVal path, ByVal ToPath)
    On Error Resume Next 
    Dim fso 
    Call handlePath(path)                                                           '获得完整路径
    Call handlePath(ToPath)                                                         '获得完整路径
    moveFile = False 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(path) = True And fso.fileExists(ToPath) = False Then          '要移动文件为真，被移动文件为假
            Call fso.moveFile(path, ToPath) 
            moveFile = True 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "MoveFile 移动文件 函数出错，Path=" & path & "，ToPath=" & ToPath) 
End Function 
'修改文件，相当于移动文件 (辅助)
Function editFile(path, ToPath)
    editFile = moveFile(path, ToPath) 
End Function 
'拷贝文件
Function copyFile(ByVal path, ByVal ToPath)
    On Error Resume Next 
    Dim fso 
    Call handlePath(path)                                                           '获得完整路径
    Call handlePath(ToPath)                                                         '获得完整路径
    copyFile = False 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(path) = True And fso.fileExists(ToPath) = False Then          '要拷贝文件为真，被拷贝文件为假
            Call fso.copyFile(path, ToPath) 
            copyFile = True 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "CopyFile 拷贝文件 函数出错，Path=" & path & "，ToPath=" & ToPath) 
End Function 
'删除文件，只读文件也可删除，谨慎使用
Function deleteFile(ByVal filePath)
    On Error Resume Next 
    Dim fso 
    Call handlePath(filePath)                                                       '获得完整路径
    deleteFile = False 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(filePath) = True Then
            Call fso.deleteFile(filePath, True) 
            deleteFile = True 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "DeleteFile 删除文件 函数出错，filePath=" & filePath) 
End Function 
'删除文件列表
Function delSelFile(RootPath, ByVal FileList)
    Dim splStr, path, c 
    c = "" 
    splStr = Split(FileList, "|") 
    For Each path In splStr
        If path <> "" Then
            '删除文件
            Call deleteFile(RootPath & path) 
            c = c & path & vbCrLf 
        End If 
    Next 
    delSelFile = c 
End Function 
Function delFileList(RootPath, ByVal FileList)
    delFileList = delSelFile(RootPath, FileList) 
End Function 
'批量删除文件(20151127) Call batchDelFile("/test/", Split("2.html|2.php|3.Asp", "|"))
Function batchDelFile(dirPath, fileArray)
    Dim fileName, filePath, c 
    For Each fileName In fileArray
        filePath = dirPath & fileName 
        c = c & filePath & vbCrLf 
        Call deleteFile(filePath) 
    Next 
    batchDelFile = c 
End Function 
'删除指定文件列表  前面有#号的文件，给删除掉 例：Call DeleteAppointFileList("/Skins/Test/Css/")
Function deleteAppointFileList(ByVal folderPath)
    'On Error Resume Next
    Dim fso, f, myFile, fileName 
    Call handlePath(folderPath)                                                     '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(folderPath) = False Then deleteAppointFileList = "" : Exit Function '文件夹不存在，退出
        Set f = fso.GetFolder(folderPath)
            For Each myFile In f.files
                fileName = baseName(myFile) 
                'Call Echo(FileName,MyFile)
                If Left(fileName, 1) = "#" Then
                    Call deleteFile(myFile) 
                End If 
                doEvents 
            Next 
        Set f = Nothing 
    Set fso = Nothing 
    deleteAppointFileList = "" 
End Function 
'获得文件大小
Function getFileSize(ByVal fileName)
    On Error Resume Next 
    Dim fso, OpenFile 
    Call handlePath(fileName)                                                       '获得完整路径
    getFileSize = 0 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(fileName) = True Then
            Set OpenFile = fso.GetFile(fileName)
                getFileSize = OpenFile.Size 
            Set OpenFile = Nothing 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "GetFileSize 获得文件大小 函数出错，FileName=" & fileName) 
End Function 
'文件大小 (辅助)
Function fileSize(fileName)
    fileSize = getFileSize(fileName) 
End Function 
'文件大小 (辅助)
Function getFSize(fileName)
    getFSize = getFileSize(fileName) 
End Function 
'获得文件创建时间
Function getFileCreateDate(ByVal fileName)
    On Error Resume Next 
    Dim fso, OpenFile 
    Call handlePath(fileName)                                                       '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(fileName) = True Then
            Set OpenFile = fso.GetFile(fileName)
                getFileCreateDate = OpenFile.DateCreated 
            Set OpenFile = Nothing 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "GetFileCreateDate 获得文件创建时间 函数出错，FileName=" & fileName) 
End Function 
'获得文件创建时间 (辅助)
Function fileCreateDate(fileName)
    fileCreateDate = getFileCreateDate(fileName) 
End Function 
'获得文件修改时间
Function getFileEditDate(ByVal fileName)
    On Error Resume Next 
    Dim fso, OpenFile 
    Call handlePath(fileName)                                                       '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(fileName) = True Then
            Set OpenFile = fso.GetFile(fileName)
                getFileEditDate = OpenFile.DateLastModified 
            Set OpenFile = Nothing 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "GetFileEditDate 获得文件修改时间 函数出错，FileName=" & fileName) 
End Function 
'获得文件修改时间 (辅助)
Function fileEditDate(fileName)
    fileEditDate = getFileEditDate(fileName) 
End Function 
'获得文件修改时间 (辅助)
Function getFileEditTime(fileName)
    getFileEditTime = getFileEditDate(fileName) 
End Function 
'获得文件访问时间
Function getDateLastAccessed(ByVal fileName)
    On Error Resume Next 
    Dim fso, OpenFile 
    Call handlePath(fileName)                                                       '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(fileName) = True Then
            Set OpenFile = fso.GetFile(fileName)
                getDateLastAccessed = OpenFile.DateLastAccessed 
            Set OpenFile = Nothing 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "GetFileEditDate 获得文件访问时间 函数出错，FileName=" & fileName) 
End Function 
'修改文件最后修改时间 如：Call ModifyFileModifydate("F:\Temp\001\Config.ini", Now() - 5)
Function modifyFileModifydate(ByVal fileName, ByVal dateTime)
    On Error Resume Next 
    Dim dirPath, ShellObj, App_Path, App_File 
    Call handlePath(fileName)                                                       '获得完整路径
    If checkFile(fileName) = False Then                                             '文件不存在，为假退出
        modifyFileModifydate = False 
        Exit Function 
    End If 
    dirPath = Left(fileName, InStrRev(fileName, "\")) 
    fileName = Right(fileName, Len(fileName) - Len(dirPath)) 
    'Call Echo(DirPath,FileName)
    '通过shell.Application修改文件的最后修改时间
    Set ShellObj = CreateObject("Shell.Application")
        Set App_Path = ShellObj.NameSpace(dirPath)
            Set App_File = App_Path.ParseName(fileName)
                App_File.Modifydate = dateTime 
                modifyFileModifydate = True 
                If Err Then Call doError(Err.Description, "ModifyFileModifydate 获得文件修改时间 函数出错，FileName=" & dirPath & fileName) 
End Function
'修改文件最后修改时间 (辅助)
Function editFileEditDate(fileName, dateTime)
    editFileEditDate = modifyFileModifydate(fileName, dateTime) 
End Function 
'判断只读文件
Function existsZhiDuFile(ByVal fileName)
    On Error Resume Next 
    Dim fso, OpenFile 
    existsZhiDuFile = False 
    Call handlePath(fileName)                                                       '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(fileName) = True Then
            Set OpenFile = fso.GetFile(fileName)
                existsZhiDuFile = OpenFile.Attributes 
            Set OpenFile = Nothing 
            If existsZhiDuFile >= 128 Then existsZhiDuFile = existsZhiDuFile - 128 
            If existsZhiDuFile >= 64 Then existsZhiDuFile = existsZhiDuFile - 64 
            If existsZhiDuFile >= 32 Then existsZhiDuFile = existsZhiDuFile - 32 
            If existsZhiDuFile >= 8 Then existsZhiDuFile = existsZhiDuFile - 8 
            If existsZhiDuFile >= 4 Then existsZhiDuFile = existsZhiDuFile - 4 
            If existsZhiDuFile >= 2 Then existsZhiDuFile = existsZhiDuFile - 2 
            existsZhiDuFile = CBool(existsZhiDuFile)                                        '让它成为布而类型
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "ExistsZhiDuFile 判断只读文件 函数出错，FileName=" & fileName) 
End Function 
'获得文件属性  '0普通文件 1只读 2为隐藏 4系统 8卷材 16文件夹 32存档 64快捷键 128压缩文件
Function getFileAttribute(ByVal fileName)
    On Error Resume Next 
    Dim fso, OpenFile 
    Call handlePath(fileName)                                                       '获得完整路径
    getFileAttribute = -1 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(fileName) = True Then
            Set OpenFile = fso.GetFile(fileName)
                getFileAttribute = OpenFile.Attributes 
            Set OpenFile = Nothing 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "GetFileAttribute 获得文件属性 函数出错，FileName=" & fileName) 
End Function 
'修改文件属性  '0普通文件 1只读 2为隐藏 4系统 8卷材 16文件夹 32存档 64快捷键 128压缩文件
Function editFileAttribute(ByVal fileName, ByVal n)
    On Error Resume Next 
    Dim fso, OpenFile 
    Call handlePath(fileName)                                                       '获得完整路径
    editFileAttribute = False 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(fileName) = True Then
            Set OpenFile = fso.GetFile(fileName)
                OpenFile.Attributes = n                                                         '32文件
            Set OpenFile = Nothing 
            editFileAttribute = True 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "EditFileAttribute 修改文件属性 函数出错，FileName=" & fileName) 
End Function 
'判断文件天数（用修改时间进行比较）
Function existsFileDay(ByVal fileName, ByVal InputDayNumb)
    On Error Resume Next 
    Dim fso, OpenFile, DayNumb 
    Call handlePath(fileName)                                                       '获得完整路径
    existsFileDay = False 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.fileExists(fileName) = True Then
            Set OpenFile = fso.GetFile(fileName)
                DayNumb = DateDiff("d", OpenFile.DateLastModified, Now())                       '之前是Date  vb.net不能用
                If DayNumb <= InputDayNumb Then existsFileDay = True 
            Set OpenFile = Nothing 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "ExistsFileDay 判断文件天数 函数出错，FileName=" & fileName) 
End Function 
'判断文件天数 (辅助)
Function checkFileDay(fileName, InputDayNumb)
    checkFileDay = existsFileDay(fileName, InputDayNumb) 
End Function 
'========================================= 文件夹 =========================================
'创建文件夹
Function createFolder(ByVal folderPath)
    On Error Resume Next 
    Dim fso 
    Call handlePath(folderPath)                                                     '获得完整路径
    createFolder = False 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(folderPath) = False Then
            fso.createFolder(folderPath) 
            createFolder = True 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "CreateFolder 创建文件夹 函数出错，FolderPath=" & folderPath & ":::" & fso.FolderExists(folderPath)) 
End Function 
'创建目录文件夹 创建一组文件夹
Function createDirFolder(ByVal folderPath)
    'On Error Resume Next
    createDirFolder = "" 
    If checkFolder(folderPath) = True Then Exit Function                            '文件夹存在，退出
    Dim splStr, i, s, NewFolder 
    NewFolder = ""                                                                  'vb.net使用报错
    Call handlePath(folderPath)                                                     '获得完整路径
    splStr = Split(folderPath, "\") 
    For i = 0 To UBound(splStr)
        s = Trim(splStr(i))                                                             '去除两边空格
        If s <> "" Then
            NewFolder = NewFolder & s & "\" 
            If i = UBound(splStr) And InStr(s, ".") > 0 Then Exit For                       '创建到最后文件夹时，要是为文件则退出循环 20141215
            If InStr(s, ":") = False Then
                'Call Echo("NewFolder",NewFolder)
                Call CreateFolder(NewFolder) 
            End If 
        End If 
    Next 
End Function 
'创建以域名命名的文件夹
Function createDomain(ByVal folderPath, ByVal httpurl)
    On Error Resume Next 
    Dim ToPath 
    Call handlePath(folderPath)                                                     '获得完整路径
    createDomain = False 
    httpurl = getWebSite(httpurl) 
    httpurl = Replace(Replace(Replace(httpurl, "http://", ""), "/", ""), ".", "_") 
    httpurl = Replace(httpurl, ":", "_")                                            '追加于2012年7月13日
    ToPath = folderPath & httpurl 
    If checkFolder(ToPath) = False Then
        Call CreateFolder(ToPath) 
        createDomain = True 
    End If 
    If Err Then Call doError(Err.Description, "CreateDoMain 创建以域名命名的文件夹 函数出错，FolderPath=" & folderPath) 
End Function 
'创建以域名命名的文件夹 (辅助)
Function createDomainFolder(folderPath, httpurl)
    createDomainFolder = createDomain(folderPath, httpurl) 
End Function 
'移动文件夹
Function moveFolder(ByVal path, ByVal ToPath)
    On Error Resume Next 
    Dim fso 
    Call handlePath(path)                                                           '获得完整路径
    Call handlePath(ToPath)                                                         '获得完整路径
    moveFolder = False 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(path) = True And fso.FolderExists(ToPath) = False Then      '要移动文件夹为真，被移动文件夹为假
            Call fso.moveFolder(path, ToPath) 
            moveFolder = True 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "MoveFolder 移动文件夹 函数出错，Path=" & path & "，ToPath=" & ToPath) 
End Function 
'修改文件夹 (辅助)
Function editFolder(path, ToPath)
    editFolder = moveFolder(path, ToPath) 
End Function 
'拷贝文件夹
Function copyFolder(ByVal path, ByVal ToPath)
    On Error Resume Next 
    Dim fso 
    Call handlePath(path)                                                           '获得完整路径
    Call handlePath(ToPath)                                                         '获得完整路径
    '完善当前最后为\时会报错(20151022)
    If Right(path, 1) = "\" Then
        path = Left(path, Len(path) - 1) 
    End If 
    If Right(ToPath, 1) = "\" Then
        ToPath = Left(ToPath, Len(ToPath) - 1) 
    End If 
    copyFolder = False 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(path) = True And fso.FolderExists(ToPath) = False Then      '要拷贝文件夹为真，被拷贝文件夹为假
            Call fso.copyFolder(path, ToPath) 
            copyFolder = True 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "CopyFolder 拷贝文件夹 函数出错，Path=" & path & "，ToPath=" & ToPath) 
End Function 
'删除文件夹，只读文件夹照样删除 使用耐谨慎
Function deleteFolder(ByVal folderPath)
    On Error Resume Next 
    Dim fso 
    Call handlePath(folderPath)                                                     '获得完整路径
    deleteFolder = False 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(folderPath) Then
            If Right(folderPath, 1) = "\" Then folderPath = Left(folderPath, Len(folderPath) - 1) '晕，后台\就报错，必需去掉它
            Call fso.deleteFolder(folderPath, True) 
            deleteFolder = True 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "DeleteFolder 删除文件夹 函数出错，FolderPath=" & folderPath) 
End Function 
'获得文件夹大小
Function getFolderSize(ByVal folderName)
    On Error Resume Next 
    Dim fso, OpenFolder 
    Call handlePath(folderName)                                                     '获得完整路径
    getFolderSize = 0 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(folderName) Then
            Set OpenFolder = fso.GetFolder(folderName)
                getFolderSize = OpenFolder.Size 
            Set OpenFolder = Nothing 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "GetFolderSize 获得文件夹大小 函数出错，FolderName=" & folderName) 
End Function 
'检查文件夹存在
Function checkFolder(ByVal folderName)
    On Error Resume Next 
    Dim fso 
    Call handlePath(folderName)                                                     '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        checkFolder = fso.FolderExists(folderName) 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "CheckFolder 检查夹文件 函数出错，FolderName=" & folderName) 
End Function 
'判断文件夹存在 (辅助)
Function existsFolder(folderName)
    existsFolder = checkFolder(folderName) 
End Function 
'获得文件夹创建时间
Function getFolderCreateDate(ByVal folderName)
    On Error Resume Next 
    Dim fso, OpenFolder 
    Call handlePath(folderName)                                                     '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(folderName) Then
            Set OpenFolder = fso.GetFolder(folderName)
                getFolderCreateDate = OpenFolder.DateCreated 
            Set OpenFolder = Nothing 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "GetFolderCreateDate 获得文件创建时间 函数出错，FolderName=" & folderName) 
End Function 
'获得文件夹修改时间
Function getFolderEditDate(ByVal folderName)
    On Error Resume Next 
    Dim fso, OpenFolder 
    Call handlePath(folderName)                                                     '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(folderName) Then
            Set OpenFolder = fso.GetFolder(folderName)
                getFolderEditDate = OpenFolder.DateLastModified 
            Set OpenFolder = Nothing 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "GetFolderEditDate 获得文件修改时间 函数出错，FolderName=" & folderName) 
End Function 
'获得文件夹修改时间  调用上面
Function getFolderTime(folderName)
    getFolderTime = getFolderEditDate(folderName) 
End Function 
'获得文件夹属性  '0普通文件 1只读 2为隐藏 4系统 8卷材 16文件夹 32存档 64快捷键 128压缩文件
Function getFolderAttribute(ByVal folderName)
    On Error Resume Next 
    Dim fso, OpenFolder 
    Call handlePath(folderName)                                                     '获得完整路径
    getFolderAttribute = 0 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(folderName) = True Then
            Set OpenFolder = fso.GetFolder(folderName)
                getFolderAttribute = OpenFolder.Attributes 
            Set OpenFolder = Nothing 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "GetFolderAttribute 获得文件属性 函数出错，FolderName=" & folderName) 
End Function 
'修改文件夹属性  '0普通文件 1只读 2为隐藏 4系统 8卷材 16文件夹 32存档 64快捷键 128压缩文件
Function editFolderAttribute(ByVal folderName, n)
    On Error Resume Next 
    Dim fso, OpenFolder 
    Call handlePath(folderName)                                                     '获得完整路径
    editFolderAttribute = False 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(folderName) Then
            Set OpenFolder = fso.GetFolder(folderName)
                OpenFolder.Attributes = n                                                       '32文件
            Set OpenFolder = Nothing 
            editFolderAttribute = True 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "EditFolderAttribute 修改文件属性 函数出错，FolderName=" & folderName) 
End Function 
'判断文件夹天数（用修改时间进行比较）
Function existsFolderDay(ByVal folderName, ByVal InputDayNumb)
    On Error Resume Next 
    Dim fso, OpenFolder, DayNumb 
    Call handlePath(folderName)                                                     '获得完整路径
    existsFolderDay = False 
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(folderName) = True Then
            Set OpenFolder = fso.GetFolder(folderName)
                DayNumb = DateDiff("d", OpenFolder.DateLastModified, Now())                     '之前是Data
                If DayNumb <= InputDayNumb Then existsFolderDay = True 
            Set OpenFolder = Nothing 
        End If 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "ExistsFolderDay 判断文件夹天数 函数出错，FolderName=" & folderName) 
End Function 
'判断文件夹天数 (辅助)
Function checkFolderDay(folderName, InputDayNumb)
    checkFolderDay = existsFolderDay(folderName, InputDayNumb) 
End Function 
'获得全部文件（夹）列表
Function getFileFolderList(ByVal folderPath, FileYes, FileTypeList, FileNameOrType, FolderType, FolderNameOrType, ByVal c)
    On Error Resume Next 
    getFileFolderList = "" 
    Dim fso, f, fc, myFile, f1, FileType 
    Call handlePath(folderPath)                                                     '获得完整路径
    FileTypeList = LCase(FileTypeList)                                              '文件类型列表小写
    Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(folderPath) Then Exit Function 
        Set f = fso.GetFolder(folderPath)
            If FileYes = True Then
                For Each myFile In f.files
                    FileType = LCase(Right(myFile.Name, Len(myFile.Name) - InStrRev(myFile.Name, "."))) 

                    If(InStr("|" & FileTypeList & "|", "|" & FileType & "|") > 0 Or FileTypeList = "" Or FileTypeList = "全部" Or FileTypeList = "*") And FileTypeList <> "0" Then'改进于20140929 文件类型判断完善
                        If InStr("|名称|取名称|", "|" & FileNameOrType & "|") > 0 Then
                            c = c & myFile.Name & vbCrLf 
                        Else
                            c = c & myFile & vbCrLf 
                        End If 
                    End If 
                    doEvents 
                Next 
            End If 

            If InStr("|" & FolderType & "|", "|文件夹|") > 0 Or InStr("|" & FolderType & "|", "|全部文件夹|") > 0 Or InStr("|" & FolderType & "|", "|全部|") > 0 Then
                Set fc = f.subFolders
                    For Each f1 In fc
                        If InStr("|" & FolderType & "|", "|hidefolderlist|") = False Then
                            If InStr("|名称|取名称|", "|" & FolderNameOrType & "|") > 0 Then
                                c = c & f1.Name & vbCrLf 
                            ElseIf LCase(FolderNameOrType) <> "no" Then
                                c = c & f1 & vbCrLf 
                            End If 
                        End If 
                        If InStr("|" & FolderType & "|", "|全部文件夹|") > 0 Or InStr("|" & FolderType & "|", "|全部|") > 0 Then
                            c = getFileFolderList(f1, FileYes, FileTypeList, FileNameOrType, FolderType, FolderNameOrType, c) & vbCrLf 
                        End If 
                        doEvents 
                    Next 
                Set fc = Nothing 
            End If 
        Set f = Nothing 
    Set fso = Nothing 
    If c <> "" Then c = Left(c, Len(c) - 2) 
    getFileFolderList = c 
    If Err Then Call doError(Err.Description, "GetFileFolderList 获得全部文件（夹）列表 函数出错，FolderPath=" & folderPath) 
End Function 
'获得当前目录下全部Jpg文件
Function getDirJpgList(ByVal folderPath)
    getDirJpgList = getDirFileList(folderPath, "jpg") 
End Function 
'获得当前目录下全部Png文件
Function getDirPngList(ByVal folderPath)
    getDirPngList = getDirFileList(folderPath, "png") 
End Function 
'获得当前目录下全部Ini文件
Function getDirIniList(ByVal folderPath)
    getDirIniList = getDirFileList(folderPath, "ini") 
End Function 
'获得当前目录下全部Txt文件
Function getDirTxtList(ByVal folderPath)
    getDirTxtList = getDirFileList(folderPath, "txt") 
End Function 
'获得当前目录下全部js文件
Function getDirJsList(ByVal folderPath)
    getDirJsList = getDirFileList(folderPath, "js") 
End Function 
'获得当前目录下全部css文件
Function getDirCssList(ByVal folderPath)
    getDirCssList = getDirFileList(folderPath, "css") 
End Function 
'获得当前目录下全部Html文件
Function getDirHtmlList(ByVal folderPath)
    getDirHtmlList = getDirFileList(folderPath, "html") 
End Function 
'获得当前目录下全部asp文件
Function getDirAspList(ByVal folderPath)
    getDirAspList = getDirFileList(folderPath, "asp") 
End Function 
'获得当前目录下全部php文件
Function getDirPhpList(ByVal folderPath)
    getDirPhpList = getDirFileList(folderPath, "php") 
End Function 
'获得当前目录下批量文件列表
Function getDirFileList(ByVal folderPath, fileTypeList)
    If fileTypeList = "" Then
        fileTypeList = "全部" 
    End If 
    getDirFileList = getFileFolderList(folderPath, True, fileTypeList, "", "", "", "") 
End Function 

'获得当前目录下全部Jpg文件名称
Function getDirJpgNameList(ByVal folderPath)
    getDirJpgNameList = getDirFileNameList(folderPath, "jpg") 
End Function 
'获得当前目录下全部Png文件名称
Function getDirPngNameList(ByVal folderPath)
    getDirPngNameList = getDirFileNameList(folderPath, "png") 
End Function 
'获得当前目录下全部Ini文件名称
Function getDirIniNameList(ByVal folderPath)
    getDirIniNameList = getDirFileNameList(folderPath, "ini") 
End Function 
'获得当前目录下全部Txt文件名称
Function getDirTxtNameList(ByVal folderPath)
    getDirTxtNameList = getDirFileNameList(folderPath, "txt") 
End Function 
'获得当前目录下全部js文件名称
Function getDirJsNameList(ByVal folderPath)
    getDirJsNameList = getDirFileNameList(folderPath, "js") 
End Function 
'获得当前目录下全部css文件名称
Function getDirCssNameList(ByVal folderPath)
    getDirCssNameList = getDirFileNameList(folderPath, "css") 
End Function 
'获得当前目录下全部Html文件名称
Function getDirHtmlNameList(ByVal folderPath)
    getDirHtmlNameList = getDirFileNameList(folderPath, "html") 
End Function 
'获得当前目录下全部asp文件名称
Function getDirAspNameList(ByVal folderPath)
    getDirAspNameList = getDirFileNameList(folderPath, "asp") 
End Function 
'获得当前目录下全部php文件名称
Function getDirPhpNameList(ByVal folderPath)
    getDirPhpNameList = getDirFileNameList(folderPath, "php") 
End Function 
'获得当前目录下全部类型文件名称
Function getDirFileNameList(ByVal folderPath, fileTypeList)
    If fileTypeList = "" Then
        fileTypeList = "全部" 
    End If 
    getDirFileNameList = getFileFolderList(folderPath, True, fileTypeList, "名称", "", "", "") 
End Function 


'获得当前目录下全部文件
Function getDirAllFileList(ByVal folderPath, fileTypeList)
    If fileTypeList = "" Then
        fileTypeList = "全部" 
    End If 
    getDirAllFileList = getFileFolderList(folderPath, True, fileTypeList, "", "全部文件夹|hidefolderlist", "", "") 
End Function 
'获得当前目录下全部文件名称
Function getDirAllFileNameList(ByVal folderPath, fileTypeList)
    If fileTypeList = "" Then
        fileTypeList = "全部" 
    End If 
    getDirAllFileNameList = getFileFolderList(folderPath, True, fileTypeList, "名称", "全部文件夹|hidefolderlist", "", "") 
End Function 

'获得当前目录下文件夹
Function getDirFolderList(ByVal folderPath)
    getDirFolderList = getFileFolderList(folderPath, False, "", "", "文件夹", "", "") 
End Function 
'获得当前目录下文件夹名称
Function getDirFolderNameList(ByVal folderPath)
    getDirFolderNameList = getFileFolderList(folderPath, False, "", True, "文件夹", "名称", "") 
End Function 

'获得当前目录下全部文件夹
Function getDirAllFolderList(ByVal folderPath)
    getDirAllFolderList = getFileFolderList(folderPath, False, "", "", "全部文件夹", "", "") 
End Function 
'获得当前目录下全部文件夹名称
Function getDirAllFolderNameList(ByVal folderPath)
    getDirAllFolderNameList = getFileFolderList(folderPath, False, "", True, "全部文件夹", "名称", "") 
End Function 


'获得全部文件（夹）数量
Function getFileFolderNumb(ByVal folderPath, sType, ByVal countNumb)
    On Error Resume Next 
    Dim fso, f, fc, f1 
    getFileFolderNumb = countNumb 
    Call handlePath(folderPath)                                                     '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(folderPath) = True Then
            Set f = fso.GetFolder(folderPath)
                Set fc = f.subFolders
                    If sType = "文件数量" Or sType = "全部文件数量" Then
                        countNumb = countNumb + f.files.Count 
                    ElseIf sType = "文件夹数量" Or sType = "全部文件夹数量" Then
                        countNumb = countNumb + fc.Count 
                    ElseIf sType = "文件夹大小" Then
                        countNumb = f.Size 
                    End If 
                    If sType = "全部文件数量" Or sType = "全部文件夹数量" Then
                        For Each f1 In fc
                            countNumb = getFileFolderNumb(f1, sType, countNumb) 
                            doEvents 
                        Next 
                    End If 
                Set f = Nothing 
            Set fc = Nothing 
        End If 
        getFileFolderNumb = countNumb 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "GetDirFolderNumb 获得当前目录文件夹数量 函数出错，FolderPath=" & folderPath) 
End Function 
'获得当前目录文件夹数量
Function getDirFolderNumb(folderPath)
    getDirFolderNumb = getFileFolderNumb(folderPath, "文件夹数量", 0) 
End Function 
'获得当前目录文件夹大小
Function getDirFolderSize(folderPath)
    getDirFolderSize = getFileFolderNumb(folderPath, "文件夹大小", 0) 
End Function 
'获得当前目录文件数量
Function getDirFileNumb(folderPath)
    getDirFileNumb = getFileFolderNumb(folderPath, "文件数量", 0) 
End Function 
'获得当前目录全部文件数量
Function getDirAllFileNumb(folderPath)
    getDirAllFileNumb = getFileFolderNumb(folderPath, "全部文件数量", 0) 
End Function 
'获得当前目录全部文件夹数量
Function getDirAllFolderNumb(folderPath)
    getDirAllFolderNumb = getFileFolderNumb(folderPath, "全部文件夹数量", 0) 
End Function 
'获得全部文件夹列表
Function getFolderList(ByVal folderPath, sType)
    'On Error Resume Next
    Dim fso, f, fc, f1, c 
    c = "" 
    sType = CStr(sType) 
    Call handlePath(folderPath)                                                     '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(folderPath) = True Then
            Set f = fso.GetFolder(folderPath)
                Set fc = f.subFolders
                    For Each f1 In fc
                        If sType = "名称" Or sType = "0" Then
                            c = c & f1.Name & vbCrLf 
                        Else
                            c = c & f1 & vbCrLf 
                        End If 
                        doEvents 
                    Next 
                Set f = Nothing 
            Set fc = Nothing 
        End If 
        If c <> "" Then c = Left(c, Len(c) - 2) 
        getFolderList = c 
    Set fso = Nothing 
    If Err Then Call doError(Err.Description, "GetDirFolderNumb 获得当前目录文件夹数量 函数出错，FolderPath=" & folderPath) 
End Function 
'获得文件列表排序
Function getDirFileSort(ByVal folderPath)
    On Error Resume Next 
    Dim fso, f, fc, myFile, id, arrayStr(99), i, c 
    Call handlePath(folderPath)                                                     '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(folderPath) = False Then Exit Function                      '文件夹不存在，退出
        Set f = fso.GetFolder(folderPath)
            For Each myFile In f.files
                If InStr(myFile.Name, "#") = False And InStr(myFile.Name, "、") > 0 Then        '#代表不显示，并且要有、符号
                    id = Replace(Left(myFile.Name, 2), "、", "") 
                    arrayStr(id) = myFile.Name 
                End If 
                doEvents 
            Next 
            For i = 0 To UBound(arrayStr)
                If arrayStr(i) <> "" Then
                    c = c & arrayStr(i) & vbCrLf 
                End If 
            Next 
        Set f = Nothing 
    Set fso = Nothing 
    getDirFileSort = c 
    If Err Then Call doError(Err.Description, "GetDirFileSort 获得文件列表排序 函数出错，FolderPath=" & folderPath) 
End Function 
'获得文件夹列表排序
Function getDirFolderSort(ByVal folderPath)
    On Error Resume Next 
    Dim fso, f, fc, myFile, id, arrayStr(99), i, c, f1 
    Call handlePath(folderPath)                                                     '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(folderPath) = False Then Exit Function                      '文件夹不存在，退出
        Set f = fso.GetFolder(folderPath)
            Set fc = f.subFolders
                For Each f1 In fc
                    If InStr(f1.Name, "#") = 0 And InStr(f1.Name, "、") > 0 Then                    '#代表不显示，并且要有、符号
                        id = Replace(Left(f1.Name, 2), "、", "") 
                        arrayStr(id) = f1.Name 
                        doEvents 
                    End If 
                Next 
                For i = 0 To UBound(arrayStr)
                    If arrayStr(i) <> "" Then
                        c = c & arrayStr(i) & vbCrLf 
                    End If 
                Next 
            Set f = Nothing 
        Set fc = Nothing 
    Set fso = Nothing 
    getDirFolderSort = c 
    If Err Then Call doError(Err.Description, "GetDirFileSort 获得文件列表排序 函数出错，FolderPath=" & folderPath) 
End Function 
'删除目录下的所有文件,2013,10,26
Sub delDirAllFile(dirPath)
    Dim fso, f, fc 
    Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.GetFolder(dirPath)
            Set fc = f.files
                For Each f In fc
                    If fso.fileExists(f) Then fso.deleteFile(f) 
                Next 
            Set fc = Nothing 
        Set f = Nothing 
    Set fso = Nothing 
End Sub 
'删除目录下的所有文件 (辅助)
Sub deleteAllFile(dirPath)
    Call DelDirAllFile(dirPath) 
End Sub 
'获得当前更新文件列表  例 Call Rw(GetUpdateFile("/Inc","html|css|jpg|Asp","全部","n,1,文件路径",""))
'暂时这么设计，设计不友好，当前的息是怎么想的 20150711
Function getUpdateFile(ByVal folderPath, FileTypeList, FolderType, setTime, ByVal c)
    'On Error Resume Next
    Dim fso, f, fc, myFile, f1, FileType, ResultTime, s, setTimeType, splxx, showMsgType 
    getUpdateFile = "" 
    setTimeType = "" 
    showMsgType = "" 
    Call handlePath(folderPath)                                                     '获得完整路径
    If InStr(setTime, ",") Then
        splxx = Split(setTime, ",") 
        setTimeType = Trim(splxx(0)) 
        setTime = Trim(splxx(1)) 
        If UBound(splxx) >= 2 Then showMsgType = Trim(splxx(2))                         '显示回显内容类型
    End If 
    If setTimeType = "" Then
        setTimeType = "s" 
    End If 
    If setTime = "" Then
        setTime = -1 
    Else
        setTime = CLng(setTime)                                                         '设置时间为数值类型
    End If 

    FileTypeList = LCase(FileTypeList)                                              '文件类型列表小写
    Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(folderPath) Then Exit Function 
        Set f = fso.GetFolder(folderPath)
            For Each myFile In f.files
                If Left(myFile.Name, 1) <> "#" Then
                    'q 季   m 月   y 一年的日数   d 日   w 一周的日数   ww 周   h 时   n 分钟   s 秒
                    '计算时间 然后再上传
                    ResultTime = DateDiff(setTimeType, myFile.DateLastModified, Now()) 
                    'call echo(MyFile & "("& setTimeType &")", ResultTime & "<=" & SetTime)
                    If ResultTime <= setTime Or setTime = -1 Then
                        FileType = LCase(Right(myFile.Name, Len(myFile.Name) - InStrRev(myFile.Name, "."))) 
                        'Call Echo(FileType,MyFile)
                        If(InStr("|" & FileTypeList & "|", "|" & FileType & "|") > 0 Or FileTypeList = "" Or FileTypeList = "全部" Or FileTypeList = "*") And FileTypeList <> "0" Then'改进于20140929 文件类型判断完善
                            If showMsgType = "文件路径" Then
                                s = myFile 
                            Else
                                s = myFile & "|" & checkCode(myFile) & "|" & printSpaceValue(myFile.Size) & "|" & format_Time(myFile.DateLastModified, 1) & "|" & ResultTime & "/" & setTime & "(" & setTime - ResultTime & ")" 
                            End If 
                            c = c & s & vbCrLf 
                        End If 
                    End If 
                End If 
                doEvents 
            Next 

            If InStr("|当前|全部|文件夹|全部文件夹|", "|" & FolderType & "|") > 0 Then
                Set fc = f.subFolders
                    For Each f1 In fc
                        If InStr("|全部|全部文件夹|", "|" & FolderType & "|") > 0 Then
                            c = getUpdateFile(f1, FileTypeList, FolderType, setTime, c) & vbCrLf 
                        End If 
                        doEvents 
                    Next 
                Set fc = Nothing 
            End If 
        Set f = Nothing 
    Set fso = Nothing 
    If c <> "" Then c = Left(c, Len(c) - 2) 
    getUpdateFile = c 
'If Err Then call doError(Err.Description, "GetUpdateFile 获得当前更新文件列表 函数出错，FolderName=" & FolderName)
End Function 
'========================================= 辅助区 =========================================
'处理成完成路径 (2013,9,27
Function handlePath(FFPath)                                                     'Path前面不加ByVal 重定义，这样是为了让前面函数里可以使用这个路径完整调用
    FFPath = Replace(FFPath, "/", "\") 
    FFPath = Replace(FFPath, "\\", "\") 
    FFPath = Replace(FFPath, "\\", "\") 
    Dim isDir                                                                       '为目录
    isDir = False 
    If Right(FFPath, 1) = "\" Then
        isDir = True 
    End If 
    If InStr(FFPath, ":") = 0 Then
        If Left(FFPath, 1) = "\" Then
            FFPath = Server.MapPath("\") & "\" & FFPath 
        Else
            FFPath = Server.MapPath(".\") & "\" & FFPath 
        End If 
    End If 
    FFPath = Replace(FFPath, "/", "\") 
    FFPath = Replace(FFPath, "\\", "\") 
    FFPath = Replace(FFPath, "\\", "\") 
    FFPath = fullPath(FFPath) 
    If isDir = True Then
        FFPath = FFPath & "\" 
    End If 
    handlePath = FFPath 
End Function 
'真正的路径  PHP里函数 为假返回空
Function realPath(ByVal FFPath)
    realPath = "" 
    If checkFile(FFPath) Then
        realPath = FFPath 
        Exit Function 
    End If 
    If checkFolder(FFPath) Then
        realPath = FFPath 
        Exit Function 
    End If 
End Function 
'完整路径
Function fullPath(ByVal FFPath)
    Dim splStr, s, c 
    c = "" 
    FFPath = Replace(FFPath, "/", "\") 
    splStr = Split(FFPath, "\") 
    For Each s In splStr
        s = Trim(s) 
        If s <> "" And s <> "." Then
            If InStr(c, "\") > 0 And s = ".." Then
                c = Mid(c, 1, InStrRev(c, "\") - 1) 
            Else
                If c <> "" And Right(c, 1) <> "\" Then c = c & "\" 
                c = c & s 
            End If 
        End If 
    Next 
    fullPath = c 
End Function 
'处理成相对路径(20150906)  如 a=handleRelativePath("",a)
Function handleRelativePath(rootPath, ByVal filePath)
    If rootPath = "" Then rootPath = "\" 
    rootPath = handlePath(rootPath) 
    filePath = Replace(filePath, rootPath, "\") 
    handleRelativePath = filePath 
End Function 
'获得目录
Function dirName(filePath)
    filePath = handleFileUrl(filePath) 
    If Right(filePath, 1) <> "\" Then
        filePath = Mid(filePath, 1, InStrRev(filePath, "\")) 
    End If 
    dirName = filePath 
End Function 
'获得目录 PHP版 20160516
Function phpDirName(filePath)
    filePath = handleFileUrl(filePath) 
    If Right(filePath, 1) <> "\" Then
        filePath = Mid(filePath, 1, InStr(filePath, "\")) 
    End If 
    phpDirName = filePath 
End Function 
'判断组件是否存在
Function isObjInstalled(obj)
    On Error Resume Next 
    isObjInstalled = False                                                          '默认组件为假
    Dim xTestObj 
    Set xTestObj = CreateObject(obj)
        If 0 = Err.Number Then isObjInstalled = True                                    '如果组件存在为真
    Set xTestObj = Nothing 
End Function 
'判断组件是否存在 (辅助)
Function existsObject(obj)
    existsObject = isObjInstalled(obj) 
End Function 
'获得文件操作组件
Function getFileObject()
    If existsObject("Scripting.FileSystemObject") Then
        getFileObject = "Scripting.FileSystemObject" 
    Else
        getFileObject = "Scripting.wang198060FileSystemObject" 
    End If 
End Function 
'获得字典组件
Function getDictionary()
    If existsObject("Scripting.Dictionary") Then
        getDictionary = "Scripting.Dictionary" 
    Else
        getDictionary = "Scripting.wang198060Dictionary" 
    End If 
End Function 
'获得数据流组件
Function getStream()
    If existsObject("Adodb.Stream") Then
        getStream = "Adodb.Stream" 
    Else
        getStream = "Adodb.wang198060Stream" 
    End If 
End Function 
'显示对应属性中文说明
Function showAttributeTip(n)
    Dim c 
    c = "" 
    If n = 0 Then c = "0普通文件，" 
    If n >= 128 Then
        n = n - 128 
        c = c & "128压缩文件，" 
    End If 
    If n >= 64 Then
        n = n - 64 
        c = c & "64快捷键，" 
    End If 
    If n >= 32 Then
        n = n - 32 
        c = c & "32存档，" 
    End If 
    If n >= 16 Then
        n = n - 16 
        c = c & "16文件夹，" 
    End If 
    If n >= 8 Then
        n = n - 8 
        c = c & "8卷材，" 
    End If 
    If n >= 4 Then
        n = n - 4 
        c = c & "4系统，" 
    End If 
    If n >= 2 Then
        n = n - 2 
        c = c & "2为隐藏，" 
    End If 
    If n >= 1 Then
        n = n - 1 
        c = c & "1只读，" 
    End If 
    If c <> "" Then c = Left(c, Len(c) - 1) 
    showAttributeTip = c 
End Function 
'KB转M函数，不够1M,自动显示K 最大为G   (追求好看与可读，舒服)
Function printSpaceValue(ByVal numb)
    Dim n 
    n = numb 
    If n >= 1073741824 Then                                                         '等于1024*1024*1024  VB中用CDec(1024)
        n = n / 1073741824 
        If InStr(n, ".") Then
            n = Left(n, InStr(n, ".") + 2)
        End If 
        n = n & "G" 
    ElseIf n >= 1048576 Then                                                        '等于1024*1024
        n = n / 1048576 
        If InStr(n, ".") Then
            n = Left(n, InStr(n, ".") + 2) 
        End If 
        n = n & "M" 
    Else
        n = FormatNumber(n / 1024,2) & "K" 				'把Fix换成FormatNumber 是因为与php那边一致 20160526
    End If
	
    If n = "0K" Then n = "0" & FormatNumber(numb/1024,2) & "K"   'n = Numb & "Bet"   '如果为0 K 就显示字的Bet数    改成与php版一样最小为K 如0.93K  20160525
    printSpaceValue = n 
End Function 
'KB转M函数 (辅助)
Function printSpaceSize(numb)
    printSpaceSize = printSpaceValue(numb) 
End Function 
'KB转M函数 (辅助)
Function printFileSize(numb)
    printFileSize = printSpaceValue(numb) 
End Function 
'获得桌面路径
Function getDeskPath()
    Dim WshShell 
    Set WshShell = CreateObject("Wscript.shell")
        getDeskPath = WshShell.SpecialFolders("Desktop") & "\" 
End Function
'创建快捷方式*.lnk
Function createUrlShortcut(SavePath, ByVal url)
    Dim c 
    c = "[InternetShortcut] " & vbCrLf 
    c = c & "URL=" & url & vbCrLf 
    c = c & "IDList= " & vbCrLf 
    c = c & "[{000214A0-0000-0000-C000-000000000046}] " & vbCrLf 
    c = c & "Prop3=19,2 " & vbCrLf 
    'Response.AddHeader "Content-Disposition", "attachment;filename="&Url&".url;"
    'Response.ContentType = "application/octet-stream"
    Call CreateFile(SavePath, c) 
    createUrlShortcut = c 
End Function 
'创建文件夹快捷方式*.lnk
Function createFolderShortcut(SavePath, ShortcutPath)
    Dim c 
    c = "[InternetShortcut]  " & vbCrLf 
    c = c & "URL=file:///" & ShortcutPath & vbCrLf 
    c = c & "HotKey=0" & vbCrLf 
    c = c & "IDList=" & vbCrLf 
    c = c & "IconFile=C:\WINDOWS\system32\SHELL32.dll" & vbCrLf 
    c = c & "IconIndex=3" & vbCrLf 
    c = c & "[{000214A0-0000-0000-C000-000000000046}]" & vbCrLf 
    c = c & "Prop3=19,9" & vbCrLf 
    c = c & "[InternetShortcut.A]" & vbCrLf 
    c = c & "URL=file:///" & ShortcutPath & vbCrLf 
    c = c & "[InternetShortcut.W]" & vbCrLf 
    c = c & "URL=file:///" & ShortcutPath & vbCrLf 
    'Response.AddHeader "Content-Disposition", "attachment;filename="&ShortcutPath&".url;"
    'Response.ContentType = "application/octet-stream"
    Call CreateFile(SavePath, c) 
    createFolderShortcut = c 
End Function 
'寻找目录里相同文件名而不同类型的文件        创建于20140710
Function findDirFileName(folderPath, fileName)
    Dim fso, f, myFile, fName 
    findDirFileName = "" 
    fileName = LCase(fileName)                                                      '转成小写
    Call handlePath(folderPath)                                                     '获得完整路径
    Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(folderPath) Then Exit Function 
        Set f = fso.GetFolder(folderPath)
            For Each myFile In f.files
                fName = LCase(Mid(myFile.Name, 1, InStrRev(myFile.Name, ".") - 1)) 
                If fileName = fName Then
                    findDirFileName = myFile.Name 
                    Exit Function 
                End If 
                doEvents 
            Next 
        Set f = Nothing 
    Set fso = Nothing 
End Function 
'寻找目录一组里相同文件名而不同类型的文件 用法 FileName=FindSubDirFileName("E:\","aa|bb|cc")
Function findSubDirFileName(folderPath, FileNameList)
    Dim splStr, fileName 
    findSubDirFileName = "" 
    splStr = Split(FileNameList, "|") 
    For Each fileName In splStr
        fileName = Trim(fileName) 
        If fileName <> "" Then
            findSubDirFileName = findDirFileName(folderPath, fileName) 
            If findSubDirFileName <> "" Then Exit Function 
        End If 
    Next 
End Function 


'获得文章填充内容 20150113 GetArticleFillContent("","",9)
Function getArticleFillContent(filePath, FileContent, nLen)
    Dim splStr, c, nId, i 
    c = "" 
    If filePath <> "" Then
        FileContent = getFText(filePath) 
    ElseIf FileContent = "" Then
        FileContent = getFText("\DataDir\自动填文章内容数据.txt") 
    End If 
    splStr = Split(FileContent, vbCrLf) 
    Randomize 
    For i = 0 To nLen
        nId = CInt(UBound(splStr) * Rnd) 
        'Call Echo("nId",nId)
        c = c & splStr(nId) 
    Next 
    getArticleFillContent = c 
End Function 

'FilePath：E:\E盘\WEB网站\至前网站\filename.asp
'GetFilePathName：filename
'RemoveFileDir：filename.asp
'getFileExtName：asp

'获得路径文件名称
Function getFilePathName(ByVal filePath)
    filePath = Replace(filePath, "\", "/") 
    If InStr(filePath, "/") > 0 Then
        filePath = Mid(filePath, InStrRev(filePath, "/") + 1) 
    End If 
    If InStr(filePath, ".") > 0 Then
        filePath = Mid(filePath, 1, InStrRev(filePath, ".") - 1) 
    End If 
    getFilePathName = filePath 
End Function 
'获取文件扩展名 后缀(辅助上面)
Function getExtension(filePath)
    getExtension = getFilePathName(filePath) 
End Function 
'去除文件路径 返回路径中的文件名部分
Function removeFileDir(ByVal filePath)
    filePath = Replace(filePath, "\", "/") 
    If InStr(filePath, "/") > 0 Then
        filePath = Mid(filePath, InStrRev(filePath, "/") + 1) 
    End If 
    removeFileDir = filePath 
End Function 
'返回路径中的文件名部分 (辅助上面)
Function baseName(ByVal filePath)
    baseName = removeFileDir(filePath) 
End Function 
'-----获取扩展名--------------------------------------
Function getFileExtName(fileName)
    Dim pos 
    pos = InStrRev(fileName, ".") 
    If pos > 0 Then
        getFileExtName = Mid(fileName, pos + 1) 
    Else
        getFileExtName = "" 
    End If 
End Function 

'0：E:\E盘\WEB网站\至前网站\filename.asp
'1：E:\E盘\WEB网站\至前网站\      可以这么调用 HandleFilePathArray(FilePath)(1)
'2：filename.asp
'3：filename
'4：asp

'文件处理成数组20150124  数组  0原文件路径 1为文件路径   2为文件名称  3为去除文件类型文件名称   4为文件类型后缀名
Function handleFilePathArray(ByVal filePath)
    Dim FileDir, fileName, FileNoTypeName, FileType 
    filePath = handleFileUrl(filePath) 

    FileDir = Mid(filePath, 1, InStrRev(filePath, "\")) 
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1) 
    If InStrRev(fileName, ".") > 0 Then
        FileNoTypeName = Mid(fileName, 1, InStrRev(fileName, ".") - 1) 
    Else
        FileNoTypeName = "" 
    End If 
    FileType = Mid(fileName, InStrRev(fileName, ".") + 1) 

    'Call Echo("FilePath",FilePath)
    'Call Echo("FileDir",FileDir)
    'Call Echo("FileName",FileName)
    'Call Echo("FileNoTypeName",FileNoTypeName)
    'Call Echo("FileType",FileType)

    handleFilePathArray = Split(filePath & vbCrLf & FileDir & vbCrLf & fileName & vbCrLf & FileNoTypeName & vbCrLf & FileType, vbCrLf) 
'HandleFilePathArray=Array(FilePath,FileDir,FileName,FileNoTypeName,FileType)
End Function 
'获得处理后的文件名称
Function getStrFileName(ByVal filePath)
    getStrFileName = handleFilePathArray(filePath)(2) 
End Function 
'获得文件名称20160105
Function getFileName(ByVal filePath)
    getFileName = handleFilePathArray(filePath)(2) 
End Function 
'获得文件属性
Function getFileAttr(ByVal filePath, sType)
    getFileAttr = "" 
    If sType = "0" Then
        getFileAttr = handleFilePathArray(filePath)(0) 
    ElseIf sType = "1" Then
        getFileAttr = handleFilePathArray(filePath)(1) 
    ElseIf sType = "name" Or "2" Then
        getFileAttr = handleFilePathArray(filePath)(2) 
    ElseIf sType = "3" Then
        getFileAttr = handleFilePathArray(filePath)(3) 
    ElseIf sType = "4" Then
        getFileAttr = handleFilePathArray(filePath)(4) 
    End If 
End Function 
'**************************************************
'函数名：getFileExt
'作  用：获取文件扩展名
'举  例：ab.use "mvc" : ctrl.use "file" : ab.c.print ctrl.file.FileExt("aaa.txt")
'**************************************************
Function getFileExt(fName)
    Dim temp : temp = Split(fName, ".") 
    getFileExt = temp(UBound(temp)) 
End Function 
'获得上一级目录(20150909)
Function getParentPath(ByVal dirPath)
    dirPath = handlePath(dirPath) 
    getParentPath = "" 
    If InStr(dirPath, "\") > 0 Then
        If Right(dirPath, 1) = "\" Then
            dirPath = Left(dirPath, Len(dirPath) - 1) 
        End If 
        getParentPath = Mid(dirPath, 1, InStrRev(dirPath, "\")) 
    End If 
End Function 
%>   



