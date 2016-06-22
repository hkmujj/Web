<!--#Include File = "Config.Asp"-->     
<% 

'登录判断
If getIP() <> "127.0.0.1" And Session("install") = "" Then
    If Request("password") = "ASPPHPCMS" Then
        Session("install") = "ok" 
        Call defaultDisplay() : Response.End() 
    ElseIf Request("password") <> "" Then
        Call echo("登录失败", "密码错误") 
    End If 
    Call rwend(displayLogin()) 
End If 

Dim ObjAccess : Set ObjAccess = new ClassAccess
    '选择
    Select Case Request("act")
        Case "Web", "web" : MDBPath = "/Data/Data.mdb" : Call ObjAccess.configAccessDatabase() : Call ObjAccess.webData()'创建表
        Case "addwebtable" : MDBPath = "/Data/Data.mdb" : Call OpenConn() : Call ObjAccess.webData()'追加表

        Case Else defaultDisplay()
    End Select
Set ObjAccess = Nothing 
'默认回显内容
Sub defaultDisplay()
    Dim s 
    s = "<b>创建数据库2016</b><hr>" 
    s = s & "<a href='?act=web&DB_PREFIX=xy_' target='_blank'>点击创建(本地Access)</a> | <a href='?act=addwebtable&DB_PREFIX=test_' target='_blank'>追加test_前缀表</a>" 
    s = s & " | <a href='?act=addwebtable' target='_blank'>追加前缀表</a>"
    Call echo("创建网站数据库", s) 
End Sub
'创建远程SqlServer数据库
Sub createRemoteSqlServerDatabase()
    MDBPath = "" 
    If 1 = 2 Then
        DatabaseType = "SqlServer" 
        SqlServerHostIP = "hds-113.hichina.com,1433" 
        SqlServerDatabaseName = "hds1130012_db" 
        SqlServerUsername = "hds1130012" 
        SqlServerPassword = "L4dN4eRd" 
    Else
        DatabaseType = "SqlServerWebData"                                               '本地
        DatabaseType = "RemoteSqlServer20150727"                                        '远程
    End If 

    Call OpenConn() 
    Call ObjAccess.webData() 
End Sub 
'类
Class ClassAccess
    Dim tableName, sql, addData, splStr, splxx, i, s, splTrue, bigClassName 
    Dim splUrl, splNavName, navClassName, splImg, templateIndex, templateHome, templateMain 
    Dim DB_PREFIX 
    '默认配置
    Sub configAccessDatabase()
        MDBPath = handlePath(MDBPath) 
        Call deleteFile(MDBPath)                                                        '先删除这个数据库
        If existsFile(MDBPath) = False Then createMdb(MDBPath) 
        Call OpenConn() 
    End Sub 
    '注意，创建数据时 Default """"  默认值用这个   不要用这种Default ''    20150506
    '网站数据库
    Sub webData()
        DB_PREFIX = Request("DB_PREFIX") 
        Call echo("DB_PREFIX", DB_PREFIX) 
        '网站栏目表-------------------------------
        tableName = DB_PREFIX & "WebColumn" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 

            sql = sql & "columnName VarChar Default """","                            '导航名称
            sql = sql & "columnEnName VarChar Default """","                          '导航名称(英文)
            sql = sql & "columnType VarChar Default """","                            '导航类型
            sql = sql & "parentId int Default -1,"                                    '父栏目ID
            sql = sql & "sortRank int Default 0,"                                     '排序编号
            sql = sql & "views int Default 0,"                                        '点击次数
            sql = sql & "adminId int Default 0,"                                      '管理员id

            sql = sql & "isDisplay YesNo,"                                            '是否显示
            sql = sql & "smallImage VarChar Default """","                            '小图片地址
            sql = sql & "bigImage VarChar Default """","                              '大图片地址
            sql = sql & "bannerImage VarChar Default """","                           '当前Banner图片

            sql = sql & "flags VarChar Default """","                                 '旗
            sql = sql & "displayTitle VarChar Default """","                          '显示标题
            sql = sql & "lableTitle VarChar Default """","                            '标签描述


            sql = sql & "webTitle VarChar Default """","                              '网站自定标题
            sql = sql & "webKeywords VarChar Default """","                           '网站关键词
            sql = sql & "webDescription VarChar Default """","                        '网站描述
            sql = sql & "folderName VarChar Default """","                            'HTMl文件夹名称
            sql = sql & "fileName VarChar Default """","                              'HTML文件名称
            sql = sql & "customAUrl VarChar Default """","                            '自定义链接网址
            sql = sql & "templatePath VarChar Default """","                          '模板路径
            sql = sql & "target VarChar Default """","                                '链接打开方式
            sql = sql & "noFollow int Default 0,"                                     '不要追踪此网页上的链接
            sql = sql & "fontColor VarChar Default """","                             '字颜色
            sql = sql & "fontB YesNo Default No,"                                     '字粗
            sql = sql & "isMakeHtml YesNo Default Yes,"                               '开启生成Html




            sql = sql & "nPageSize int Default 10,"                                   '每页显示条数
            sql = sql & "sortSql VarChar Default """","                               '排序SQL
            sql = sql & "isHtml YesNo Default false,"                                 '是否为html
            sql = sql & "isOnHtml YesNo Default Yes,"                                 '是否开启生成Html

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间


            sql = sql & "noteBody Text Default """","                                 '笔记信息

            sql = sql & "aboutcontent Text Default """","                             '简单介绍
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 

        '文章信息表-------------------------------
        tableName = DB_PREFIX & "ArticleDetail" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "parentId int Default -1,"                                    '父栏目ID
            sql = sql & "sortRank int Default 0,"                                     '排序编号
            sql = sql & "views int Default 0,"                                        '点击次数
            sql = sql & "adminId int Default 0,"                                      '管理员id

            sql = sql & "smallImage VarChar Default """","                            '小图片地址
            sql = sql & "bigImage VarChar Default """","                              '大图片地址
            sql = sql & "bannerImage VarChar Default """","                           '当前Banner图片
            sql = sql & "downloadFile VarChar Default """","                          '下载文件
            sql = sql & "smallimageAlt VarChar Default """","                         '小图alt
            sql = sql & "bigimageAlt VarChar Default """","                           '大图alt
            sql = sql & "bannerimageAlt VarChar Default """","                        'Banneralt


            sql = sql & "title VarChar Default """","                                 '标题
            sql = sql & "titleColor VarChar Default """","                            '标题颜色
            sql = sql & "titleAlt VarChar Default """","                              '标题alt
            sql = sql & "lableTitle VarChar Default """","                            '标签描述

            sql = sql & "isThrough YesNo Default No,"                                 '通过审核

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间

            sql = sql & "occasions VarChar Default """","                             '场合
            sql = sql & "hotline VarChar Default """","                               '热线
            sql = sql & "model VarChar Default """","                                 '型号
            sql = sql & "author VarChar Default """","                                '作者
            sql = sql & "articleSource VarChar Default """","                         '文章来源

            sql = sql & "price Float Default 0,"                                        '产品价格
            sql = sql & "newPrice Float Default 0,"                                     '产品新价格20150729
            sql = sql & "memberPrice Float Default 0,"                                  '会员价格
            sql = sql & "sold Int Default 0,"                                         '已售出

            sql = sql & "memberType VarChar Default """","                            '会员类型
            sql = sql & "memberUser VarChar Default """","                            '会员名
            sql = sql & "hits Int Default 0,"                                         '点击数



            sql = sql & "productAbout Text Default """","                             '产品介绍
            sql = sql & "articleDescription Text Default """","                       '摘要(文章简要描述)20141225

            sql = sql & "httpUrl VarChar Default """","                               '网址
            sql = sql & "recordUrl VarChar Default """","                             '远程网址
            sql = sql & "webTitle VarChar Default """","                              '网站自定标题
            sql = sql & "webKeywords VarChar Default """","                           '网站关键词
            sql = sql & "webDescription VarChar Default """","                        '网站描述
            sql = sql & "folderName VarChar Default """","                            'HTMl文件夹名称
            sql = sql & "fileName VarChar Default """","                              'HTML文件名称
            sql = sql & "templatePath VarChar Default """","                          '模板路径
            sql = sql & "target VarChar Default """","                                '链接打开方式
            sql = sql & "customAUrl VarChar Default """","                            '自定义链接网址
            sql = sql & "fontColor VarChar Default """","                             '字颜色
            sql = sql & "noFollow int Default 0,"                                     '不要追踪此网页上的链接

            sql = sql & "flags VarChar Default """","                                 '旗

            sql = sql & "isHtml YesNo Default false,"                                 '是否为html
            sql = sql & "isOnHtml YesNo Default Yes,"                                 '开启生成Html



            sql = sql & "articleInfoStyle VarChar Default """","                      '列表内容样式
            sql = sql & "articleInfoPhotoWidth VarChar Default """","                 '内容页图片宽
            sql = sql & "articleInfoPhotoHeight VarChar Default """","                '内容页图片高 0为引用上面的 -1为空


            sql = sql & "relatedTags VarChar Default """","                           '相关标签20141219
            sql = sql & "weight Int Default 0,"                                       '权重：     (越小越靠前)



            sql = sql & "noteBody Text Default """","                                 '笔记信息
            sql = sql & "aboutcontent Text Default """","                             '简单介绍
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 

        '单页表-------------------------------
        tableName = DB_PREFIX & "OnePage" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "title VarChar Default """","                                 '标题
            sql = sql & "displayTitle VarChar Default """","                          '显示标题
            sql = sql & "adminId int Default 0,"                                      '管理员id

            sql = sql & "webTitle VarChar Default """","                              '网站自定标题
            sql = sql & "webKeywords VarChar Default """","                           '网站关键词
            sql = sql & "webDescription VarChar Default """","                        '网站描述
            sql = sql & "folderName VarChar Default """","                            'HTMl文件夹名称
            sql = sql & "fileName VarChar Default """","                              'HTML文件名称
            sql = sql & "customAUrl VarChar Default """","                            '自定义链接网址
            sql = sql & "templatePath VarChar Default """","                          '模板路径
            sql = sql & "target VarChar Default """","                                '链接打开方式
            sql = sql & "fontColor VarChar Default """","                             '字颜色
            sql = sql & "fontB YesNo Default No,"                                     '字粗
            sql = sql & "noFollow int Default 0,"                                     '不要追踪此网页上的链接
            sql = sql & "sortRank Int Default 0,"                                     '顺序
            sql = sql & "views int Default 0,"                                        '点击次数
            sql = sql & "isRecommend YesNo Default Yes,"                              '推荐
            sql = sql & "lableTitle VarChar Default """","                            '标签描述
            sql = sql & "banner VarChar Default """","                                '当前Banner图片

            sql = sql & "isHtml YesNo Default false,"                                 '是否为html
            sql = sql & "isOnHtml YesNo Default Yes,"                                 '开启生成Html

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间

            sql = sql & "noteBody Text Default """","                                 '笔记信息
            sql = sql & "aboutcontent Text Default """","                             '简单介绍
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 

        '文章评论表 20160129-------------------------------
        tableName = DB_PREFIX & "TableComment" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "UserID Int Default 0,"                                       '用户ID
            sql = sql & "itemID Int Default 0,"                                       '项目ID 产品ID
            sql = sql & "tableName VarChar Default """","                             '表名称
            sql = sql & "userName VarChar Default """","                              '用户名称
            sql = sql & "title VarChar Default """","                                 '评论标题
            sql = sql & "EMail VarChar Default """","                                 '邮箱
            sql = sql & "tel VarChar Default """","                                   '电话ID
            sql = sql & "ip VarChar Default """","                                    '评论者IP

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间

            sql = sql & "reply Text Default """","                                    '回复内容
            sql = sql & "noteBody Text Default """","                                 '笔记信息

            sql = sql & "isThrough Int Default 0,"                                    '是否通过
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
		
        '网站信息表-------------------------------
        tableName = DB_PREFIX & "WebSite" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "WebSiteUrl VarChar Default """","                            '网站地址
            sql = sql & "webSiteBottom Text Default """","                            '底部信息 WebSiteBottom
            sql = sql & "WebSiteFlow int Default 999,"                                '网站统计数
            sql = sql & "WebSiteFlowStyle int Default 1,"                             '网站统计样式
            sql = sql & "WebSiteFlowMedian int Default 6,"                            '网站统计几位数Median中位数
            sql = sql & "ProductList VarChar Default """","                           '产品列表
            sql = sql & "NewsList VarChar Default """","                              '新闻列表
            sql = sql & "NewsDid VarChar Default """","                               '新闻大类
            '201109追加
            sql = sql & "TZ51La VarChar Default """","                                '统计网站
            sql = sql & "UserEmail VarChar Default """","                             '用户Email
            sql = sql & "ProductDid VarChar Default """","                            '产品大类
            sql = sql & "TemplateIndex Text Default '首页',"                          '模板首页
            sql = sql & "TemplateHome Text Default '内页',"                           '模板内页
            sql = sql & "TemplateMain Text Default '详细页',"                         '模板详细页
            sql = sql & "TemplateMain2 Text Default '详细页',"                        '模板详细页2
            sql = sql & "TemplateMain3 Text Default '详细页',"                        '模板详细页3
            sql = sql & "UseNumb Text Default """","                                  '网站使用次数
            '20111118追加
            sql = sql & "WebRecord VarChar Default """","                             '网站记录
            sql = sql & "ContentWebRecord Text Default """","                         '全部网站记录
            '20111128追加
            sql = sql & "UseHttpUrl VarChar Default """","                            '当前使用网址
            sql = sql & "TempUseHttpUrl Text Default """","                           '存储使用过的网址
            sql = sql & "WebDate DateTime Default Date(),"                            '网站时间
            '2013,12,18
            sql = sql & "webTitle VarChar Default """","                              '网站自定标题
            sql = sql & "webKeywords VarChar Default """","                           '网站关键词
            sql = sql & "webDescription VarChar Default """","                        '网站描述

            sql = sql & "WebTemplate VarChar Default '\Templates',"                   '网站模板路径
            sql = sql & "WebSkins VarChar Default '\Skins\Default\',"                 '网站样式路径'暂时不要用

            sql = sql & "WebFolderName VarChar Default 'test3',"                      '网站文件夹名称 new20150721


            sql = sql & "WebImages VarChar Default '\Skins\Default\Images\',"         '网站Images路径
            sql = sql & "WebCss VarChar Default '\Skins\Default\Css\',"               '网站Css路径
            sql = sql & "WebJs VarChar Default '\Skins\Default\Js\',"                 '网站Js路径

            sql = sql & "AddWebSite YesNo Default No,"                                'URL前添加域名
            sql = sql & "UpdateHtml YesNo Default No,"                                '修改后更新HTML
            sql = sql & "isHtmlFormatting YesNo Default No,"                          '是否为HTML格式化
            sql = sql & "isWebLabelClose YesNo Default No,"                           '是否为闭合网页标签
            sql = sql & "isCnToEn YesNo Default No,"                                  '是否为文本名汉字转拼音
            sql = sql & "flags VarChar Default """","                                 '旗


            '20150114
            sql = sql & "ModuleSkins VarChar Default """","                           '模块皮肤
            sql = sql & "FindTpl VarChar Default """","                               '查找要替换模板内容
            sql = sql & "ReplaceTpl VarChar Default """","                            '替换模板指定内容，批量

            sql = sql & "WebCodeFindTpl VarChar Default """","                        '查找要替换网页内容
            sql = sql & "WebCodeReplaceTpl VarChar Default """","                     '替换网页指定内容，批量

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间

            '最后垫底
            sql = sql & "WebHtml VarChar Default 0)"                                  '网站运行状态，如ASP，HTML
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 

        '管理员表-------------------------------
        tableName = DB_PREFIX & "Admin" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "userName VarChar Default """","                              '账号
            sql = sql & "pwd VarChar Default """","                                   '密码
            sql = sql & "pseudonym VarChar Default """","                             '笔名


            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间

            sql = sql & "regIP VarChar Default """","                                 '注册IP
            sql = sql & "upIP VarChar Default """","                                  '修改IP


            sql = sql & "quanXian Text Default """","                                 '管理权限
            sql = sql & "verificationMode Int Default 1,"                             '验证方式，1为Session
            sql = sql & "adminLevel VarChar Default """","                            '级别
            sql = sql & "channel VarChar Default """","                               '频道
            sql = sql & "mtest YesNo Default No,"                                     '修改后更新HTML
            sql = sql & "flags VarChar Default """","                                 '旗        只有这个后台做了处理


            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
        '搜索统计表-------------------------------
        tableName = DB_PREFIX & "SearchStat" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "title VarChar Default """","                                 '搜索关键词
            sql = sql & "isThrough YesNo Default No,"                                 '通过
            sql = sql & "webTitle VarChar Default """","                              '网站自定标题
            sql = sql & "webKeywords VarChar Default """","                           '网站关键词
            sql = sql & "webDescription VarChar Default """","                        '网站描述
            sql = sql & "folderName VarChar Default """","                            'HTMl文件夹名称
            sql = sql & "fileName VarChar Default """","                              'HTML文件名称
            sql = sql & "customAUrl VarChar Default """","                            '自定义链接网址
            sql = sql & "templatePath VarChar Default """","                          '模板路径
            sql = sql & "target VarChar Default """","                                '链接打开方式
            sql = sql & "isHtml YesNo Default false,"                                 '是否为html
            sql = sql & "isOnHtml YesNo Default Yes,"                                 '是否开启生成Html
            sql = sql & "views int Default 0,"                                        '点击次数
            sql = sql & "author VarChar Default """","                                '作者

            sql = sql & "sortRank int Default 0,"                                     '排序编号
            sql = sql & "fontColor VarChar Default """","                             '字颜色
            sql = sql & "noFollow int Default 0,"                                     '不要追踪此网页上的链接
            sql = sql & "flags VarChar Default """","                                 '旗

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
            sql = sql & "aboutcontent Text Default """","                             '简单介绍
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
		'友情链接表-------------------------------
        tableName = DB_PREFIX & "FriendLink" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 

            sql = sql & "adminId int Default 0,"                                      '管理员id
            sql = sql & "title VarChar Default """","                                 '标题
            sql = sql & "titleColor VarChar Default """","                            '标题颜色
            sql = sql & "lableTitle VarChar Default """","                            '标签描述
            sql = sql & "httpUrl VarChar Default """","                               '网址
            sql = sql & "sortRank int Default 0,"                                     '排序编号

            sql = sql & "titleAlt VarChar Default """","                              '标题alt
            sql = sql & "smallImage VarChar Default """","                            '小图片地址
            sql = sql & "smallimageAlt VarChar Default """","                         '小图alt


            sql = sql & "isThrough YesNo Default No,"                                 '通过审核
            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间


            sql = sql & "webTitle VarChar Default """","                              '网站自定标题
            sql = sql & "webKeywords VarChar Default """","                           '网站关键词
            sql = sql & "webDescription VarChar Default """","                        '网站描述
            sql = sql & "folderName VarChar Default """","                            'HTMl文件夹名称
            sql = sql & "fileName VarChar Default """","                              'HTML文件名称
            sql = sql & "templatePath VarChar Default """","                          '模板路径
            sql = sql & "target VarChar Default """","                                '链接打开方式
            sql = sql & "customAUrl VarChar Default """","                            '自定义链接网址
            sql = sql & "fontColor VarChar Default """","                             '字颜色
            sql = sql & "noFollow int Default 0,"                                     '不要追踪此网页上的链接
            sql = sql & "flags VarChar Default """","                                 '旗

            sql = sql & "isHtml YesNo Default false,"                                 '是否为html
            sql = sql & "isOnHtml YesNo Default Yes,"                                 '开启生成Html

            sql = sql & "weight Int Default 0,"                                       '权重：     (越小越靠前)

            sql = sql & "noteBody Text Default """","                                 '笔记信息
            sql = sql & "aboutcontent Text Default """","                             '简单介绍
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
        '招聘表-------------------------------
        tableName = DB_PREFIX & "Job" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "title VarChar Default """","                                 '职位名称
            sql = sql & "sex VarChar Default """","                                   '性别
            sql = sql & "age VarChar Default """","                                   '年龄
            sql = sql & "education VarChar Default """","                             '教育
            sql = sql & "workArea VarChar Default """","                              '工作区
            sql = sql & "monthly VarChar Default """","                               '月薪待遇
            sql = sql & "startDateTime VarChar Default """","                         '发布日期
            sql = sql & "endDateTime VarChar Default """","                           '截止日期
            sql = sql & "webTitle VarChar Default """","                              '网站自定标题
            sql = sql & "webKeywords VarChar Default """","                           '网站关键词
            sql = sql & "webDescription VarChar Default """","                        '网站描述
            sql = sql & "folderName VarChar Default """","                            'HTMl文件夹名称
            sql = sql & "fileName VarChar Default """","                              'HTML文件名称
            sql = sql & "templatePath VarChar Default """","                          '模板路径
            sql = sql & "target VarChar Default """","                                '链接打开方式

            sql = sql & "fontColor VarChar Default """","                             '字颜色
            sql = sql & "noFollow int Default 0,"                                     '不要追踪此网页上的链接
            sql = sql & "flags VarChar Default """","                                 '旗
            sql = sql & "sortRank int Default 0,"                                     '排序编号
            sql = sql & "titleAlt VarChar Default """","                              '标题alt
            sql = sql & "smallImage VarChar Default """","                            '小图片地址
            sql = sql & "smallimageAlt VarChar Default """","                         '小图alt
            sql = sql & "isThrough Int Default 0,"                                    '是否通过

            sql = sql & "isHtml YesNo Default false,"                                 '是否为html
            sql = sql & "isOnHtml YesNo Default Yes,"                                 '开启生成Html

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
            sql = sql & "noteBody Text Default """","                                 '笔记信息
            sql = sql & "aboutcontent Text Default """","                             '简单介绍
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
        '留言表-------------------------------
        tableName = DB_PREFIX & "GuestBook" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "columnid VarChar Default """","                              '栏目ID
            sql = sql & "parentId int Default -1,"                                    '父栏目ID
            sql = sql & "title VarChar Default """","                                 '标题
            sql = sql & "guestName VarChar Default """","                             '姓名
            sql = sql & "tel VarChar Default """","                                   '电话
            sql = sql & "fax VarChar Default """","                                   '传真
            sql = sql & "EMail VarChar Default """","                                 '邮箱
            sql = sql & "mobile VarChar Default """","                                '手机
            sql = sql & "QQ VarChar Default """","                                    'QQ号码
            sql = sql & "MSN VarChar Default """","                                   'MSN号码
            sql = sql & "company VarChar Default """","                               '公司
            sql = sql & "address VarChar Default """","                               '地址
            sql = sql & "postcode VarChar Default """","                              '邮编
            sql = sql & "ip VarChar Default """","                                    'IP地址
            sql = sql & "webTitle VarChar Default """","                              '网站自定标题
            sql = sql & "webKeywords VarChar Default """","                           '网站关键词
            sql = sql & "webDescription VarChar Default """","                        '网站描述
            sql = sql & "folderName VarChar Default """","                            'HTMl文件夹名称
            sql = sql & "fileName VarChar Default """","                              'HTML文件名称
            sql = sql & "customAUrl VarChar Default """","                            '自定义链接网址
            sql = sql & "templatePath VarChar Default """","                          '模板路径
            sql = sql & "target VarChar Default """","                                '链接打开方式
            sql = sql & "fontColor VarChar Default """","                             '字颜色

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
            sql = sql & "isThrough Int Default 0,"                                    '是否通过


            sql = sql & "reply Text Default """","                                    '回复内容
            sql = sql & "replyip Text Default """","                                  '回复IP
            sql = sql & "replydatetime Text Default """","                            '回复时间

            sql = sql & "noteBody Text Default """","                                 '笔记信息
            sql = sql & "aboutcontent Text Default """","                             '简单介绍
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
        '反馈表-------------------------------
        tableName = DB_PREFIX & "Feedback" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "columnid VarChar Default """","                              '栏目ID
            sql = sql & "title VarChar Default """","                                 '标题
            sql = sql & "feedbacktype VarChar Default """","                          '反馈类型
            sql = sql & "guestName VarChar Default """","                             '姓名
            sql = sql & "tel VarChar Default """","                                   '电话
            sql = sql & "fax VarChar Default """","                                   '传真
            sql = sql & "EMail VarChar Default """","                                 '邮箱
            sql = sql & "mobile VarChar Default """","                                '手机
            sql = sql & "QQ VarChar Default """","                                    'QQ号码
            sql = sql & "MSN VarChar Default """","                                   'MSN号码
            sql = sql & "company VarChar Default """","                               '公司
            sql = sql & "address VarChar Default """","                               '地址
            sql = sql & "postcode VarChar Default """","                              '邮编
            sql = sql & "ip VarChar Default """","                                    'IP地址
            sql = sql & "webTitle VarChar Default """","                              '网站自定标题
            sql = sql & "webKeywords VarChar Default """","                           '网站关键词
            sql = sql & "webDescription VarChar Default """","                        '网站描述
            sql = sql & "folderName VarChar Default """","                            'HTMl文件夹名称
            sql = sql & "fileName VarChar Default """","                              'HTML文件名称
            sql = sql & "customAUrl VarChar Default """","                            '自定义链接网址
            sql = sql & "templatePath VarChar Default """","                          '模板路径
            sql = sql & "target VarChar Default """","                                '链接打开方式
            sql = sql & "fontColor VarChar Default """","                             '字颜色

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
            sql = sql & "isThrough Int Default 0,"                                    '是否通过


            sql = sql & "reply Text Default """","                                    '回复内容
            sql = sql & "replyip Text Default """","                                  '回复IP
            sql = sql & "replydatetime Text Default """","                            '回复时间

            sql = sql & "noteBody Text Default """","                                 '笔记信息
            sql = sql & "aboutcontent Text Default """","                             '简单介绍
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
        '后台菜单表-------------------------------
        tableName = DB_PREFIX & "ListMenu" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "title VarChar Default """","                                 '标题
            sql = sql & "parentId int Default -1,"                                    '父栏目ID
            sql = sql & "sortRank int Default 0,"                                     '排序编号
            sql = sql & "lablename VarChar Default """","                             '标签名称
            sql = sql & "customAUrl VarChar Default """","                            '自定义链接网址

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
            sql = sql & "isDisplay YesNo,"                                            '显示与隐藏
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
        '在线QQ表------------------------------
        tableName = DB_PREFIX & "LineQQ" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "BigClassName VarChar Default """","                          '大类名称
            sql = sql & "title VarChar Default """","                                 '标题
            sql = sql & "QQ VarChar Default """","                                    'QQ号码
            sql = sql & "IsOnlineChat YesNo Default No,"                              '是否为在线聊天
            sql = sql & "IsAddFriend YesNo Default Yes,"                              '是否为加为好友

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
            sql = sql & "bodycontent Text)"                                           '更多详细内容
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
        '投票表-------------------------------
        tableName = DB_PREFIX & "LineVote" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "title VarChar Default """","                                 '投票标题
            sql = sql & "Option1 VarChar Default """","                               '投票选项一
            sql = sql & "Option2 VarChar Default """","                               '投票选项二
            sql = sql & "Option3 VarChar Default """","                               '投票选项三
            sql = sql & "Option4 VarChar Default """","                               '投票选项四
            sql = sql & "Option5 VarChar Default """","                               '投票选项五
            sql = sql & "Option6 VarChar Default """","                               '投票选项六
            sql = sql & "Num1 Int,"                                                   '投票选项数一
            sql = sql & "Num2 Int,"                                                   '投票选项数二
            sql = sql & "Num3 Int,"                                                   '投票选项数三
            sql = sql & "Num4 Int,"                                                   '投票选项数四
            sql = sql & "Num5 Int,"                                                   '投票选项数五
            sql = sql & "Num6 Int,"                                                   '投票选项数六
            sql = sql & "isDisplay YesNo Default Yes,"                                '投票显示关闭
            sql = sql & "webTitle VarChar Default """","                              '网站自定标题
            sql = sql & "webKeywords VarChar Default """","                           '网站关键词
            sql = sql & "webDescription VarChar Default """","                        '网站描述
            sql = sql & "folderName VarChar Default """","                            'HTMl文件夹名称
            sql = sql & "fileName VarChar Default """","                              'HTML文件名称
            sql = sql & "banner VarChar Default """","                                '当前Banner图片
            sql = sql & "templatePath VarChar Default """","                          '模板路径
            sql = sql & "target VarChar Default """","                                '链接打开方式
            sql = sql & "fontColor VarChar Default """","                             '字颜色
            sql = sql & "fontB YesNo Default No,"                                     '字粗
            sql = sql & "OnHtml YesNo Default Yes,"                                   '开启生成Html

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
            sql = sql & "VoteType YesNo Default Yes)"                                 '投票类型
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
        '会员表-------------------------------
        tableName = DB_PREFIX & "Member" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "userType VarChar Default """","                          '会员类型
            sql = sql & "userName VarChar Default """","                              '会员账号
            sql = sql & "pwd VarChar Default """","                                   '会员密码
            sql = sql & "yunPwd VarChar Default """","                                '原密码
            sql = sql & "sex VarChar Default '男',"                                   '性别
            sql = sql & "age Int Default 18,"                                         '年龄
            sql = sql & "tel VarChar Default """","                                   '电话
            sql = sql & "phone VarChar Default """","                                 '手机
            sql = sql & "fax VarChar Default """","                                   '传真
            sql = sql & "email VarChar Default """","                                 '邮箱
            sql = sql & "postcode VarChar Default """","                              '邮编
            sql = sql & "address VarChar Default """","                               '地址
            sql = sql & "company VarChar Default """","                               '公司

            sql = sql & "regIP VarChar Default """","                                 '注册IP
            sql = sql & "loginIp VarChar Default """","                                 'IP

            sql = sql & "loginCount Int Default 0,"                                    '登录次数
            sql = sql & "loginTime DateTime Default Date(),"                          '登录时间
            sql = sql & "lastLoginTime DateTime Default Date(),"                      '上次登录时间
			
			

			'QQ信息
            sql = sql & "openId VarChar Default """","                                 'openId
            sql = sql & "accessToken VarChar Default """","                                 'accessToken
            sql = sql & "nickname VarChar Default """","                                 'nickname
            sql = sql & "qqphoto VarChar Default """","                               'QQ头像
            sql = sql & "useryear Int Default 18,"                                         '那一年出生

 
            sql = sql & "Province VarChar Default """","                              '省份
            sql = sql & "City VarChar Default """","                                  '城市
            sql = sql & "Area VarChar Default """","                                  '区域20150305



            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
            sql = sql & "isThrough Int Default 0,"                                    '是否通过
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
        '付款信息表-------------------------------
        tableName = DB_PREFIX & "Payment" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "UserName VarChar Default """","                              '会员账号
            sql = sql & "MemberId VarChar Default """","                              '会员ID编号
            sql = sql & "Sex VarChar Default '男',"                                   '性别
            sql = sql & "Age Int Default 18,"                                         '年龄
            sql = sql & "Tel VarChar Default """","                                   '电话
            sql = sql & "Phone VarChar Default """","                                 '手机
            sql = sql & "Fax VarChar Default """","                                   '传真
            sql = sql & "Email VarChar Default """","                                 '邮箱
            sql = sql & "postcode VarChar Default """","                              '邮编
            sql = sql & "Address VarChar Default """","                               '地址
            sql = sql & "Company VarChar Default """","                               '公司
            sql = sql & "QQMSN VarChar Default """","                                 '公司
            sql = sql & "Ip VarChar Default """","                                    'IP

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
            sql = sql & "isThrough Int Default 0)"                                    '是否通过
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
        '购物信息表-------------------------------
        tableName = DB_PREFIX & "Previeworder" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "MemberId VarChar Default """","                              '会员编号账号
            sql = sql & "OrderId VarChar Default """","                               '订单ID编号
            sql = sql & "ProductId VarChar Default """","                             '产品ID编号
            sql = sql & "title VarChar Default '男',"                                 '产品名称
            sql = sql & "Total VarChar Default """","                                 '些产品总数
            sql = sql & "Price Int Default 18,"                                       '此产品单价
            sql = sql & "ProductSum VarChar Default """","                            '此产品总

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
            sql = sql & "isThrough YesNo Default No)"                                   '是否发货
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
        '产品评论表-------------------------------
        tableName = DB_PREFIX & "ProductComment" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "UserName VarChar Default """","                              '评论者
            sql = sql & "title VarChar Default """","                                 '评论标题
            sql = sql & "Pid Int,"                                                    '信息ID
            sql = sql & "PTitle VarChar Default """","                                '信息标题
            sql = sql & "bodycontent Text Default """","                              '评论内容
            sql = sql & "Sort Int Default 1,"                                         '评论排序
            sql = sql & "Ip VarChar Default """","                                    '评论者IP

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
            sql = sql & "isThrough YesNo Default No)"                                   '是否显示评论
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
        '网站统计表 20160203-------------------------------
        tableName = DB_PREFIX & "WebSiteStat" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "visiturl Text Default """","                                 '来访网址
            sql = sql & "viewurl Text Default """","                                  '浏览网址
            sql = sql & "browser VarChar Default """","                               '浏览器版本
            sql = sql & "operatingSystem VarChar Default """","                       '操作系统
            sql = sql & "screenWH VarChar Default """","                              '屏幕宽高
            sql = sql & "moreInfo Text Default """","                                 '更多信息
            sql = sql & "viewDateTime DateTime Default Now(),"                        '访问时间
            sql = sql & "ip VarChar Default """","                                    '访问IP

            sql = sql & "dateClass VarChar Default """","                             '日期分类

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间



            sql = sql & "isThrough Int Default 0,"                                    '是否通过
            sql = sql & "noteInfo Text Default """","                                 '笔记信息
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
        '系统日志表-------------------------------
        tableName = DB_PREFIX & "SystemLog" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "msgStr Text Default """","                                   '信息内容
            sql = sql & "tableName VarChar Default """","                             '表名 也可以说module模块
            sql = sql & "url Text Default """","                                   '动作
            sql = sql & "adminId int Default 0,"                                      '管理员id
            sql = sql & "adminName VarChar Default 0,"                                '管理员名称
            sql = sql & "IP VarChar Default """","                                    'IP地址


            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
		'采集网站表-------------------------------
        tableName = DB_PREFIX & "CaiWeb" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "BigClassName VarChar Default """","                          '大类名称		
            sql = sql & "httpurl VarChar Default """","                            	  '网址
            sql = sql & "morePageUrl VarChar Default """","                           '更多页配置  {*}		为i页
            sql = sql & "charset VarChar Default """","                            	  '字符编码
            sql = sql & "thisPage int Default 0,"                                      '当前页数
            sql = sql & "countPage int Default 0,"                                      '总页数 
            sql = sql & "sType VarChar Default """","                          		  '类型	
            sql = sql & "sortRank int Default 0,"                                     '排序编号

            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
            sql = sql & "IP VarChar Default """","                                    'IP地址
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If
        '采集配置表-------------------------------
        tableName = DB_PREFIX & "CaiConfig" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "BigClassName VarChar Default """","                          '大类名称
            sql = sql & "sType VarChar Default """","                          		  '类型
            sql = sql & "startStr VarChar Default """","                          	  '开始截取字符
            sql = sql & "endStr VarChar Default """","                          	  '结束截取字符			
            sql = sql & "startAddStr VarChar Default """","                           '开始截取字符前面追加字符
            sql = sql & "endAddStr VarChar Default """","                          	  '结束截取字符后台追加字符
			
            sql = sql & "sortRank int Default 0,"                                     '排序编号 


            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
            sql = sql & "IP VarChar Default """","                                    'IP地址
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If
        '采集数据表-------------------------------
        tableName = DB_PREFIX & "CaiData" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "BigClassName VarChar Default """","                          '大类名称
            sql = sql & "sType VarChar Default """","                          		  '类型

            sql = sql & "sortRank int Default 0,"                                     '排序编号
            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
            sql = sql & "IP VarChar Default """","                                    'IP地址
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If
		
        '布局表 20151228-------------------------------
        tableName = DB_PREFIX & "webLayout" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "layoutName VarChar Default """","                            '布局名称
            sql = sql & "layoutList VarChar Default """","                            '布局列表
			
            sql = sql & "sourceStr VarChar Default """","                          '被替换内容
            sql = sql & "replaceStr VarChar Default """","                          '统一替换内容
			
            sql = sql & "actionContent Text Default """","                            '动作内容

            sql = sql & "sortRank int Default 0,"                                     '排序编号
            sql = sql & "isDisplay YesNo,"                                            '是否显示
            sql = sql & "author VarChar Default """","                                '作者
            sql = sql & "views int Default 0,"                                        '点击次数
            sql = sql & "adminId int Default 0,"                                      '管理员id


            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间

            sql = sql & "aboutcontent Text Default """","                             '简单介绍
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
        '模块表 20151228-------------------------------
        tableName = DB_PREFIX & "webModule" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "moduleType VarChar Default """","                            '模块类型
            sql = sql & "moduleName VarChar Default """","                            '模块名称
            sql = sql & "title VarChar Default """","                                 '标题
            sql = sql & "sortRank int Default 0,"                                     '排序编号
            sql = sql & "isDisplay YesNo,"                                            '是否显示
            sql = sql & "author VarChar Default """","                                '作者
            sql = sql & "views int Default 0,"                                        '点击次数
            sql = sql & "adminId int Default 0,"                                      '管理员id


            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间

            sql = sql & "aboutcontent Text Default """","                             '简单介绍
            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If 
        '竞价表(百度竞价) 20151228-------------------------------
        tableName = DB_PREFIX & "Bidding" 
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key," 
            sql = sql & "searchWords VarChar Default """","                           '搜索词
            sql = sql & "webKeywords VarChar Default """","                           '网站关键词
            sql = sql & "showReason VarChar Default """","                            '展现理由
            sql = sql & "nComputerSearch int Default 0,"                              '电脑搜索量
            sql = sql & "nMoblieSearch int Default 0,"                                '移动搜索量
            sql = sql & "nCountSearch int Default 0,"                                 '总体搜索量

            sql = sql & "nWordPrice int Default 0,"                                   '新词左侧（上方）指导价
            sql = sql & "nDegree int Default 0,"                                      '竞争激烈程度

            sql = sql & "sortRank Int Default 0,"                                     '顺序
            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间


            sql = sql & "bodycontent Text Default """")"                              '备注
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql 
        End If
		'网站URL扫描表 20160428-------------------------------
        tableName = DB_PREFIX & "WebUrlScan"
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key,"
            sql = sql & "bigClassName VarChar Default """","                           '大类名称
            sql = sql & "linkType VarChar Default """","                           '链接类型
            sql = sql & "website VarChar Default """","                           '域名
            sql = sql & "title VarChar Default """","                            '标题
            sql = sql & "httpurl VarChar Default """","                            '网址
            sql = sql & "totitle VarChar Default """","                            '来访标题
            sql = sql & "tohttpurl VarChar Default """","                            '来访网址
            sql = sql & "webState int Default 0,"                                   '状态
            sql = sql & "openSpeed int Default 0,"                                   '打开时间
            sql = sql & "charset VarChar Default """","                            '字符编码
            sql = sql & "content_type VarChar Default """","                            '内容类型		
            sql = sql & "link_count Int Default 0,"                                     '链接总数
				
            sql = sql & "webSize int Default 0,"                                   '网页大小
            sql = sql & "sortRank Int Default 0,"                                     '顺序
            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
            sql = sql & "isThrough YesNo Default No)"                                 '通过审核
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql
        End If
		'网站域名表 20160511-------------------------------
        tableName = DB_PREFIX & "WebDomain"
        If showIsTable(tableName) = False Then
            sql = "Create Table " & tableName & " (Id int Identity(0,1) Primary Key,"
            sql = sql & "bigClassName VarChar Default """","                           '大类名称
            sql = sql & "website VarChar Default """","                           '域名
            sql = sql & "webTitle VarChar Default """","                              '网站自定标题
            sql = sql & "webKeywords VarChar Default """","                           '网站关键词
            sql = sql & "webDescription VarChar Default """","                        '网站描述
			
            sql = sql & "webState int Default 0,"                                   '状态
            sql = sql & "openSpeed int Default 0,"                                   '打开时间
            sql = sql & "charset VarChar Default """","                            '字符编码
            sql = sql & "content_type VarChar Default """","                            '内容类型
            sql = sql & "server_name VarChar Default """","                            '服务器名称
			
			
            sql = sql & "isAsp YesNo Default No,"                                  '可运行Asp程序
            sql = sql & "isAspx YesNo Default No,"                                  '可运行Aspx程序
            sql = sql & "isPhp YesNo Default No,"                                  '可运行Php程序
            sql = sql & "isJsp YesNo Default No,"                                  '可运行Jsp程序
            sql = sql & "isHtm YesNo Default No,"                                  '可运行Htm程序
            sql = sql & "isHtml YesNo Default No,"                                  '可运行Html程序
            sql = sql & "nlinks int Default 0,"                                 '外链数
            sql = sql & "links Text Default """","                                 '外链列表
			
			
            sql = sql & "homepagelist VarChar Default """","                                 '旗
			
            sql = sql & "flags VarChar Default """","                                 '旗
            sql = sql & "webSize int Default 0,"                                      '网页大小
            sql = sql & "sortRank Int Default 0,"                                     '顺序			
            sql = sql & "addDateTime DateTime Default Now(),"                         '创建时间
            sql = sql & "upDateTime DateTime Default Date(),"                         '修改时间
			
            sql = sql & "isDomain YesNo Default No,"                                  '是否为域名
            sql = sql & "isThrough YesNo Default No)"                                 '通过审核
            If MDBPath = "" Then sql = handleSqlServer(sql)                           '把Access数据库类型转成SqlServer数据库类型
            conn.Execute sql
        End If

        addData = True 
        If addData = True Then
            'Add Admin
            splStr = Split("admin|test|guest", "|") 
            For Each s In splStr
                rs.Open "Select * From " & DB_PREFIX & "Admin Where UserName='" & s & "'", conn, 1, 3 
                If rs.EOF Then
                    rs.AddNew 
                    rs("UserName") = s 
                    rs("Pwd") = myMD5(s) 
                    If s = "admin" Then
                        rs("flags") = "|*|" 
                    Else
                        rs("flags") = "|显示站点配置|显示网站栏目|显示分类信息|显示评论|显示单页管理|显示后台管理员|显示搜索统计|显示网站统计|显示生成HTML|显示生成SiteMap|显示生成Robots|显示模板管理|" 
                    End If 
                    rs.Update 
                End If : rs.Close 
            Next 

            'Add WebSite
            rs.Open "Select * From " & DB_PREFIX & "WebSite", conn, 1, 3 
            If rs.EOF Then
                rs.AddNew 
                rs("webTitle") = "网站标题" 
                rs("webKeywords") = "网站关键词" 
                rs("webDescription") = "网站描述" 
                rs("WebSiteUrl") = "http://www.baidu.com" 
                rs("WebSiteBottom") = "版权所有 上海子映有限公司 电话：021-66666666 传真：021-8888888<br>联系QQ：313801120 邮箱：313801120@qq.com" 
                rs("TemplateIndex") = templateIndex 
                rs("TemplateHome") = templateHome 
                rs("TemplateMain") = templateMain 
                rs("ModuleSkins") = ""                                                          '皮肤为空
                rs.Update 
            End If : rs.Close 
            'add menu

            Dim parentid, title, lableName, content, tempS, url, isdisplay, nCount 
            content = getftext("/web/后台菜单配置.ini") 
            content = Replace(content, vbTab, "    ") 
            splStr = Split(content, vbCrLf) 
            nCount = 0 
            For Each s In splStr
                tempS = s 
                s = Trim(s)
                If tempS <> "" and left(phptrim(s)&" ",1)<>"#" Then
                    nCount = nCount + 1         '总数
                    If Left(tempS, 4) = "    " Then
                    Else
                        parentid = "-1" 
                    End If 

                    If LCase(tempS) = "end" Then
                        Exit For 
                    ElseIf s <> "" Then
                        title = Mid(s & " ", 1, InStr(s & " ", " ") - 1) 
                        lableName = getStrCut(s, "lablename='", "'", 2) 
                        url = getStrCut(s, "url='", "'", 2) 
                        isdisplay = getStrCut(s, "isdisplay='", "'", 2) 
                        If isdisplay = "" Then
                            isdisplay = 1 
                        End If 
                        if title<>"" then
							Call echo("lablename", lableName) 
							
							conn.Execute("insert into " & DB_PREFIX & "ListMenu (title,parentid,sortrank,lablename,isdisplay,customaurl) values('" & title & "'," & parentid & "," & nCount & ",'" & lableName & "'," & isdisplay & ",'" & url & "')") 
							If parentid = "-1" Then
								rs.Open "select * from " & DB_PREFIX & "ListMenu where title='" & title & "'", conn, 1, 1 
								If Not rs.EOF Then
									parentid = rs("id") 
								End If : rs.Close 
							End If 
	
							Call echo(title, s) 
							Call echo(url, lableName) 
						end if
                    End If 
                End If 
            Next 


        End If 

    End Sub 
End Class 


'http://127.0.0.1"& adminDir &"TableManage.Asp?TableName=ZNRobot&ClassName=获得方案&ThisMDBPath=/../网站备份\数据库/ZNRobot.mdb

Function displayLogin()
    %>     
        <form id="form1" name="form1" method="post" action="?act=login">  
          密码：  
          <input type="password" name="password" id="password" />  
          <input type="submit" name="button" id="button" value="登录" />  
        </form>  
<% 
End Function 
%>  
