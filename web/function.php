<?PHP

//调用function文件函数
function callFunction(){
    switch ( @$_REQUEST['stype'] ){
        case 'updateWebsiteStat' ; updateWebsiteStat() ;break;//更新网站统计
        case 'clearWebsiteStat' ; clearWebsiteStat() ;break;//清空网站统计
        case 'updateTodayWebStat' ; updateTodayWebStat() ;break;//更新网站今天统计
        case 'websiteDetail' ; websiteDetail() ;break;//详细网站统计
        case 'displayAccessDomain' ; displayAccessDomain();										//显示访问域名
        break;
        default ; eerr('function1页里没有动作', @$_REQUEST['stype']);
    }
}

//显示访问域名
function displayAccessDomain(){
    $visitWebSite='';$visitWebSiteList='';$urlList='';$nOK='';
    $GLOBALS['conn=']=OpenConn();
    $nOK=0;
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'websitestat');
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        $visitWebSite=strtolower(getWebSite($rs['visiturl']));
        //call echo("visitWebSite",visitWebSite)
        if( instr(vbCrlf() . $visitWebSiteList . vbCrlf(),vbCrlf() . $visitWebSite . vbCrlf())==false ){
            if( $visitWebSite<>strtolower(getWebSite(webDoMain())) ){
                $visitWebSiteList=$visitWebSiteList . $visitWebSite . vbCrlf();
                $nOK=$nOK+1;
                $urlList=$urlList . $nOK . '、<a href=\'' . $rs['visiturl'] . '\' target=\'_blank\'>' . $rs['visiturl'] . '</a><br>';
            }
        }
    }
    ASPEcho('显示访问域名','操作完成 <a href=\'javascript:history.go(-1)\'>点击返回</a>');
    rwend($visitWebSiteList . '<br><hr><br>' . $urlList);
}
//获得处理后表列表 20160313
function getHandleTableList(){
    $s=''; $lableStr ='';
    $lableStr= '表列表[' . @$_REQUEST['mdbpath'] . ']';
    if( $GLOBALS['WEB_CACHEContent']== '' ){
        $GLOBALS['WEB_CACHEContent']= getftext($GLOBALS['WEB_CACHEFile']);
    }
    $s= getConfigContentBlock($GLOBALS['WEB_CACHEContent'], '#' . $lableStr . '#');
    if( $s== '' ){
        $s= strtolower(getTableList());
        $s= '|' . Replace($s, vbCrlf(), '|') . '|';
        $GLOBALS['WEB_CACHEContent']= setConfigFileBlock($GLOBALS['WEB_CACHEFile'], $s, '#' . $lableStr . '#');
        if( $GLOBALS['isCacheTip']==true ){
            ASPEcho('缓冲', $lableStr);
        }
    }
    $getHandleTableList= $s;
    return @$getHandleTableList;
}

//获得处理的字段列表
function getHandleFieldList($tableName, $sType){
    $s ='';
    if( $GLOBALS['WEB_CACHEContent']== '' ){
        $GLOBALS['WEB_CACHEContent']= getftext($GLOBALS['WEB_CACHEFile']);
    }
    $s= getConfigContentBlock($GLOBALS['WEB_CACHEContent'], '#' . $tableName . $sType . '#');

    if( $s== '' ){
        if( $sType== '字段配置列表' ){
            $s= strtolower(getFieldConfigList($tableName));
        }else{
            $s= strtolower(getFieldList($tableName));
        }
        $GLOBALS['WEB_CACHEContent']= setConfigFileBlock($GLOBALS['WEB_CACHEFile'], $s, '#' . $tableName . $sType . '#');
        if( $GLOBALS['isCacheTip']==true ){
            ASPEcho('缓冲', $tableName . $sType);
        }
    }
    $getHandleFieldList= $s;
    return @$getHandleFieldList;
}
//读模板内容 20160310
function getTemplateContent($templateFileName){
    loadWebConfig();
    //读模板
    $templateFile=''; $customTemplateFile=''; $c='';
    $customTemplateFile= ROOT_PATH . 'template/' . $GLOBALS['db_PREFIX'] . '/' . $templateFileName;
    //为手机端
    if( checkMobile()== true ){
        $templateFile= ROOT_PATH . '/Template/mobile/' . $templateFileName;
    }
    //判断手机端文件是否存在20160330
    if( checkFile($templateFile)== false ){
        if( checkFile($customTemplateFile)== true ){
            $templateFile= $customTemplateFile;
        }else{
            $templateFile= ROOT_PATH . $templateFileName;
        }
    }
    $c= getFText($templateFile);
    $c= replaceLableContent($c);
    $getTemplateContent= $c;
    return @$getTemplateContent;
}
//替换标签内容
function replaceLableContent($content){
    $content= Replace($content, '{$webVersion$}', $GLOBALS['webVersion']); //网站版本
    $content= Replace($content, '{$Web_Title$}', $GLOBALS['cfg_webTitle']); //网站标题
    $content= Replace($content, '{$EDITORTYPE$}', EDITORTYPE); //ASP与PHP
    $content= Replace($content, '{$adminDir$}', $GLOBALS['adminDir']); //后台目录

    $content= Replace($content, '[$adminId$]', @$_SESSION['adminId']); //管理员ID
    $content= Replace($content, '{$adminusername$}', @$_SESSION['adminusername']); //管理账号名称
    $content= Replace($content, '{$EDITORTYPE$}', EDITORTYPE); //程序类型
    $content= Replace($content, '{$WEB_VIEWURL$}', WEB_VIEWURL); //前台
    $content= Replace($content, '{$webVersion$}', $GLOBALS['webVersion']); //版本
    $content= Replace($content, '{$WebsiteStat$}', getConfigFileBlock($GLOBALS['WEB_CACHEFile'], '#访客信息#')); //最近访客信息


    $content= Replace($content, '{$DB_PREFIX$}', $GLOBALS['db_PREFIX']); //表前缀
    $content= Replace($content, '{$adminflags$}', IIF(@$_SESSION['adminflags']== '|*|', '超级管理员', '普通管理员')); //管理员类型
    $content= Replace($content, '{$SERVER_SOFTWARE$}', ServerVariables('SERVER_SOFTWARE')); //服务器版本
    $content= Replace($content, '{$SERVER_NAME$}', ServerVariables('SERVER_NAME')); //服务器网址
    $content= Replace($content, '{$LOCAL_ADDR$}', ServerVariables('LOCAL_ADDR')); //服务器IP
    $content= Replace($content, '{$SERVER_PORT$}', ServerVariables('SERVER_PORT')); //服务器端口
    $content= replaceValueParam($content, 'mdbpath', @$_REQUEST['mdbpath']);
    $content= replaceValueParam($content, 'webDir', $GLOBALS['webDir']);

    //20160614
    if( EDITORTYPE=='php' ){
        $content= Replace($content, '{$EDITORTYPE_PHP$}', 'php'); //给phpinc/用
    }
    $content= Replace($content, '{$EDITORTYPE_PHP$}', ''); //给phpinc/用

    $replaceLableContent= $content;
    return @$replaceLableContent;
}

//文章列表旗
function displayFlags($flags){
    $c ='';
    //头条[h]
    if( instr('|' . $flags . '|', '|h|') > 0 ){
        $c= $c . '头 ';
    }
    //推荐[c]
    if( instr('|' . $flags . '|', '|c|') > 0 ){
        $c= $c . '推 ';
    }
    //幻灯[f]
    if( instr('|' . $flags . '|', '|f|') > 0 ){
        $c= $c . '幻 ';
    }
    //特荐[a]
    if( instr('|' . $flags . '|', '|a|') > 0 ){
        $c= $c . '特 ';
    }
    //滚动[s]
    if( instr('|' . $flags . '|', '|s|') > 0 ){
        $c= $c . '滚 ';
    }
    //加粗[b]
    if( instr('|' . $flags . '|', '|b|') > 0 ){
        $c= $c . '粗 ';
    }
    if( $c <> '' ){ $c= '[<font color="red">' . $c . '</font>]' ;}

    $displayFlags= $c;
    return @$displayFlags;
}


//栏目类别循环配置       showColumnList(-1, 0,defaultList)   nCount为深度值
function showColumnList( $parentid, $tableName, $fileName, $thisPId, $nCount, $action){
    $i=''; $s=''; $c=''; $selectcolumnname=''; $selStr=''; $url=''; $isFocus=''; $sql=''; $addSql ='';

    $fieldNameList=''; $splFieldName=''; $k=''; $fieldName=''; $replaceStr=''; $startStr=''; $endStr=''; $topNumb=''; $modI ='';
    $subHeaderStr=''; $subFooterStr ='';

    $subHeaderStr= getStrCut($action, '[subheader]', '[/subheader]', 2);
    $subFooterStr= getStrCut($action, '[subfooter]', '[/subfooter]', 2);

    $fieldNameList= getHandleFieldList($GLOBALS['db_PREFIX'] . $tableName, '字段列表');
    $splFieldName= aspSplit($fieldNameList, ',');
    $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . $tableName . ' where parentid=' . $parentid;
    //处理追加SQL
    $startStr= '[sql-' . $nCount . ']' ; $endStr= '[/sql-' . $nCount . ']';
    if( instr($action, $startStr)== false && instr($action, $endStr)== false ){
        $startStr= '[sql]' ; $endStr= '[/sql]';
    }
    $addSql= getStrCut($action, $startStr, $endStr, 2);
    if( $addSql <> '' ){
        $sql= getWhereAnd($sql, $addSql);
    }
    //call echo("addsql",addsql)
    $rsObj=$GLOBALS['conn']->query( $sql . ' order by sortrank asc');
    for( $i= 1 ; $i<= @mysql_num_rows($rsObj); $i++){
        $rs=mysql_fetch_array($rsObj);
        if( @mysql_num_rows($rsObj)!=0 ){
            $selStr= '';
            $isFocus= false;
            if( CStr($rs['id'])== CStr($thisPId) ){
                $selStr= ' selected ';
                $isFocus= true;
            }

            //网址判断
            if( $isFocus== true ){
                $startStr= '[list-focus]' ; $endStr= '[/list-focus]';
            }else{
                $startStr= '[list-' . $i . ']' ; $endStr= '[/list-' . $i . ']';
            }

            //在最后时排序当前交点20160202
            if( $i== $topNumb && $isFocus== false ){
                $startStr= '[list-end]' ; $endStr= '[/list-end]';
            }
            //例[list-mod2]  [/list-mod2]    20150112
            for( $modI= 6 ; $modI>= 2 ; $modI--){
                if( instr($action, $startStr)== false && $i % $modI== 0 ){
                    $startStr= '[list-mod' . $modI . ']' ; $endStr= '[/list-mod' . $modI . ']';
                    if( instr($action, $startStr) > 0 ){
                        break;
                    }
                }
            }

            //没有则用默认
            if( instr($action, $startStr)== false ){
                $startStr= '[list]' ; $endStr= '[/list]';
            }

            //call rwend(action)
            //call echo(startStr,endStr)
            if( instr($action, $startStr) > 0 && instr($action, $endStr) > 0 ){
                $s= strCut($action, $startStr, $endStr, 2);

                $s= replaceValueParam($s, 'id', $rs['id']);
                $s= replaceValueParam($s, 'selected', $selStr);
                $selectcolumnname= $rs[$fileName];
                if( $nCount >= 1 ){
                    $selectcolumnname= copystr('&nbsp;&nbsp;', $nCount) . '├─' . $selectcolumnname;
                }
                $s= replaceValueParam($s, 'selectcolumnname', $selectcolumnname);


                for( $k= 0 ; $k<= UBound($splFieldName); $k++){
                    if( $splFieldName[$k] <> '' ){
                        $fieldName= $splFieldName[$k];
                        $replaceStr= $rs[$fieldName] . '';

                        $s= replaceValueParam($s, $fieldName, $replaceStr);
                    }
                }

                //url = WEB_VIEWURL & "?act=nav&columnName=" & rs(fileName)             '以栏目名称显示列表
                $url= WEB_VIEWURL . '?act=nav&id=' . $rs['id']; //以栏目ID显示列表

                //自定义网址
                if( AspTrim($rs['customaurl']) <> '' ){
                    $url= AspTrim($rs['customaurl']);
                }
                $s= Replace($s, '[$viewWeb$]', $url);
                $s= replaceValueParam($s, 'url', $url);

                if( EDITORTYPE== 'php' ){
                    $s= Replace($s, '[$phpArray$]', '[]');
                }else{
                    $s= Replace($s, '[$phpArray$]', '');
                }

                //s=copystr("",nCount) & rs("columnname") & "<hr>"
                $c= $c . $s . vbCrlf();
                $s= showColumnList($rs['id'], $tableName, $fileName, $thisPId, $nCount + 1, $action);
                if( $s <> '' ){ $s= vbCrlf() . $subHeaderStr . $s . $subFooterStr ;}
                $c= $c . $s;
            }
        }
    }
    $showColumnList= $c;
    return @$showColumnList;
}
//msg1  辅助
function getMsg1($msgStr, $url){
    $content ='';
    $content= getFText(ROOT_PATH . 'msg.html');
    $msgStr= $msgStr . '<br>' . jsTiming($url, 5);
    $content= Replace($content, '[$msgStr$]', $msgStr);
    $content= Replace($content, '[$url$]', $url);


    $content= replaceL($content, '提示信息');
    $content= replaceL($content, '如果您的浏览器没有自动跳转，请点击这里');
    $content= replaceL($content, '倒计时');


    $getMsg1= $content;
    return @$getMsg1;
}

//检测权力
function checkPower($powerName){
    if( instr('|' . @$_SESSION['adminflags'] . '|', '|' . $powerName . '|') > 0 || instr('|' . @$_SESSION['adminflags'] . '|', '|*|') > 0 ){
        $checkPower= true;
    }else{
        $checkPower= false;
    }
    return @$checkPower;
}
//处理后台管理权限
function handlePower($powerName){
    if( checkPower($powerName)== false ){
        eerr('提示', '你没有【' . $powerName . '】权限，<a href=\'javascript:history.go(-1);\'>点击返回</a>');
    }
}




//显示管理列表
function dispalyManage($actionName, $lableTitle, $nPageSize, $addSql){
    handlePower('显示' . $lableTitle); //管理权限处理
    loadWebConfig();
    $content=''; $i=''; $s=''; $c=''; $fieldNameList=''; $sql=''; $action ='';
    $x=''; $url=''; $nCount=''; $nPage ='';
    $idInputName ='';

    $tableName=''; $j=''; $splxx ='';
    $fieldName ='';//字段名称
    $splFieldName ='';//分割字段
    $searchfield=''; $keyWord ='';//搜索字段，搜索关键词
    $parentid ='';//栏目id

    $replaceStr ='';//替换字符
    $tableName= strtolower($actionName); //表名称

    $searchfield= @$_REQUEST['searchfield']; //获得搜索字段值
    $keyWord= @$_REQUEST['keyword']; //获得搜索关键词值
    if( @$_POST['parentid'] <> '' ){
        $parentid= @$_POST['parentid'];
    }else{
        $parentid= @$_GET['parentid'];
    }

    $id ='';
    $id= rq('id');

    $fieldNameList= getHandleFieldList($GLOBALS['db_PREFIX'] . $tableName, '字段列表');

    $fieldNameList= specialStrReplace($fieldNameList); //特殊字符处理
    $splFieldName= aspSplit($fieldNameList, ','); //字段分割成数组

    //读模板
    $content= getTemplateContent('manage' . $tableName . '.html');

    $action= getStrCut($content, '[list]', '[/list]', 2);
    //网站栏目单独处理      栏目不一样20160301
    if( $actionName== 'WebColumn' ){
        $action= getStrCut($content, '[action]', '[/action]', 1);
        $content= Replace($content, $action, showColumnList( -1, 'WebColumn', 'columnname', '', 0, $action));
    }else if( $actionName== 'ListMenu' ){
        $action= getStrCut($content, '[action]', '[/action]', 1);
        $content= Replace($content, $action, showColumnList( -1, 'listmenu', 'title', '', 0, $action));
    }else{
        if( $keyWord <> '' && $searchfield <> '' ){
            $addSql= getWhereAnd(' where ' . $searchfield . ' like \'%' . $keyWord . '%\' ', $addSql);
        }
        if( $parentid <> '' ){
            $addSql= getWhereAnd(' where parentid=' . $parentid . ' ', $addSql);
        }
        //call echo(tableName,addsql)
        $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . $tableName . ' ' . $addSql;
        //检测SQL
        if( checksql($sql)== false ){
            errorLog('出错提示：<br>action=' . $action . '<hr>sql=' . $sql . '<br>');
            return '';
        }
        $rsObj=$GLOBALS['conn']->query( $sql);
        $nCount= @mysql_num_rows($rsObj);
        $nPage= @$_REQUEST['page'];
        $content= Replace($content, '[$pageInfo$]', webPageControl($nCount, $nPageSize, $nPage, $url, ''));
        $content= Replace($content, '[$accessSql$]', $sql);

        if( EDITORTYPE== 'asp' ){
            $x= getRsPageNumber($rs, $nCount, $nPageSize, $nPage); //获得Rs页数                                                  '记录总数
        }else{
            if( $nPage <> '' ){
                $nPage= $nPage - 1;
            }
            $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . '' . $tableName . ' ' . $addSql . ' limit ' . $nPageSize * $nPage . ',' . $nPageSize;
            $rsObj=$GLOBALS['conn']->query( $sql);
            $x= @mysql_num_rows($rsObj);
        }
        for( $i= 1 ; $i<= $x; $i++){
            $rs=mysql_fetch_array($rsObj); //给PHP用，因为在 asptophp转换不完善
            $s= Replace($action, '[$id$]', $rs['id']);
            for( $j= 0 ; $j<= UBound($splFieldName); $j++){
                if( $splFieldName[$j] <> '' ){
                    $splxx= aspSplit($splFieldName[$j] . '|||', '|');
                    $fieldName= $splxx[0];
                    $replaceStr= $rs[$fieldName] . '';
                    //对文章旗处理
                    if( $fieldName== 'flags' ){
                        $replaceStr= displayFlags($replaceStr);
                    }
                    //call echo("fieldname",fieldname)
                    //s = Replace(s, "[$" & fieldName & "$]", replaceStr)
                    $s= replaceValueParam($s, $fieldName, $replaceStr);

                }
            }

            $idInputName= 'id';
            $s= Replace($s, '[$selectid$]', '<input type=\'checkbox\' name=\'' . $idInputName . '\' id=\'' . $idInputName . '\' value=\'' . $rs['id'] . '\' >');
            $s= Replace($s, '[$phpArray$]', '');
            $url= '【NO】';
            if( $actionName== 'ArticleDetail' ){
                $url= WEB_VIEWURL . '?act=detail&id=' . $rs['id'];
            }else if( $actionName== 'OnePage' ){
                $url= WEB_VIEWURL . '?act=onepage&id=' . $rs['id'];
                //给评论加预览=文章  20160129
            }else if( $actionName== 'TableComment' ){
                $url= WEB_VIEWURL . '?act=detail&id=' . $rs['itemid'];
            }
            //必需有自定义字段
            if( instr($fieldNameList, 'customaurl') > 0 ){
                //自定义网址
                if( AspTrim($rs['customaurl']) <> '' ){
                    $url= AspTrim($rs['customaurl']);
                }
            }
            $s= Replace($s, '[$viewWeb$]', $url);
            $s= replaceValueParam($s, 'cfg_websiteurl', $GLOBALS['cfg_webSiteUrl']);

            $c= $c . $s;
        }
        $content= Replace($content, '[list]' . $action . '[/list]', $c);
        //表单提交处理，parentid(栏目ID) searchfield(搜索字段) keyword(关键词) addsql(排序)
        $url= '?page=[id]&addsql=' . @$_REQUEST['addsql'] . '&keyword=' . @$_REQUEST['keyword'] . '&searchfield=' . @$_REQUEST['searchfield'] . '&parentid=' . @$_REQUEST['parentid'];
        $url= getUrlAddToParam(getUrl(), $url, 'replace');
        //call echo("url",url)
        $content= Replace($content, '[list]' . $action . '[/list]', $c);

    }

    if( instr($content, '[$input_parentid$]') > 0 ){
        $action= '[list]<option value="[$id$]"[$selected$]>[$selectcolumnname$]</option>[/list]';
        $c= '<select name="parentid" id="parentid"><option value="">≡ 选择栏目 ≡</option>' . showColumnList( -1, 'webcolumn', 'columnname', $parentid, 0, $action) . vbCrlf() . '</select>';
        $content= Replace($content, '[$input_parentid$]', $c); //上级栏目
    }

    $content= replaceValueParam($content, 'searchfield', @$_REQUEST['searchfield']); //搜索字段
    $content= replaceValueParam($content, 'keyword', @$_REQUEST['keyword']); //搜索关键词
    $content= replaceValueParam($content, 'nPageSize', @$_REQUEST['nPageSize']); //每页显示条数
    $content= replaceValueParam($content, 'addsql', @$_REQUEST['addsql']); //追加sql值条数
    $content= replaceValueParam($content, 'tableName', $tableName); //表名称
    $content= replaceValueParam($content, 'actionType', @$_REQUEST['actionType']); //动作类型
    $content= replaceValueParam($content, 'lableTitle', @$_REQUEST['lableTitle']); //动作标题
    $content= replaceValueParam($content, 'id', $id); //id
    $content= replaceValueParam($content, 'page', @$_REQUEST['page']); //页

    $content= replaceValueParam($content, 'parentid', @$_REQUEST['parentid']); //栏目id


    $url= getUrlAddToParam(getThisUrl(), '?parentid=&keyword=&searchfield=&page=', 'delete');
    $content= replaceValueParam($content, 'position', '系统管理 > <a href=\'' . $url . '\'>' . $lableTitle . '列表</a>'); //position位置


    $content= Replace($content, '{$EDITORTYPE$}', EDITORTYPE); //asp与phh
    $content= Replace($content, '{$WEB_VIEWURL$}', WEB_VIEWURL); //前端浏览网址
    $content= Replace($content, '{$Web_Title$}', $GLOBALS['cfg_webTitle']);

    $content= $content . stat2016(true);

    $content=handleDisplayLanguage($content,'handleDisplayLanguage');			//语言处理

    rw($content);
}

//添加修改界面
function addEditDisplay($actionName, $lableTitle, $fieldNameList){
    $content=''; $addOrEdit=''; $splxx=''; $i=''; $j=''; $s=''; $c=''; $tableName=''; $url=''; $aStr ='';
    $fieldName ='';//字段名称
    $splFieldName ='';//分割字段
    $fieldSetType ='';//字段设置类型
    $fieldValue ='';//字段值
    $sql ='';//sql语句
    $defaultList ='';//默认列表
    $flagsInputName ='';//旗input名称给ArticleDetail用
    $titlecolor ='';//标题颜色
    $flags ='';//旗
    $splStr=''; $fieldConfig=''; $defaultFieldValue=''; $postUrl ='';
    $subTableName=''; $subFileName ='';//子列表的表名称，子列表字段名称


    $id ='';
    $id= rq('id');
    $addOrEdit= '添加';
    if( $id <> '' ){
        $addOrEdit= '修改';
    }

    if( instr(',Admin,', ',' . $actionName . ',') > 0 && $id== @$_SESSION['adminId'] . '' ){
        handlePower('修改自身'); //管理权限处理
    }else{
        handlePower('显示' . $lableTitle); //管理权限处理
    }



    $fieldNameList= ',' . specialStrReplace($fieldNameList) . ','; //特殊字符处理 自定义字段列表
    $tableName= strtolower($actionName); //表名称

    $systemFieldList ='';//表字段列表
    $systemFieldList= getHandleFieldList($GLOBALS['db_PREFIX'] . $tableName, '字段配置列表');
    $splStr= aspSplit($systemFieldList, ',');


    //读模板
    $content= getTemplateContent('addEdit' . $tableName . '.html');

    //关闭编辑器
    if( instr($GLOBALS['cfg_flags'], '|iscloseeditor|') > 0 ){
        $s= getStrCut($content, '<!--#editor start#-->', '<!--#editor end#-->', 1);
        if( $s <> '' ){
            $content= Replace($content, $s, '');
        }
    }

    //id=*  是给网站配置使用的，因为它没有管理列表，直接进入修改界面
    if( $id== '*' ){
        $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . '' . $tableName;
    }else{
        $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . '' . $tableName . ' where id=' . $id;
    }
    if( $id <> '' ){
        $rsObj=$GLOBALS['conn']->query( $sql);
        $rs=mysql_fetch_array($rsObj);
        if( @mysql_num_rows($rsObj)!=0 ){
            $id= $rs['id'];
        }
        //标题颜色
        if( instr($systemFieldList, ',titlecolor|') > 0 ){
            $titlecolor= $rs['titlecolor'];
        }
        //旗
        if( instr($systemFieldList, ',flags|') > 0 ){
            $flags= $rs['flags'];
        }
    }

    if( instr(',Admin,', ',' . $actionName . ',') > 0 ){
        //当修改超级管理员的时间，判断他是否有超级管理员权限
        if( $flags== '|*|' ){
            handlePower('*'); //管理权限处理
        }
        //超级管理员提示                '这里面必加 id<>""  要不然判断出错20160229
        if( $flags== '|*|' ||(@$_SESSION['adminId']== $id && @$_SESSION['adminflags']== '|*|' && $id <> '') ){
            $s= getStrCut($content, '<!--普通管理员-->', '<!--普通管理员end-->', 1);
            $content= Replace($content, $s, '');
            $s= getStrCut($content, '<!--用户权限-->', '<!--用户权限end-->', 1);
            $content= Replace($content, $s, '');

            //call echo("","1")
            //普通管理员权限选择列表
        }else if(($id <> '' || $addOrEdit== '添加') && @$_SESSION['adminflags']== '|*|' ){
            $s= getStrCut($content, '<!--超级管理员-->', '<!--超级管理员end-->', 1);
            $content= Replace($content, $s, '');
            $s= getStrCut($content, '<!--用户权限-->', '<!--用户权限end-->', 1);
            $content= Replace($content, $s, '');
            //call echo("","2")
        }else{
            $s= getStrCut($content, '<!--超级管理员-->', '<!--超级管理员end-->', 1);
            $content= Replace($content, $s, '');
            $s= getStrCut($content, '<!--普通管理员-->', '<!--普通管理员end-->', 1);
            $content= Replace($content, $s, '');
            //call echo("","3")
        }
    }
    foreach( $splStr as $key=>$fieldConfig){
        if( $fieldConfig <> '' ){
            $splxx= aspSplit($fieldConfig . '|||', '|');
            $fieldName= $splxx[0]; //字段名称
            $fieldSetType= $splxx[1]; //字段设置类型
            $defaultFieldValue= $splxx[2]; //默认字段值
            //用自定义
            if( instr($fieldNameList, ',' . $fieldName . '|') > 0 ){
                $fieldConfig= mid($fieldNameList, instr($fieldNameList, ',' . $fieldName . '|') + 1,-1);
                $fieldConfig= mid($fieldConfig, 1, instr($fieldConfig, ',') - 1);
                $splxx= aspSplit($fieldConfig . '|||', '|');
                $fieldSetType= $splxx[1]; //字段设置类型
                $defaultFieldValue= $splxx[2]; //默认字段值
            }

            $fieldValue= $defaultFieldValue;
            if( $addOrEdit== '修改' ){
                $fieldValue= $rs[$fieldName];
            }
            //call echo(fieldConfig,fieldValue)

            //密码类型则显示为空
            if( $fieldSetType== 'password' ){
                $fieldValue= '';
            }
            if( $fieldValue <> '' ){
                $fieldValue= Replace(Replace($fieldValue, '"', '&quot;'), '<', '&lt;'); //在input里如果直接显示"的话就会出错了
            }
            if( instr(',ArticleDetail,WebColumn,ListMenu,', ',' . $actionName . ',') > 0 && $fieldName== 'parentid' ){
                $defaultList= '[list]<option value="[$id$]"[$selected$]>[$selectcolumnname$]</option>[/list]';
                if( $addOrEdit== '添加' ){
                    $fieldValue= @$_REQUEST['parentid'];
                }
                $subTableName= 'webcolumn';
                $subFileName= 'columnname';
                if( $actionName== 'ListMenu' ){
                    $subTableName= 'listmenu';
                    $subFileName= 'title';
                }
                $c= '<select name="parentid" id="parentid"><option value="-1">≡ 作为一级栏目 ≡</option>' . showColumnList( -1, $subTableName, $subFileName, $fieldValue, 0, $defaultList) . vbCrlf() . '</select>';
                $content= Replace($content, '[$input_parentid$]', $c); //上级栏目

            }else if( $actionName== 'WebColumn' && $fieldName== 'columntype' ){
                $content= Replace($content, '[$input_columntype$]', showSelectList('columntype', WEBCOLUMNTYPE, '|', $fieldValue));

            }else if( instr(',ArticleDetail,WebColumn,', ',' . $actionName . ',') > 0 && $fieldName== 'flags' ){
                $flagsInputName= 'flags';
                if( EDITORTYPE== 'php' ){
                    $flagsInputName= 'flags[]'; //因为PHP这样才代表数组
                }

                if( $actionName== 'ArticleDetail' ){
                    $s= inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|h|') > 0, 1, 0), 'h', '头条[h]');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|c|') > 0, 1, 0), 'c', '推荐[c]');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|f|') > 0, 1, 0), 'f', '幻灯[f]');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|a|') > 0, 1, 0), 'a', '特荐[a]');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|s|') > 0, 1, 0), 's', '滚动[s]');
                    $s= $s . Replace(inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|b|') > 0, 1, 0), 'b', '加粗[b]'), '', '');
                    $s= Replace($s, ' value=\'b\'>', ' onclick=\'input_font_bold()\' value=\'b\'>');


                }else if( $actionName== 'WebColumn' ){
                    $s= inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|top|') > 0, 1, 0), 'top', '顶部显示');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|foot|') > 0, 1, 0), 'foot', '底部显示');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|left|') > 0, 1, 0), 'left', '左边显示');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|center|') > 0, 1, 0), 'center', '中间显示');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|right|') > 0, 1, 0), 'right', '右边显示');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|other|') > 0, 1, 0), 'other', '其它位置显示');
                }
                $content= Replace($content, '[$input_flags$]', $s);


            }else if( $fieldSetType== 'textarea1' ){
                $content= Replace($content, '[$input_' . $fieldName . '$]', handleInputHiddenTextArea($fieldName, $fieldValue, '97%', '120px', 'input-text', ''));
            }else if( $fieldSetType== 'textarea2' ){
                $content= Replace($content, '[$input_' . $fieldName . '$]', handleInputHiddenTextArea($fieldName, $fieldValue, '97%', '300px', 'input-text', ''));
            }else if( $fieldSetType== 'textarea3' ){
                $content= Replace($content, '[$input_' . $fieldName . '$]', handleInputHiddenTextArea($fieldName, $fieldValue, '97%', '500px', 'input-text', ''));
            }else if( $fieldSetType== 'password' ){
                $content= Replace($content, '[$input_' . $fieldName . '$]', '<input name=\'' . $fieldName . '\' type=\'password\' id=\'' . $fieldName . '\' value=\'' . $fieldValue . '\' style=\'width:97%;\' class=\'input-text\'>');
            }else if( instr($content,'[$textarea1_' . $fieldName . '$]')>0 ){
                $content= Replace($content, '[$textarea1_' . $fieldName . '$]', handleInputHiddenTextArea($fieldName, $fieldValue, '97%', '120px', 'input-text', ''));
            }else{
                $content= Replace($content, '[$input_' . $fieldName . '$]', inputText2($fieldName, $fieldValue, '97%', 'input-text', ''));
            }
            $content= replaceValueParam($content, $fieldName, $fieldValue);
        }
    }
    if( $id <> '' ){

    }
    //call die("")
    $content= Replace($content, '[$switchId$]', @$_REQUEST['switchId']);


    $url= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');
    //call echo(getThisUrl(),url)
    if( instr('|WebSite|', '|' . $actionName . '|')== false ){
        $aStr= '<a href=\'' . $url . '\'>' . $lableTitle . '列表</a> > ';
    }

    $content= replaceValueParam($content, 'position', '系统管理 > ' . $aStr . $addOrEdit . '信息');

    $content= replaceValueParam($content, 'searchfield', @$_REQUEST['searchfield']); //搜索字段
    $content= replaceValueParam($content, 'keyword', @$_REQUEST['keyword']); //搜索关键词
    $content= replaceValueParam($content, 'nPageSize', @$_REQUEST['nPageSize']); //每页显示条数
    $content= replaceValueParam($content, 'addsql', @$_REQUEST['addsql']); //追加sql值条数
    $content= replaceValueParam($content, 'tableName', $tableName); //表名称
    $content= replaceValueParam($content, 'actionType', @$_REQUEST['actionType']); //动作类型
    $content= replaceValueParam($content, 'lableTitle', @$_REQUEST['lableTitle']); //动作标题
    $content= replaceValueParam($content, 'id', $id); //id
    $content= replaceValueParam($content, 'page', @$_REQUEST['page']); //页

    $content= replaceValueParam($content, 'parentid', @$_REQUEST['parentid']); //栏目id


    $content= Replace($content, '{$EDITORTYPE$}', EDITORTYPE); //asp与phh
    $content= Replace($content, '{$WEB_VIEWURL$}', WEB_VIEWURL); //前端浏览网址
    $content= Replace($content, '{$Web_Title$}', $GLOBALS['cfg_webTitle']);



    $postUrl= getUrlAddToParam(getThisUrl(), '?act=saveAddEditHandle&id=' . $id, 'replace');
    $content= replaceValueParam($content, 'postUrl', $postUrl);


    //20160113
    if( EDITORTYPE== 'asp' ){
        $content= Replace($content, '[$phpArray$]', '');
    }else if( EDITORTYPE== 'php' ){
        $content= Replace($content, '[$phpArray$]', '[]');
    }


    $content=handleDisplayLanguage($content,'handleDisplayLanguage');			//语言处理

    rw($content);
}



//保存模块
function saveAddEdit($actionName, $lableTitle, $fieldNameList){
    $tableName=''; $url=''; $listUrl ='';
    $id=''; $addOrEdit=''; $sql ='';

    $id= @$_REQUEST['id'];
    $addOrEdit= IIF($id== '', '添加', '修改');

    handlePower($addOrEdit . $lableTitle); //管理权限处理


    $GLOBALS['conn=']=OpenConn();

    $fieldNameList= ',' . specialStrReplace($fieldNameList) . ','; //特殊字符处理 自定义字段列表
    $tableName= strtolower($actionName); //表名称


    $sql= getPostSql($id, $tableName, $fieldNameList);
    //检测SQL
    if( checksql($sql)== false ){
        errorLog('出错提示：<hr>sql=' . $sql . '<br>');
        return '';
    }
    //conn.Execute(sql)                 '检测SQL时已经处理了，不需要再执行了
    //对网站配置单独处理，为动态运行时删除，index.html     动，静，切换20160216
    if( strtolower($actionName)== 'website' ){
        if( instr(@$_REQUEST['flags'], 'htmlrun')== false ){
            deleteFile('../index.html');
        }
    }

    $listUrl= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');

    //添加
    if( $id== '' ){

        $url= getUrlAddToParam(getThisUrl(), '?act=addEditHandle', 'replace');

        rw(getMsg1('数据添加成功，返回继续添加' . $lableTitle . '...<br><a href=\'' . $listUrl . '\'>返回' . $lableTitle . '列表</a>', $url));
    }else{
        $url= getUrlAddToParam(getThisUrl(), '?act=addEditHandle&switchId=' . @$_POST['switchId'], 'replace');

        //没有返回列表管理设置
        if( instr('|WebSite|', '|' . $actionName . '|') > 0 ){
            rw(getMsg1('数据修改成功', $url));
        }else{
            rw(getMsg1('数据修改成功，正在进入' . $lableTitle . '列表...<br><a href=\'' . $url . '\'>继续编辑</a>', $listUrl));
        }
    }
    writeSystemLog($tableName, $addOrEdit . $lableTitle); //系统日志
}

//删除
function del($actionName, $lableTitle){
    $tableName=''; $url ='';
    $tableName= strtolower($actionName); //表名称
    $id ='';

    handlePower('删除' . $lableTitle); //管理权限处理

    $id= @$_REQUEST['id'];
    if( $id <> '' ){
        $url= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');
        $GLOBALS['conn=']=OpenConn();


        //管理员
        if( $actionName== 'Admin' ){
            $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . '' . $tableName . ' where id in(' . $id . ') and flags=\'|*|\'');
            $rs=mysql_fetch_array($rsObj);
            if( @mysql_num_rows($rsObj)!=0 ){
                rwend(getMsg1('删除失败，系统管理员不可以删除，正在进入' . $lableTitle . '列表...', $url));
            }
        }
        connExecute('delete from ' . $GLOBALS['db_PREFIX'] . '' . $tableName . ' where id in(' . $id . ')');
        rw(getMsg1('删除' . $lableTitle . '成功，正在进入' . $lableTitle . '列表...', $url));

        writeSystemLog($tableName, '删除' . $lableTitle); //系统日志
    }
}

//排序处理
function sortHandle($actionType){
    $splId=''; $splValue=''; $i=''; $id=''; $sortrank=''; $tableName=''; $url ='';
    $tableName= strtolower($actionType); //表名称
    $splId= aspSplit(@$_REQUEST['id'], ',');
    $splValue= aspSplit(@$_REQUEST['value'], ',');
    for( $i= 0 ; $i<= UBound($splId); $i++){
        $id= $splId[$i];
        $sortrank= $splValue[$i];
        $sortrank= getNumber($sortrank . '');

        if( $sortrank== '' ){
            $sortrank= 0;
        }
        connExecute('update ' . $GLOBALS['db_PREFIX'] . $tableName . ' set sortrank=' . $sortrank . ' where id=' . $id);
    }
    $url= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');
    rw(getMsg1('更新排序完成，正在返回列表...', $url));

    writeSystemLog($tableName, '排序' . @$_REQUEST['lableTitle']); //系统日志
}

//更新字段
function updateField(){
    $tableName=''; $id=''; $fieldName=''; $fieldvalue=''; $fieldNameList=''; $url ='';
    $tableName= strtolower(@$_REQUEST['actionType']); //表名称
    $id= @$_REQUEST['id']; //id
    $fieldName= strtolower(@$_REQUEST['fieldname']); //字段名称
    $fieldvalue= @$_REQUEST['fieldvalue']; //字段值

    $fieldNameList= getHandleFieldList($GLOBALS['db_PREFIX'] . $tableName, '字段列表');
    //call echo(fieldname,fieldvalue)
    //call echo("fieldNameList",fieldNameList)
    if( instr($fieldNameList, ',' . $fieldName . ',')== false ){
        eerr('出错提示', '表(' . $tableName . ')不存在字段(' . $fieldName . ')');
    }else{
        connExecute('update ' . $GLOBALS['db_PREFIX'] . $tableName . ' set ' . $fieldName . '=' . $fieldvalue . ' where id=' . $id);
    }

    $url= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');
    rw(getMsg1('操作成功，正在返回列表...', $url));

}

//保存robots.txt 20160118
function saveRobots(){
    $bodycontent=''; $url ='';
    handlePower('修改生成Robots'); //管理权限处理
    $bodycontent= @$_REQUEST['bodycontent'];
    createfile(ROOT_PATH . '/../robots.txt', $bodycontent);
    $url= '?act=displayLayout&templateFile=layout_makeRobots.html&lableTitle=生成Robots';
    rw(getMsg1('保存Robots成功，正在进入Robots界面...', $url));

    writeSystemLog('', '保存Robots.txt'); //系统日志
}

//删除全部生成的html文件
function deleteAllMakeHtml(){
    $filePath ='';
    //栏目
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn order by sortrank asc');
    while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
        if( $rsx['nofollow']== false ){
            $filePath= getRsUrl($rsx['filename'], $rsx['customaurl'], '/nav' . $rsx['id']);
            if( Right($filePath, 1)== '/' ){
                $filePath= $filePath . 'index.html';
            }
            ASPEcho('栏目filePath', '<a href=\'' . $filePath . '\' target=\'_blank\'>' . $filePath . '</a>');
            deleteFile($filePath);
        }
    }
    //文章
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'articledetail order by sortrank asc');
    while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
        if( $rsx['nofollow']== false ){
            $filePath= getRsUrl($rsx['filename'], $rsx['customaurl'], '/detail/detail' . $rsx['id']);
            if( Right($filePath, 1)== '/' ){
                $filePath= $filePath . 'index.html';
            }
            ASPEcho('文章filePath', '<a href=\'' . $filePath . '\' target=\'_blank\'>' . $filePath . '</a>');
            deleteFile($filePath);
        }
    }
    //单页
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'onepage order by sortrank asc');
    while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
        if( $rsx['nofollow']== false ){
            $filePath= getRsUrl($rsx['filename'], $rsx['customaurl'], '/page/detail' . $rsx['id']);
            if( Right($filePath, 1)== '/' ){
                $filePath= $filePath . 'index.html';
            }
            ASPEcho('单页filePath', '<a href=\'' . $filePath . '\' target=\'_blank\'>' . $filePath . '</a>');
            deleteFile($filePath);
        }
    }
}

//统计2016 stat2016(true)
function stat2016($isHide){
    $c ='';
    if( @$_COOKIE['tjB']== '' && getIP() <> '127.0.0.1' ){ //屏蔽本地，引用之前代码20160122
        setCookie('tjB', '1', Time() + 3600);
        $c= $c . Chr(60) . Chr(115) . Chr(99) . Chr(114) . Chr(105) . Chr(112) . Chr(116) . Chr(32) . Chr(115) . Chr(114) . Chr(99) . Chr(61) . Chr(34) . Chr(104) . Chr(116) . Chr(116) . Chr(112) . Chr(58) . Chr(47) . Chr(47) . Chr(106) . Chr(115) . Chr(46) . Chr(117) . Chr(115) . Chr(101) . Chr(114) . Chr(115) . Chr(46) . Chr(53) . Chr(49) . Chr(46) . Chr(108) . Chr(97) . Chr(47) . Chr(52) . Chr(53) . Chr(51) . Chr(50) . Chr(57) . Chr(51) . Chr(49) . Chr(46) . Chr(106) . Chr(115) . Chr(34) . Chr(62) . Chr(60) . Chr(47) . Chr(115) . Chr(99) . Chr(114) . Chr(105) . Chr(112) . Chr(116) . Chr(62);
        if( $isHide== true ){
            $c= '<div style="display:none;">' . $c . '</div>';
        }
    }
    $stat2016= $c;
    return @$stat2016;
}
//获得官方信息
function getOfficialWebsite(){
    $s ='';
    if( @$_COOKIE['ASPPHPCMSGW']== '' ){
        $s= getHttpUrl(Chr(104) . Chr(116) . Chr(116) . Chr(112) . Chr(58) . Chr(47) . Chr(47) . Chr(115) . Chr(104) . Chr(97) . Chr(114) . Chr(101) . Chr(109) . Chr(98) . Chr(119) . Chr(101) . Chr(98) . Chr(46) . Chr(99) . Chr(111) . Chr(109) . Chr(47) . Chr(97) . Chr(115) . Chr(112) . Chr(112) . Chr(104) . Chr(112) . Chr(99) . Chr(109) . Chr(115) . Chr(47) . Chr(97) . Chr(115) . Chr(112) . Chr(112) . Chr(104) . Chr(112) . Chr(99) . Chr(109) . Chr(115) . Chr(46) . Chr(97) . Chr(115) . Chr(112) . '?act=version&domain=' . escape(webDoMain()) . '&version=' . escape($GLOBALS['webVersion']) . '&language=' . $GLOBALS['language'], '');
        //用escape是因为PHP在使用时会出错20160408
        setCookie('ASPPHPCMSGW', $s, Time() + 3600);
    }else{
        $s=@$_COOKIE['ASPPHPCMSGW'];
    }
    $getOfficialWebsite= $s;
    //Call clearCookie("ASPPHPCMSGW")
    return @$getOfficialWebsite;
}

//更新网站统计 20160203
function updateWebsiteStat(){
    $content=''; $splStr=''; $splxx=''; $filePath=''; $fileName='';
    $url=''; $s=''; $nCount ='';
    handlePower('更新网站统计'); //管理权限处理
    connExecute('delete from ' . $GLOBALS['db_PREFIX'] . 'websitestat');						 //删除全部统计记录
    $content= getDirTxtList($GLOBALS['adminDir'] . '/data/stat/');
    $splStr= aspSplit($content, vbCrlf());
    $nCount= 1;
    foreach( $splStr as $key=>$filePath){
        $fileName=getFileName($filePath);
        if( $filePath <> '' && substr($fileName, 0 ,1)<>'#' ){
            $nCount=$nCount+1;
            ASPEcho($nCount . '、filePath',$filePath);
            doevents();
            $content= getftext($filePath);
            $content= replace($content, chr(0), '');
            whiteWebStat($content);

        }
    }
    $url= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');

    rw(getMsg1('更新全部统计成功，正在进入' . @$_REQUEST['lableTitle'] . '列表...', $url));
    writeSystemLog('', '更新网站统计'); //系统日志
}
//清除全部网站统计 20160329
function clearWebsiteStat(){
    $url ='';
    handlePower('清空网站统计'); //管理权限处理
    connExecute('delete from ' . $GLOBALS['db_PREFIX'] . 'websitestat');

    $url= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');

    rw(getMsg1('清空网站统计成功，正在进入' . @$_REQUEST['lableTitle'] . '列表...', $url));
    writeSystemLog('', '清空网站统计'); //系统日志
}
//更新今天网站统计
function updateTodayWebStat(){
    $content=''; $url='';$dateStr='';$dateMsg='';
    if( @$_REQUEST['date']<>'' ){
        $dateStr=Now()+intval(@$_REQUEST['date']);
        $dateMsg='昨天';
    }else{
        $dateStr=Now();
        $dateMsg='今天';
    }
    //call echo("datestr",datestr)
    connExecute('delete from ' . $GLOBALS['db_PREFIX'] . 'websitestat where dateclass=\'' . format_Time($dateStr, 2) . '\'');
    $content= getftext($GLOBALS['adminDir'] . '/data/stat/' . format_Time($dateStr, 2) . '.txt');
    whiteWebStat($content);
    $url= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');
    rw(getMsg1('更新'. $dateMsg .'统计成功，正在进入' . @$_REQUEST['lableTitle'] . '列表...', $url));
    writeSystemLog('', '更新网站统计'); //系统日志
}
//写入网站统计信息
function whiteWebStat($content){
    $splStr=''; $splxx=''; $filePath='';$nCount='';
    $url=''; $s=''; $visitUrl=''; $viewUrl=''; $viewdatetime=''; $ip=''; $browser=''; $operatingsystem=''; $cookie=''; $screenwh=''; $moreInfo=''; $ipList=''; $dateClass ='';
    $splxx= aspSplit($content, vbCrlf() . '-------------------------------------------------' . vbCrlf());
    $nCount=0;
    foreach( $splxx as $key=>$s){
        if( instr($s, '当前：') > 0 ){
            $nCount=$nCount+1;
            $s= vbCrlf() . $s . vbCrlf();
            $dateClass= ADSql(getFileAttr($filePath, '3'));
            $visitUrl= ADSql(getStrCut($s, vbCrlf() . '来访', vbCrlf(), 0));
            $viewUrl= ADSql(getStrCut($s, vbCrlf() . '当前：', vbCrlf(), 0));
            $viewdatetime= ADSql(getStrCut($s, vbCrlf() . '时间：', vbCrlf(), 0));
            $ip= ADSql(getStrCut($s, vbCrlf() . 'IP:', vbCrlf(), 0));
            $browser= ADSql(getStrCut($s, vbCrlf() . 'browser: ', vbCrlf(), 0));
            $operatingsystem= ADSql(getStrCut($s, vbCrlf() . 'operatingsystem=', vbCrlf(), 0));
            $cookie= ADSql(getStrCut($s, vbCrlf() . 'Cookies=', vbCrlf(), 0));
            $screenwh= ADSql(getStrCut($s, vbCrlf() . 'Screen=', vbCrlf(), 0));
            $moreInfo= ADSql(getStrCut($s, vbCrlf() . '用户信息=', vbCrlf(), 0));
            $browser= ADSql(getBrType($moreInfo));
            if( instr(vbCrlf() . $ipList . vbCrlf(), vbCrlf() . $ip . vbCrlf())== false ){
                $ipList= $ipList . $ip . vbCrlf();
            }

            $viewdatetime=replace($viewdatetime,'来访','00');
            if( isDate($viewdatetime)==false ){
                $viewdatetime='1988/07/12 10:10:10';
            }

            $screenwh= substr($screenwh, 0 , 20);
            if( 1== 2 ){
                ASPEcho('编号',$nCount);
                ASPEcho('dateClass', $dateClass);
                ASPEcho('visitUrl', $visitUrl);
                ASPEcho('viewUrl', $viewUrl);
                ASPEcho('viewdatetime', $viewdatetime);
                ASPEcho('IP', $ip);
                ASPEcho('browser', $browser);
                ASPEcho('operatingsystem', $operatingsystem);
                ASPEcho('cookie', $cookie);
                ASPEcho('screenwh', $screenwh);
                ASPEcho('moreInfo', $moreInfo);
                hr();
            }
            connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'websitestat (visiturl,viewurl,browser,operatingsystem,screenwh,moreinfo,viewdatetime,ip,dateclass) values(\'' . $visitUrl . '\',\'' . $viewUrl . '\',\'' . $browser . '\',\'' . $operatingsystem . '\',\'' . $screenwh . '\',\'' . $moreInfo . '\',\'' . $viewdatetime . '\',\'' . $ip . '\',\'' . $dateClass . '\')');
        }
    }
}

//详细网站统计
function websiteDetail(){
    $content=''; $splxx=''; $filePath ='';
    $s=''; $ip=''; $ipList ='';
    $nIP=''; $nPV=''; $i=''; $timeStr=''; $c ='';

    handlePower('网站统计详细'); //管理权限处理

    for( $i= 1 ; $i<= 30; $i++){
        $timeStr= getHandleDate(($i - 1) * - 1); //format_Time(Now() - i + 1, 2)
        $filePath= $GLOBALS['adminDir'] . '/data/stat/' . $timeStr . '.txt';
        $content= getftext($filePath);
        $splxx= aspSplit($content, vbCrlf() . '-------------------------------------------------' . vbCrlf());
        $nIP= 0;
        $nPV= 0;
        $ipList= '';
        foreach( $splxx as $key=>$s){
            if( instr($s, '当前：') > 0 ){
                $s= vbCrlf() . $s . vbCrlf();
                $ip= ADSql(getStrCut($s, vbCrlf() . 'IP:', vbCrlf(), 0));
                $nPV= $nPV + 1;
                if( instr(vbCrlf() . $ipList . vbCrlf(), vbCrlf() . $ip . vbCrlf())== false ){
                    $ipList= $ipList . $ip . vbCrlf();
                    $nIP= $nIP + 1;
                }
            }
        }
        ASPEcho($timeStr, 'IP(' . $nIP . ') PV(' . $nPV . ')');
        if( $i < 4 ){
            $c= $c . $timeStr . ' IP(' . $nIP . ') PV(' . $nPV . ')' . '<br>';
        }
    }

    setConfigFileBlock($GLOBALS['WEB_CACHEFile'], $c, '#访客信息#');
    writeSystemLog('', '详细网站统计'); //系统日志

}

//显示指定布局
function displayLayout(){
    $content=''; $lableTitle=''; $templateFile='';
    handlePower('显示' . $lableTitle); //管理权限处理
    //读模板
    $lableTitle= @$_REQUEST['lableTitle'];
    $templateFile=@$_REQUEST['templateFile'];

    $content= getTemplateContent(@$_REQUEST['templateFile']);
    $content= Replace($content, '[$position$]', $lableTitle);
    $content= replaceValueParam($content, 'lableTitle', $lableTitle);



    if( $templateFile=='layout_makeRobots.html' ){
        $content= Replace($content, '[$bodycontent$]', getftext('/robots.txt'));
    }else if( $templateFile=='layout_adminMap.html' ){
        $content= replaceValueParam($content, 'adminmapbody', getAdminMap());
    }else if( $templateFile=='layout_manageTemplates.html' ){
        $content= displayTemplatesList($content);
    }else if( $templateFile=='layout_manageMakeHtml.html' ){
        $content= replaceValueParam($content, 'columnList', getMakeColumnList());


    }


    $content=handleDisplayLanguage($content,'handleDisplayLanguage');			//语言处理
    rw($content);
}
//获得生成栏目列表
function getMakeColumnList(){
    $c ='';
    //栏目
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn order by sortrank asc');
    while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
        if( $rsx['nofollow']== false ){
            $c= $c . '<option value="' . $rsx['id'] . '">' . $rsx['columnname'] . '</option>' . vbCrlf();
        }
    }
    $getMakeColumnList= $c;
    return @$getMakeColumnList;
}

//获得后台地图
function getAdminMap(){
    $s=''; $c=''; $url=''; $addSql ='';
    if( @$_SESSION['adminflags'] <> '|*|' ){
        $addSql= ' and isDisplay<>0 ';
    }
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'listmenu where parentid=-1 ' . $addSql . ' order by sortrank');
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        $c= $c . '<div class="map-menu fl"><ul>' . vbCrlf();
        $c= $c . '<li class="title">' . $rs['title'] . '</li><div>' . vbCrlf();
        $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'listmenu where parentid=' . $rs['id'] . ' ' . $addSql . '  order by sortrank');
        while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
            $url= phptrim($rsx['customaurl']);
            if( $rsx['lablename'] <> '' ){
                $url= $url . '&lableTitle=' . $rsx['lablename'];
            }
            $c= $c . '<li><a href="' . $url . '">' . $rsx['title'] . '</a></li>' . vbCrlf();
        }
        $c= $c . '</div></ul></div>' . vbCrlf();
    }
    $c= replaceLableContent($c);
    $getAdminMap= $c;
    return @$getAdminMap;
}

//获得后台一级菜单列表
function getAdminOneMenuList(){
    $c=''; $focusStr=''; $addSql=''; $sql ='';
    if( @$_SESSION['adminflags'] <> '|*|' ){
        $addSql= ' and isDisplay<>0 ';
    }
    $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . 'listmenu where parentid=-1 ' . $addSql . ' order by sortrank';
    //检测SQL
    if( checksql($sql)== false ){
        errorLog('出错提示：<br>function=getAdminOneMenuList<hr>sql=' . $sql . '<br>');
        return '';
    }
    $rsObj=$GLOBALS['conn']->query( $sql);
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        $focusStr= '';
        if( $c== '' ){
            $focusStr= ' class="focus"';
        }
        $c= $c . '<li' . $focusStr . '>' . $rs['title'] . '</li>' . vbCrlf();
    }
    $c= replaceLableContent($c);
    $getAdminOneMenuList= $c;
    return @$getAdminOneMenuList;
}
//获得后台菜单列表
function getAdminMenuList(){
    $s=''; $c=''; $url=''; $selStr=''; $addSql=''; $sql ='';
    if( @$_SESSION['adminflags'] <> '|*|' ){
        $addSql= ' and isDisplay<>0 ';
    }
    $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . 'listmenu where parentid=-1 ' . $addSql . ' order by sortrank';
    //检测SQL
    if( checksql($sql)== false ){
        errorLog('出错提示：<br>function=getAdminMenuList<hr>sql=' . $sql . '<br>');
        return '';
    }
    $rsObj=$GLOBALS['conn']->query( $sql);
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        $selStr= 'didoff';
        if( $c== '' ){
            $selStr= 'didon';
        }

        $c= $c . '<ul class="navwrap">' . vbCrlf();
        $c= $c . '<li class="' . $selStr . '">' . $rs['title'] . '</li>' . vbCrlf();


        $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'listmenu where parentid=' . $rs['id'] . '  ' . $addSql . ' order by sortrank');
        while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
            $url= phptrim($rsx['customaurl']);
            $c= $c . ' <li class="item" onClick="window1(\'' . $url . '\',\'' . $rsx['lablename'] . '\');">' . $rsx['title'] . '</li>' . vbCrlf();

        }
        $c= $c . '</ul>' . vbCrlf();
    }
    $c= replaceLableContent($c);
    $getAdminMenuList= $c;
    return @$getAdminMenuList;
}
//处理模板列表
function displayTemplatesList($content){
    $templatesFolder=''; $templatePath=''; $templatePath2=''; $templateName=''; $defaultList=''; $folderList=''; $splStr=''; $s=''; $c ='';
    $splTemplatesFolder ='';
    //加载网址配置
    loadWebConfig();

    $defaultList= getStrCut($content, '[list]', '[/list]', 2);

    $splTemplatesFolder= aspSplit('/Templates/|/Templates2015/|/Templates2016/', '|');
    foreach( $splTemplatesFolder as $key=>$templatesFolder){
        if( $templatesFolder <> '' ){
            $folderList= getDirFolderNameList($templatesFolder);
            $splStr= aspSplit($folderList, vbCrlf());
            foreach( $splStr as $key=>$templateName){
                if( $templateName <> '' && instr('#_', substr($templateName, 0 , 1))== false ){
                    $templatePath= $templatesFolder . $templateName . '/';
                    $templatePath2= $templatePath;
                    $s= $defaultList;
                    if( $GLOBALS['cfg_webtemplate']== $templatePath ){
                        $templateName= '<font color=red>' . $templateName . '</font>';
                        $templatePath2= '<font color=red>' . $templatePath2 . '</font>';
                        $s= Replace($s, '启用</a>', '</a>');
                    }else{
                        $s= Replace($s, '恢复数据</a>', '</a>');
                    }
                    $s= replaceValueParam($s, 'templatename', $templateName);
                    $s= replaceValueParam($s, 'templatepath', $templatePath);
                    $s= replaceValueParam($s, 'templatepath2', $templatePath2);
                    $c= $c . $s . vbCrlf();
                }
            }
        }
    }
    $content= Replace($content, '[list]' . $defaultList . '[/list]', $c);
    $displayTemplatesList= $content;
    return @$displayTemplatesList;
}
//应用模板
function isOpenTemplate(){
    $templatePath=''; $templateName=''; $editValueStr=''; $url ='';

    handlePower('启用模板'); //管理权限处理

    $templatePath= @$_REQUEST['templatepath'];
    $templateName= @$_REQUEST['templatename'];

    if( getRecordCount($GLOBALS['db_PREFIX'] . 'website', '')== 0 ){
        connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'website(webtitle) values(\'测试\')');
    }


    $editValueStr= 'webtemplate=\'' . $templatePath . '\',webimages=\'' . $templatePath . 'Images/\'';
    $editValueStr= $editValueStr . ',webcss=\'' . $templatePath . 'css/\',webjs=\'' . $templatePath . 'Js/\'';
    connExecute('update ' . $GLOBALS['db_PREFIX'] . 'website set ' . $editValueStr);
    $url= '?act=displayLayout&templateFile=layout_manageTemplates.html&lableTitle=模板管理';



    rw(getMsg1('启用模板成功，正在进入模板管理界面...', $url));
    writeSystemLog('', '应用模板' . $templatePath); //系统日志
}
//执行SQL
function executeSQL(){
    $sqlvalue ='';
    $sqlvalue= 'delete from ' . $GLOBALS['db_PREFIX'] . 'WebSiteStat';
    if( @$_REQUEST['sqlvalue'] <> '' ){
        $sqlvalue= @$_REQUEST['sqlvalue'];
        $GLOBALS['conn=']=OpenConn();
        //检测SQL
        if( checksql($sqlvalue)== false ){
            errorLog('出错提示：<br>sql=' . $sqlvalue . '<br>');
            return '';
        }
        ASPEcho('执行SQL语句成功', $sqlvalue);
    }
    if( @$_SESSION['adminusername']== 'ASPPHPCMS' ){
        rw('<form id="form1" name="form1" method="post" action="?act=executeSQL"  onSubmit="if(confirm(\'你确定要操作吗？\\n操作后将不可恢复\')){return true}else{return false}">SQL<input name="sqlvalue" type="text" id="sqlvalue" value="' . $sqlvalue . '" size="80%" /><input type="submit" name="button" id="button" value="执行" /></form>');
    }else{
        rw('你没有权限执行SQL语句');
    }
}





?>