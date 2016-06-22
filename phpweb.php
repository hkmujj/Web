<?php 


define('WEBPATH', $_SERVER ['DOCUMENT_ROOT'].DIRECTORY_SEPARATOR);				//网站主目录
 

//系统核心程序
require_once './PHP2/ImageWaterMark/Include/ASP.php';
require_once './PHP2/ImageWaterMark/Include/sys_FSO.php';
require_once './PHP2/ImageWaterMark/Include/sys_Url.php';
require_once './PHP2/ImageWaterMark/Include/sys_Cai.php';
require_once './PHP2/ImageWaterMark/Include/Conn.php';
require_once './PHP2/ImageWaterMark/Include/MySqlClass.php';
//引用inc
require_once './PHP2/Web/Inc/2014_Array.php';
require_once './PHP2/Web/Inc/2014_Author.php';
require_once './PHP2/Web/Inc/2014_Css.php';
require_once './PHP2/Web/Inc/2014_Js.php'; 
require_once './PHP2/Web/Inc/2015_APGeneral.php';
require_once './PHP2/Web/Inc/2015_Color.php';
require_once './PHP2/Web/Inc/2015_Formatting.php';
require_once './PHP2/Web/Inc/2015_Param.php';
require_once './PHP2/Web/Inc/2015_ToMyPHP.php';
require_once './PHP2/Web/Inc/2015_NewWebFunction.php';
require_once './PHP2/Web/Inc/2016_SaveData.php'; 
require_once './PHP2/Web/Inc/2016_WebControl.php'; 
require_once './PHP2/Web/Inc/ASPPHPAccess.php'; 
//require_once './PHP2/Web/Inc/2015_ToPhpCms.php';
require_once './PHP2/Web/Inc/Cai.php';
require_once './PHP2/Web/Inc/Check.php';
require_once './PHP2/Web/Inc/Common.php'; 
require_once './PHP2/Web/Inc/Incpage.php';
require_once './PHP2/Web/Inc/Print.php';
require_once './PHP2/Web/Inc/StringNumber.php';
require_once './PHP2/Web/Inc/Time.php';
require_once './PHP2/Web/Inc/URL.php';;
require_once './PHP2/Web/Inc/EncDec.php';
require_once './PHP2/Web/Inc/IE.php';
require_once './PHP2/Web/Inc/html.php';
require_once './PHP2/Web/Inc/2016_Log.php'; 
require_once './PHP2/Web/Inc/SystemInfo.php'; 
require_once './PHP2/Web/Inc/2014_Template.php'; 
require_once './PHP2/Web/Inc/FunHTML.php';  
require_once './Web/setAccess.php';
require_once './Web/function.php';
require_once './Web/config.php';
  
//end 引用inc
 

$ReadBlockList='';
$ModuleReplaceArray=''; //替换模块数组，暂时没用，但是要留着，要不出错了  



  
//=========

$code ='';//html代码
$templateName ='';//模板名称
$cfg_webSiteUrl=''; $cfg_webTemplate=''; $cfg_webImages=''; $cfg_webCss=''; $cfg_webJs=''; $cfg_webTitle=''; $cfg_webKeywords=''; $cfg_webDescription=''; $cfg_webSiteBottom=''; $cfg_flags ='';
$glb_columnName=''; $glb_columnId=''; $glb_id=''; $glb_columnType=''; $glb_columnENType=''; $glb_table=''; $glb_detailTitle=''; $glb_flags ='';
$webTemplate ='';//网站模板路径
$glb_url=''; $glb_filePath ='';//当前链接网址,和文件路径
$glb_isonhtml ='';//是否生成静态网页
$glb_locationType ='';//位置类型

$glb_bodyContent ='';//主体内容
$glb_artitleAuthor ='';//文章作者
$glb_artitleAdddatetime ='';//文章添加时间
$glb_upArticle ='';//上一篇文章
$glb_downArticle ='';//下一篇文章
$glb_aritcleRelatedTags ='';//文章标签组
$glb_aritcleSmallImage=''; $glb_aritcleBigImage ='';//文章小图与文章大图
$glb_searchKeyWord ='';//搜索关键词

$isMakeHtml ='';//是否生成网页
//处理动作   ReplaceValueParam为控制字符显示方式
function handleAction($content){
    $startStr=''; $endStr=''; $ActionList=''; $splStr=''; $action=''; $s=''; $HandYes ='';
    $startStr= '{$' ; $endStr= '$}';
    $ActionList= getArray($content, $startStr, $endStr, true, true);
    //Call echo("ActionList ", ActionList)
    $splStr= aspSplit($ActionList, '$Array$');
    foreach( $splStr as $key=>$s){
        $action= AspTrim($s);
        $action= handleInModule($action, 'start'); //处理\'替换掉
        if( $action <> '' ){
            $action= AspTrim(mid($action, 3, Len($action) - 4)) . ' ';
            //call echo("s",s)
            $HandYes= true; //处理为真
            //{VB #} 这种是放在图片路径里，目的是为了在VB里不处理这个路径
            if( checkFunValue($action, '# ')== true ){
                $action= '';
                //测试
            }else if( checkFunValue($action, 'GetLableValue ')== true ){
                $action= XY_getLableValue($action);
                //标题在搜索引擎里列表
            }else if( checkFunValue($action, 'TitleInSearchEngineList ')== true ){
                $action= XY_TitleInSearchEngineList($action);

                //加载文件
            }else if( checkFunValue($action, 'Include ')== true ){
                $action= XY_Include($action);
                //栏目列表
            }else if( checkFunValue($action, 'ColumnList ')== true ){
                $action= XY_AP_ColumnList($action);
                //文章列表
            }else if( checkFunValue($action, 'ArticleList ')== true || checkFunValue($action, 'CustomInfoList ')== true ){
                $action= XY_AP_ArticleList($action);
                //评论列表
            }else if( checkFunValue($action, 'CommentList ')== true ){
                $action= XY_AP_CommentList($action);
                //搜索统计列表
            }else if( checkFunValue($action, 'SearchStatList ')== true ){
                $action= XY_AP_SearchStatList($action);
                //友情链接列表
            }else if( checkFunValue($action, 'Links ')== true ){
                $action= XY_AP_Links($action);

                //显示单页内容
            }else if( checkFunValue($action, 'GetOnePageBody ')== true || checkFunValue($action, 'MainInfo ')== true ){
                $action= XY_AP_GetOnePageBody($action);
                //显示文章内容
            }else if( checkFunValue($action, 'GetArticleBody ')== true ){
                $action= XY_AP_GetArticleBody($action);
                //显示栏目内容
            }else if( checkFunValue($action, 'GetColumnBody ')== true ){
                $action= XY_AP_GetColumnBody($action);

                //获得栏目URL
            }else if( checkFunValue($action, 'GetColumnUrl ')== true ){
                $action= XY_GetColumnUrl($action);
                //获得文章URL
            }else if( checkFunValue($action, 'GetArticleUrl ')== true ){
                $action= XY_GetArticleUrl($action);
                //获得单页URL
            }else if( checkFunValue($action, 'GetOnePageUrl ')== true ){
                $action= XY_GetOnePageUrl($action);

                //显示包裹块 作用不大
            }else if( checkFunValue($action, 'DisplayWrap ')== true ){
                $action= XY_DisplayWrap($action);
                //显示布局
            }else if( checkFunValue($action, 'Layout ')== true ){
                $action= XY_Layout($action);
                //显示模块
            }else if( checkFunValue($action, 'Module ')== true ){
                $action= XY_Module($action);
                //获得内容模块 20150108
            }else if( checkFunValue($action, 'GetContentModule ')== true ){
                $action= XY_ReadTemplateModule($action);
                //读模板样式并设置标题与内容   软件里有个栏目Style进行设置
            }else if( checkFunValue($action, 'ReadColumeSetTitle ')== true ){
                $action= XY_ReadColumeSetTitle($action);

                //显示JS渲染ASP/PHP/VB等程序的编辑器
            }else if( checkFunValue($action, 'displayEditor ')== true ){
                $action= displayEditor($action);

                //Js版网站统计
            }else if( checkFunValue($action, 'JsWebStat ')== true ){
                $action= XY_JsWebStat($action);

                //------------------- 链接区 -----------------------
                //普通链接A
            }else if( checkFunValue($action, 'HrefA ')== true ){
                $action= XY_HrefA($action);

                //栏目菜单(引用后台栏目程序)
            }else if( checkFunValue($action, 'ColumnMenu ')== true ){
                $action= XY_AP_ColumnMenu($action);

                //网站底部
            }else if( checkFunValue($action, 'WebSiteBottom ')== true ){
                $action= XY_AP_WebSiteBottom($action);
                //显示网站栏目 20160331
            }else if( checkFunValue($action, 'DisplayWebColumn ')== true ){
                $action= XY_DisplayWebColumn($action);
                //URL加密
            }else if( checkFunValue($action, 'escape ')== true ){
                $action= XY_escape($action);
                //URL解密
            }else if( checkFunValue($action, 'unescape ')== true ){
                $action= XY_unescape($action);
                //asp与php版本
            }else if( checkFunValue($action, 'EDITORTYPE ')== true ){
                $action= XY_EDITORTYPE($action);


                //暂时不屏蔽
            }else if( checkFunValue($action, 'copyTemplateMaterial ')== true ){
                $action= '';
            }else if( checkFunValue($action, 'clearCache ')== true ){
                $action= '';
            }else{
                $HandYes= false; //处理为假
            }
            //注意这样，有的则不显示 晕 And IsNul(action)=False
            if( isNul($action)== true ){ $action= '' ;}
            if( $HandYes== true ){
                $content= Replace($content, $s, $action);
            }
        }
    }
    $handleAction= $content;
    return @$handleAction;
}

//显示网站栏目 新版 把之前网站 导航程序改进过来的
function XY_DisplayWebColumn($action){
    $i=''; $c=''; $s=''; $url=''; $sql=''; $dropDownMenu=''; $focusType=''; $addSql ='';
    $isConcise ='';//简洁显示20150212
    $styleId=''; $styleValue ='';//样式ID与样式内容
    $cssNameAddId ='';
    $shopnavidwrap ='';//是否显示栏目ID包

    $styleId= PHPTrim(RParam($action, 'styleID'));
    $styleValue= PHPTrim(RParam($action, 'styleValue'));
    $addSql= PHPTrim(RParam($action, 'addSql'));
    $shopnavidwrap= PHPTrim(RParam($action, 'shopnavidwrap'));
    //If styleId <> "" Then
    //Call ReadNavCSS(styleId, styleValue)
    //End If

    //为数字类型 则自动提取样式内容  20150615
    if( checkStrIsNumberType($styleValue) ){
        $cssNameAddId= '_' . $styleValue; //Css名称追加Id编号
    }
    $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn';
    //追加sql
    if( $addSql <> '' ){
        $sql= getWhereAnd($sql, $addSql);
    }
    if( checkSql($sql)== false ){ eerr('Sql', $sql) ;}
    $rsObj=$GLOBALS['conn']->query( $sql);
    $dropDownMenu= strtolower(RParam($action, 'DropDownMenu'));
    $focusType= strtolower(RParam($action, 'FocusType'));
    $isConcise= strtolower(RParam($action, 'isConcise'));
    if( $isConcise== 'true' ){
        $isConcise= false;
    }else{
        $isConcise= true;
    }
    if( $isConcise== true ){ $c= $c . copyStr(' ', 4) . '<li class=left></li>' . vbCrlf() ;}
    for( $i= 1 ; $i<= @mysql_num_rows($rsObj); $i++){

        $rs=mysql_fetch_array($rsObj); //给PHP用，因为在 asptophp转换不完善
        $url= getColumnUrl($rs['columnname'], 'name');
        if( $rs['columnname']== $GLOBALS['glb_columnName'] ){
            if( $focusType== 'a' ){
                $s= copyStr(' ', 8) . '<li class=focus><a href="' . $url . '">' . $rs['columname'] . '</a>';
            }else{
                $s= copyStr(' ', 8) . '<li class=focus>' . $rs['columnname'];
            }
        }else{
            $s= copyStr(' ', 8) . '<li><a href="' . $url . '">' . $rs['columnname'] . '</a>';
        }
        $url= WEB_ADMINURL . '?act=addEditHandle&actionType=WebColumn&lableTitle=网站栏目&nPageSize=10&page=&id=' . $rs['id'] . '&n=' . getRnd(11);
        $s= handleDisplayOnlineEditDialog($url, $s, '', 'div|li|span'); //处理是否添加在线修改管理器

        $c= $c . $s;

        //小类

        $c= $c . copyStr(' ', 8) . '</li>' . vbCrlf();

        if( $isConcise== true ){ $c= $c . copyStr(' ', 8) . '<li class=line></li>' . vbCrlf() ;}
    }
    if( $isConcise== true ){ $c= $c . copyStr(' ', 8) . '<li class=right></li>' . vbCrlf() ;}

    if( $styleId <> '' ){
        $c= '<ul class=\'nav' . $styleId . $cssNameAddId . '\'>' . vbCrlf() . $c . vbCrlf() . '</ul>' . vbCrlf();
    }
    if( $shopnavidwrap== '1' || $shopnavidwrap== 'true' ){
        $c= '<div id=\'nav' . $styleId . $cssNameAddId . '\'>' . vbCrlf() . $c . vbCrlf() . '</div>' . vbCrlf();
    }

    $XY_DisplayWebColumn= $c;
    return @$XY_DisplayWebColumn;
}

//替换全局变量 {$cfg_websiteurl$}
function replaceGlobleVariable( $content){
    $content= handleRGV($content, '{$cfg_webSiteUrl$}', $GLOBALS['cfg_webSiteUrl']); //网址
    $content= handleRGV($content, '{$cfg_webTemplate$}', $GLOBALS['cfg_webTemplate']); //模板
    $content= handleRGV($content, '{$cfg_webImages$}', $GLOBALS['cfg_webImages']); //图片路径
    $content= handleRGV($content, '{$cfg_webCss$}', $GLOBALS['cfg_webCss']); //css路径
    $content= handleRGV($content, '{$cfg_webJs$}', $GLOBALS['cfg_webJs']); //js路径
    $content= handleRGV($content, '{$cfg_webTitle$}', $GLOBALS['cfg_webTitle']); //网站标题
    $content= handleRGV($content, '{$cfg_webKeywords$}', $GLOBALS['cfg_webKeywords']); //网站关键词
    $content= handleRGV($content, '{$cfg_webDescription$}', $GLOBALS['cfg_webDescription']); //网站描述
    $content= handleRGV($content, '{$cfg_webSiteBottom$}', $GLOBALS['cfg_webSiteBottom']); //网站描述

    $content= handleRGV($content, '{$glb_columnId$}', $GLOBALS['glb_columnId']); //栏目Id
    $content= handleRGV($content, '{$glb_columnName$}', $GLOBALS['glb_columnName']); //栏目名称
    $content= handleRGV($content, '{$glb_columnType$}', $GLOBALS['glb_columnType']); //栏目类型
    $content= handleRGV($content, '{$glb_columnENType$}', $GLOBALS['glb_columnENType']); //栏目英文类型

    $content= handleRGV($content, '{$glb_Table$}', $GLOBALS['glb_table']); //表
    $content= handleRGV($content, '{$glb_Id$}', $GLOBALS['glb_id']); //id


    //兼容旧版本 渐渐把它去掉
    $content= handleRGV($content, '{$WebImages$}', $GLOBALS['cfg_webImages']); //图片路径
    $content= handleRGV($content, '{$WebCss$}', $GLOBALS['cfg_webCss']); //css路径
    $content= handleRGV($content, '{$WebJs$}', $GLOBALS['cfg_webJs']); //js路径
    $content= handleRGV($content, '{$Web_Title$}', $GLOBALS['cfg_webTitle']);
    $content= handleRGV($content, '{$Web_KeyWords$}', $GLOBALS['cfg_webKeywords']);
    $content= handleRGV($content, '{$Web_Description$}', $GLOBALS['cfg_webDescription']);


    $content= handleRGV($content, '{$EDITORTYPE$}', EDITORTYPE); //后缀
    $content= handleRGV($content, '{$WEB_VIEWURL$}', WEB_VIEWURL); //首页显示网址
    //文章用到
    $content= handleRGV($content, '{$glb_artitleAuthor$}', $GLOBALS['glb_artitleAuthor']); //文章作者
    $content= handleRGV($content, '{$glb_artitleAdddatetime$}', $GLOBALS['glb_artitleAdddatetime']); //文章添加时间
    $content= handleRGV($content, '{$glb_upArticle$}', $GLOBALS['glb_upArticle']); //上一篇文章
    $content= handleRGV($content, '{$glb_downArticle$}', $GLOBALS['glb_downArticle']); //下一篇文章
    $content= handleRGV($content, '{$glb_aritcleRelatedTags$}', $GLOBALS['glb_aritcleRelatedTags']); //文章标签组
    $content= handleRGV($content, '{$glb_aritcleBigImage$}', $GLOBALS['glb_aritcleBigImage']); //文章大图
    $content= handleRGV($content, '{$glb_aritcleSmallImage$}', $GLOBALS['glb_aritcleSmallImage']); //文章小图
    $content= handleRGV($content, '{$glb_searchKeyWord$}', $GLOBALS['glb_searchKeyWord']); //首页显示网址


    $replaceGlobleVariable= $content;
    return @$replaceGlobleVariable;
}

//处理替换
function handleRGV( $content, $findStr, $replaceStr){
    $lableName ='';
    //对[$$]处理
    $lableName= mid($findStr, 3, Len($findStr) - 4) . ' ';
    $lableName= mid($lableName, 1, instr($lableName, ' ') - 1);
    $content= replaceValueParam($content, $lableName, $replaceStr);
    $content= replaceValueParam($content, strtolower($lableName), $replaceStr);
    //直接替换{$$}这种方式，兼容之前网站
    $content= Replace($content, $findStr, $replaceStr);
    $content= Replace($content, strtolower($findStr), $replaceStr);
    $handleRGV= $content;
    return @$handleRGV;
}

//加载网址配置
function loadWebConfig(){
    $templatedir ='';
    $GLOBALS['conn=']=OpenConn();
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'website');
    $rs=mysql_fetch_array($rsObj);
    if( @mysql_num_rows($rsObj)!=0 ){
        $GLOBALS['cfg_webSiteUrl']= phptrim($rs['websiteurl']); //网址
        $GLOBALS['cfg_webTemplate']= $GLOBALS['webDir'] . phptrim($rs['webtemplate']); //模板路径
        $GLOBALS['cfg_webImages']= $GLOBALS['webDir'] . phptrim($rs['webimages']); //图片路径
        $GLOBALS['cfg_webCss']= $GLOBALS['webDir'] . phptrim($rs['webcss']); //css路径
        $GLOBALS['cfg_webJs']= $GLOBALS['webDir'] . phptrim($rs['webjs']); //js路径
        $GLOBALS['cfg_webTitle']= $rs['webtitle']; //网址标题
        $GLOBALS['cfg_webKeywords']= $rs['webkeywords']; //网站关键词
        $GLOBALS['cfg_webDescription']= $rs['webdescription']; //网站描述
        $GLOBALS['cfg_webSiteBottom']= $rs['websitebottom']; //网站地底
        $GLOBALS['cfg_flags']= $rs['flags']; //旗

        //改换模板20160202
        if( @$_REQUEST['templatedir'] <> '' ){
            //删除绝对目录前面的目录，不需要那个东西20160414
            $templatedir= replace(handlePath(@$_REQUEST['templatedir']),handlePath('/'),'/');
            //call eerr("templatedir",templatedir)

            if((instr($templatedir, ':') > 0 || instr($templatedir, '..') > 0) && getIP() <> '127.0.0.1' ){
                eerr('提示', '模板目录有非法字符');
            }
            $templatedir= handlehttpurl(Replace($templatedir, handlePath('/'), '/'));

            $GLOBALS['cfg_webImages']= Replace($GLOBALS['cfg_webImages'], $GLOBALS['cfg_webTemplate'], $templatedir);
            $GLOBALS['cfg_webCss']= Replace($GLOBALS['cfg_webCss'], $GLOBALS['cfg_webTemplate'], $templatedir);
            $GLOBALS['cfg_webJs']= Replace($GLOBALS['cfg_webJs'], $GLOBALS['cfg_webTemplate'], $templatedir);
            $GLOBALS['cfg_webTemplate']= $templatedir;
        }
        $GLOBALS['webTemplate']= $GLOBALS['cfg_webTemplate'];
    }
}

//网站位置 待完善
function thisPosition($content){
    $c ='';
    $c= '<a href="/">首页</a>';
    if( $GLOBALS['glb_columnName'] <> '' ){
        $c= $c . ' >> <a href="' . getColumnUrl($GLOBALS['glb_columnName'], 'name') . '">' . $GLOBALS['glb_columnName'] . '</a>';
    }
    //20160330
    if( $GLOBALS['glb_locationType']== 'detail' ){
        $c= $c . ' >> 查看内容';
    }
    //call echo("glb_locationType",glb_locationType)

    $content= Replace($content, '[$detailPosition$]', $c);
    $content= Replace($content, '[$detailTitle$]', $GLOBALS['glb_detailTitle']);
    $content= Replace($content, '[$detailContent$]', $GLOBALS['glb_bodyContent']);

    $thisPosition= $content;
    return @$thisPosition;
}

//显示管理列表
function getDetailList($action, $content, $actionName, $lableTitle, $fieldNameList, $nPageSize, $nPage, $addSql){
    $GLOBALS['conn=']=OpenConn();
    $defaultList=''; $i=''; $s=''; $c=''; $tableName=''; $j=''; $splxx=''; $sql ='';
    $x=''; $url=''; $nCount ='';
    $pageInfo ='';

    $fieldName ='';//字段名称
    $splFieldName ='';//分割字段

    $replaceStr ='';//替换字符
    $tableName= strtolower($actionName); //表名称
    $listFileName ='';//列表文件名称
    $listFileName= RParam($action, 'listFileName');
    $abcolorStr ='';//A加粗和颜色
    $atargetStr ='';//A链接打开方式
    $atitleStr ='';//A链接的title20160407
    $anofollowStr ='';//A链接的nofollow

    $id ='';
    $id= rq('id');
    checkIDSQL(@$_REQUEST['id']);

    if( $fieldNameList== '*' ){
        $fieldNameList= getHandleFieldList($GLOBALS['db_PREFIX'] . $tableName, '字段列表');
    }

    $fieldNameList= specialStrReplace($fieldNameList); //特殊字符处理
    $splFieldName= aspSplit($fieldNameList, ','); //字段分割成数组


    $defaultList= getStrCut($content, '[list]', '[/list]', 2);
    $pageInfo= getStrCut($content, '[page]', '[/page]', 1);
    if( $pageInfo <> '' ){
        $content= Replace($content, $pageInfo, '');
    }

    $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . $tableName . ' ' . $addSql;
    //检测SQL
    if( checksql($sql)== false ){
        errorLog('出错提示：<br>sql=' . $sql . '<br>');
        return '';
    }
    $rsObj=$GLOBALS['conn']->query( $sql);
    $nCount= @mysql_num_rows($rsObj);

    //为动态翻页网址
    if( $GLOBALS['isMakeHtml']== true ){
        $url= '';
        if( Len($listFileName) > 5 ){
            $url= mid($listFileName, 1, Len($listFileName) - 5) . '[id].html';
            $url= urlAddHttpUrl($GLOBALS['cfg_webSiteUrl'], $url);
        }
    }else{
        $url= getUrlAddToParam(getUrl(), '?page=[id]', 'replace');
    }
    $content= Replace($content, '[$pageInfo$]', webPageControl($nCount, $nPageSize, $nPage, $url, $pageInfo));

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

        $s= $defaultList;
        $s= Replace($s, '[$id$]', $rs['id']);
        for( $j= 0 ; $j<= UBound($splFieldName); $j++){
            if( $splFieldName[$j] <> '' ){
                $splxx= aspSplit($splFieldName[$j] . '|||', '|');
                $fieldName= $splxx[0];
                $replaceStr= $rs[$fieldName] . '';
                $s= replaceValueParam($s, $fieldName, $replaceStr);
            }

            if( $GLOBALS['isMakeHtml']== true ){
                $url= getHandleRsUrl($rs['filename'], $rs['customaurl'], '/detail/detail' . $rs['id']);
            }else{
                $url= handleWebUrl('?act=detail&id=' . $rs['id']);
                if( $rs['customaurl'] <> '' ){
                    $url= $rs['customaurl'];
                }
            }

            //A链接添加颜色
            $abcolorStr= '';
            if( instr($fieldNameList, ',titlecolor,') > 0 ){
                //A链接颜色
                if( $rs['titlecolor'] <> '' ){
                    $abcolorStr= 'color:' . $rs['titlecolor'] . ';';
                }
            }
            if( instr($fieldNameList, ',flags,') > 0 ){
                //A链接加粗
                if( instr($rs['flags'], '|b|') > 0 ){
                    $abcolorStr= $abcolorStr . 'font-weight:bold;';
                }
            }
            if( $abcolorStr <> '' ){
                $abcolorStr= ' style="' . $abcolorStr . '"';
            }

            //打开方式2016
            if( instr($fieldNameList, ',target,') > 0 ){
                $atargetStr= IIF($rs['target'] <> '', ' target="' . $rs['target'] . '"', '');
            }

            //A的title
            if( instr($fieldNameList, ',title,') > 0 ){
                $atitleStr= IIF($rs['title'] <> '', ' title="' . $rs['title'] . '"', '');
            }

            //A的nofollow
            if( instr($fieldNameList, ',nofollow,') > 0 ){
                $anofollowStr= IIF($rs['nofollow'] <> 0, ' rel="nofollow"', '');
            }



            $s= replaceValueParam($s, 'url', $url);
            $s= replaceValueParam($s, 'abcolor', $abcolorStr); //A链接加颜色与加粗
            $s= replaceValueParam($s, 'atitle', $atitleStr); //A链接title
            $s= replaceValueParam($s, 'anofollow', $anofollowStr); //A链接nofollow
            $s= replaceValueParam($s, 'atarget', $atargetStr); //A链接打开方式


        }
        //文章列表加在线编辑
        $url= WEB_ADMINURL . '?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page=&parentid=&id=' . $rs['id'] . '&n=' . getRnd(11);
        $s= handleDisplayOnlineEditDialog($url, $s, '', 'div|li|span');

        $c= $c . $s;
    }
    $content= Replace($content, '[list]' . $defaultList . '[/list]', $c);

    if( $GLOBALS['isMakeHtml']== true ){
        $url= '';
        if( Len($listFileName) > 5 ){
            $url= mid($listFileName, 1, Len($listFileName) - 5) . '[id].html';
            $url= urlAddHttpUrl($GLOBALS['cfg_webSiteUrl'], $url);
        }
    }else{
        $url= getUrlAddToParam(getUrl(), '?page=[id]', 'replace');
    }

    $getDetailList= $content;
    return @$getDetailList;
}


//****************************************************
//默认列表模板
function defaultListTemplate(){
    $c=''; $templateHtml=''; $listTemplate=''; $lableName=''; $startStr=''; $endStr ='';

    $templateHtml= getFText($GLOBALS['cfg_webTemplate'] . '/' . $GLOBALS['templateName']);

    $lableName= 'list';
    $startStr= '<!--#' . $lableName . ' start#-->';
    $endStr= '<!--#' . $lableName . ' end#-->';
    //call rwend(templateHtml)
    if( instr($templateHtml, $startStr) > 0 && instr($templateHtml, $endStr) > 0 ){
        $listTemplate= strCut($templateHtml, $startStr, $endStr, 2);
    }else{
        $startStr= '<!--#' . $lableName;
        $endStr= '#-->';
        if( instr($templateHtml, $startStr) > 0 && instr($templateHtml, $endStr) > 0 ){
            $listTemplate= strCut($templateHtml, $startStr, $endStr, 2);
        }
    }
    if( $listTemplate== '' ){
        $c= '<ul class="list">' . vbCrlf();
        $c= $c . '[list]    <li><a href="[$url$]"[$atitle$][$atarget$][$abcolor$][$anofollow$]>[$title$]</a><span class="time">[$adddatetime format_time=\'7\'$]</span></li>' . vbCrlf();
        $c= $c . '[/list]' . vbCrlf();
        $c= $c . '</ul>' . vbCrlf();
        $c= $c . '<div class="clear10"></div>' . vbCrlf();
        $c= $c . '<div>[$pageInfo$]</div>' . vbCrlf();
        $listTemplate= $c;
    }

    $defaultListTemplate= $listTemplate;
    return @$defaultListTemplate;
}


//记录表前缀
if( @$_REQUEST['db_PREFIX'] <> '' ){
    $db_PREFIX= @$_REQUEST['db_PREFIX'];
}else if( @$_SESSION['db_PREFIX'] <> '' ){
    $db_PREFIX= @$_SESSION['db_PREFIX'];
}
//加载网址配置
loadWebConfig();
$isMakeHtml= false; //默认生成HTML为关闭
if( @$_REQUEST['isMakeHtml']== '1' || @$_REQUEST['isMakeHtml']== 'true' ){
    $isMakeHtml= true;
}
$templateName= @$_REQUEST['templateName']; //模板名称

//保存数据处理页
switch ( @$_REQUEST['act'] ){
    case 'savedata' ; saveData(@$_REQUEST['stype']) ; die(); //保存数据
    break;//'站长统计 | 今日IP[653] | 今日PV[9865] | 当前在线[65]')
    case 'webstat' ; webStat($adminDir . '/Data/Stat/');die(); //网站统计
    break;
    case 'saveSiteMap' ; $isMakeHtml=true;saveSiteMap() ;die(); //保存sitemap.xml
    break;
    case 'handleAction';
    if( @$_REQUEST['ishtml']=='1' ){
        $isMakeHtml= true;
    }
    rwend(handleAction(@$_REQUEST['content']));		//处理动作

}

//生成html
if( @$_REQUEST['act']== 'makehtml' ){
    ASPEcho('makehtml', 'makehtml');
    $isMakeHtml= true;
    makeWebHtml(' action actionType=\'' . @$_REQUEST['act'] . '\' columnName=\'' . @$_REQUEST['columnName'] . '\' id=\'' . @$_REQUEST['id'] . '\' ');
    createFileGBK('index.html', $code);

    //复制Html到网站
}else if( @$_REQUEST['act']== 'copyHtmlToWeb' ){
    copyHtmlToWeb();
    //全部生成
}else if( @$_REQUEST['act']== 'makeallhtml' ){
    makeAllHtml('', '', @$_REQUEST['id']);

    //生成当前页面
}else if( @$_REQUEST['isMakeHtml'] <> '' && @$_REQUEST['isSave'] <> '' ){

    handlePower('生成当前HTML页面'); //管理权限处理
    writeSystemLog('', '生成当前HTML页面'); //系统日志

    $isMakeHtml= true;


    checkIDSQL(@$_REQUEST['id']);
    rw(makeWebHtml(' action actionType=\'' . @$_REQUEST['act'] . '\' columnName=\'' . @$_REQUEST['columnName'] . '\' columnType=\'' . @$_REQUEST['columnType'] . '\' id=\'' . @$_REQUEST['id'] . '\' npage=\'' . @$_REQUEST['page'] . '\' '));
    $glb_filePath= Replace($glb_url, $cfg_webSiteUrl, '');
    if( Right($glb_filePath, 1)== '/' ){
        $glb_filePath= $glb_filePath . 'index.html';
    }else if( $glb_filePath== '' && $glb_columnType== '首页' ){
        $glb_filePath= 'index.html';
    }
    //文件不为空  并且开启生成html
    if( $glb_filePath <> '' && $glb_isonhtml== true ){
        createDirFolder(getFileAttr($glb_filePath, '1'));
        createFileGBK($glb_filePath, $code);
        if( @$_REQUEST['act']== 'detail' ){
            connExecute('update ' . $db_PREFIX . 'ArticleDetail set ishtml=true where id=' . @$_REQUEST['id']);
        }else if( @$_REQUEST['act']== 'nav' ){
            if( @$_REQUEST['id'] <> '' ){
                connExecute('update ' . $db_PREFIX . 'WebColumn set ishtml=true where id=' . @$_REQUEST['id']);
            }else{
                connExecute('update ' . $db_PREFIX . 'WebColumn set ishtml=true where columnname=\'' . @$_REQUEST['columnName'] . '\'');
            }
        }
        ASPEcho('生成文件路径', '<a href="' . $glb_filePath . '" target=\'_blank\'>' . $glb_filePath . '</a>');

        //新闻则批量生成 20160216
        if( $glb_columnType== '新闻' ){
            makeAllHtml('', '', $glb_columnId);
        }

    }

    //全部生成
}else if( @$_REQUEST['act']== 'Search' ){
    rw(makeWebHtml('actionType=\'Search\' npage=\'1\' '));
}else{
    if( strtolower(@$_REQUEST['issave'])== '1' ){
        makeAllHtml(@$_REQUEST['columnType'], @$_REQUEST['columnName'], @$_REQUEST['columnId']);
    }else{
        checkIDSQL(@$_REQUEST['id']);
        rw(makeWebHtml(' action actionType=\'' . @$_REQUEST['act'] . '\' columnName=\'' . @$_REQUEST['columnName'] . '\' columnType=\'' . @$_REQUEST['columnType'] . '\' id=\'' . @$_REQUEST['id'] . '\' npage=\'' . @$_REQUEST['page'] . '\' '));
    }
}
//检测ID是否SQL安全
function checkIDSQL($id){
    if( checkNumber($id)== false && $id <> '' ){
        eerr('提示', 'id中有非法字符');
    }
}




//http://127.0.0.1/aspweb.asp?act=nav&columnName=ASP
//http://127.0.0.1/aspweb.asp?act=detail&id=75
//生成html静态页
function makeWebHtml($action){
    $actionType=''; $npagesize=''; $npage=''; $url=''; $addSql=''; $sortSql ='';
    $actionType= RParam($action, 'actionType');
    $npage= RParam($action, 'npage');
    $npage= getnumber($npage);
    if( $npage== '' ){
        $npage= 1;
    }else{
        $npage= intval($npage);
    }
    //导航
    if( $actionType== 'nav' ){
        $GLOBALS['glb_columnType']= RParam($action, 'columnType');
        $GLOBALS['glb_columnName']= RParam($action, 'columnName');
        $GLOBALS['glb_columnId']= RParam($action, 'columnId');
        if( $GLOBALS['glb_columnId']== '' ){
            $GLOBALS['glb_columnId']= RParam($action, 'id');
        }
        if( $GLOBALS['glb_columnType'] <> '' ){
            $addSql= 'where columnType=\'' . $GLOBALS['glb_columnType'] . '\'';
        }
        if( $GLOBALS['glb_columnName'] <> '' ){
            $addSql= getWhereAnd($addSql, 'where columnName=\'' . $GLOBALS['glb_columnName'] . '\'');
        }
        if( $GLOBALS['glb_columnId'] <> '' ){
            $addSql= getWhereAnd($addSql, 'where id=' . $GLOBALS['glb_columnId'] . '');
        }
        //call echo("addsql",addsql)
        $rsObj=$GLOBALS['conn']->query( 'Select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn ' . $addSql);
        $rs=mysql_fetch_array($rsObj);
        if( @mysql_num_rows($rsObj)!=0 ){
            $GLOBALS['glb_columnId']= $rs['id'];
            $GLOBALS['glb_columnName']= $rs['columnname'];
            $GLOBALS['glb_columnType']= $rs['columntype'];
            $GLOBALS['glb_bodyContent']= $rs['bodycontent'];
            $GLOBALS['glb_detailTitle']= $GLOBALS['glb_columnName'];
            $GLOBALS['glb_flags']= $rs['flags'];
            $npagesize= $rs['npagesize']; //每页显示条数
            $GLOBALS['glb_isonhtml']= $rs['isonhtml']; //是否生成静态网页
            $sortSql= ' ' . $rs['sortsql']; //排序SQL

            if( $rs['webtitle'] <> '' ){
                $GLOBALS['cfg_webTitle']= $rs['webtitle']; //网址标题
            }
            if( $rs['webkeywords'] <> '' ){
                $GLOBALS['cfg_webKeywords']= $rs['webkeywords']; //网站关键词
            }
            if( $rs['webdescription'] <> '' ){
                $GLOBALS['cfg_webDescription']= $rs['webdescription']; //网站描述
            }
            if( $GLOBALS['templateName']== '' ){
                if( AspTrim($rs['templatepath']) <> '' ){
                    $GLOBALS['templateName']= $rs['templatepath'];
                }else if( $rs['columntype'] <> '首页' ){
                    $GLOBALS['templateName']= getDateilTemplate($rs['id'], 'List');
                }
            }
        }
        $GLOBALS['glb_columnENType']= handleColumnType($GLOBALS['glb_columnType']);
        $GLOBALS['glb_url']= getColumnUrl($GLOBALS['glb_columnName'], 'name');

        //文章类列表
        if( instr('|产品|新闻|视频|下载|案例|', '|' . $GLOBALS['glb_columnType'] . '|') > 0 ){
            $GLOBALS['glb_bodyContent']= getDetailList($action, defaultListTemplate(), 'ArticleDetail', '栏目列表', '*', $npagesize, $npage, 'where parentid=' . $GLOBALS['glb_columnId'] . $sortSql);
            //留言类列表
        }else if( instr('|留言|', '|' . $GLOBALS['glb_columnType'] . '|') > 0 ){
            $GLOBALS['glb_bodyContent']= getDetailList($action, defaultListTemplate(), 'GuestBook', '留言列表', '*', $npagesize, $npage, ' where isthrough<>0 ' . $sortSql);
        }else if( $GLOBALS['glb_columnType']== '文本' ){
            //航行栏目加管理
            if( @$_REQUEST['gl']== 'edit' ){
                $GLOBALS['glb_bodyContent']= '<span>' . $GLOBALS['glb_bodyContent'] . '</span>';
            }
            $url= WEB_ADMINURL . '?act=addEditHandle&actionType=WebColumn&lableTitle=网站栏目&nPageSize=10&page=&id=' . $GLOBALS['glb_columnId'] . '&n=' . getRnd(11);
            $GLOBALS['glb_bodyContent']= handleDisplayOnlineEditDialog($url, $GLOBALS['glb_bodyContent'], '', 'span');

        }
        //细节
    }else if( $actionType== 'detail' ){
        $GLOBALS['glb_locationType']= 'detail';
        $rsObj=$GLOBALS['conn']->query( 'Select * from ' . $GLOBALS['db_PREFIX'] . 'articledetail where id=' . RParam($action, 'id'));
        $rs=mysql_fetch_array($rsObj);
        if( @mysql_num_rows($rsObj)!=0 ){
            $GLOBALS['glb_columnName']= getColumnName($rs['parentid']);
            $GLOBALS['glb_detailTitle']= $rs['title'];
            $GLOBALS['glb_flags']= $rs['flags'];
            $GLOBALS['glb_isonhtml']= $rs['isonhtml']; //是否生成静态网页
            $GLOBALS['glb_id']= $rs['id']; //文章ID
            if( $GLOBALS['isMakeHtml']== true ){
                $GLOBALS['glb_url']= getHandleRsUrl($rs['filename'], $rs['customaurl'], '/detail/detail' . $rs['id']);
            }else{
                $GLOBALS['glb_url']= handleWebUrl('?act=detail&id=' . $rs['id']);
            }

            if( $rs['webtitle'] <> '' ){
                $GLOBALS['cfg_webTitle']= $rs['webtitle']; //网址标题
            }
            if( $rs['webkeywords'] <> '' ){
                $GLOBALS['cfg_webKeywords']= $rs['webkeywords']; //网站关键词
            }
            if( $rs['webdescription'] <> '' ){
                $GLOBALS['cfg_webDescription']= $rs['webdescription']; //网站描述
            }

            $GLOBALS['glb_artitleAuthor']= $rs['author'];
            $GLOBALS['glb_artitleAdddatetime']= $rs['adddatetime'];
            $GLOBALS['glb_upArticle']= upArticle($rs['parentid'], 'sortrank', $rs['sortrank']);
            $GLOBALS['glb_downArticle']= downArticle($rs['parentid'], 'sortrank', $rs['sortrank']);
            $GLOBALS['glb_aritcleRelatedTags']= aritcleRelatedTags($rs['relatedtags']);
            $GLOBALS['glb_aritcleSmallImage']= $rs['smallimage'];
            $GLOBALS['glb_aritcleBigImage']= $rs['bigimage'];

            //文章内容
            //glb_bodyContent = "<div class=""articleinfowrap"">[$articleinfowrap$]</div>" & rs("bodycontent") & "[$relatedtags$]<ul class=""updownarticlewrap"">[$updownArticle$]</ul>"
            //上一篇文章，下一篇文章
            //glb_bodyContent = Replace(glb_bodyContent, "[$updownArticle$]", upArticle(rs("parentid"), "sortrank", rs("sortrank")) & downArticle(rs("parentid"), "sortrank", rs("sortrank")))
            //glb_bodyContent = Replace(glb_bodyContent, "[$articleinfowrap$]", "来源：" & rs("author") & " &nbsp; 发布时间：" & format_Time(rs("adddatetime"), 1))
            //glb_bodyContent = Replace(glb_bodyContent, "[$relatedtags$]", aritcleRelatedTags(rs("relatedtags")))

            $GLOBALS['glb_bodyContent']= $rs['bodycontent'];

            //文章详细加控制
            if( @$_REQUEST['gl']== 'edit' ){
                $GLOBALS['glb_bodyContent']= '<span>' . $GLOBALS['glb_bodyContent'] . '</span>';
            }
            $url= WEB_ADMINURL . '?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page=&parentid=&id=' . RParam($action, 'id') . '&n=' . getRnd(11);
            $GLOBALS['glb_bodyContent']= handleDisplayOnlineEditDialog($url, $GLOBALS['glb_bodyContent'], '', 'span');

            if( $GLOBALS['templateName']== '' ){
                if( AspTrim($rs['templatepath']) <> '' ){
                    $GLOBALS['templateName']= $rs['templatepath'];
                }else{
                    $GLOBALS['templateName']= getDateilTemplate($rs['parentid'], 'Detail');
                }
            }

        }

        //单页
    }else if( $actionType== 'onepage' ){
        $rsObj=$GLOBALS['conn']->query( 'Select * from ' . $GLOBALS['db_PREFIX'] . 'onepage where id=' . RParam($action, 'id'));
        $rs=mysql_fetch_array($rsObj);
        if( @mysql_num_rows($rsObj)!=0 ){
            $GLOBALS['glb_detailTitle']= $rs['title'];
            $GLOBALS['glb_isonhtml']= $rs['isonhtml']; //是否生成静态网页
            if( $GLOBALS['isMakeHtml']== true ){
                $GLOBALS['glb_url']= getHandleRsUrl($rs['filename'], $rs['customaurl'], '/page/page' . $rs['id']);
            }else{
                $GLOBALS['glb_url']= handleWebUrl('?act=detail&id=' . $rs['id']);
            }

            if( $rs['webtitle'] <> '' ){
                $GLOBALS['cfg_webTitle']= $rs['webtitle']; //网址标题
            }
            if( $rs['webkeywords'] <> '' ){
                $GLOBALS['cfg_webKeywords']= $rs['webkeywords']; //网站关键词
            }
            if( $rs['webdescription'] <> '' ){
                $GLOBALS['cfg_webDescription']= $rs['webdescription']; //网站描述
            }
            //内容
            $GLOBALS['glb_bodyContent']= $rs['bodycontent'];


            //文章详细加控制
            if( @$_REQUEST['gl']== 'edit' ){
                $GLOBALS['glb_bodyContent']= '<span>' . $GLOBALS['glb_bodyContent'] . '</span>';
            }
            $url= WEB_ADMINURL . '?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page=&parentid=&id=' . RParam($action, 'id') . '&n=' . getRnd(11);
            $GLOBALS['glb_bodyContent']= handleDisplayOnlineEditDialog($url, $GLOBALS['glb_bodyContent'], '', 'span');


            if( $GLOBALS['templateName']== '' ){
                if( AspTrim($rs['templatepath']) <> '' ){
                    $GLOBALS['templateName']= $rs['templatepath'];
                }else{
                    $GLOBALS['templateName']= 'Main_Model.html';
                    //call echo(templateName,"templateName")
                }
            }

        }

        //搜索
    }else if( $actionType== 'Search' ){
        $GLOBALS['templateName']= 'Main_Model.html';
        $GLOBALS['glb_searchKeyWord']= @$_REQUEST['wd'];
        $addSql= ' where title like \'%' . $GLOBALS['glb_searchKeyWord'] . '%\'';
        $npagesize= 20;
        //call echo(npagesize, npage)
        $GLOBALS['glb_bodyContent']= getDetailList($action, defaultListTemplate(), 'ArticleDetail', '网站栏目', '*', $npagesize, $npage, $addSql);

        //加载等待
    }else if( $actionType== 'loading' ){
        rwend('页面正在加载中。。。');
    }
    //模板为空，则用默认首页模板
    if( $GLOBALS['templateName']== '' ){
        $GLOBALS['templateName']= 'Index_Model.html'; //默认模板
    }
    //检测当前路径是否有模板
    if( instr($GLOBALS['templateName'], '/')== false ){
        $GLOBALS['templateName']= $GLOBALS['cfg_webTemplate'] . '/' . $GLOBALS['templateName'];
    }
    //call echo("templateName",templateName)
    $GLOBALS['code']= getftext($GLOBALS['templateName']);


    $GLOBALS['code']= handleAction($GLOBALS['code']); //处理动作
    $GLOBALS['code']= thisPosition($GLOBALS['code']); //位置
    $GLOBALS['code']= replaceGlobleVariable($GLOBALS['code']); //替换全局标签
    $GLOBALS['code']= handleAction($GLOBALS['code']); //处理动作    '再来一次，处理数据内容里动作

    $GLOBALS['code']= handleAction($GLOBALS['code']); //处理动作
    $GLOBALS['code']= handleAction($GLOBALS['code']); //处理动作
    $GLOBALS['code']= thisPosition($GLOBALS['code']); //位置
    $GLOBALS['code']= replaceGlobleVariable($GLOBALS['code']); //替换全局标签
    $GLOBALS['code']= delTemplateMyNote($GLOBALS['code']); //删除无用内容

    //格式化HTML
    if( instr($GLOBALS['cfg_flags'], '|formattinghtml|') > 0 ){
        //code = HtmlFormatting(code)        '简单
        $GLOBALS['code']= handleHtmlFormatting($GLOBALS['code'], false, 0, '删除空行'); //自定义
        //格式化HTML第二种
    }else if( instr($GLOBALS['cfg_flags'], '|formattinghtmltow|') > 0 ){
        $GLOBALS['code']= htmlFormatting($GLOBALS['code']); //简单
        $GLOBALS['code']= handleHtmlFormatting($GLOBALS['code'], false, 0, '删除空行'); //自定义
        //压缩HTML
    }else if( instr($GLOBALS['cfg_flags'], '|ziphtml|') > 0 ){
        $GLOBALS['code']= ziphtml($GLOBALS['code']);

    }
    //闭合标签
    if( instr($GLOBALS['cfg_flags'], '|labelclose|') > 0 ){
        $GLOBALS['code']= handleCloseHtml($GLOBALS['code'], true, ''); //图片自动加alt  "|*|",
    }

    //在线编辑20160127
    if( rq('gl')== 'edit' ){
        if( instr($GLOBALS['code'], '</head>') > 0 ){
            if( instr($GLOBALS['code'], 'jquery.Min.js')== false ){
                $GLOBALS['code']= Replace($GLOBALS['code'], '</head>', '<script src="/Jquery/jquery.Min.js"></script></head>');
            }
            $GLOBALS['code']= Replace($GLOBALS['code'], '</head>', '<script src="/Jquery/Callcontext_menu.js"></script></head>');
        }
        if( instr($GLOBALS['code'], '<body>') > 0 ){
            //Code = Replace(Code,"<body>", "<body onLoad=""ContextMenu.intializeContextMenu()"">")
        }
    }
    //call echo(templateName,templateName)
    $makeWebHtml= $GLOBALS['code'];
    return @$makeWebHtml;
}

//获得默认细节模板页
function getDateilTemplate($parentid, $templateType){
    $templateName ='';
    $templateName= 'Main_Model.html';
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn where id=' . $parentid);
    $rsx=mysql_fetch_array($rsxObj);
    if( @mysql_num_rows($rsxObj)!=0 ){
        //call echo("columntype",rsx("columntype"))
        if( $rsx['columntype']== '新闻' ){
            //新闻细节页
            if( checkFile($GLOBALS['cfg_webTemplate'] . '/News_' . $templateType . '.html')== true ){
                $templateName= 'News_' . $templateType . '.html';
            }
        }else if( $rsx['columntype']== '产品' ){
            //产品细节页
            if( checkFile($GLOBALS['cfg_webTemplate'] . '/Product_' . $templateType . '.html')== true ){
                $templateName= 'Product_' . $templateType . '.html';
            }
        }else if( $rsx['columntype']== '下载' ){
            //下载细节页
            if( checkFile($GLOBALS['cfg_webTemplate'] . '/Down_' . $templateType . '.html')== true ){
                $templateName= 'Down_' . $templateType . '.html';
            }

        }else if( $rsx['columntype']== '视频' ){
            //视频细节页
            if( checkFile($GLOBALS['cfg_webTemplate'] . '/Video_' . $templateType . '.html')== true ){
                $templateName= 'Video_' . $templateType . '.html';
            }
        }else if( $rsx['columntype']== '留言' ){
            //视频细节页
            if( checkFile($GLOBALS['cfg_webTemplate'] . '/GuestBook_' . $templateType . '.html')== true ){
                $templateName= 'Video_' . $templateType . '.html';
            }
        }else if( $rsx['columntype']== '文本' ){
            //视频细节页
            if( checkFile($GLOBALS['cfg_webTemplate'] . '/Page_' . $templateType . '.html')== true ){
                $templateName= 'Page_' . $templateType . '.html';
            }
        }
    }
    //call echo(templateType,templateName)
    $getDateilTemplate= $templateName;

    return @$getDateilTemplate;
}


//生成全部html页面
function makeAllHtml($columnType, $columnName, $columnId){
    $action=''; $s=''; $i=''; $nPageSize=''; $nCountSize=''; $nPage=''; $addSql=''; $url=''; $articleSql ='';
    handlePower('生成全部HTML页面'); //管理权限处理
    writeSystemLog('', '生成全部HTML页面'); //系统日志

    $GLOBALS['isMakeHtml']= true;
    //栏目
    ASPEcho('栏目', '');
    if( $columnType <> '' ){
        $addSql= 'where columnType=\'' . $columnType . '\'';
    }
    if( $columnName <> '' ){
        $addSql= getWhereAnd($addSql, 'where columnName=\'' . $columnName . '\'');
    }
    if( $columnId <> '' ){
        $addSql= getWhereAnd($addSql, 'where id in(' . $columnId . ')');
    }
    $rssObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn ' . $addSql . ' order by sortrank asc');
    while( $rss= $GLOBALS['conn']->fetch_array($rssObj)){
        $GLOBALS['glb_columnName']= '';
        //开启生成html
        if( $rss['isonhtml']== true ){
            if( instr('|产品|新闻|视频|下载|案例|留言|反馈|招聘|订单|', '|' . $rss['columntype'] . '|') > 0 ){
                if( $rss['columntype']== '留言' ){
                    $nCountSize= getRecordCount($GLOBALS['db_PREFIX'] . 'guestbook', ''); //记录数
                }else{
                    $nCountSize= getRecordCount($GLOBALS['db_PREFIX'] . 'articledetail', ' where parentid=' . $rss['id']); //记录数
                }
                $nPageSize= $rss['npagesize'];
                $nPage= getPageNumb(intval($nCountSize), intval($nPageSize));
                if( $nPage <= 0 ){
                    $nPage= 1;
                }
                for( $i= 1 ; $i<= $nPage; $i++){
                    $url= getHandleRsUrl($rss['filename'], $rss['customaurl'], '/nav' . $rss['id']);
                    $GLOBALS['glb_filePath']= Replace($url, $GLOBALS['cfg_webSiteUrl'], '');
                    if( Right($GLOBALS['glb_filePath'], 1)== '/' || $GLOBALS['glb_filePath']== '' ){
                        $GLOBALS['glb_filePath']= $GLOBALS['glb_filePath'] . 'index.html';
                    }
                    //call echo("glb_filePath",glb_filePath)
                    $action= ' action actionType=\'nav\' columnName=\'' . $rss['columnname'] . '\' npage=\'' . $i . '\' listfilename=\'' . $GLOBALS['glb_filePath'] . '\' ';
                    //call echo("action",action)
                    makeWebHtml($action);
                    if( $i > 1 ){
                        $GLOBALS['glb_filePath']= mid($GLOBALS['glb_filePath'], 1, Len($GLOBALS['glb_filePath']) - 5) . $i . '.html';
                    }
                    $s= '<a href="' . $GLOBALS['glb_filePath'] . '" target=\'_blank\'>' . $GLOBALS['glb_filePath'] . '</a>(' . $rss['isonhtml'] . ')';
                    ASPEcho($action, $s);
                    if( $GLOBALS['glb_filePath'] <> '' ){
                        createDirFolder(getFileAttr($GLOBALS['glb_filePath'], '1'));
                        createFileGBK($GLOBALS['glb_filePath'], $GLOBALS['code']);
                    }
                    doevents();
                    $GLOBALS['templateName']= ''; //清空模板文件名称
                }
            }else{
                $action= ' action actionType=\'nav\' columnName=\'' . $rss['columnname'] . '\'';
                makeWebHtml($action);
                $GLOBALS['glb_filePath']= Replace(getColumnUrl($rss['columnname'], 'name'), $GLOBALS['cfg_webSiteUrl'], '');
                if( Right($GLOBALS['glb_filePath'], 1)== '/' || $GLOBALS['glb_filePath']== '' ){
                    $GLOBALS['glb_filePath']= $GLOBALS['glb_filePath'] . 'index.html';
                }
                $s= '<a href="' . $GLOBALS['glb_filePath'] . '" target=\'_blank\'>' . $GLOBALS['glb_filePath'] . '</a>(' . $rss['isonhtml'] . ')';
                ASPEcho($action, $s);
                if( $GLOBALS['glb_filePath'] <> '' ){
                    createDirFolder(getFileAttr($GLOBALS['glb_filePath'], '1'));
                    createFileGBK($GLOBALS['glb_filePath'], $GLOBALS['code']);
                }
                doevents();
                $GLOBALS['templateName']= '';
            }
            connExecute('update ' . $GLOBALS['db_PREFIX'] . 'WebColumn set ishtml=true where id=' . $rss['id']); //更新导航为生成状态
        }
    }

    //单独处理指定栏目对应文章
    if( $columnId <> '' ){
        $articleSql= 'select * from ' . $GLOBALS['db_PREFIX'] . 'articledetail where parentid=' . $columnId . ' order by sortrank asc';
        //批量处理文章
    }else if( $addSql== '' ){
        $articleSql= 'select * from ' . $GLOBALS['db_PREFIX'] . 'articledetail order by sortrank asc';
    }
    if( $articleSql <> '' ){
        //文章
        ASPEcho('文章', '');
        $rssObj=$GLOBALS['conn']->query( $articleSql);
        while( $rss= $GLOBALS['conn']->fetch_array($rssObj)){
            $GLOBALS['glb_columnName']= '';
            $action= ' action actionType=\'detail\' columnName=\'' . $rss['parentid'] . '\' id=\'' . $rss['id'] . '\'';
            //call echo("action",action)
            makeWebHtml($action);
            $GLOBALS['glb_filePath']= Replace($GLOBALS['glb_url'], $GLOBALS['cfg_webSiteUrl'], '');
            if( Right($GLOBALS['glb_filePath'], 1)== '/' ){
                $GLOBALS['glb_filePath']= $GLOBALS['glb_filePath'] . 'index.html';
            }
            $s= '<a href="' . $GLOBALS['glb_filePath'] . '" target=\'_blank\'>' . $GLOBALS['glb_filePath'] . '</a>(' . $rss['isonhtml'] . ')';
            ASPEcho($action, $s);
            //文件不为空  并且开启生成html
            if( $GLOBALS['glb_filePath'] <> '' && $rss['isonhtml']== true ){
                createDirFolder(getFileAttr($GLOBALS['glb_filePath'], '1'));
                createFileGBK($GLOBALS['glb_filePath'], $GLOBALS['code']);
                connExecute('update ' . $GLOBALS['db_PREFIX'] . 'ArticleDetail set ishtml=true where id=' . $rss['id']); //更新文章为生成状态
            }
            $GLOBALS['templateName']= ''; //清空模板文件名称
        }
    }

    if( $addSql== '' ){
        //单页
        ASPEcho('单页', '');
        $rssObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'onepage order by sortrank asc');
        while( $rss= $GLOBALS['conn']->fetch_array($rssObj)){
            $GLOBALS['glb_columnName']= '';
            $action= ' action actionType=\'onepage\' id=\'' . $rss['id'] . '\'';
            //call echo("action",action)
            makeWebHtml($action);
            $GLOBALS['glb_filePath']= Replace($GLOBALS['glb_url'], $GLOBALS['cfg_webSiteUrl'], '');
            if( Right($GLOBALS['glb_filePath'], 1)== '/' ){
                $GLOBALS['glb_filePath']= $GLOBALS['glb_filePath'] . 'index.html';
            }
            $s= '<a href="' . $GLOBALS['glb_filePath'] . '" target=\'_blank\'>' . $GLOBALS['glb_filePath'] . '</a>(' . $rss['isonhtml'] . ')';
            ASPEcho($action, $s);
            //文件不为空  并且开启生成html
            if( $GLOBALS['glb_filePath'] <> '' && $rss['isonhtml']== true ){
                createDirFolder(getFileAttr($GLOBALS['glb_filePath'], '1'));
                createFileGBK($GLOBALS['glb_filePath'], $GLOBALS['code']);
                connExecute('update ' . $GLOBALS['db_PREFIX'] . 'onepage set ishtml=true where id=' . $rss['id']); //更新单页为生成状态
            }
            $GLOBALS['templateName']= ''; //清空模板文件名称
        }

    }


}

//复制html到网站
function copyHtmlToWeb(){
    $webDir='';$toWebDir=''; $toFilePath=''; $filePath=''; $fileName=''; $fileList=''; $splStr=''; $content=''; $s=''; $s1=''; $c=''; $webImages=''; $webCss=''; $webJs=''; $splJs ='';
    $webFolderName=''; $jsFileList=''; $setFileCode=''; $nErrLevel=''; $jsFilePath='';

    $setFileCode= @$_REQUEST['setcode']; //设置文件保存编码

    handlePower('复制生成HTML页面'); //管理权限处理
    writeSystemLog('', '复制生成HTML页面'); //系统日志

    $webFolderName= $GLOBALS['cfg_webTemplate'];
    if( substr($webFolderName, 0 , 1)== '/' ){
        $webFolderName= mid($webFolderName, 2,-1);
    }
    if( Right($webFolderName, 1)== '/' ){
        $webFolderName= mid($webFolderName, 1, Len($webFolderName) - 1);
    }
    if( instr($webFolderName, '/') > 0 ){
        $webFolderName= mid($webFolderName, instr($webFolderName, '/') + 1,-1);
    }
    $webDir= '/htmladmin/' . $webFolderName . '/';
    $toWebDir='/htmlw' . 'eb/viewweb/';
    createDirFolder($toWebDir);
    $toWebDir= $toWebDir . pinYin2($webFolderName) . '/';

    deleteFolder($toWebDir);				//删除
    createFolder('/htmlweb/web');		//创建文件夹 防止web文件夹不存在20160504
    deleteFolder($webDir);
    createDirFolder($webDir);
    $webImages= $webDir . 'Images/';
    $webCss= $webDir . 'Css/';
    $webJs= $webDir . 'Js/';
    copyFolder($GLOBALS['cfg_webImages'], $webImages);
    copyFolder($GLOBALS['cfg_webCss'], $webCss);
    createFolder($webJs); //创建Js文件夹


    //处理Js文件夹
    $splJs= aspSplit(getDirJsList($webJs), vbCrlf());
    foreach( $splJs as $key=>$filePath){
        if( $filePath <> '' ){
            $toFilePath= $webJs . getFileName($filePath);
            ASPEcho('js', $filePath);
            moveFile($filePath, $toFilePath);
        }
    }
    //处理Css文件夹
    $splStr= aspSplit(getDirCssList($webCss), vbCrlf());
    foreach( $splStr as $key=>$filePath){
        if( $filePath <> '' ){
            $content= getftext($filePath);
            $content= Replace($content, $GLOBALS['cfg_webImages'], '../images/');

            $content= deleteCssNote($content);
            $content=phptrim($content);
            //设置为utf-8编码 20160527
            if( strtolower($setFileCode)=='utf-8' ){
                $content=replace($content,'gb2312','utf-8');
            }
            writeToFile($filePath, $content, $setFileCode);
            ASPEcho('css', $GLOBALS['cfg_webImages']);
        }
    }
    //复制栏目HTML
    $GLOBALS['isMakeHtml']= true;
    $rssObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn where isonhtml=true');
    while( $rss= $GLOBALS['conn']->fetch_array($rssObj)){
        $GLOBALS['glb_filePath']= Replace(getColumnUrl($rss['columnname'], 'name'), $GLOBALS['cfg_webSiteUrl'], '');
        if( Right($GLOBALS['glb_filePath'], 1)== '/' || Right($GLOBALS['glb_filePath'], 1)== '' ){
            $GLOBALS['glb_filePath']= $GLOBALS['glb_filePath'] . 'index.html';
        }
        if( Right($GLOBALS['glb_filePath'], 5)== '.html' ){
            if( Right($GLOBALS['glb_filePath'], 11)== '/index.html' ){
                $fileList= $fileList . $GLOBALS['glb_filePath'] . vbCrlf();
            }else{
                $fileList= $GLOBALS['glb_filePath'] . vbCrlf() . $fileList;
            }
            $fileName= Replace($GLOBALS['glb_filePath'], '/', '_');
            $toFilePath= $webDir . $fileName;
            copyfile($GLOBALS['glb_filePath'], $toFilePath);
            ASPEcho('导航', $GLOBALS['glb_filePath']);
        }
    }
    //复制文章HTML
    $rssObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'articledetail where isonhtml=true');
    while( $rss= $GLOBALS['conn']->fetch_array($rssObj)){
        $GLOBALS['glb_url']= getHandleRsUrl($rss['filename'], $rss['customaurl'], '/detail/detail' . $rss['id']);
        $GLOBALS['glb_filePath']= Replace($GLOBALS['glb_url'], $GLOBALS['cfg_webSiteUrl'], '');
        if( Right($GLOBALS['glb_filePath'], 1)== '/' || Right($GLOBALS['glb_filePath'], 1)== '' ){
            $GLOBALS['glb_filePath']= $GLOBALS['glb_filePath'] . 'index.html';
        }
        if( Right($GLOBALS['glb_filePath'], 5)== '.html' ){
            if( Right($GLOBALS['glb_filePath'], 11)== '/index.html' ){
                $fileList= $fileList . $GLOBALS['glb_filePath'] . vbCrlf();
            }else{
                $fileList= $GLOBALS['glb_filePath'] . vbCrlf() . $fileList;
            }
            $fileName= Replace($GLOBALS['glb_filePath'], '/', '_');
            $toFilePath= $webDir . $fileName;
            copyfile($GLOBALS['glb_filePath'], $toFilePath);
            ASPEcho('文章' . $rss['title'], $GLOBALS['glb_filePath']);
        }
    }
    //复制单面HTML
    $rssObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'onepage where isonhtml=true');
    while( $rss= $GLOBALS['conn']->fetch_array($rssObj)){
        $GLOBALS['glb_url']= getHandleRsUrl($rss['filename'], $rss['customaurl'], '/page/page' . $rss['id']);
        $GLOBALS['glb_filePath']= Replace($GLOBALS['glb_url'], $GLOBALS['cfg_webSiteUrl'], '');
        if( Right($GLOBALS['glb_filePath'], 1)== '/' || Right($GLOBALS['glb_filePath'], 1)== '' ){
            $GLOBALS['glb_filePath']= $GLOBALS['glb_filePath'] . 'index.html';
        }
        if( Right($GLOBALS['glb_filePath'], 5)== '.html' ){
            if( Right($GLOBALS['glb_filePath'], 11)== '/index.html' ){
                $fileList= $fileList . $GLOBALS['glb_filePath'] . vbCrlf();
            }else{
                $fileList= $GLOBALS['glb_filePath'] . vbCrlf() . $fileList;
            }
            $fileName= Replace($GLOBALS['glb_filePath'], '/', '_');
            $toFilePath= $webDir . $fileName;
            copyfile($GLOBALS['glb_filePath'], $toFilePath);
            ASPEcho('单页' . $rss['title'], $GLOBALS['glb_filePath']);
        }
    }
    //批量处理html文件列表
    //call echo(cfg_webSiteUrl,cfg_webTemplate)
    //call rwend(fileList)
    $sourceUrl=''; $replaceUrl ='';
    $splStr= aspSplit($fileList, vbCrlf());
    foreach( $splStr as $key=>$filePath){
        if( $filePath <> '' ){
            $filePath= $webDir . Replace($filePath, '/', '_');
            ASPEcho('filePath', $filePath);
            $content= getftext($filePath);

            foreach( $splStr as $key=>$s){
                $s1= $s;
                if( Right($s1, 11)== '/index.html' ){
                    $s1= substr($s1, 0 , Len($s1) - 11) . '/';
                }
                $sourceUrl= $GLOBALS['cfg_webSiteUrl'] . $s1;
                $replaceUrl= $GLOBALS['cfg_webSiteUrl'] . Replace($s, '/', '_');
                //Call echo(sourceUrl, replaceUrl) 							'屏蔽  否则大量显示20160613
                $content= Replace($content, $sourceUrl, $replaceUrl);
            }
            $content= Replace($content, $GLOBALS['cfg_webSiteUrl'], ''); //删除网址
            $content= Replace($content, $GLOBALS['cfg_webTemplate'], ''); //删除模板路径
            //content=nullLinkAddDefaultName(content)
            foreach( $splJs as $key=>$s){
                if( $s <> '' ){
                    $fileName= getFileName($s);
                    $content= Replace($content, 'Images/' . $fileName, 'js/' . $fileName);
                }
            }
            if( instr($content, '/Jquery/Jquery.Min.js') > 0 ){
                $content= Replace($content, '/Jquery/Jquery.Min.js', 'js/Jquery.Min.js');
                copyfile('/Jquery/Jquery.Min.js', $webJs . '/Jquery.Min.js');
            }
            $content= Replace($content, '<a href="" ', '<a href="index.html" '); //让首页加index.html
            createFileGBK($filePath, $content);
        }
    }

    //把复制网站夹下的images/文件夹下的js移到js/文件夹下  20160315
    $htmlFileList=''; $splHtmlFile=''; $splJsFile=''; $htmlFilePath=''; $jsFileName ='';
    $jsFileList= getDirJsNameList($webImages);
    $htmlFileList= getDirHtmlList($webDir);
    $splHtmlFile= aspSplit($htmlFileList, vbCrlf());
    $splJsFile= aspSplit($jsFileList, vbCrlf());
    foreach( $splHtmlFile as $key=>$htmlFilePath){
        $content= getftext($htmlFilePath);
        foreach( $splJsFile as $key=>$jsFileName){
            $content= regExp_Replace($content, 'Images/' . $jsFileName, 'js/' . $jsFileName);
        }

        $nErrLevel= 0;
        $content= handleHtmlFormatting($content, false, $nErrLevel, '|删除空行|'); //|删除空行|
        $content= handleCloseHtml($content, true, ''); //闭合标签
        $nErrLevel= checkHtmlFormatting($content);
        if( checkHtmlFormatting($content)== false ){
            eerr($htmlFilePath . '(格式化错误)', $nErrLevel);		 //注意
        }
        //设置为utf-8编码
        if( strtolower($setFileCode)=='utf-8' ){
            $content=replace($content,'<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />','<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />');
        }
        $content=phptrim($content);
        writeToFile($htmlFilePath, $content, $setFileCode);
    }
    //images下js移动到js下
    foreach( $splJsFile as $key=>$jsFileName){
        $jsFilePath=$webImages . $jsFileName;
        $content= getftext($jsFilePath);
        $content=phptrim($content);
        writeToFile($webJs . $jsFileName, $content, $setFileCode);
        deleteFile($jsFilePath);
    }

    copyFolder($webDir,$toWebDir);
    //使htmlWeb文件夹用php压缩
    if( @$_REQUEST['isMakeZip']=='1' ){
        makeHtmlWebToZip($webDir);
    }
    //使网站用xml打包20160612
    if( @$_REQUEST['isMakeXml']=='1' ){
        makeHtmlWebToXmlZip('/htmladmin/', $webFolderName);
    }
}

//使htmlWeb文件夹用php压缩
function makeHtmlWebToZip($webDir){
    $content=''; $splStr=''; $filePath=''; $c=''; $fileArray=''; $fileName=''; $fileType=''; $isTrue ='';
    $webFolderName ='';
    $cleanFileList ='';
    $splStr= aspSplit($webDir, '/');
    $webFolderName= $splStr[2];
    //call eerr(webFolderName,webDir)
    $content= getFileFolderList($webDir, true, '全部', '', '全部文件夹', '', '');
    $splStr= aspSplit($content, vbCrlf());
    foreach( $splStr as $key=>$filePath){
        if( checkfolder($filePath)== false ){
            $fileArray= handleFilePathArray($filePath);
            $fileName= strtolower($fileArray[2]);
            $fileType= strtolower($fileArray[4]);
            $fileName= remoteNumber($fileName);
            $isTrue= true;

            if( instr('|' . $cleanFileList . '|', '|' . $fileName . '|') > 0 && $fileType== 'html' ){
                $isTrue= false;
            }
            if( $isTrue== true ){
                //call echo(fileType,fileName)
                if( $c <> '' ){ $c= $c . '|' ;}
                $c= $c . Replace($filePath, handlePath('/'), '');
                $cleanFileList= $cleanFileList . $fileName . '|';
            }
        }
    }
    rw($c);
    $c= $c . '|||||';
    createFileGBK('htmlweb/1.txt', $c);
    ASPEcho('<hr>cccccccccccc', $c);
    //先判断这个文件存在20160309
    if( checkFile('/myZIP.php')== true ){
        ASPEcho('', XMLPost(getHost() . '/myZIP.php?webFolderName=' . $webFolderName , 'content=' . escape($c)));
    }

}
//使网站用xml打包20160612
function makeHtmlWebToXmlZip($newWebDir,$rootDir){
    $xmlFileName='';$xmlSize='';
    $xmlFileName= getIP() . '_update.xml';

    //newWebDir="\Templates2015\"
    //rootDir="\sharembweb\"

    $objXmlZIP =''; $objXmlZIP= new xmlZIP();
    $objXmlZIPrun(handlePath($newWebDir), handlePath($newWebDir . $rootDir), false, $xmlFileName);
    ASPEcho(handlePath($newWebDir), handlePath($newWebDir . $rootDir));
    $objXmlZIP= $GLOBALS['Nothing'];
    doevents();
    $xmlSize=getFSize($xmlFileName);
    $xmlSize=printSpaceValue($xmlSize);
    ASPEcho('下载xml打包文件', '<a href=/tools/downfile.asp?act=download&downfile=' . xorEnc('/' . $xmlFileName, 31380) . ' title=\'点击下载\'>点击下载' . $xmlFileName . '('. $xmlSize .')</a>');
}


//生成更新sitemap.xml 20160118
function saveSiteMap(){
    $isWebRunHtml ='';//是否为html方式显示网站
    $changefreg ='';//更新频率
    $priority ='';//优先级
    $c=''; $url ='';
    handlePower('修改生成SiteMap'); //管理权限处理

    $changefreg= @$_REQUEST['changefreg'];
    $priority= @$_REQUEST['priority'];
    loadWebConfig(); //加载配置
    //call eerr("cfg_flags",cfg_flags)
    if( instr($GLOBALS['cfg_flags'], '|htmlrun|') > 0 ){
        $isWebRunHtml= true;
    }else{
        $isWebRunHtml= false;
    }

    $c= $c . '<?xml version="1.0" encoding="UTF-8"?>' . vbCrlf();
    $c= $c . "\t" . '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">' . vbCrlf();

    //栏目
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn order by sortrank asc');
    while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
        if( $rsx['nofollow']== false ){
            $c= $c . copystr("\t", 2) . '<url>' . vbCrlf();
            if( $isWebRunHtml== true ){
                $url= getRsUrl($rsx['filename'], $rsx['customaurl'], '/nav' . $rsx['id']);
                $url=handleAction($url);
            }else{
                $url= escape('?act=nav&columnName=' . $rsx['columnname']);
            }
            $url= urlAddHttpUrl($GLOBALS['cfg_webSiteUrl'], $url);
            //call echo(cfg_webSiteUrl,url)

            $c= $c . copystr("\t", 3) . '<loc>' . $url . '</loc>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<lastmod>' . format_Time($rsx['updatetime'], 2) . '</lastmod>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<changefreq>' . $changefreg . '</changefreq>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<priority>' . $priority . '</priority>' . vbCrlf();
            $c= $c . copystr("\t", 2) . '</url>' . vbCrlf();
            ASPEcho('栏目', '<a href="' . $url . '" target=\'_blank\'>' . $url . '</a>');
        }
    }

    //文章
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'articledetail order by sortrank asc');
    while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
        if( $rsx['nofollow']== false ){
            $c= $c . copystr("\t", 2) . '<url>' . vbCrlf();
            if( $isWebRunHtml== true ){
                $url= getRsUrl($rsx['filename'], $rsx['customaurl'], '/detail/detail' . $rsx['id']);
                $url=handleAction($url);
            }else{
                $url= '?act=detail&id=' . $rsx['id'];
            }
            $url= urlAddHttpUrl($GLOBALS['cfg_webSiteUrl'], $url);
            //call echo(cfg_webSiteUrl,url)

            $c= $c . copystr("\t", 3) . '<loc>' . $url . '</loc>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<lastmod>' . format_Time($rsx['updatetime'], 2) . '</lastmod>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<changefreq>' . $changefreg . '</changefreq>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<priority>' . $priority . '</priority>' . vbCrlf();
            $c= $c . copystr("\t", 2) . '</url>' . vbCrlf();
            ASPEcho('文章', '<a href="' . $url . '" target=\'_blank\'>' . $url . '</a>');
        }
    }

    //单页
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'onepage order by sortrank asc');
    while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
        if( $rsx['nofollow']== false ){
            $c= $c . copystr("\t", 2) . '<url>' . vbCrlf();
            if( $isWebRunHtml== true ){
                $url= getRsUrl($rsx['filename'], $rsx['customaurl'], '/page/detail' . $rsx['id']);
                $url=handleAction($url);
            }else{
                $url= '?act=onepage&id=' . $rsx['id'];
            }
            $url= urlAddHttpUrl($GLOBALS['cfg_webSiteUrl'], $url);
            //call echo(cfg_webSiteUrl,url)

            $c= $c . copystr("\t", 3) . '<loc>' . $url . '</loc>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<lastmod>' . format_Time($rsx['updatetime'], 2) . '</lastmod>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<changefreq>' . $changefreg . '</changefreq>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<priority>' . $priority . '</priority>' . vbCrlf();
            $c= $c . copystr("\t", 2) . '</url>' . vbCrlf();
            ASPEcho('单页', '<a href="' . $url . '" target=\'_blank\'>' . $url . '</a>');
        }
    }


    $c= $c . "\t" . '</urlset>' . vbCrlf();

    loadWebConfig();

    createfile('sitemap.xml', $c);
    ASPEcho('生成sitemap.xml文件成功', '<a href=\'/sitemap.xml\' target=\'_blank\'>点击预览sitemap.xml</a>');

    //判断是否生成sitemap.html
    if( @$_REQUEST['issitemaphtml']== '1' ){
        $c= '';
        //第二种
        //栏目
        $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn order by sortrank asc');
        while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
            if( $rsx['nofollow']== false ){
                if( $isWebRunHtml== true ){
                    $url= getRsUrl($rsx['filename'], $rsx['customaurl'], '/nav' . $rsx['id']);
                    $url=handleAction($url);
                }else{
                    $url= escape('?act=nav&columnName=' . $rsx['columnname']);
                }
                $url= urlAddHttpUrl($GLOBALS['cfg_webSiteUrl'], $url);

                $c= $c . '<li style="width:20%;"><a href="' . $url . '">' . $rsx['columnname'] . '</a><ul>' . vbCrlf();

                //文章
                $rssObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'articledetail where parentId=' . $rsx['id'] . ' order by sortrank asc');
                while( $rss= $GLOBALS['conn']->fetch_array($rssObj)){
                    if( $rss['nofollow']== false ){
                        if( $isWebRunHtml== true ){
                            $url= getRsUrl($rss['filename'], $rss['customaurl'], '/detail/detail' . $rss['id']);
                            $url=handleAction($url);
                        }else{
                            $url= '?act=detail&id=' . $rss['id'];
                        }
                        $url= urlAddHttpUrl($GLOBALS['cfg_webSiteUrl'], $url);
                        $c= $c . '<li style="width:20%;"><a href="' . $url . '" target="_blank">' . $rss['title'] . '</a>' . vbCrlf();
                    }
                }
                $c= $c . '</ul></li>' . vbCrlf();


            }
        }

        //单面
        $c= $c . '<li style="width:20%;"><a href="javascript:;">单面列表</a><ul>' . vbCrlf();
        $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'onepage order by sortrank asc');
        while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
            if( $rsx['nofollow']== false ){
                $c= $c . copystr("\t", 2) . '<url>' . vbCrlf();
                if( $isWebRunHtml== true ){
                    $url= getRsUrl($rsx['filename'], $rsx['customaurl'], '/page/detail' . $rsx['id']);
                    $url=handleAction($url);
                }else{
                    $url= '?act=onepage&id=' . $rsx['id'];
                }
                $c= $c . '<li style="width:20%;"><a href="' . $url . '" target="_blank">' . $rsx['title'] . '</a>' . vbCrlf();
            }
        }
        $c= $c . '</ul></li>' . vbCrlf();

        $templateContent ='';
        $templateContent= getftext($GLOBALS['adminDir'] . '/template_SiteMap.html');


        $templateContent= Replace($templateContent, '{$content$}', $c);
        $templateContent= Replace($templateContent, '{$Web_Title$}', $GLOBALS['cfg_webTitle']);


        createfile('sitemap.html', $templateContent);
        ASPEcho('生成sitemap.html文件成功', '<a href=\'/sitemap.html\' target=\'_blank\'>点击预览sitemap.html</a>');
    }
    writeSystemLog('', '保存sitemap.xml'); //系统日志
}
?>


