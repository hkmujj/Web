<?PHP
//ϵͳ���ĳ���
require_once './PHP2/ImageWaterMark/Include/ASP.php';
require_once './PHP2/ImageWaterMark/Include/sys_FSO.php';
require_once './PHP2/ImageWaterMark/Include/sys_Url.php';
require_once './PHP2/ImageWaterMark/Include/sys_Cai.php';
require_once './PHP2/ImageWaterMark/Include/Conn.php';
require_once './PHP2/ImageWaterMark/Include/MySqlClass.php';
//����inc
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
require_once './Web/setAccess.php';
require_once './Web/function.php';
require_once './Web/config.php';
define('WEBPATH', $_SERVER ['DOCUMENT_ROOT'].DIRECTORY_SEPARATOR);				//��վ��Ŀ¼

$db_PREFIX='xy_';

switch ( @$_REQUEST['act'] ){ 
    case 'checkDomainName' ;aspEcho(@$_REQUEST['httpurl'],checkDomainName(@$_REQUEST['httpurl'])) ;break;//��������Ƿ����
    case 'displayEvalList' ;displayEvalList();break;//��ַ�б�
    default ;displayDefault(); //��ʾĬ��
}

//��ʾĬ�� 
function displayDefault(){
	aspecho('��������Ƿ����','<a href="?act=checkDomainName&httpurl=http://127.0.0.1/4.asp">(http://127.0.0.1/4.asp)</a>');
	aspecho('��������Ƿ����','<a href="?act=checkDomainName&httpurl=http://www.baidu.com/">(http://www.baidu.com/)</a>');
	aspecho('��������Ƿ����','<a href="?act=checkDomainName&httpurl=http://www.wzl99.com/">(http://www.wzl99.com/)</a>');
	aspecho('��ַ�б�','<a href="?act=displayEvalList">��ַ�б�</a>');
	
	
	
}
//ev��ַ�б�
function displayEvalList(){
	$filePath=handlePath('\VB����\Template\ModuleTemplate\��ѧ����\ev\tempurl.txt');
	//aspecho('$filePath',$filePath);
	$content=getftext($filePath);
	$splstr=aspSplit($content,vbCrlf());
	$urlList='';$webSiteList='';
	foreach( $splstr as $key=>$url){
		if($url!='' && instr(vbCrlf() . $urlList . vbCrlf(),vbCrlf() . $url . vbCrlf())==false ){
			aspecho($url,$url);
			$urlList+=$url.vbCrlf();
			doevents();
			$website=getWebSite($url);
			if($website!='' && instr(vbCrlf() . $webSiteList . vbCrlf(),vbCrlf() . $website . vbCrlf())==false ){			
				$webSiteList.=$website.vbCrlf();
				
				
							
				$GLOBALS['conn=']=OpenConn();
				$rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webdomain where website=\''. $website .'\'');
				if( @mysql_num_rows($rsObj)==0 ){
					//����д�Ǹ�תPHPʱ����
					//connExecute('update ' . $GLOBALS['db_PREFIX'] . 'webdomain  set isthrough=true,website=\''. $website .'\'');
					connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'webdomain(isdomain,isthrough,website) values(false,true,\''.$website.'\')');
					aspecho("д�����ݿ�ɹ�", $website);
			
				}
						 
				
			}
			
		}
	}
	
	echo('='.$webSiteList);
}



















//
/*
$arr;
$arr=handleXmlGet("http://sharembweb.com/");
aspecho('url',$arr[0]);
//aspecho('url',$arr);
*/

/*
print_r($_POST);
	foreach( $_POST as $key=>$s){
	aspecho('key',$key);
}
*/  























/*
$httpurl='http://www.ufoer.com/';
aspecho($httpurl,getHttpUrlState($httpurl));
$c=get_url_content($httpurl);
$c=toGB2312Char($c);
echo($c);
*/
?> 
