<?PHP
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
require_once './Web/setAccess.php';
require_once './Web/function.php';
require_once './Web/config.php';
 

?> 


<?php
/* 
网页截图功能，必须安装IE+CutyCapt
url:要截图的网页
out：图片保存路径
path：CutyCapt路径
cmd:CutyCapt执行命令
比如:http://你php路径.php?url=http://niutuku.com/
*/
//$url=$_GET["url"];
$url="http://sharembweb.com/";
$imgname=str_replace('http://','',$url);
$imgname=str_replace('https://','',$imgname);
$imgname=str_replace('.','-',$imgname);
$out = 'E:/E盘/WEB网站/11.png';
$path = 'E:/E盘/WEB网站/至前网站/CutyCapt.exe';
$cmd = "$path --url=$url --out=$out";
echo $cmd;
system($cmd);
?>
 