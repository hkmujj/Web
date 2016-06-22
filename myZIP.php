<?php 
/*
require_once './PHP2/ImageWaterMark/Include/ASP.php';
require_once './PHP2/ImageWaterMark/Include/sys_FSO.php';
require_once './PHP2/ImageWaterMark/Include/sys_Url.php';
require_once './PHP2/ImageWaterMark/Include/testInc/Time.php';
*/

$content=unescape($_POST["content"]); 
$webFolderName=$_REQUEST["webFolderName"];
$zipFileName=$webFolderName . "_" . Format_Time(now(),12).".zip";			//zip文件名称

//zip目录
if(isset($_REQUEST['zipDir'])){
	$zipDir=$_REQUEST['zipDir'];
}else{
	$zipDir="htmlweb/";
}
//zip文件路径删除字符
if(isset($_REQUEST['replaceZipFilePath'])){
	$replaceZipFilePath=unescape($_REQUEST['replaceZipFilePath']);
	echo('$replaceZipFilePath111111111=' . $replaceZipFilePath . '<hr>');
}else{
	$replaceZipFilePath="htmlweb\\". $webFolderName."\\";
}
//是否打印回显信息
if(isset($_REQUEST['isPrintEchoMsg'])){
	$isPrintEchoMsg=unescape($_REQUEST['isPrintEchoMsg']);
}else{
	$isPrintEchoMsg=true;
}


$zipFilePath=$zipDir .$zipFileName;										//zip文件路径
echo('<a href="'.$zipFilePath .'" target=_blank>点击下载('. $zipFileName .')</a>');
echo('<hr>第二种下载文件<a href="downfile.asp?downfile='.$zipFilePath .'" target=_blank>点击下载('. $zipFileName .')</a>');

$nCount=0;
fclose(fopen($zipFilePath,'w'));
$zip=new ZipArchive();
if($zip->open($zipFilePath,ZipArchive::OVERWRITE)===TRUE){
	//$zip->addFile('newFileName.asp');
	//$zip->addFile('aabc.txt');
	//$zip->addFile('中国人.txt');
	//$zip->addFile('ev/Md5.Asp');
	//$zip->addFile('ttemp');
	$splstr=aspSplit($content,"|");
	foreach( $splstr as $filePath){
		if($filePath!=""){
			if(checkFile($filePath)==true){
				$nCount++;
				$newFilePath=replace($filePath, 'htmladmin\\', "");
				$newFilePath=replace($newFilePath, $replaceZipFilePath, "");
				//echo($filePath.','.$replaceZipFilePath.'=============='.$newFilePath.'<br>');				
				if($isPrintEchoMsg==true){		 	 
					echo($nCount . '、'.$filePath .' ==>>   '. $newFilePath .'<hr>');
				}
				$zip->addFile($filePath,$newFilePath);
			}
		}
	}
		
	$zip->close();	 
	echo("ok");
}


//======================================函数区

//转字符
function Cstr($str){
	return (string)$str;
}
//时间处理
function  Format_Time($s_Time, $n_Flag){
     $Y=""; $M=""; $D=""; $H=""; $Mi=""; $S ="";
    $Format_Time = "" ;
    if( IsDate($s_Time) == false ){  return @$Format_Time; ;}
    $Y = Cstr(Year($s_Time)) ;
    $M = Cstr(Month($s_Time)) ;
    if( strlen($M) == 1 ){ $M = "0" . $M ;}
    $D = Cstr(Day($s_Time)) ;
    if( strlen($D) == 1 ){ $D = "0" . $D ;}
    $H = Cstr(Hour($s_Time)) ;
    if( strlen($H) == 1 ){ $H = "0" . $H ;}
    $Mi = Cstr(Minute($s_Time)) ;
    if( strlen($Mi) == 1 ){ $Mi = "0" . $Mi ;}
    $S = Cstr(Second($s_Time)) ;
    if( strlen($S) == 1 ){ $S = "0" . $S ;}
     switch ( $n_Flag ){
        case 1;
//yyyy-mm-dd hh:mm:ss
            $Format_Time = $Y . "-" . $M . "-" . $D . " " . $H . ":" . $Mi . ":" . $S ;break;
        case 2;
//yyyy-mm-dd
            $Format_Time = $Y . "-" . $M . "-" . $D ;break;
        case 3;
//hh:mm:ss
            $Format_Time = $H . ":" . $Mi . ":" . $S ;break;
        case 4;
//yyyy年mm月dd日
            $Format_Time = $Y . "年" . $M . "月" . $D . "日" ;break;
        case 5;
//yyyymmdd
            $Format_Time = $Y . $M . $D ;break;
        case 6;
//yyyymmddhhmmss
            $Format_Time = $Y . $M . $D . $H . $Mi . $S ;break;
        case 7;
//mm-dd
            $Format_Time = $M . "-" . $D ;break;
        case 8;
//yyyy年mm月dd日
            $Format_Time = $Y . "年" . $M . "月" . $D . "日" . " " . $H . ":" . $Mi . ":" . $S ;break;
        case 9;
//yyyy年mm月dd日H时mi分S秒 早上
            $Format_Time = $Y . "年" . $M . "月" . $D . "日" . " " . $H . "时" . $Mi . "分" . $S . "秒，" . GetDayStatus($H, 1) ;break;
        case 10;
//yyyy年mm月dd日H时
            $Format_Time = $Y . "年" . $M . "月" . $D . "日" . $H . "时" ;break;
        case 11;
//yyyy年mm月dd日H时mi分S秒
            $Format_Time = $Y . "年" . $M . "月" . $D . "日" . " " . $H . "时" . $Mi . "分" . $S . "秒" ;break;
        case 12;
//yyyy年mm月dd日H时mi分
            $Format_Time = $Y . "年" . $M . "月" . $D . "日" . " " . $H . "时" . $Mi . "分" ;break;
        case 13;
//yyyy年mm月dd日H时mi分 早上
            $Format_Time = $M . "月" . $D . "日" . " " . $H . ":" . $Mi . " " . GetDayStatus($H, 0) ;break;
        case 14;
//yyyy年mm月dd日
            $Format_Time = $Y . "/" . $M . "/" . $D ;break;
        case 15;
//yyyy年mm月 第1周
            $Format_Time = $Y . "年" . $M . "月 第" . GetCountPage($D,7) . "周";
     }
 return @$Format_Time;} 
//判断时间
function isDate($timeStr){
	if(instr($timeStr,"-")>0 || instr($timeStr,"\/")>0 || instr($timeStr," ")>0){
		return true;
	}else{
		return false;
	}
}
//获得年
function Year($timeStr){
	return getYMDHMS($timeStr,0);
}
//获得月
function Month($timeStr){
	return getYMDHMS($timeStr,1);
}
//获得日
function Day($timeStr){
	return getYMDHMS($timeStr,2);
}
//获得时
function Hour($timeStr){
	return getYMDHMS($timeStr,3);
}
//获得分
function Minute($timeStr){
	return getYMDHMS($timeStr,4);
}
//获得秒
function Second($timeStr){
	return getYMDHMS($timeStr,5);
}
//查找字符所在位置
function InStr($content,$search){
	if( $search!=""){
		if(strstr($content,$search)){
			return strpos($content,$search)+1;
		}else{
			return 0;
		}
	}else{
		return 0;
	}
}
//获得年月日时分钞
function getYMDHMS( $timeStr,$sType){
	 $splstr="";
	$timeStr=replace(replace(replace(replace(replace(replace($timeStr,"年","-"),"月","-"),"日","-"),"时","-"),"分","-"),"秒","-");
	$timeStr=replace(replace(replace(replace(replace($timeStr," ","-"),":","-"),"/","-"),"--","-"),"--","-") . "------";
	$splstr=aspSplit($timeStr,"-");
	$nYear = $splstr[0];
	$nMonth = $splstr[1];
	$nDay = $splstr[2];
	$nHour = $splstr[3];
	$nMinute = $splstr[4];
	$nSecond = $splstr[5];
	if( len($nYear)==1 ){ $nYear="0" . $nYear;}
	if( len($nMonth)==1 ){ $nMonth="0" . $nMonth;}
	if( len($nDay)==1 ){ $nDay="0" . $nDay;}
	if( len($nHour)==1 ){ $nHour="0" . $nHour;}
	if( len($nMinute)==1 ){ $nMinute="0" . $nMinute;}
	if( len($nSecond)==1 ){ $nSecond="0" . $nSecond ;}

	if( $nHour=="" ){ $nHour="00";}
	if( $nMinute=="" ){ $nMinute="00";}
	if( $nSecond=="" ){ $nSecond="00";}

	$sType=CStr($sType);
	if( $sType=="年" || $sType=="0" ){
		$getYMDHMS=$nYear;
	}else if( $sType=="月" || $sType=="1" ){
		$getYMDHMS=$nMonth;
	}else if( $sType=="日" || $sType=="2" ){
		$getYMDHMS=$nDay;
	}else if( $sType=="时" || $sType=="3" ){
		$getYMDHMS=$nHour;
	}else if( $sType=="分" || $sType=="4" ){
		$getYMDHMS=$nMinute;
	}else if( $sType=="秒" || $sType=="5" ){
		$getYMDHMS=$nSecond;
	 }

 return @$getYMDHMS;}
 //分割
function aspSplit($contnet,$splStr){
	return explode($splStr,$contnet);
}
//字符长度
function len($content){
	return strlen($content);				//采用这种
	//return strlen($content);
	//	return mb_strlen($content,'gb2312');
	$split = 1;	
	$n = 0;
	$array = array ();
	//echo (strlen ( $content ) . "<hr>");
	for($i = 0; $i < strlen ( $content );) {
		$value = ord ( $content [$i] );
		if ($value > 127) {
			if ($value >= 192 && $value <= 223)
				$split = 2;
			elseif ($value >= 224 && $value <= 239)
				$split = 3;
			elseif ($value >= 240 && $value <= 247)
				$split = 4;
		} else {
			$split = 1;
		}
		$key = NULL;
		for($j = 0; $j < $split; $j ++, $i ++) {
			$key .= $content [$i];
			$n ++;
		} 
		array_push ( $array, $key );
	} 
	return Count ( $array );
}
//获得时间
function now(){
	$s=date('Y/m/d H:i:s');
	$s=replace($s,"/0","/");
	return $s;
}
//给ASP用 替换内容
function replace($c,$findStr,$replaceStr){
	return str_replace($findStr, $replaceStr, $c);
}
//解密escape gb2312编码
function unEscape($str) {
    $str = rawurldecode($str);
    preg_match_all("/%u.{4}|&#x.{4};|&#d+;|.+/U",$str,$r);
    $ar = $r[0];
    foreach($ar as $k=>$v) {
        if(substr($v,0,2) == "%u")
        $ar[$k] = iconv("UCS-2","GBK",pack("H4",substr($v,-4)));
        elseif(substr($v,0,3) == "&#x")
        $ar[$k] = iconv("UCS-2","GBK",pack("H4",substr($v,3,-1)));
        elseif(substr($v,0,2) == "&#") {
            $ar[$k] = iconv("UCS-2","GBK",pack("n",substr($v,2,-1)));
        }
    }
    return join("",$ar);
}
// 检测文件
function checkFile($file) {
	return is_file ( $file );
}
?>
