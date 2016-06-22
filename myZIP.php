<?php 
/*
require_once './PHP2/ImageWaterMark/Include/ASP.php';
require_once './PHP2/ImageWaterMark/Include/sys_FSO.php';
require_once './PHP2/ImageWaterMark/Include/sys_Url.php';
require_once './PHP2/ImageWaterMark/Include/testInc/Time.php';
*/

$content=unescape($_POST["content"]); 
$webFolderName=$_REQUEST["webFolderName"];
$zipFileName=$webFolderName . "_" . Format_Time(now(),12).".zip";			//zip�ļ�����

//zipĿ¼
if(isset($_REQUEST['zipDir'])){
	$zipDir=$_REQUEST['zipDir'];
}else{
	$zipDir="htmlweb/";
}
//zip�ļ�·��ɾ���ַ�
if(isset($_REQUEST['replaceZipFilePath'])){
	$replaceZipFilePath=unescape($_REQUEST['replaceZipFilePath']);
	echo('$replaceZipFilePath111111111=' . $replaceZipFilePath . '<hr>');
}else{
	$replaceZipFilePath="htmlweb\\". $webFolderName."\\";
}
//�Ƿ��ӡ������Ϣ
if(isset($_REQUEST['isPrintEchoMsg'])){
	$isPrintEchoMsg=unescape($_REQUEST['isPrintEchoMsg']);
}else{
	$isPrintEchoMsg=true;
}


$zipFilePath=$zipDir .$zipFileName;										//zip�ļ�·��
echo('<a href="'.$zipFilePath .'" target=_blank>�������('. $zipFileName .')</a>');
echo('<hr>�ڶ��������ļ�<a href="downfile.asp?downfile='.$zipFilePath .'" target=_blank>�������('. $zipFileName .')</a>');

$nCount=0;
fclose(fopen($zipFilePath,'w'));
$zip=new ZipArchive();
if($zip->open($zipFilePath,ZipArchive::OVERWRITE)===TRUE){
	//$zip->addFile('newFileName.asp');
	//$zip->addFile('aabc.txt');
	//$zip->addFile('�й���.txt');
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
					echo($nCount . '��'.$filePath .' ==>>   '. $newFilePath .'<hr>');
				}
				$zip->addFile($filePath,$newFilePath);
			}
		}
	}
		
	$zip->close();	 
	echo("ok");
}


//======================================������

//ת�ַ�
function Cstr($str){
	return (string)$str;
}
//ʱ�䴦��
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
//yyyy��mm��dd��
            $Format_Time = $Y . "��" . $M . "��" . $D . "��" ;break;
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
//yyyy��mm��dd��
            $Format_Time = $Y . "��" . $M . "��" . $D . "��" . " " . $H . ":" . $Mi . ":" . $S ;break;
        case 9;
//yyyy��mm��dd��Hʱmi��S�� ����
            $Format_Time = $Y . "��" . $M . "��" . $D . "��" . " " . $H . "ʱ" . $Mi . "��" . $S . "�룬" . GetDayStatus($H, 1) ;break;
        case 10;
//yyyy��mm��dd��Hʱ
            $Format_Time = $Y . "��" . $M . "��" . $D . "��" . $H . "ʱ" ;break;
        case 11;
//yyyy��mm��dd��Hʱmi��S��
            $Format_Time = $Y . "��" . $M . "��" . $D . "��" . " " . $H . "ʱ" . $Mi . "��" . $S . "��" ;break;
        case 12;
//yyyy��mm��dd��Hʱmi��
            $Format_Time = $Y . "��" . $M . "��" . $D . "��" . " " . $H . "ʱ" . $Mi . "��" ;break;
        case 13;
//yyyy��mm��dd��Hʱmi�� ����
            $Format_Time = $M . "��" . $D . "��" . " " . $H . ":" . $Mi . " " . GetDayStatus($H, 0) ;break;
        case 14;
//yyyy��mm��dd��
            $Format_Time = $Y . "/" . $M . "/" . $D ;break;
        case 15;
//yyyy��mm�� ��1��
            $Format_Time = $Y . "��" . $M . "�� ��" . GetCountPage($D,7) . "��";
     }
 return @$Format_Time;} 
//�ж�ʱ��
function isDate($timeStr){
	if(instr($timeStr,"-")>0 || instr($timeStr,"\/")>0 || instr($timeStr," ")>0){
		return true;
	}else{
		return false;
	}
}
//�����
function Year($timeStr){
	return getYMDHMS($timeStr,0);
}
//�����
function Month($timeStr){
	return getYMDHMS($timeStr,1);
}
//�����
function Day($timeStr){
	return getYMDHMS($timeStr,2);
}
//���ʱ
function Hour($timeStr){
	return getYMDHMS($timeStr,3);
}
//��÷�
function Minute($timeStr){
	return getYMDHMS($timeStr,4);
}
//�����
function Second($timeStr){
	return getYMDHMS($timeStr,5);
}
//�����ַ�����λ��
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
//���������ʱ�ֳ�
function getYMDHMS( $timeStr,$sType){
	 $splstr="";
	$timeStr=replace(replace(replace(replace(replace(replace($timeStr,"��","-"),"��","-"),"��","-"),"ʱ","-"),"��","-"),"��","-");
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
	if( $sType=="��" || $sType=="0" ){
		$getYMDHMS=$nYear;
	}else if( $sType=="��" || $sType=="1" ){
		$getYMDHMS=$nMonth;
	}else if( $sType=="��" || $sType=="2" ){
		$getYMDHMS=$nDay;
	}else if( $sType=="ʱ" || $sType=="3" ){
		$getYMDHMS=$nHour;
	}else if( $sType=="��" || $sType=="4" ){
		$getYMDHMS=$nMinute;
	}else if( $sType=="��" || $sType=="5" ){
		$getYMDHMS=$nSecond;
	 }

 return @$getYMDHMS;}
 //�ָ�
function aspSplit($contnet,$splStr){
	return explode($splStr,$contnet);
}
//�ַ�����
function len($content){
	return strlen($content);				//��������
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
//���ʱ��
function now(){
	$s=date('Y/m/d H:i:s');
	$s=replace($s,"/0","/");
	return $s;
}
//��ASP�� �滻����
function replace($c,$findStr,$replaceStr){
	return str_replace($findStr, $replaceStr, $c);
}
//����escape gb2312����
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
// ����ļ�
function checkFile($file) {
	return is_file ( $file );
}
?>
