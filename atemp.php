<?php 


define('WEBPATH', $_SERVER ['DOCUMENT_ROOT'].DIRECTORY_SEPARATOR);				//��վ��Ŀ¼
 

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
require_once './PHP2/Web/Inc/PinYin.php'; 

require_once './Web/setAccess.php';

require_once './Web/function.php';
require_once './Web/config.php';
  
//end ����inc

$ReadBlockList='';
$ModuleReplaceArray=''; //�滻ģ�����飬��ʱû�ã�����Ҫ���ţ�Ҫ��������  
 
 
//=========
?><?PHP


//����function2�ļ�����
function callFunction2(){
    switch ( @$_REQUEST['stype'] ){
        case 'runScanWebUrl' ; runScanWebUrl() ;break;//����ɨ����ַ
        case 'scanCheckDomain' ; scanCheckDomain() ;break;//���������Ч
        case 'bantchImportDomain' ; bantchImportDomain() ;break;//������������
        case 'scanDomainHomePage' ; scanDomainHomePage()								;break;//ɨ��������ҳ
        case 'isthroughTrue' ; isthroughTrue();											//�����ȫ��Ϊ��
        break;
        case 'function2test' ; function2test()											;break;//����
        default ; eerr('function2ҳ��û�ж���', @$_REQUEST['stype']);
    }
}

//����
function function2test(){
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webdomain where isdomain=true');
    ASPEcho('��', @mysql_num_rows($rsObj));
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        ASPEcho($rs['isdomain'],$rs['website']);
    }
}

//�����ȫ��Ϊ��
function isthroughTrue(){
    $GLOBALS['conn=']=OpenConn();
    connExecute('update ' . $GLOBALS['db_PREFIX'] . 'webdomain set isthrough=true');
    ASPEcho('�������', '<a href=\'?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����\'>OK</a>');
}

//ɨ��������ҳ
function scanDomainHomePage(){
    $url=''; $nSetTime=''; $isdomain=''; $htmlDir=''; $txtFilePath='';$homePageList='';$nThis='';
    $splstr='';$s='';$c='';$website='';$nState='';
    $isAsp='';$isAspx='';$isPhp='';$isJsp='';$c2='';
    $isAsp=0;$isAspx=0;$isPhp=0;$isJsp=0;
    $nThis=@$_REQUEST['nThis'];
    if( $nThis=='' ){
        $nThis=0;
    }else{
        $nThis=intval($nThis);
    }

    $nSetTime= 3;
    $GLOBALS['conn=']=OpenConn();
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webdomain where website<>\'\' and isthrough=true and isdomain=true');
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        $nThis=$nThis+1;
        ASPEcho($nThis, $rs['website']);
        doevents( );
        $htmlDir= '/../��վUrlScan/������ҳ/';
        createDirFolder($htmlDir);
        $txtFilePath= $htmlDir . '/' . setFileName($rs['website']) . '.txt';
        if( checkFile($txtFilePath)== true ){
            $c= phptrim(getFText($txtFilePath));
            $isAsp=getstrcut($c,'isAsp=',vbCrlf(),1);
            $isAspx=getstrcut($c,'isAspx=',vbCrlf(),1);
            $isPhp=getstrcut($c,'isPhp=',vbCrlf(),1);
            $isJsp=getstrcut($c,'isJsp=',vbCrlf(),1);
            ASPEcho('����', '����');
            $nSetTime=1;
        }else{
            $website=getwebsite($rs['website']);
            if( $website=='' ){
                eerr('����Ϊ��',$GLOBALS['httpurl']);
            }
            $splstr=array('index.asp','index.aspx','index.php','index.jsp','index.htm','index.html','default.asp','default.aspx','default.jsp','default.htm','default.html');
            $c2='';
            $homePageList='';
            foreach( $splstr as $key=>$s){
                $url=$website . $s;
                $nState=getHttpUrlState($url);
                ASPEcho($url,$nState . '   ('. getHttpUrlStateAbout($nState) .')');
                doevents();
                if( ($s=='index.asp' || $s=='default.asp') && ($nState=='200' || $nState=='302') ){
                    $isAsp=1;
                }else if( ($s=='index.aspx' || $s=='default.aspx') && ($nState=='200' || $nState=='302') ){
                    $isAspx=1;
                }else if( ($s=='index.php' || $s=='default.php') && ($nState=='200' || $nState=='302') ){
                    $isPhp=1;
                }else if( ($s=='index.jsp' || $s=='default.jsp') && ($nState=='200' || $nState=='302') ){
                    $isJsp=1;
                }
                if( $nState=='200' || $nState=='302' ){
                    $homePageList=$homePageList . $s . '|';
                }
                $c2=$c2 . $s . '=' . $nState . vbCrlf();
            }
            $c= 'isAsp=' . $isAsp . vbCrlf();
            $c= $c . 'isAspx=' . $isAspx . vbCrlf();
            $c= $c . 'isPhp=' . $isPhp . vbCrlf();
            $c= $c . 'isJsp=' . $isJsp . vbCrlf() . $c2;


            createFile($txtFilePath, $c);
            ASPEcho('����', '����');
        }
        //����д�Ǹ�תPHPʱ����
        connExecute('update ' . $GLOBALS['db_PREFIX'] . 'webdomain  set isasp='. $isAsp .',isaspx='. $isAspx .',isphp='. $isPhp .',isjsp='. $isJsp .',isthrough=false,homepagelist=\''. $homePageList .'\',updatetime=\'' . Now() . '\'  where id=' . $rs['id'] . '');



        $url= getUrlAddToParam(getThisUrl(), '?nThis='. $nThis .'&N=' . getRnd(11), 'replace');
        rw(jsTiming($url, $nSetTime));
        die();
    }
    ASPEcho('�������', '<a href=\'?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����\'>OK����('. $nThis .')��</a>');
}

//������������
function bantchImportDomain(){
    $content=''; $splStr=''; $url=''; $webSite=''; $nOK ='';
    $content= strtolower(@$_POST['bodycontent']);
    $splStr= aspSplit($content, vbCrlf());
    $nOK= 0;
    $GLOBALS['conn=']=OpenConn();
    foreach( $splStr as $key=>$url){
        $webSite= getwebsite($url);
        if( $webSite <> '' ){
            $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webdomain where website=\'' . $webSite . '\'');
            $rs=mysql_fetch_array($rsObj);
            if( @mysql_num_rows($rsObj)==0 ){
                connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'webdomain(website,isthrough,isdomain) values(\'' . $webSite . '\',true,false)');
                ASPEcho('��ӳɹ�', $webSite);
                $nOK= $nOK + 1;
            }else{
                ASPEcho('website', $webSite);
            }
        }
    }
    ASPEcho('�������', '<a href=\'?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����\'>OK ��(' . $nOK . ')��</a>');
}

//���������Ч
function scanCheckDomain(){
    $url=''; $nSetTime=''; $isdomain=''; $htmlDir=''; $txtFilePath=''; $nThis='';
    $nSetTime= 3;
    $nThis=@$_REQUEST['nThis'];
    if( $nThis=='' ){
        $nThis=0;
    }else{
        $nThis=intval($nThis);
    }
    $GLOBALS['conn=']=OpenConn();
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webdomain where isthrough=true');
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        $nThis=$nThis+1;
        ASPEcho($nThis, $rs['website']);
        doevents( );
        $htmlDir= '/../��վUrlScan/����/';
        createDirFolder($htmlDir);
        $txtFilePath= $htmlDir . '/' . setFileName($rs['website']) . '.txt';
        if( checkFile($txtFilePath)== true ){
            $isdomain= phptrim(getFText($txtFilePath));
            ASPEcho('����', '����');
            $nSetTime=1;
        }else{
            $isdomain= IIF(checkDomainName($rs['website']), 1, 0);
            createFile($txtFilePath, $isdomain);
            ASPEcho('����', '����');
        }
        //����д�Ǹ�תPHPʱ����
        connExecute('update ' . $GLOBALS['db_PREFIX'] . 'webdomain  set isthrough=false,isdomain=' . $isdomain . ',updatetime=\'' . Now() . '\'  where id=' . $rs['id'] . '');



        $url= getUrlAddToParam(getThisUrl(), '?nThis='. $nThis .'&N=' . getRnd(11), 'replace');
        rw(jsTiming($url, $nSetTime));
        die();
    }
    ASPEcho('�������', '<a href=\'?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����\'>OK����('. $nThis .')��</a>');
}

//ɨ����ַ
function runScanWebUrl(){
    $nSetTime=''; $setCharSet=''; $httpUrl=''; $url=''; $selectWeb ='';$nThis='';
    $setCharSet= 'gb2312'; //gb2312
    //http://www.dfz9.com/
    //http://www.maiside.net/
    //http://sharembweb.com/
    //http://www.ufoer.com/
    $httpUrl= 'http://sharembweb.com/';
    //selectWeb="ufoer"
    if( $selectWeb== 'ufoer' ){
        $httpUrl= 'http://www.ufoer.com/';
        $setCharSet= 'utf-8';
    }

    $nThis=@$_REQUEST['nThis'];
    if( $nThis=='' ){
        $nThis=0;
    }else{
        $nThis=intval($nThis);
    }

    $GLOBALS['conn=']=OpenConn();
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'weburlscan');
    $rs=mysql_fetch_array($rsObj);
    if( @mysql_num_rows($rsObj)==0 ){
        connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'weburlscan(httpurl,title,isthrough,charset) values(\'' . $httpUrl . '\',\'home\',true,\'' . $setCharSet . '\')');
    }
    //ѭ��
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'weburlscan where isThrough=true');
    $rsx=mysql_fetch_array($rsxObj);
    if( @mysql_num_rows($rsxObj)!=0 ){
        $nThis=$nThis+1;
        ASPEcho($nThis, $rsx['httpurl']);
        doevents( );
        $nSetTime= scanUrl($rsx['httpurl'], $rsx['title'], $rsx['charset']);
        //����д�Ǹ�תPHPʱ����
        connExecute('update ' . $GLOBALS['db_PREFIX'] . 'weburlscan  set isthrough=false  where id=' . $rsx['id'] . '');

        $url= getUrlAddToParam(getThisUrl(), '?nThis='. $nThis .'&N=' . getRnd(11), 'replace');

        rw(jsTiming($url, $nSetTime));
        die();
    }
    ASPEcho('�������', '<a href=\'?act=dispalyManageHandle&actionType=WebUrlScan&addsql=order by id desc&lableTitle=��ַɨ��\'>OK����('. $nThis .')��</a>');
    //���뱨��
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'weburlscan where webstate=404');
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        ASPEcho('<a href=\'' . $rs['httpurl'] . '\' target=\'_blank\'>' . $rs['httpurl'] . '</a>', '<a href=\'' . $rs['tohttpurl'] . '\' target=\'_blank\'>' . $rs['tohttpurl'] . '</a>');
    }
}
//ɨ����ַ
function scanUrl($httpUrl, $toTitle, $codeset){
    $splStr=''; $i=''; $s=''; $content=''; $PubAHrefList=''; $PubATitleList=''; $splUrl=''; $spltitle=''; $title=''; $url=''; $htmlDir=''; $htmlFilePath=''; $nOK=''; $dataArray=''; $webState=''; $u=''; $iniDir=''; $iniFilePath ='';$websize='';
    $nSetTime=''; $startTime=''; $openSpeed=''; $isLocal=''; $isThrough='';
    $htmlDir= '/../��վUrlScan/' . setFileName(getwebsite($httpUrl));
    createDirFolder($htmlDir);
    $htmlFilePath= $htmlDir . '/' . setFileName($httpUrl) . '.html';
    $iniDir= $htmlDir . '/conifg';
    createfolder($iniDir);
    $iniFilePath= $iniDir . '/' . setFileName($httpUrl) . '.txt';

    //httpurl="http://maiside.net/"

    $webState= 0;
    $nSetTime= 1;
    $openSpeed= 0;
    if( checkFile($htmlFilePath)== false ){
        $startTime= Now();
        ASPEcho('codeset', $codeset);
        $dataArray= handleXmlGet($httpUrl, $codeset);
        $content= $dataArray[0];
        $content=toGB2312Char($content); //��PHP�ã�ת��gb2312�ַ�

        $webState= $dataArray[1];
        $openSpeed= DateDiff('s', $startTime, Now());
        //content=gethttpurl(httpurl,codeset)
        //call createfile(htmlFilePath,content)
        writeToFile($htmlFilePath, $content, $codeset);
        createFile($iniFilePath, $webState . vbCrlf() . $openSpeed);
        $nSetTime= 3;
        $isLocal= 0;
    }else{
        //content=getftext(htmlFilePath)
        $content= reaFile($htmlFilePath, $codeset);
        $content=toGB2312Char($content); //��PHP�ã�ת��gb2312�ַ�
        $splStr= aspSplit(getftext($iniFilePath), vbCrlf());
        $webState= intval($splStr[0]);
        $openSpeed= intval($splStr[0]);
        $isLocal= 1;
    }
    $websize=getFSize($htmlFilePath);
    if( $websize=='' ){
        $websize=0;
    }
    ASPEcho('isLocal', $isLocal);
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'weburlscan where httpurl=\'' . $httpUrl . '\'');
    $rs=mysql_fetch_array($rsObj);
    if( @mysql_num_rows($rsObj)==0 ){
        connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'weburlscan(httpurl,title,charset) values(\'' . $httpUrl . '\',\'' . $toTitle . '\',\'' . $codeset . '\')');
    }
    connExecute('update ' . $GLOBALS['db_PREFIX'] . 'weburlscan  set webstate=' . $webState . ',websize=' . $websize . ',openspeed=' . $openSpeed . ',charset=\'' . $codeset . '\'  where httpurl=\'' . $httpUrl . '\'');

    //strLen(content)  �������������׼

    $s= getContentAHref('', $content, $PubAHrefList, $PubATitleList);
    $s= handleScanUrlList($httpUrl, $s);

    //call echo("httpurl",httpurl)
    //call echo("s",s)
    //call echo("PubATitleList",PubATitleList)
    $nOK= 0;
    $splUrl= aspSplit($PubAHrefList, vbCrlf());
    $spltitle= aspSplit($PubATitleList, vbCrlf());
    for( $i= 1 ; $i<= UBound($splUrl); $i++){
        $title= $spltitle[$i];
        $url= $splUrl[$i];
        //ȥ��#�ź�̨���ַ�20160506
        if( instr($url, '#') > 0 ){
            $url= mid($url, 1, instr($url, '#') - 1);
        }
        if( $url== '' ){
            if( $title <> '' ){
                ASPEcho('��ַΪ��', $title);
            }
        }else{
            $url= handleScanUrlList($httpUrl, $url);
            $url= handleWithWebSiteList($httpUrl, $url);
            if( $url <> '' ){
                $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'weburlscan where httpurl=\'' . $url . '\'');
                $rs=mysql_fetch_array($rsObj);
                if( @mysql_num_rows($rsObj)==0 ){
                    $u= strtolower($url);
                    if( instr($u, 'tools/downfile.asp') > 0 || instr($u, '/url.asp?') > 0 || instr($u, '/aspweb.asp?') > 0 || instr($u, '/phpweb.php?') > 0 || $u== 'http://www.maiside.net/qq/' || instr($u, 'mailto:') > 0 || instr($u, 'tel:') > 0 || instr($u, '.html?replytocom') > 0 ){//.html?replytocom  ��ͨ��վ
                        $isThrough= 0;
                    }else{
                        $isThrough= 1; //����true ��Ϊд�����ݻ�������
                    }
                    connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'weburlscan(tohttpurl,totitle,httpurl,title,isthrough,charset) values(\'' . $httpUrl . '\',\'' . $toTitle . '\',\'' . $url . '\',\'' . substr($title, 0 , 255) . '\',' . $isThrough . ',\'' . $codeset . '\')');
                    $nOK= $nOK + 1;
                    ASPEcho($i, $url);
                }else{
                    ASPEcho($title, $url);
                }
            }
        }
    }

    $scanUrl= $nSetTime;
    return @$scanUrl;
}


?>


