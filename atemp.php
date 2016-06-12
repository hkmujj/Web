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

switch ( @$_REQUEST['act'] ){
    case 'scanurl' ; scanHttpUrl(@$_REQUEST['url']);break;
    case 'displayUrlList' ; displayUrlList();break;
    case 'displayLisks' ; displayLisks();		//��ʾ����
    break;
    case 'testUrl' ; testUrl();break;
    default ; displayDefaultLayout();
}

//��ʾĬ�����
function displayDefaultLayout(){
    ASPEcho('��ʾURL�б�','<a href=\'?act=displayUrlList\'>�������</a>');
    ASPEcho('�����б�','<a href=\'?act=displayLisks\'>�������</a>');
    ASPEcho('������ַ','<a href=\'?act=testUrl\'>�������</a>');
}

//��÷�������ַ
function getServerUrl( $nIndex){
    $serverName ='';
    $serverName= array('bb', 'cc', 'dd', 'ee', 'ff', 'gg','z1','z2','z3','z4','z5','z6','');
    $nIndex= $nIndex % UBound($serverName);
    $getServerUrl= 'http://' . $serverName[$nIndex] . '/atemp.' . EDITORTYPE . '?act=scanurl';
    return @$getServerUrl;
}

//��ʾ����
function displayLisks(){
    $i=''; $url=''; $s ='';$urlList='';$c='';
    $GLOBALS['conn=']=OpenConn();
    $rsObj=$GLOBALS['conn']->query( 'select * from xy_webdomain where links<>\'\'');
    ASPEcho('���м�¼', @mysql_num_rows($rsObj));

    deletefile('1.txt');
    $rsObj=$GLOBALS['conn']->query( 'select top 130999 * from xy_webdomain where links<>\'\' order by nlinks desc');
    for( $i= 1 ; $i<= @mysql_num_rows($rsObj); $i++){
        $rs=mysql_fetch_array($rsObj);
        ASPEcho($rs['nlinks'], '<a href=\'' . $rs['website'] . '\' target=\'_blank\'>' . $rs['website'] . '</a>');
        //		call rwend(rs("links"))
        createAddFile2('1.txt',handleScanUrl($rs['links']));
    }
}
//����ɨ����ַ�б�
function handleScanUrl($urlList){
    $splstr='';$url='';$c='';
    $splstr=aspSplit($urlList,vbCrlf());
    foreach( $splstr as $key=>$url){
        //call echo(url,checkScanUrl(url))
        if( instr(vbCrlf() . $c . vbCrlf(), vbCrlf() . $url . vbCrlf())==false && checkScanUrl($url)==true ){
            $c=$c . $url . vbCrlf();
        }
    }
    $handleScanUrl=$c;
    return @$handleScanUrl;
}
//����Ƿ�Ϊ��ɨ����ַ
function checkScanUrl($url){
    $website='';$c='';$splstr='';$s='';
    $website=strtolower(getwebsite($url));
    $website=replace($website,'/','.');		//���⴦��
    $c='alipay|sina|qq|google|baidu|so|bing|sogou|yahoo|haosou|youdao|163|360|39|ifeng|hao123|51|letv|sohu|pptv|taobao|tv|renren|gov|github|admin5|chinaz';
    $c=$c . '|cnzz|mycodes|hicode|asp300|662p|codesky|php|asp|jsp|thinkphp|dedecms|jb51|codefans|aspjzy|ip138|';

    $splstr=aspSplit($c,'|');
    foreach( $splstr as $key=>$s){
        if( $s <>'' ){
            if( instr($website,'.' . $s . '.com.')>0 || instr($website,'.' . $s . '.cn.')>0 || instr($website,'.' . $s . '.net.')>0 || instr($website,'.' . $s . '.la.')>0 ){
                $checkScanUrl=false;
                return @$checkScanUrl;
            }
        }
    }
    $checkScanUrl=true;
    return @$checkScanUrl;
}

//��ʾ��ַ�б�
function displayUrlList(){
    $i=''; $url=''; $s ='';$urlList='';$c='';
    $GLOBALS['conn=']=OpenConn();
    $rsObj=$GLOBALS['conn']->query( 'select * from xy_webdomain where isthrough<>0');
    ASPEcho('���м�¼', @mysql_num_rows($rsObj));

    $rsObj=$GLOBALS['conn']->query( 'select top 130999 * from xy_webdomain where isthrough<>0');
    for( $i= 1 ; $i<= @mysql_num_rows($rsObj); $i++){
        $rs=mysql_fetch_array($rsObj);
        $url= getServerUrl($i) . '&url=' . $rs['website'];

        if( $c<>'' ){ $c=$c . ',';}
        $c=$c . '\''. $url .'\'';
        //s = "<iframe src='" & url & "' height='100' width='100%' frameborder='1' scrolling='yes'></iframe>" & vbcrlf
        //Call rw(s)
        //Call echo(url, rs("website"))
    }
    rw(batchJsHandle($c,12));
}

//������ַ
function testUrl(){
    $httpurl ='';
    $httpurl= 'http://127.0.0.1/1.html';
    $httpurl= 'http://sharembweb.com';
    scanHttpUrl($httpurl);
}
//ɨ����ַ
function scanHttpUrl($httpurl){
    $rootDir=''; $webSite=''; $folderDir=''; $configFile=''; $configContent=''; $webstate='';$tempWebstate=''; $msgStr=''; $content=''; $splStr=''; $s=''; $url ='';
    $htmlFile=''; $websize=''; $webtitle=''; $webkeywords=''; $webdescription=''; $isdomain=''; $startTime=''; $openspeed='';$homepagelist='';$lists='';$PubAHrefList=''; $PubATitleList='';
    $isasp='';$isaspx='';$isphp='';$isjsp='';$ishtm='';$ishtml='';$links='';$nlinks='';
    $isasp=0;$isaspx=0;$isphp=0;$isjsp=0;$ishtm=0;$ishtml=0;
    $rootDir= '/../��վUrlScan/httpurl2016/';
    createDirFolder($rootDir);

    $webSite= getWebSite($httpurl);
    ASPEcho('��ַ', $httpurl);
    $s= '<a href=\''. $webSite .'\' target=\'_blank\'>' . $webSite . '</a>';
    ASPEcho('����', $s);
    $folderDir= $rootDir . setFileName($webSite) . '/' . setFileName($httpurl) . '/';
    createDirFolder($folderDir);
    $configFile= $folderDir . 'config.txt';

    $configContent= vbCrlf() . getftext($configFile) . vbCrlf();

    $isdomain= getStrCut($configContent, vbCrlf() . 'isdomain=', vbCrlf(), 2);
    if( $isdomain== '' ){
        $isdomain= IIF(checkDomainName($webSite), '1', '0');
        createAddFile2($configFile, 'isdomain=' . $isdomain . vbCrlf());
    }
    ASPEcho('������Ч��', $isdomain);
    if( $isdomain== '0' ){
        $GLOBALS['conn=']=OpenConn();
        connexecute('update xy_webdomain set isthrough=0 where website=\''. $httpurl .'\'');
        eerr('������Ч', $webSite);
    }

    $webstate= getStrCut($configContent, vbCrlf() . 'webstate=', vbCrlf(), 2);
    $openspeed= getStrCut($configContent, vbCrlf() . 'openspeed=', vbCrlf(), 2);
    $msgStr= '���ػ���';
    if( $webstate== '' ){
        $startTime= Now();
        $webstate= getHttpUrlState($httpurl);
        $openspeed= DateDiff('s', $startTime, Now());
        createAddFile2($configFile, 'webstate=' . $webstate . vbCrlf());
        createAddFile2($configFile, 'openspeed=' . $openspeed . vbCrlf());
        $msgStr= '��ȡ����';
    }
    ASPEcho('��״̬��' . $msgStr, $webstate . '(' . getHttpUrlStateAbout($webstate) . ')');
    ASPEcho('openspeed', $openspeed);

    $htmlFile= $folderDir . setFileName($httpurl) . '.html';
    if( checkFile($htmlFile)== false ){
        $content= getHttpUrl($httpurl, '');
        $content=toGB2312Char($content); //��PHP�ã�ת��gb2312�ַ�
        if( $content=='' ){ $content='��Ϊ�ա�';}
        createFile($htmlFile, $content);
    }
    $content= getFText($htmlFile);
    $websize= getFSize($htmlFile);

    $webtitle= getHtmlValue($content, 'webtitle');
    $webkeywords= getHtmlValue($content, 'webkeywords');
    $webdescription= getHtmlValue($content, 'webdescription');
    ASPEcho('webtitle', $webtitle);
    ASPEcho('webkeywords', $webkeywords);
    ASPEcho('webdescription', $webdescription);
    ASPEcho('��ҳ��С', printSpaceSize($websize));

    $nlinks=0;
    $lists= getContentAHref('', $content, $PubAHrefList, $PubATitleList);
    $lists= handleDifferenceWebSiteList($httpurl, $lists);		//��ò�ͬ��ַ�б�
    $splstr=aspSplit($lists,vbCrlf());
    foreach( $splstr		 as $key=>$url){
        $url=getwebsite($url);
        if( $url<>'' && instr(vbCrlf() . $links . vbCrlf(),vbCrlf() . $lists . vbCrlf())==false ){
            $links=$links . $url . vbCrlf();
            $nlinks=$nlinks+1;
        }
    }
    $links=ADSql($links);
    ASPEcho($nlinks,$links);

    //��ҳ״̬
    $splStr= array('index.asp', 'index.aspx', 'index.php', 'index.jsp', 'index.htm', 'index.html', 'default.asp', 'default.aspx', 'default.jsp', 'default.htm', 'default.html');
    foreach( $splStr as $key=>$s){
        $url= $webSite . $s;
        $tempWebstate= getStrCut($configContent, $s . '=', vbCrlf(), 2);
        if( $tempWebstate== '' ){
            $tempWebstate= getHttpUrlState($url);
            createAddFile2($configFile, $s . '=' . $tempWebstate . vbCrlf());
        }
        if( ($s=='index.asp' || $s=='default.asp') && ($tempWebstate=='200' || $tempWebstate=='301') ){
            $isasp=1;
        }
        if( ($s=='index.aspx' || $s=='default.aspx') && ($tempWebstate=='200' || $tempWebstate=='301') ){
            $isaspx=1;
        }
        if( ($s=='index.php' || $s=='default.php') && ($tempWebstate=='200' || $tempWebstate=='301') ){
            $isphp=1;
        }
        if( ($s=='index.jsp' || $s=='default.jsp') && ($tempWebstate=='200' || $tempWebstate=='301') ){
            $isjsp=1;
        }
        if( ($s=='index.htm' || $s=='default.htm') && ($tempWebstate=='200' || $tempWebstate=='301') ){
            $ishtm=1;
        }
        if( ($s=='index.html' || $s=='default.html') && ($tempWebstate=='200' || $tempWebstate=='301') ){
            $ishtml=1;
        }
        if( $tempWebstate=='200' ){
            $homepagelist=$homepagelist . $s . '|';
        }
        ASPEcho($s, $tempWebstate);
    }
    $GLOBALS['conn=']=OpenConn();
    $s=',homepagelist=\''. $homepagelist .'\',isasp='. $isasp .',isaspx='. $isaspx .',isphp='. $isphp .',isjsp='. $isjsp .',ishtm='. $ishtm .',ishtml='. $ishtml .',links=\''. $links .'\',nlinks='. $nlinks .'';
    connexecute('update xy_webdomain set isthrough=0,webstate='. $webstate .',openspeed='. $openspeed .',websize='. $websize .',isdomain='. $isdomain . $s .' where website=\''. $httpurl .'\'');
}

function batchJsHandle($urlList,$nThread){
    ?>
    <script type="text/javascript" src="/Jquery/Jquery.Min.js"></script>
    </head>
    <body>
    <div id="msg"></div>
    <?PHP
    $i='';$s='';$c='';
    for( $i=1 ; $i<= $nThread; $i++){
        rw('<iframe src=\'\' id=\'iframe'. $i .'\' height=\'100\' width=\'100%\' frameborder=\'1\' scrolling=\'yes\'></iframe>' . vbCrlf());
        if( $i>1 ){
            $s='index++' . vbCrlf() . 'callIframe("iframe'. $i .'",urlArray[index],callback)';
        }
        $c=$c . $s . vbCrlf();
    }
    ?>
    <script>
    var urlArray=Array(<?=$urlList?>);
    var index=0
    var nRun=0
    function callIframe(id,url,callback) {
        $('iframe#'+id).attr('src', url);
        $('iframe#'+id).load(function(){
            callback(id,this);
        });
    }
    function callback(id,e){
        if(index<urlArray.length-1){
            //alert(index + '<' + (urlArray.length-1))
            index++

            callIframe(id,urlArray[index],callback)
        }
        nRun++
        $("#msg").html( index + '/' + (urlArray.length-1) + ",���д�����"+ nRun +"��" )

        if(index >=(urlArray.length-1) && urlArray.length>0){
            $("#msg").html( $("#msg").html() + "���ö�ʱ����8 ���ˢ�µ�ǰҳ��")
            var picTimer = setInterval(function() {
                location.reload()
                clearInterval(picTimer);

            },8000); //��4000�����Զ����ŵļ������λ������
        }

    }
    callIframe("iframe1",urlArray[index],callback)
    <?=$c?>
    </script>
    <? }?>


