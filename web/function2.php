<?PHP


//����function2�ļ�����
function callFunction2(){
    switch ( @$_REQUEST['stype'] ){
        case 'runScanWebUrl' ; runScanWebUrl() ;break;//����ɨ����ַ
        case 'scanCheckDomain' ; scanCheckDomain() ;break;//���������Ч
        case 'bantchImportDomain' ; bantchImportDomain() ;break;//������������
        case 'scanDomainHomePage' ; scanDomainHomePage()								;break;//ɨ��������ҳ
        case 'scanDomainHomePageSize' ; scanDomainHomePageSize()								;break;//ɨ��������ҳ��С�����
        case 'isthroughTrue' ; isthroughTrue()											;break;//�����ȫ��Ϊ��
        case 'printOKWebSite' ; printOKWebSite()										;break;//��ӡ��Ч��ַ
        case 'printAspServerWebSite' ; printAspServerWebSite();										//��ӡasp������վ
        break;
        case 'clearAllData' ; clearAllData();										//���ȫ������
        break;
        case 'function2test' ; function2test()											;break;//����
        default ; eerr('function2ҳ��û�ж���', @$_REQUEST['stype']);
    }
}

//����
function function2test(){
    $GLOBALS['conn=']=OpenConn();
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webdomain where isdomain=true');
    ASPEcho('��', @mysql_num_rows($rsObj));
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        ASPEcho($rs['isdomain'],$rs['website']);
    }
}
//���ȫ������
function clearAllData(){
    $GLOBALS['conn=']=OpenConn();
    connExecute('delete from ' . $GLOBALS['db_PREFIX'] . 'webdomain');
    ASPEcho('�������', '<a href=\'?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����\'>OK</a>');
}
//��ӡ��Ч��ַ
function printOKWebSite(){
    $GLOBALS['conn=']=OpenConn();
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webdomain where isdomain=true');
    ASPEcho('��', @mysql_num_rows($rsObj));
    ASPEcho('�������', '<a href=\'?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����\'>OK</a>');
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        //call echo(rs("isdomain"),rs("website"))
        rw($rs['website'] . '<br>');
    }
}
//��ӡasp������վ
function printAspServerWebSite(){
    $GLOBALS['conn=']=OpenConn();
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webdomain where isasp=true and (isaspx=false and isphp=false)');
    ASPEcho('��', @mysql_num_rows($rsObj));
    ASPEcho('�������', '<a href=\'?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����\'>OK</a>');
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        //call echo(rs("isdomain"),rs("website"))
        rw($rs['website'] . '<br>');
    }
}

//�����ȫ��Ϊ��
function isthroughTrue(){
    $GLOBALS['conn=']=OpenConn();
    connExecute('update ' . $GLOBALS['db_PREFIX'] . 'webdomain set isthrough=true');
    ASPEcho('�������', '<a href=\'?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����\'>OK</a>');
}

//ɨ����ҳ��С
function scanDomainHomePageSize(){
    $url=''; $nSetTime=''; $isdomain=''; $htmlDir=''; $txtFilePath='';$homePageList='';$nThis='';$nCount='';
    $splstr='';$s='';$c='';$website='';$nState='';$websize='';$content='';$startTime='';$webtitle='';$webkeywords='';$webdescription='';

    $nThis=@$_REQUEST['nThis'];
    if( $nThis=='' ){
        $nThis=0;
    }else{
        $nThis=intval($nThis);
    }

    $nSetTime= 3;
    $GLOBALS['conn=']=OpenConn();
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webdomain where website<>\'\' and websize=0 and isdomain=true');
    $nCount=@$_REQUEST['nCount'];
    if( $nCount=='' ){
        $nCount= @mysql_num_rows($rsObj);
    }
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        $nThis=$nThis+1;
        ASPEcho($nThis . '/' . $nCount, $rs['website']);
        doevents( );
        $htmlDir= '/../��վUrlScan/������ҳ��С/';
        createDirFolder($htmlDir);
        $txtFilePath= $htmlDir . '/' . setFileName($rs['website']) . '.txt';
        if( checkFile($txtFilePath)== true ){
            ASPEcho('����', '����');
            $nSetTime=1;
        }else{
            $website=getwebsite($rs['website']);
            if( $website=='' ){
                eerr('����Ϊ��',$GLOBALS['httpurl']);
            }
            $content=getHttpPage($website,$rs['charset']);
            $content=toGB2312Char($content); //��PHP�ã�ת��gb2312�ַ�
            if( $content=='' ){
                $content=' ';
            }

            createFile($txtFilePath, $content);
            ASPEcho('����', '����');
        }
        $content=getftext($txtFilePath);
        $webtitle=getHtmlValue($content,'webtitle');
        $webkeywords=getHtmlValue($content,'webkeywords');
        $webdescription=getHtmlValue($content,'webdescription');


        $websize=getfsize($txtFilePath);
        ASPEcho('webtitle',$webtitle);
        //����д�Ǹ�תPHPʱ����
        connExecute('update ' . $GLOBALS['db_PREFIX'] . 'webdomain  set webtitle=\''. ADSql($webtitle) .'\',webkeywords=\''. $webkeywords .'\',webdescription=\''. $webdescription .'\',websize='. $websize .',isthrough=false,updatetime=\'' . Now() . '\'  where id=' . $rs['id'] . '');

        $startTime=@$_REQUEST['startTime'];
        if( $startTime=='' ){
            $startTime=now();
        }

        rw(VBRunTimer($startTime) . '<hr>');
        $url= getUrlAddToParam(getThisUrl(), '?nThis='. $nThis .'&nCount='. $nCount .'&startTime='. $startTime .'&N=' . getRnd(11), 'replace');

        rw(jsTiming($url, $nSetTime));
        die();
    }
    ASPEcho('�������', '<a href=\'?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����\'>OK����('. $nThis .')��</a>');
}

//ɨ��������ҳ
function scanDomainHomePage(){
    $url=''; $nSetTime=''; $isdomain=''; $htmlDir=''; $txtFilePath='';$homePageList='';$nThis='';$nCount='';
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
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webdomain where website<>\'\' and homepagelist=\'\' and isdomain=true');
    $nCount=@$_REQUEST['nCount'];
    if( $nCount=='' ){
        $nCount= @mysql_num_rows($rsObj);
    }
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        $nThis=$nThis+1;
        ASPEcho($nThis . '/' . $nCount, $rs['website']);
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

            if( $homePageList=='' ){
                $homePageList='��';
            }

            createFile($txtFilePath, $c);
            ASPEcho('����', '����');
        }
        //����д�Ǹ�תPHPʱ����
        connExecute('update ' . $GLOBALS['db_PREFIX'] . 'webdomain  set isasp='. $isAsp .',isaspx='. $isAspx .',isphp='. $isPhp .',isjsp='. $isJsp .',isthrough=false,homepagelist=\''. $homePageList .'\',updatetime=\'' . Now() . '\'  where id=' . $rs['id'] . '');

        $GLOBALS['startTime']=@$_REQUEST['startTime'];
        if( $GLOBALS['startTime']=='' ){
            $GLOBALS['startTime']=now();
        }

        rw(VBRunTimer($GLOBALS['startTime']) . '<hr>');
        $url= getUrlAddToParam(getThisUrl(), '?nThis='. $nThis .'&nCount='. $nCount .'&startTime='. $GLOBALS['startTime'] .'&N=' . getRnd(11), 'replace');

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
    $url=''; $nSetTime=''; $isdomain=''; $htmlDir=''; $txtFilePath=''; $nThis='';$nCount='';$startTime='';
    $nSetTime= 3;
    $nThis=@$_REQUEST['nThis'];
    if( $nThis=='' ){
        $nThis=0;
    }else{
        $nThis=intval($nThis);
    }
    $GLOBALS['conn=']=OpenConn();
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webdomain where isthrough=true');
    $nCount=@$_REQUEST['nCount'];
    if( $nCount=='' ){
        $nCount= @mysql_num_rows($rsObj);
    }
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        $nThis=$nThis+1;
        ASPEcho($nThis . '/' . $nCount, $rs['website']);
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
            createFile($txtFilePath, $isdomain . ' ');			 //��ֹPHP��д�벻��ȥ 0 �������
            ASPEcho('����', '����' . $txtFilePath . '('. checkFile($txtFilePath) .')=' . $isdomain);
        }
        //����д�Ǹ�תPHPʱ����
        connExecute('update ' . $GLOBALS['db_PREFIX'] . 'webdomain  set isthrough=false,isdomain=' . $isdomain . ',updatetime=\'' . Now() . '\'  where id=' . $rs['id'] . '');

        $startTime=@$_REQUEST['startTime'];
        if( $startTime=='' ){
            $startTime=now();
        }

        rw(VBRunTimer($startTime) . '<hr>');
        $url= getUrlAddToParam(getThisUrl(), '?nThis='. $nThis .'&nCount='. $nCount .'&startTime='. $startTime .'&N=' . getRnd(11), 'replace');

        rw(jsTiming($url, $nSetTime));
        die();
    }
    ASPEcho('�������', '<a href=\'?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����\'>OK����('. $nThis .')��</a>');
}

//ɨ����ַ
function runScanWebUrl(){
    $nSetTime=''; $setCharSet=''; $httpUrl=''; $url=''; $selectWeb ='';$nThis='';$nCount='';$startTime='';
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
    $nCount=@$_REQUEST['nCount'];
    if( $nCount=='' ){
        $nCount= @mysql_num_rows($rsObj);
    }
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

        $startTime=@$_REQUEST['startTime'];
        if( $startTime=='' ){
            $startTime=now();
        }

        VBRunTimer($startTime);
        $url= getUrlAddToParam(getThisUrl(), '?nThis='. $nThis .'&nCount='. $nCount .'&startTime='. $startTime .'&N=' . getRnd(11), 'replace');

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