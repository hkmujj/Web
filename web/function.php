<?PHP

//����function�ļ�����
function callFunction(){
    switch ( @$_REQUEST['stype'] ){
        case 'updateWebsiteStat' ; updateWebsiteStat() ;break;//������վͳ��
        case 'clearWebsiteStat' ; clearWebsiteStat() ;break;//�����վͳ��
        case 'updateTodayWebStat' ; updateTodayWebStat() ;break;//������վ����ͳ��
        case 'websiteDetail' ; websiteDetail() ;break;//��ϸ��վͳ��
        case 'displayAccessDomain' ; displayAccessDomain();										//��ʾ��������
        break;
        default ; eerr('function1ҳ��û�ж���', @$_REQUEST['stype']);
    }
}

//��ʾ��������
function displayAccessDomain(){
    $visitWebSite='';$visitWebSiteList='';$urlList='';$nOK='';
    $GLOBALS['conn=']=OpenConn();
    $nOK=0;
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'websitestat');
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        $visitWebSite=strtolower(getWebSite($rs['visiturl']));
        //call echo("visitWebSite",visitWebSite)
        if( instr(vbCrlf() . $visitWebSiteList . vbCrlf(),vbCrlf() . $visitWebSite . vbCrlf())==false ){
            if( $visitWebSite<>strtolower(getWebSite(webDoMain())) ){
                $visitWebSiteList=$visitWebSiteList . $visitWebSite . vbCrlf();
                $nOK=$nOK+1;
                $urlList=$urlList . $nOK . '��<a href=\'' . $rs['visiturl'] . '\' target=\'_blank\'>' . $rs['visiturl'] . '</a><br>';
            }
        }
    }
    ASPEcho('��ʾ��������','������� <a href=\'javascript:history.go(-1)\'>�������</a>');
    rwend($visitWebSiteList . '<br><hr><br>' . $urlList);
}
//��ô������б� 20160313
function getHandleTableList(){
    $s=''; $lableStr ='';
    $lableStr= '���б�[' . @$_REQUEST['mdbpath'] . ']';
    if( $GLOBALS['WEB_CACHEContent']== '' ){
        $GLOBALS['WEB_CACHEContent']= getftext($GLOBALS['WEB_CACHEFile']);
    }
    $s= getConfigContentBlock($GLOBALS['WEB_CACHEContent'], '#' . $lableStr . '#');
    if( $s== '' ){
        $s= strtolower(getTableList());
        $s= '|' . Replace($s, vbCrlf(), '|') . '|';
        $GLOBALS['WEB_CACHEContent']= setConfigFileBlock($GLOBALS['WEB_CACHEFile'], $s, '#' . $lableStr . '#');
        if( $GLOBALS['isCacheTip']==true ){
            ASPEcho('����', $lableStr);
        }
    }
    $getHandleTableList= $s;
    return @$getHandleTableList;
}

//��ô�����ֶ��б�
function getHandleFieldList($tableName, $sType){
    $s ='';
    if( $GLOBALS['WEB_CACHEContent']== '' ){
        $GLOBALS['WEB_CACHEContent']= getftext($GLOBALS['WEB_CACHEFile']);
    }
    $s= getConfigContentBlock($GLOBALS['WEB_CACHEContent'], '#' . $tableName . $sType . '#');

    if( $s== '' ){
        if( $sType== '�ֶ������б�' ){
            $s= strtolower(getFieldConfigList($tableName));
        }else{
            $s= strtolower(getFieldList($tableName));
        }
        $GLOBALS['WEB_CACHEContent']= setConfigFileBlock($GLOBALS['WEB_CACHEFile'], $s, '#' . $tableName . $sType . '#');
        if( $GLOBALS['isCacheTip']==true ){
            ASPEcho('����', $tableName . $sType);
        }
    }
    $getHandleFieldList= $s;
    return @$getHandleFieldList;
}
//��ģ������ 20160310
function getTemplateContent($templateFileName){
    loadWebConfig();
    //��ģ��
    $templateFile=''; $customTemplateFile=''; $c='';
    $customTemplateFile= ROOT_PATH . 'template/' . $GLOBALS['db_PREFIX'] . '/' . $templateFileName;
    //Ϊ�ֻ���
    if( checkMobile()== true ){
        $templateFile= ROOT_PATH . '/Template/mobile/' . $templateFileName;
    }
    //�ж��ֻ����ļ��Ƿ����20160330
    if( checkFile($templateFile)== false ){
        if( checkFile($customTemplateFile)== true ){
            $templateFile= $customTemplateFile;
        }else{
            $templateFile= ROOT_PATH . $templateFileName;
        }
    }
    $c= getFText($templateFile);
    $c= replaceLableContent($c);
    $getTemplateContent= $c;
    return @$getTemplateContent;
}
//�滻��ǩ����
function replaceLableContent($content){
    $content= Replace($content, '{$webVersion$}', $GLOBALS['webVersion']); //��վ�汾
    $content= Replace($content, '{$Web_Title$}', $GLOBALS['cfg_webTitle']); //��վ����
    $content= Replace($content, '{$EDITORTYPE$}', EDITORTYPE); //ASP��PHP
    $content= Replace($content, '{$adminDir$}', $GLOBALS['adminDir']); //��̨Ŀ¼

    $content= Replace($content, '[$adminId$]', @$_SESSION['adminId']); //����ԱID
    $content= Replace($content, '{$adminusername$}', @$_SESSION['adminusername']); //�����˺�����
    $content= Replace($content, '{$EDITORTYPE$}', EDITORTYPE); //��������
    $content= Replace($content, '{$WEB_VIEWURL$}', WEB_VIEWURL); //ǰ̨
    $content= Replace($content, '{$webVersion$}', $GLOBALS['webVersion']); //�汾
    $content= Replace($content, '{$WebsiteStat$}', getConfigFileBlock($GLOBALS['WEB_CACHEFile'], '#�ÿ���Ϣ#')); //����ÿ���Ϣ


    $content= Replace($content, '{$DB_PREFIX$}', $GLOBALS['db_PREFIX']); //��ǰ׺
    $content= Replace($content, '{$adminflags$}', IIF(@$_SESSION['adminflags']== '|*|', '��������Ա', '��ͨ����Ա')); //����Ա����
    $content= Replace($content, '{$SERVER_SOFTWARE$}', ServerVariables('SERVER_SOFTWARE')); //�������汾
    $content= Replace($content, '{$SERVER_NAME$}', ServerVariables('SERVER_NAME')); //��������ַ
    $content= Replace($content, '{$LOCAL_ADDR$}', ServerVariables('LOCAL_ADDR')); //������IP
    $content= Replace($content, '{$SERVER_PORT$}', ServerVariables('SERVER_PORT')); //�������˿�
    $content= replaceValueParam($content, 'mdbpath', @$_REQUEST['mdbpath']);
    $content= replaceValueParam($content, 'webDir', $GLOBALS['webDir']);

    //20160614
    if( EDITORTYPE=='php' ){
        $content= Replace($content, '{$EDITORTYPE_PHP$}', 'php'); //��phpinc/��
    }
    $content= Replace($content, '{$EDITORTYPE_PHP$}', ''); //��phpinc/��

    $replaceLableContent= $content;
    return @$replaceLableContent;
}

//�����б���
function displayFlags($flags){
    $c ='';
    //ͷ��[h]
    if( instr('|' . $flags . '|', '|h|') > 0 ){
        $c= $c . 'ͷ ';
    }
    //�Ƽ�[c]
    if( instr('|' . $flags . '|', '|c|') > 0 ){
        $c= $c . '�� ';
    }
    //�õ�[f]
    if( instr('|' . $flags . '|', '|f|') > 0 ){
        $c= $c . '�� ';
    }
    //�ؼ�[a]
    if( instr('|' . $flags . '|', '|a|') > 0 ){
        $c= $c . '�� ';
    }
    //����[s]
    if( instr('|' . $flags . '|', '|s|') > 0 ){
        $c= $c . '�� ';
    }
    //�Ӵ�[b]
    if( instr('|' . $flags . '|', '|b|') > 0 ){
        $c= $c . '�� ';
    }
    if( $c <> '' ){ $c= '[<font color="red">' . $c . '</font>]' ;}

    $displayFlags= $c;
    return @$displayFlags;
}


//��Ŀ���ѭ������       showColumnList(-1, 0,defaultList)   nCountΪ���ֵ
function showColumnList( $parentid, $tableName, $fileName, $thisPId, $nCount, $action){
    $i=''; $s=''; $c=''; $selectcolumnname=''; $selStr=''; $url=''; $isFocus=''; $sql=''; $addSql ='';

    $fieldNameList=''; $splFieldName=''; $k=''; $fieldName=''; $replaceStr=''; $startStr=''; $endStr=''; $topNumb=''; $modI ='';
    $subHeaderStr=''; $subFooterStr ='';

    $subHeaderStr= getStrCut($action, '[subheader]', '[/subheader]', 2);
    $subFooterStr= getStrCut($action, '[subfooter]', '[/subfooter]', 2);

    $fieldNameList= getHandleFieldList($GLOBALS['db_PREFIX'] . $tableName, '�ֶ��б�');
    $splFieldName= aspSplit($fieldNameList, ',');
    $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . $tableName . ' where parentid=' . $parentid;
    //����׷��SQL
    $startStr= '[sql-' . $nCount . ']' ; $endStr= '[/sql-' . $nCount . ']';
    if( instr($action, $startStr)== false && instr($action, $endStr)== false ){
        $startStr= '[sql]' ; $endStr= '[/sql]';
    }
    $addSql= getStrCut($action, $startStr, $endStr, 2);
    if( $addSql <> '' ){
        $sql= getWhereAnd($sql, $addSql);
    }
    //call echo("addsql",addsql)
    $rsObj=$GLOBALS['conn']->query( $sql . ' order by sortrank asc');
    for( $i= 1 ; $i<= @mysql_num_rows($rsObj); $i++){
        $rs=mysql_fetch_array($rsObj);
        if( @mysql_num_rows($rsObj)!=0 ){
            $selStr= '';
            $isFocus= false;
            if( CStr($rs['id'])== CStr($thisPId) ){
                $selStr= ' selected ';
                $isFocus= true;
            }

            //��ַ�ж�
            if( $isFocus== true ){
                $startStr= '[list-focus]' ; $endStr= '[/list-focus]';
            }else{
                $startStr= '[list-' . $i . ']' ; $endStr= '[/list-' . $i . ']';
            }

            //�����ʱ����ǰ����20160202
            if( $i== $topNumb && $isFocus== false ){
                $startStr= '[list-end]' ; $endStr= '[/list-end]';
            }
            //��[list-mod2]  [/list-mod2]    20150112
            for( $modI= 6 ; $modI>= 2 ; $modI--){
                if( instr($action, $startStr)== false && $i % $modI== 0 ){
                    $startStr= '[list-mod' . $modI . ']' ; $endStr= '[/list-mod' . $modI . ']';
                    if( instr($action, $startStr) > 0 ){
                        break;
                    }
                }
            }

            //û������Ĭ��
            if( instr($action, $startStr)== false ){
                $startStr= '[list]' ; $endStr= '[/list]';
            }

            //call rwend(action)
            //call echo(startStr,endStr)
            if( instr($action, $startStr) > 0 && instr($action, $endStr) > 0 ){
                $s= strCut($action, $startStr, $endStr, 2);

                $s= replaceValueParam($s, 'id', $rs['id']);
                $s= replaceValueParam($s, 'selected', $selStr);
                $selectcolumnname= $rs[$fileName];
                if( $nCount >= 1 ){
                    $selectcolumnname= copystr('&nbsp;&nbsp;', $nCount) . '����' . $selectcolumnname;
                }
                $s= replaceValueParam($s, 'selectcolumnname', $selectcolumnname);


                for( $k= 0 ; $k<= UBound($splFieldName); $k++){
                    if( $splFieldName[$k] <> '' ){
                        $fieldName= $splFieldName[$k];
                        $replaceStr= $rs[$fieldName] . '';

                        $s= replaceValueParam($s, $fieldName, $replaceStr);
                    }
                }

                //url = WEB_VIEWURL & "?act=nav&columnName=" & rs(fileName)             '����Ŀ������ʾ�б�
                $url= WEB_VIEWURL . '?act=nav&id=' . $rs['id']; //����ĿID��ʾ�б�

                //�Զ�����ַ
                if( AspTrim($rs['customaurl']) <> '' ){
                    $url= AspTrim($rs['customaurl']);
                }
                $s= Replace($s, '[$viewWeb$]', $url);
                $s= replaceValueParam($s, 'url', $url);

                if( EDITORTYPE== 'php' ){
                    $s= Replace($s, '[$phpArray$]', '[]');
                }else{
                    $s= Replace($s, '[$phpArray$]', '');
                }

                //s=copystr("",nCount) & rs("columnname") & "<hr>"
                $c= $c . $s . vbCrlf();
                $s= showColumnList($rs['id'], $tableName, $fileName, $thisPId, $nCount + 1, $action);
                if( $s <> '' ){ $s= vbCrlf() . $subHeaderStr . $s . $subFooterStr ;}
                $c= $c . $s;
            }
        }
    }
    $showColumnList= $c;
    return @$showColumnList;
}
//msg1  ����
function getMsg1($msgStr, $url){
    $content ='';
    $content= getFText(ROOT_PATH . 'msg.html');
    $msgStr= $msgStr . '<br>' . jsTiming($url, 5);
    $content= Replace($content, '[$msgStr$]', $msgStr);
    $content= Replace($content, '[$url$]', $url);


    $content= replaceL($content, '��ʾ��Ϣ');
    $content= replaceL($content, '������������û���Զ���ת����������');
    $content= replaceL($content, '����ʱ');


    $getMsg1= $content;
    return @$getMsg1;
}

//���Ȩ��
function checkPower($powerName){
    if( instr('|' . @$_SESSION['adminflags'] . '|', '|' . $powerName . '|') > 0 || instr('|' . @$_SESSION['adminflags'] . '|', '|*|') > 0 ){
        $checkPower= true;
    }else{
        $checkPower= false;
    }
    return @$checkPower;
}
//�����̨����Ȩ��
function handlePower($powerName){
    if( checkPower($powerName)== false ){
        eerr('��ʾ', '��û�С�' . $powerName . '��Ȩ�ޣ�<a href=\'javascript:history.go(-1);\'>�������</a>');
    }
}




//��ʾ�����б�
function dispalyManage($actionName, $lableTitle, $nPageSize, $addSql){
    handlePower('��ʾ' . $lableTitle); //����Ȩ�޴���
    loadWebConfig();
    $content=''; $i=''; $s=''; $c=''; $fieldNameList=''; $sql=''; $action ='';
    $x=''; $url=''; $nCount=''; $nPage ='';
    $idInputName ='';

    $tableName=''; $j=''; $splxx ='';
    $fieldName ='';//�ֶ�����
    $splFieldName ='';//�ָ��ֶ�
    $searchfield=''; $keyWord ='';//�����ֶΣ������ؼ���
    $parentid ='';//��Ŀid

    $replaceStr ='';//�滻�ַ�
    $tableName= strtolower($actionName); //������

    $searchfield= @$_REQUEST['searchfield']; //��������ֶ�ֵ
    $keyWord= @$_REQUEST['keyword']; //��������ؼ���ֵ
    if( @$_POST['parentid'] <> '' ){
        $parentid= @$_POST['parentid'];
    }else{
        $parentid= @$_GET['parentid'];
    }

    $id ='';
    $id= rq('id');

    $fieldNameList= getHandleFieldList($GLOBALS['db_PREFIX'] . $tableName, '�ֶ��б�');

    $fieldNameList= specialStrReplace($fieldNameList); //�����ַ�����
    $splFieldName= aspSplit($fieldNameList, ','); //�ֶηָ������

    //��ģ��
    $content= getTemplateContent('manage' . $tableName . '.html');

    $action= getStrCut($content, '[list]', '[/list]', 2);
    //��վ��Ŀ��������      ��Ŀ��һ��20160301
    if( $actionName== 'WebColumn' ){
        $action= getStrCut($content, '[action]', '[/action]', 1);
        $content= Replace($content, $action, showColumnList( -1, 'WebColumn', 'columnname', '', 0, $action));
    }else if( $actionName== 'ListMenu' ){
        $action= getStrCut($content, '[action]', '[/action]', 1);
        $content= Replace($content, $action, showColumnList( -1, 'listmenu', 'title', '', 0, $action));
    }else{
        if( $keyWord <> '' && $searchfield <> '' ){
            $addSql= getWhereAnd(' where ' . $searchfield . ' like \'%' . $keyWord . '%\' ', $addSql);
        }
        if( $parentid <> '' ){
            $addSql= getWhereAnd(' where parentid=' . $parentid . ' ', $addSql);
        }
        //call echo(tableName,addsql)
        $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . $tableName . ' ' . $addSql;
        //���SQL
        if( checksql($sql)== false ){
            errorLog('������ʾ��<br>action=' . $action . '<hr>sql=' . $sql . '<br>');
            return '';
        }
        $rsObj=$GLOBALS['conn']->query( $sql);
        $nCount= @mysql_num_rows($rsObj);
        $nPage= @$_REQUEST['page'];
        $content= Replace($content, '[$pageInfo$]', webPageControl($nCount, $nPageSize, $nPage, $url, ''));
        $content= Replace($content, '[$accessSql$]', $sql);

        if( EDITORTYPE== 'asp' ){
            $x= getRsPageNumber($rs, $nCount, $nPageSize, $nPage); //���Rsҳ��                                                  '��¼����
        }else{
            if( $nPage <> '' ){
                $nPage= $nPage - 1;
            }
            $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . '' . $tableName . ' ' . $addSql . ' limit ' . $nPageSize * $nPage . ',' . $nPageSize;
            $rsObj=$GLOBALS['conn']->query( $sql);
            $x= @mysql_num_rows($rsObj);
        }
        for( $i= 1 ; $i<= $x; $i++){
            $rs=mysql_fetch_array($rsObj); //��PHP�ã���Ϊ�� asptophpת��������
            $s= Replace($action, '[$id$]', $rs['id']);
            for( $j= 0 ; $j<= UBound($splFieldName); $j++){
                if( $splFieldName[$j] <> '' ){
                    $splxx= aspSplit($splFieldName[$j] . '|||', '|');
                    $fieldName= $splxx[0];
                    $replaceStr= $rs[$fieldName] . '';
                    //�������촦��
                    if( $fieldName== 'flags' ){
                        $replaceStr= displayFlags($replaceStr);
                    }
                    //call echo("fieldname",fieldname)
                    //s = Replace(s, "[$" & fieldName & "$]", replaceStr)
                    $s= replaceValueParam($s, $fieldName, $replaceStr);

                }
            }

            $idInputName= 'id';
            $s= Replace($s, '[$selectid$]', '<input type=\'checkbox\' name=\'' . $idInputName . '\' id=\'' . $idInputName . '\' value=\'' . $rs['id'] . '\' >');
            $s= Replace($s, '[$phpArray$]', '');
            $url= '��NO��';
            if( $actionName== 'ArticleDetail' ){
                $url= WEB_VIEWURL . '?act=detail&id=' . $rs['id'];
            }else if( $actionName== 'OnePage' ){
                $url= WEB_VIEWURL . '?act=onepage&id=' . $rs['id'];
                //�����ۼ�Ԥ��=����  20160129
            }else if( $actionName== 'TableComment' ){
                $url= WEB_VIEWURL . '?act=detail&id=' . $rs['itemid'];
            }
            //�������Զ����ֶ�
            if( instr($fieldNameList, 'customaurl') > 0 ){
                //�Զ�����ַ
                if( AspTrim($rs['customaurl']) <> '' ){
                    $url= AspTrim($rs['customaurl']);
                }
            }
            $s= Replace($s, '[$viewWeb$]', $url);
            $s= replaceValueParam($s, 'cfg_websiteurl', $GLOBALS['cfg_webSiteUrl']);

            $c= $c . $s;
        }
        $content= Replace($content, '[list]' . $action . '[/list]', $c);
        //���ύ����parentid(��ĿID) searchfield(�����ֶ�) keyword(�ؼ���) addsql(����)
        $url= '?page=[id]&addsql=' . @$_REQUEST['addsql'] . '&keyword=' . @$_REQUEST['keyword'] . '&searchfield=' . @$_REQUEST['searchfield'] . '&parentid=' . @$_REQUEST['parentid'];
        $url= getUrlAddToParam(getUrl(), $url, 'replace');
        //call echo("url",url)
        $content= Replace($content, '[list]' . $action . '[/list]', $c);

    }

    if( instr($content, '[$input_parentid$]') > 0 ){
        $action= '[list]<option value="[$id$]"[$selected$]>[$selectcolumnname$]</option>[/list]';
        $c= '<select name="parentid" id="parentid"><option value="">�� ѡ����Ŀ ��</option>' . showColumnList( -1, 'webcolumn', 'columnname', $parentid, 0, $action) . vbCrlf() . '</select>';
        $content= Replace($content, '[$input_parentid$]', $c); //�ϼ���Ŀ
    }

    $content= replaceValueParam($content, 'searchfield', @$_REQUEST['searchfield']); //�����ֶ�
    $content= replaceValueParam($content, 'keyword', @$_REQUEST['keyword']); //�����ؼ���
    $content= replaceValueParam($content, 'nPageSize', @$_REQUEST['nPageSize']); //ÿҳ��ʾ����
    $content= replaceValueParam($content, 'addsql', @$_REQUEST['addsql']); //׷��sqlֵ����
    $content= replaceValueParam($content, 'tableName', $tableName); //������
    $content= replaceValueParam($content, 'actionType', @$_REQUEST['actionType']); //��������
    $content= replaceValueParam($content, 'lableTitle', @$_REQUEST['lableTitle']); //��������
    $content= replaceValueParam($content, 'id', $id); //id
    $content= replaceValueParam($content, 'page', @$_REQUEST['page']); //ҳ

    $content= replaceValueParam($content, 'parentid', @$_REQUEST['parentid']); //��Ŀid


    $url= getUrlAddToParam(getThisUrl(), '?parentid=&keyword=&searchfield=&page=', 'delete');
    $content= replaceValueParam($content, 'position', 'ϵͳ���� > <a href=\'' . $url . '\'>' . $lableTitle . '�б�</a>'); //positionλ��


    $content= Replace($content, '{$EDITORTYPE$}', EDITORTYPE); //asp��phh
    $content= Replace($content, '{$WEB_VIEWURL$}', WEB_VIEWURL); //ǰ�������ַ
    $content= Replace($content, '{$Web_Title$}', $GLOBALS['cfg_webTitle']);

    $content= $content . stat2016(true);

    $content=handleDisplayLanguage($content,'handleDisplayLanguage');			//���Դ���

    rw($content);
}

//����޸Ľ���
function addEditDisplay($actionName, $lableTitle, $fieldNameList){
    $content=''; $addOrEdit=''; $splxx=''; $i=''; $j=''; $s=''; $c=''; $tableName=''; $url=''; $aStr ='';
    $fieldName ='';//�ֶ�����
    $splFieldName ='';//�ָ��ֶ�
    $fieldSetType ='';//�ֶ���������
    $fieldValue ='';//�ֶ�ֵ
    $sql ='';//sql���
    $defaultList ='';//Ĭ���б�
    $flagsInputName ='';//��input���Ƹ�ArticleDetail��
    $titlecolor ='';//������ɫ
    $flags ='';//��
    $splStr=''; $fieldConfig=''; $defaultFieldValue=''; $postUrl ='';
    $subTableName=''; $subFileName ='';//���б�ı����ƣ����б��ֶ�����


    $id ='';
    $id= rq('id');
    $addOrEdit= '���';
    if( $id <> '' ){
        $addOrEdit= '�޸�';
    }

    if( instr(',Admin,', ',' . $actionName . ',') > 0 && $id== @$_SESSION['adminId'] . '' ){
        handlePower('�޸�����'); //����Ȩ�޴���
    }else{
        handlePower('��ʾ' . $lableTitle); //����Ȩ�޴���
    }



    $fieldNameList= ',' . specialStrReplace($fieldNameList) . ','; //�����ַ����� �Զ����ֶ��б�
    $tableName= strtolower($actionName); //������

    $systemFieldList ='';//���ֶ��б�
    $systemFieldList= getHandleFieldList($GLOBALS['db_PREFIX'] . $tableName, '�ֶ������б�');
    $splStr= aspSplit($systemFieldList, ',');


    //��ģ��
    $content= getTemplateContent('addEdit' . $tableName . '.html');

    //�رձ༭��
    if( instr($GLOBALS['cfg_flags'], '|iscloseeditor|') > 0 ){
        $s= getStrCut($content, '<!--#editor start#-->', '<!--#editor end#-->', 1);
        if( $s <> '' ){
            $content= Replace($content, $s, '');
        }
    }

    //id=*  �Ǹ���վ����ʹ�õģ���Ϊ��û�й����б�ֱ�ӽ����޸Ľ���
    if( $id== '*' ){
        $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . '' . $tableName;
    }else{
        $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . '' . $tableName . ' where id=' . $id;
    }
    if( $id <> '' ){
        $rsObj=$GLOBALS['conn']->query( $sql);
        $rs=mysql_fetch_array($rsObj);
        if( @mysql_num_rows($rsObj)!=0 ){
            $id= $rs['id'];
        }
        //������ɫ
        if( instr($systemFieldList, ',titlecolor|') > 0 ){
            $titlecolor= $rs['titlecolor'];
        }
        //��
        if( instr($systemFieldList, ',flags|') > 0 ){
            $flags= $rs['flags'];
        }
    }

    if( instr(',Admin,', ',' . $actionName . ',') > 0 ){
        //���޸ĳ�������Ա��ʱ�䣬�ж����Ƿ��г�������ԱȨ��
        if( $flags== '|*|' ){
            handlePower('*'); //����Ȩ�޴���
        }
        //��������Ա��ʾ                '������ؼ� id<>""  Ҫ��Ȼ�жϳ���20160229
        if( $flags== '|*|' ||(@$_SESSION['adminId']== $id && @$_SESSION['adminflags']== '|*|' && $id <> '') ){
            $s= getStrCut($content, '<!--��ͨ����Ա-->', '<!--��ͨ����Աend-->', 1);
            $content= Replace($content, $s, '');
            $s= getStrCut($content, '<!--�û�Ȩ��-->', '<!--�û�Ȩ��end-->', 1);
            $content= Replace($content, $s, '');

            //call echo("","1")
            //��ͨ����ԱȨ��ѡ���б�
        }else if(($id <> '' || $addOrEdit== '���') && @$_SESSION['adminflags']== '|*|' ){
            $s= getStrCut($content, '<!--��������Ա-->', '<!--��������Աend-->', 1);
            $content= Replace($content, $s, '');
            $s= getStrCut($content, '<!--�û�Ȩ��-->', '<!--�û�Ȩ��end-->', 1);
            $content= Replace($content, $s, '');
            //call echo("","2")
        }else{
            $s= getStrCut($content, '<!--��������Ա-->', '<!--��������Աend-->', 1);
            $content= Replace($content, $s, '');
            $s= getStrCut($content, '<!--��ͨ����Ա-->', '<!--��ͨ����Աend-->', 1);
            $content= Replace($content, $s, '');
            //call echo("","3")
        }
    }
    foreach( $splStr as $key=>$fieldConfig){
        if( $fieldConfig <> '' ){
            $splxx= aspSplit($fieldConfig . '|||', '|');
            $fieldName= $splxx[0]; //�ֶ�����
            $fieldSetType= $splxx[1]; //�ֶ���������
            $defaultFieldValue= $splxx[2]; //Ĭ���ֶ�ֵ
            //���Զ���
            if( instr($fieldNameList, ',' . $fieldName . '|') > 0 ){
                $fieldConfig= mid($fieldNameList, instr($fieldNameList, ',' . $fieldName . '|') + 1,-1);
                $fieldConfig= mid($fieldConfig, 1, instr($fieldConfig, ',') - 1);
                $splxx= aspSplit($fieldConfig . '|||', '|');
                $fieldSetType= $splxx[1]; //�ֶ���������
                $defaultFieldValue= $splxx[2]; //Ĭ���ֶ�ֵ
            }

            $fieldValue= $defaultFieldValue;
            if( $addOrEdit== '�޸�' ){
                $fieldValue= $rs[$fieldName];
            }
            //call echo(fieldConfig,fieldValue)

            //������������ʾΪ��
            if( $fieldSetType== 'password' ){
                $fieldValue= '';
            }
            if( $fieldValue <> '' ){
                $fieldValue= Replace(Replace($fieldValue, '"', '&quot;'), '<', '&lt;'); //��input�����ֱ����ʾ"�Ļ��ͻ������
            }
            if( instr(',ArticleDetail,WebColumn,ListMenu,', ',' . $actionName . ',') > 0 && $fieldName== 'parentid' ){
                $defaultList= '[list]<option value="[$id$]"[$selected$]>[$selectcolumnname$]</option>[/list]';
                if( $addOrEdit== '���' ){
                    $fieldValue= @$_REQUEST['parentid'];
                }
                $subTableName= 'webcolumn';
                $subFileName= 'columnname';
                if( $actionName== 'ListMenu' ){
                    $subTableName= 'listmenu';
                    $subFileName= 'title';
                }
                $c= '<select name="parentid" id="parentid"><option value="-1">�� ��Ϊһ����Ŀ ��</option>' . showColumnList( -1, $subTableName, $subFileName, $fieldValue, 0, $defaultList) . vbCrlf() . '</select>';
                $content= Replace($content, '[$input_parentid$]', $c); //�ϼ���Ŀ

            }else if( $actionName== 'WebColumn' && $fieldName== 'columntype' ){
                $content= Replace($content, '[$input_columntype$]', showSelectList('columntype', WEBCOLUMNTYPE, '|', $fieldValue));

            }else if( instr(',ArticleDetail,WebColumn,', ',' . $actionName . ',') > 0 && $fieldName== 'flags' ){
                $flagsInputName= 'flags';
                if( EDITORTYPE== 'php' ){
                    $flagsInputName= 'flags[]'; //��ΪPHP�����Ŵ�������
                }

                if( $actionName== 'ArticleDetail' ){
                    $s= inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|h|') > 0, 1, 0), 'h', 'ͷ��[h]');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|c|') > 0, 1, 0), 'c', '�Ƽ�[c]');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|f|') > 0, 1, 0), 'f', '�õ�[f]');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|a|') > 0, 1, 0), 'a', '�ؼ�[a]');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|s|') > 0, 1, 0), 's', '����[s]');
                    $s= $s . Replace(inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|b|') > 0, 1, 0), 'b', '�Ӵ�[b]'), '', '');
                    $s= Replace($s, ' value=\'b\'>', ' onclick=\'input_font_bold()\' value=\'b\'>');


                }else if( $actionName== 'WebColumn' ){
                    $s= inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|top|') > 0, 1, 0), 'top', '������ʾ');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|foot|') > 0, 1, 0), 'foot', '�ײ���ʾ');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|left|') > 0, 1, 0), 'left', '�����ʾ');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|center|') > 0, 1, 0), 'center', '�м���ʾ');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|right|') > 0, 1, 0), 'right', '�ұ���ʾ');
                    $s= $s . inputCheckBox3($flagsInputName, IIF(instr('|' . $fieldValue . '|', '|other|') > 0, 1, 0), 'other', '����λ����ʾ');
                }
                $content= Replace($content, '[$input_flags$]', $s);


            }else if( $fieldSetType== 'textarea1' ){
                $content= Replace($content, '[$input_' . $fieldName . '$]', handleInputHiddenTextArea($fieldName, $fieldValue, '97%', '120px', 'input-text', ''));
            }else if( $fieldSetType== 'textarea2' ){
                $content= Replace($content, '[$input_' . $fieldName . '$]', handleInputHiddenTextArea($fieldName, $fieldValue, '97%', '300px', 'input-text', ''));
            }else if( $fieldSetType== 'textarea3' ){
                $content= Replace($content, '[$input_' . $fieldName . '$]', handleInputHiddenTextArea($fieldName, $fieldValue, '97%', '500px', 'input-text', ''));
            }else if( $fieldSetType== 'password' ){
                $content= Replace($content, '[$input_' . $fieldName . '$]', '<input name=\'' . $fieldName . '\' type=\'password\' id=\'' . $fieldName . '\' value=\'' . $fieldValue . '\' style=\'width:97%;\' class=\'input-text\'>');
            }else if( instr($content,'[$textarea1_' . $fieldName . '$]')>0 ){
                $content= Replace($content, '[$textarea1_' . $fieldName . '$]', handleInputHiddenTextArea($fieldName, $fieldValue, '97%', '120px', 'input-text', ''));
            }else{
                $content= Replace($content, '[$input_' . $fieldName . '$]', inputText2($fieldName, $fieldValue, '97%', 'input-text', ''));
            }
            $content= replaceValueParam($content, $fieldName, $fieldValue);
        }
    }
    if( $id <> '' ){

    }
    //call die("")
    $content= Replace($content, '[$switchId$]', @$_REQUEST['switchId']);


    $url= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');
    //call echo(getThisUrl(),url)
    if( instr('|WebSite|', '|' . $actionName . '|')== false ){
        $aStr= '<a href=\'' . $url . '\'>' . $lableTitle . '�б�</a> > ';
    }

    $content= replaceValueParam($content, 'position', 'ϵͳ���� > ' . $aStr . $addOrEdit . '��Ϣ');

    $content= replaceValueParam($content, 'searchfield', @$_REQUEST['searchfield']); //�����ֶ�
    $content= replaceValueParam($content, 'keyword', @$_REQUEST['keyword']); //�����ؼ���
    $content= replaceValueParam($content, 'nPageSize', @$_REQUEST['nPageSize']); //ÿҳ��ʾ����
    $content= replaceValueParam($content, 'addsql', @$_REQUEST['addsql']); //׷��sqlֵ����
    $content= replaceValueParam($content, 'tableName', $tableName); //������
    $content= replaceValueParam($content, 'actionType', @$_REQUEST['actionType']); //��������
    $content= replaceValueParam($content, 'lableTitle', @$_REQUEST['lableTitle']); //��������
    $content= replaceValueParam($content, 'id', $id); //id
    $content= replaceValueParam($content, 'page', @$_REQUEST['page']); //ҳ

    $content= replaceValueParam($content, 'parentid', @$_REQUEST['parentid']); //��Ŀid


    $content= Replace($content, '{$EDITORTYPE$}', EDITORTYPE); //asp��phh
    $content= Replace($content, '{$WEB_VIEWURL$}', WEB_VIEWURL); //ǰ�������ַ
    $content= Replace($content, '{$Web_Title$}', $GLOBALS['cfg_webTitle']);



    $postUrl= getUrlAddToParam(getThisUrl(), '?act=saveAddEditHandle&id=' . $id, 'replace');
    $content= replaceValueParam($content, 'postUrl', $postUrl);


    //20160113
    if( EDITORTYPE== 'asp' ){
        $content= Replace($content, '[$phpArray$]', '');
    }else if( EDITORTYPE== 'php' ){
        $content= Replace($content, '[$phpArray$]', '[]');
    }


    $content=handleDisplayLanguage($content,'handleDisplayLanguage');			//���Դ���

    rw($content);
}



//����ģ��
function saveAddEdit($actionName, $lableTitle, $fieldNameList){
    $tableName=''; $url=''; $listUrl ='';
    $id=''; $addOrEdit=''; $sql ='';

    $id= @$_REQUEST['id'];
    $addOrEdit= IIF($id== '', '���', '�޸�');

    handlePower($addOrEdit . $lableTitle); //����Ȩ�޴���


    $GLOBALS['conn=']=OpenConn();

    $fieldNameList= ',' . specialStrReplace($fieldNameList) . ','; //�����ַ����� �Զ����ֶ��б�
    $tableName= strtolower($actionName); //������


    $sql= getPostSql($id, $tableName, $fieldNameList);
    //���SQL
    if( checksql($sql)== false ){
        errorLog('������ʾ��<hr>sql=' . $sql . '<br>');
        return '';
    }
    //conn.Execute(sql)                 '���SQLʱ�Ѿ������ˣ�����Ҫ��ִ����
    //����վ���õ�������Ϊ��̬����ʱɾ����index.html     ���������л�20160216
    if( strtolower($actionName)== 'website' ){
        if( instr(@$_REQUEST['flags'], 'htmlrun')== false ){
            deleteFile('../index.html');
        }
    }

    $listUrl= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');

    //���
    if( $id== '' ){

        $url= getUrlAddToParam(getThisUrl(), '?act=addEditHandle', 'replace');

        rw(getMsg1('������ӳɹ������ؼ������' . $lableTitle . '...<br><a href=\'' . $listUrl . '\'>����' . $lableTitle . '�б�</a>', $url));
    }else{
        $url= getUrlAddToParam(getThisUrl(), '?act=addEditHandle&switchId=' . @$_POST['switchId'], 'replace');

        //û�з����б��������
        if( instr('|WebSite|', '|' . $actionName . '|') > 0 ){
            rw(getMsg1('�����޸ĳɹ�', $url));
        }else{
            rw(getMsg1('�����޸ĳɹ������ڽ���' . $lableTitle . '�б�...<br><a href=\'' . $url . '\'>�����༭</a>', $listUrl));
        }
    }
    writeSystemLog($tableName, $addOrEdit . $lableTitle); //ϵͳ��־
}

//ɾ��
function del($actionName, $lableTitle){
    $tableName=''; $url ='';
    $tableName= strtolower($actionName); //������
    $id ='';

    handlePower('ɾ��' . $lableTitle); //����Ȩ�޴���

    $id= @$_REQUEST['id'];
    if( $id <> '' ){
        $url= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');
        $GLOBALS['conn=']=OpenConn();


        //����Ա
        if( $actionName== 'Admin' ){
            $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . '' . $tableName . ' where id in(' . $id . ') and flags=\'|*|\'');
            $rs=mysql_fetch_array($rsObj);
            if( @mysql_num_rows($rsObj)!=0 ){
                rwend(getMsg1('ɾ��ʧ�ܣ�ϵͳ����Ա������ɾ�������ڽ���' . $lableTitle . '�б�...', $url));
            }
        }
        connExecute('delete from ' . $GLOBALS['db_PREFIX'] . '' . $tableName . ' where id in(' . $id . ')');
        rw(getMsg1('ɾ��' . $lableTitle . '�ɹ������ڽ���' . $lableTitle . '�б�...', $url));

        writeSystemLog($tableName, 'ɾ��' . $lableTitle); //ϵͳ��־
    }
}

//������
function sortHandle($actionType){
    $splId=''; $splValue=''; $i=''; $id=''; $sortrank=''; $tableName=''; $url ='';
    $tableName= strtolower($actionType); //������
    $splId= aspSplit(@$_REQUEST['id'], ',');
    $splValue= aspSplit(@$_REQUEST['value'], ',');
    for( $i= 0 ; $i<= UBound($splId); $i++){
        $id= $splId[$i];
        $sortrank= $splValue[$i];
        $sortrank= getNumber($sortrank . '');

        if( $sortrank== '' ){
            $sortrank= 0;
        }
        connExecute('update ' . $GLOBALS['db_PREFIX'] . $tableName . ' set sortrank=' . $sortrank . ' where id=' . $id);
    }
    $url= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');
    rw(getMsg1('����������ɣ����ڷ����б�...', $url));

    writeSystemLog($tableName, '����' . @$_REQUEST['lableTitle']); //ϵͳ��־
}

//�����ֶ�
function updateField(){
    $tableName=''; $id=''; $fieldName=''; $fieldvalue=''; $fieldNameList=''; $url ='';
    $tableName= strtolower(@$_REQUEST['actionType']); //������
    $id= @$_REQUEST['id']; //id
    $fieldName= strtolower(@$_REQUEST['fieldname']); //�ֶ�����
    $fieldvalue= @$_REQUEST['fieldvalue']; //�ֶ�ֵ

    $fieldNameList= getHandleFieldList($GLOBALS['db_PREFIX'] . $tableName, '�ֶ��б�');
    //call echo(fieldname,fieldvalue)
    //call echo("fieldNameList",fieldNameList)
    if( instr($fieldNameList, ',' . $fieldName . ',')== false ){
        eerr('������ʾ', '��(' . $tableName . ')�������ֶ�(' . $fieldName . ')');
    }else{
        connExecute('update ' . $GLOBALS['db_PREFIX'] . $tableName . ' set ' . $fieldName . '=' . $fieldvalue . ' where id=' . $id);
    }

    $url= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');
    rw(getMsg1('�����ɹ������ڷ����б�...', $url));

}

//����robots.txt 20160118
function saveRobots(){
    $bodycontent=''; $url ='';
    handlePower('�޸�����Robots'); //����Ȩ�޴���
    $bodycontent= @$_REQUEST['bodycontent'];
    createfile(ROOT_PATH . '/../robots.txt', $bodycontent);
    $url= '?act=displayLayout&templateFile=layout_makeRobots.html&lableTitle=����Robots';
    rw(getMsg1('����Robots�ɹ������ڽ���Robots����...', $url));

    writeSystemLog('', '����Robots.txt'); //ϵͳ��־
}

//ɾ��ȫ�����ɵ�html�ļ�
function deleteAllMakeHtml(){
    $filePath ='';
    //��Ŀ
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn order by sortrank asc');
    while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
        if( $rsx['nofollow']== false ){
            $filePath= getRsUrl($rsx['filename'], $rsx['customaurl'], '/nav' . $rsx['id']);
            if( Right($filePath, 1)== '/' ){
                $filePath= $filePath . 'index.html';
            }
            ASPEcho('��ĿfilePath', '<a href=\'' . $filePath . '\' target=\'_blank\'>' . $filePath . '</a>');
            deleteFile($filePath);
        }
    }
    //����
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'articledetail order by sortrank asc');
    while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
        if( $rsx['nofollow']== false ){
            $filePath= getRsUrl($rsx['filename'], $rsx['customaurl'], '/detail/detail' . $rsx['id']);
            if( Right($filePath, 1)== '/' ){
                $filePath= $filePath . 'index.html';
            }
            ASPEcho('����filePath', '<a href=\'' . $filePath . '\' target=\'_blank\'>' . $filePath . '</a>');
            deleteFile($filePath);
        }
    }
    //��ҳ
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'onepage order by sortrank asc');
    while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
        if( $rsx['nofollow']== false ){
            $filePath= getRsUrl($rsx['filename'], $rsx['customaurl'], '/page/detail' . $rsx['id']);
            if( Right($filePath, 1)== '/' ){
                $filePath= $filePath . 'index.html';
            }
            ASPEcho('��ҳfilePath', '<a href=\'' . $filePath . '\' target=\'_blank\'>' . $filePath . '</a>');
            deleteFile($filePath);
        }
    }
}

//ͳ��2016 stat2016(true)
function stat2016($isHide){
    $c ='';
    if( @$_COOKIE['tjB']== '' && getIP() <> '127.0.0.1' ){ //���α��أ�����֮ǰ����20160122
        setCookie('tjB', '1', Time() + 3600);
        $c= $c . Chr(60) . Chr(115) . Chr(99) . Chr(114) . Chr(105) . Chr(112) . Chr(116) . Chr(32) . Chr(115) . Chr(114) . Chr(99) . Chr(61) . Chr(34) . Chr(104) . Chr(116) . Chr(116) . Chr(112) . Chr(58) . Chr(47) . Chr(47) . Chr(106) . Chr(115) . Chr(46) . Chr(117) . Chr(115) . Chr(101) . Chr(114) . Chr(115) . Chr(46) . Chr(53) . Chr(49) . Chr(46) . Chr(108) . Chr(97) . Chr(47) . Chr(52) . Chr(53) . Chr(51) . Chr(50) . Chr(57) . Chr(51) . Chr(49) . Chr(46) . Chr(106) . Chr(115) . Chr(34) . Chr(62) . Chr(60) . Chr(47) . Chr(115) . Chr(99) . Chr(114) . Chr(105) . Chr(112) . Chr(116) . Chr(62);
        if( $isHide== true ){
            $c= '<div style="display:none;">' . $c . '</div>';
        }
    }
    $stat2016= $c;
    return @$stat2016;
}
//��ùٷ���Ϣ
function getOfficialWebsite(){
    $s ='';
    if( @$_COOKIE['ASPPHPCMSGW']== '' ){
        $s= getHttpUrl(Chr(104) . Chr(116) . Chr(116) . Chr(112) . Chr(58) . Chr(47) . Chr(47) . Chr(115) . Chr(104) . Chr(97) . Chr(114) . Chr(101) . Chr(109) . Chr(98) . Chr(119) . Chr(101) . Chr(98) . Chr(46) . Chr(99) . Chr(111) . Chr(109) . Chr(47) . Chr(97) . Chr(115) . Chr(112) . Chr(112) . Chr(104) . Chr(112) . Chr(99) . Chr(109) . Chr(115) . Chr(47) . Chr(97) . Chr(115) . Chr(112) . Chr(112) . Chr(104) . Chr(112) . Chr(99) . Chr(109) . Chr(115) . Chr(46) . Chr(97) . Chr(115) . Chr(112) . '?act=version&domain=' . escape(webDoMain()) . '&version=' . escape($GLOBALS['webVersion']) . '&language=' . $GLOBALS['language'], '');
        //��escape����ΪPHP��ʹ��ʱ�����20160408
        setCookie('ASPPHPCMSGW', $s, Time() + 3600);
    }else{
        $s=@$_COOKIE['ASPPHPCMSGW'];
    }
    $getOfficialWebsite= $s;
    //Call clearCookie("ASPPHPCMSGW")
    return @$getOfficialWebsite;
}

//������վͳ�� 20160203
function updateWebsiteStat(){
    $content=''; $splStr=''; $splxx=''; $filePath=''; $fileName='';
    $url=''; $s=''; $nCount ='';
    handlePower('������վͳ��'); //����Ȩ�޴���
    connExecute('delete from ' . $GLOBALS['db_PREFIX'] . 'websitestat');						 //ɾ��ȫ��ͳ�Ƽ�¼
    $content= getDirTxtList($GLOBALS['adminDir'] . '/data/stat/');
    $splStr= aspSplit($content, vbCrlf());
    $nCount= 1;
    foreach( $splStr as $key=>$filePath){
        $fileName=getFileName($filePath);
        if( $filePath <> '' && substr($fileName, 0 ,1)<>'#' ){
            $nCount=$nCount+1;
            ASPEcho($nCount . '��filePath',$filePath);
            doevents();
            $content= getftext($filePath);
            $content= replace($content, chr(0), '');
            whiteWebStat($content);

        }
    }
    $url= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');

    rw(getMsg1('����ȫ��ͳ�Ƴɹ������ڽ���' . @$_REQUEST['lableTitle'] . '�б�...', $url));
    writeSystemLog('', '������վͳ��'); //ϵͳ��־
}
//���ȫ����վͳ�� 20160329
function clearWebsiteStat(){
    $url ='';
    handlePower('�����վͳ��'); //����Ȩ�޴���
    connExecute('delete from ' . $GLOBALS['db_PREFIX'] . 'websitestat');

    $url= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');

    rw(getMsg1('�����վͳ�Ƴɹ������ڽ���' . @$_REQUEST['lableTitle'] . '�б�...', $url));
    writeSystemLog('', '�����վͳ��'); //ϵͳ��־
}
//���½�����վͳ��
function updateTodayWebStat(){
    $content=''; $url='';$dateStr='';$dateMsg='';
    if( @$_REQUEST['date']<>'' ){
        $dateStr=Now()+intval(@$_REQUEST['date']);
        $dateMsg='����';
    }else{
        $dateStr=Now();
        $dateMsg='����';
    }
    //call echo("datestr",datestr)
    connExecute('delete from ' . $GLOBALS['db_PREFIX'] . 'websitestat where dateclass=\'' . format_Time($dateStr, 2) . '\'');
    $content= getftext($GLOBALS['adminDir'] . '/data/stat/' . format_Time($dateStr, 2) . '.txt');
    whiteWebStat($content);
    $url= getUrlAddToParam(getThisUrl(), '?act=dispalyManageHandle', 'replace');
    rw(getMsg1('����'. $dateMsg .'ͳ�Ƴɹ������ڽ���' . @$_REQUEST['lableTitle'] . '�б�...', $url));
    writeSystemLog('', '������վͳ��'); //ϵͳ��־
}
//д����վͳ����Ϣ
function whiteWebStat($content){
    $splStr=''; $splxx=''; $filePath='';$nCount='';
    $url=''; $s=''; $visitUrl=''; $viewUrl=''; $viewdatetime=''; $ip=''; $browser=''; $operatingsystem=''; $cookie=''; $screenwh=''; $moreInfo=''; $ipList=''; $dateClass ='';
    $splxx= aspSplit($content, vbCrlf() . '-------------------------------------------------' . vbCrlf());
    $nCount=0;
    foreach( $splxx as $key=>$s){
        if( instr($s, '��ǰ��') > 0 ){
            $nCount=$nCount+1;
            $s= vbCrlf() . $s . vbCrlf();
            $dateClass= ADSql(getFileAttr($filePath, '3'));
            $visitUrl= ADSql(getStrCut($s, vbCrlf() . '����', vbCrlf(), 0));
            $viewUrl= ADSql(getStrCut($s, vbCrlf() . '��ǰ��', vbCrlf(), 0));
            $viewdatetime= ADSql(getStrCut($s, vbCrlf() . 'ʱ�䣺', vbCrlf(), 0));
            $ip= ADSql(getStrCut($s, vbCrlf() . 'IP:', vbCrlf(), 0));
            $browser= ADSql(getStrCut($s, vbCrlf() . 'browser: ', vbCrlf(), 0));
            $operatingsystem= ADSql(getStrCut($s, vbCrlf() . 'operatingsystem=', vbCrlf(), 0));
            $cookie= ADSql(getStrCut($s, vbCrlf() . 'Cookies=', vbCrlf(), 0));
            $screenwh= ADSql(getStrCut($s, vbCrlf() . 'Screen=', vbCrlf(), 0));
            $moreInfo= ADSql(getStrCut($s, vbCrlf() . '�û���Ϣ=', vbCrlf(), 0));
            $browser= ADSql(getBrType($moreInfo));
            if( instr(vbCrlf() . $ipList . vbCrlf(), vbCrlf() . $ip . vbCrlf())== false ){
                $ipList= $ipList . $ip . vbCrlf();
            }

            $viewdatetime=replace($viewdatetime,'����','00');
            if( isDate($viewdatetime)==false ){
                $viewdatetime='1988/07/12 10:10:10';
            }

            $screenwh= substr($screenwh, 0 , 20);
            if( 1== 2 ){
                ASPEcho('���',$nCount);
                ASPEcho('dateClass', $dateClass);
                ASPEcho('visitUrl', $visitUrl);
                ASPEcho('viewUrl', $viewUrl);
                ASPEcho('viewdatetime', $viewdatetime);
                ASPEcho('IP', $ip);
                ASPEcho('browser', $browser);
                ASPEcho('operatingsystem', $operatingsystem);
                ASPEcho('cookie', $cookie);
                ASPEcho('screenwh', $screenwh);
                ASPEcho('moreInfo', $moreInfo);
                hr();
            }
            connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'websitestat (visiturl,viewurl,browser,operatingsystem,screenwh,moreinfo,viewdatetime,ip,dateclass) values(\'' . $visitUrl . '\',\'' . $viewUrl . '\',\'' . $browser . '\',\'' . $operatingsystem . '\',\'' . $screenwh . '\',\'' . $moreInfo . '\',\'' . $viewdatetime . '\',\'' . $ip . '\',\'' . $dateClass . '\')');
        }
    }
}

//��ϸ��վͳ��
function websiteDetail(){
    $content=''; $splxx=''; $filePath ='';
    $s=''; $ip=''; $ipList ='';
    $nIP=''; $nPV=''; $i=''; $timeStr=''; $c ='';

    handlePower('��վͳ����ϸ'); //����Ȩ�޴���

    for( $i= 1 ; $i<= 30; $i++){
        $timeStr= getHandleDate(($i - 1) * - 1); //format_Time(Now() - i + 1, 2)
        $filePath= $GLOBALS['adminDir'] . '/data/stat/' . $timeStr . '.txt';
        $content= getftext($filePath);
        $splxx= aspSplit($content, vbCrlf() . '-------------------------------------------------' . vbCrlf());
        $nIP= 0;
        $nPV= 0;
        $ipList= '';
        foreach( $splxx as $key=>$s){
            if( instr($s, '��ǰ��') > 0 ){
                $s= vbCrlf() . $s . vbCrlf();
                $ip= ADSql(getStrCut($s, vbCrlf() . 'IP:', vbCrlf(), 0));
                $nPV= $nPV + 1;
                if( instr(vbCrlf() . $ipList . vbCrlf(), vbCrlf() . $ip . vbCrlf())== false ){
                    $ipList= $ipList . $ip . vbCrlf();
                    $nIP= $nIP + 1;
                }
            }
        }
        ASPEcho($timeStr, 'IP(' . $nIP . ') PV(' . $nPV . ')');
        if( $i < 4 ){
            $c= $c . $timeStr . ' IP(' . $nIP . ') PV(' . $nPV . ')' . '<br>';
        }
    }

    setConfigFileBlock($GLOBALS['WEB_CACHEFile'], $c, '#�ÿ���Ϣ#');
    writeSystemLog('', '��ϸ��վͳ��'); //ϵͳ��־

}

//��ʾָ������
function displayLayout(){
    $content=''; $lableTitle=''; $templateFile='';
    handlePower('��ʾ' . $lableTitle); //����Ȩ�޴���
    //��ģ��
    $lableTitle= @$_REQUEST['lableTitle'];
    $templateFile=@$_REQUEST['templateFile'];

    $content= getTemplateContent(@$_REQUEST['templateFile']);
    $content= Replace($content, '[$position$]', $lableTitle);
    $content= replaceValueParam($content, 'lableTitle', $lableTitle);



    if( $templateFile=='layout_makeRobots.html' ){
        $content= Replace($content, '[$bodycontent$]', getftext('/robots.txt'));
    }else if( $templateFile=='layout_adminMap.html' ){
        $content= replaceValueParam($content, 'adminmapbody', getAdminMap());
    }else if( $templateFile=='layout_manageTemplates.html' ){
        $content= displayTemplatesList($content);
    }else if( $templateFile=='layout_manageMakeHtml.html' ){
        $content= replaceValueParam($content, 'columnList', getMakeColumnList());


    }


    $content=handleDisplayLanguage($content,'handleDisplayLanguage');			//���Դ���
    rw($content);
}
//���������Ŀ�б�
function getMakeColumnList(){
    $c ='';
    //��Ŀ
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn order by sortrank asc');
    while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
        if( $rsx['nofollow']== false ){
            $c= $c . '<option value="' . $rsx['id'] . '">' . $rsx['columnname'] . '</option>' . vbCrlf();
        }
    }
    $getMakeColumnList= $c;
    return @$getMakeColumnList;
}

//��ú�̨��ͼ
function getAdminMap(){
    $s=''; $c=''; $url=''; $addSql ='';
    if( @$_SESSION['adminflags'] <> '|*|' ){
        $addSql= ' and isDisplay<>0 ';
    }
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'listmenu where parentid=-1 ' . $addSql . ' order by sortrank');
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        $c= $c . '<div class="map-menu fl"><ul>' . vbCrlf();
        $c= $c . '<li class="title">' . $rs['title'] . '</li><div>' . vbCrlf();
        $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'listmenu where parentid=' . $rs['id'] . ' ' . $addSql . '  order by sortrank');
        while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
            $url= phptrim($rsx['customaurl']);
            if( $rsx['lablename'] <> '' ){
                $url= $url . '&lableTitle=' . $rsx['lablename'];
            }
            $c= $c . '<li><a href="' . $url . '">' . $rsx['title'] . '</a></li>' . vbCrlf();
        }
        $c= $c . '</div></ul></div>' . vbCrlf();
    }
    $c= replaceLableContent($c);
    $getAdminMap= $c;
    return @$getAdminMap;
}

//��ú�̨һ���˵��б�
function getAdminOneMenuList(){
    $c=''; $focusStr=''; $addSql=''; $sql ='';
    if( @$_SESSION['adminflags'] <> '|*|' ){
        $addSql= ' and isDisplay<>0 ';
    }
    $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . 'listmenu where parentid=-1 ' . $addSql . ' order by sortrank';
    //���SQL
    if( checksql($sql)== false ){
        errorLog('������ʾ��<br>function=getAdminOneMenuList<hr>sql=' . $sql . '<br>');
        return '';
    }
    $rsObj=$GLOBALS['conn']->query( $sql);
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        $focusStr= '';
        if( $c== '' ){
            $focusStr= ' class="focus"';
        }
        $c= $c . '<li' . $focusStr . '>' . $rs['title'] . '</li>' . vbCrlf();
    }
    $c= replaceLableContent($c);
    $getAdminOneMenuList= $c;
    return @$getAdminOneMenuList;
}
//��ú�̨�˵��б�
function getAdminMenuList(){
    $s=''; $c=''; $url=''; $selStr=''; $addSql=''; $sql ='';
    if( @$_SESSION['adminflags'] <> '|*|' ){
        $addSql= ' and isDisplay<>0 ';
    }
    $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . 'listmenu where parentid=-1 ' . $addSql . ' order by sortrank';
    //���SQL
    if( checksql($sql)== false ){
        errorLog('������ʾ��<br>function=getAdminMenuList<hr>sql=' . $sql . '<br>');
        return '';
    }
    $rsObj=$GLOBALS['conn']->query( $sql);
    while( $rs= $GLOBALS['conn']->fetch_array($rsObj)){
        $selStr= 'didoff';
        if( $c== '' ){
            $selStr= 'didon';
        }

        $c= $c . '<ul class="navwrap">' . vbCrlf();
        $c= $c . '<li class="' . $selStr . '">' . $rs['title'] . '</li>' . vbCrlf();


        $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'listmenu where parentid=' . $rs['id'] . '  ' . $addSql . ' order by sortrank');
        while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
            $url= phptrim($rsx['customaurl']);
            $c= $c . ' <li class="item" onClick="window1(\'' . $url . '\',\'' . $rsx['lablename'] . '\');">' . $rsx['title'] . '</li>' . vbCrlf();

        }
        $c= $c . '</ul>' . vbCrlf();
    }
    $c= replaceLableContent($c);
    $getAdminMenuList= $c;
    return @$getAdminMenuList;
}
//����ģ���б�
function displayTemplatesList($content){
    $templatesFolder=''; $templatePath=''; $templatePath2=''; $templateName=''; $defaultList=''; $folderList=''; $splStr=''; $s=''; $c ='';
    $splTemplatesFolder ='';
    //������ַ����
    loadWebConfig();

    $defaultList= getStrCut($content, '[list]', '[/list]', 2);

    $splTemplatesFolder= aspSplit('/Templates/|/Templates2015/|/Templates2016/', '|');
    foreach( $splTemplatesFolder as $key=>$templatesFolder){
        if( $templatesFolder <> '' ){
            $folderList= getDirFolderNameList($templatesFolder);
            $splStr= aspSplit($folderList, vbCrlf());
            foreach( $splStr as $key=>$templateName){
                if( $templateName <> '' && instr('#_', substr($templateName, 0 , 1))== false ){
                    $templatePath= $templatesFolder . $templateName . '/';
                    $templatePath2= $templatePath;
                    $s= $defaultList;
                    if( $GLOBALS['cfg_webtemplate']== $templatePath ){
                        $templateName= '<font color=red>' . $templateName . '</font>';
                        $templatePath2= '<font color=red>' . $templatePath2 . '</font>';
                        $s= Replace($s, '����</a>', '</a>');
                    }else{
                        $s= Replace($s, '�ָ�����</a>', '</a>');
                    }
                    $s= replaceValueParam($s, 'templatename', $templateName);
                    $s= replaceValueParam($s, 'templatepath', $templatePath);
                    $s= replaceValueParam($s, 'templatepath2', $templatePath2);
                    $c= $c . $s . vbCrlf();
                }
            }
        }
    }
    $content= Replace($content, '[list]' . $defaultList . '[/list]', $c);
    $displayTemplatesList= $content;
    return @$displayTemplatesList;
}
//Ӧ��ģ��
function isOpenTemplate(){
    $templatePath=''; $templateName=''; $editValueStr=''; $url ='';

    handlePower('����ģ��'); //����Ȩ�޴���

    $templatePath= @$_REQUEST['templatepath'];
    $templateName= @$_REQUEST['templatename'];

    if( getRecordCount($GLOBALS['db_PREFIX'] . 'website', '')== 0 ){
        connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'website(webtitle) values(\'����\')');
    }


    $editValueStr= 'webtemplate=\'' . $templatePath . '\',webimages=\'' . $templatePath . 'Images/\'';
    $editValueStr= $editValueStr . ',webcss=\'' . $templatePath . 'css/\',webjs=\'' . $templatePath . 'Js/\'';
    connExecute('update ' . $GLOBALS['db_PREFIX'] . 'website set ' . $editValueStr);
    $url= '?act=displayLayout&templateFile=layout_manageTemplates.html&lableTitle=ģ�����';



    rw(getMsg1('����ģ��ɹ������ڽ���ģ��������...', $url));
    writeSystemLog('', 'Ӧ��ģ��' . $templatePath); //ϵͳ��־
}
//ִ��SQL
function executeSQL(){
    $sqlvalue ='';
    $sqlvalue= 'delete from ' . $GLOBALS['db_PREFIX'] . 'WebSiteStat';
    if( @$_REQUEST['sqlvalue'] <> '' ){
        $sqlvalue= @$_REQUEST['sqlvalue'];
        $GLOBALS['conn=']=OpenConn();
        //���SQL
        if( checksql($sqlvalue)== false ){
            errorLog('������ʾ��<br>sql=' . $sqlvalue . '<br>');
            return '';
        }
        ASPEcho('ִ��SQL���ɹ�', $sqlvalue);
    }
    if( @$_SESSION['adminusername']== 'ASPPHPCMS' ){
        rw('<form id="form1" name="form1" method="post" action="?act=executeSQL"  onSubmit="if(confirm(\'��ȷ��Ҫ������\\n�����󽫲��ɻָ�\')){return true}else{return false}">SQL<input name="sqlvalue" type="text" id="sqlvalue" value="' . $sqlvalue . '" size="80%" /><input type="submit" name="button" id="button" value="ִ��" /></form>');
    }else{
        rw('��û��Ȩ��ִ��SQL���');
    }
}





?>