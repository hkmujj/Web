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
require_once './PHP2/Web/Inc/2014_Template.php'; 
require_once './PHP2/Web/Inc/FunHTML.php';  
require_once './Web/setAccess.php';
require_once './Web/function.php';
require_once './Web/config.php';
  
//end ����inc
 

$ReadBlockList='';
$ModuleReplaceArray=''; //�滻ģ�����飬��ʱû�ã�����Ҫ���ţ�Ҫ��������  



  
//=========

$code ='';//html����
$templateName ='';//ģ������
$cfg_webSiteUrl=''; $cfg_webTemplate=''; $cfg_webImages=''; $cfg_webCss=''; $cfg_webJs=''; $cfg_webTitle=''; $cfg_webKeywords=''; $cfg_webDescription=''; $cfg_webSiteBottom=''; $cfg_flags ='';
$glb_columnName=''; $glb_columnId=''; $glb_id=''; $glb_columnType=''; $glb_columnENType=''; $glb_table=''; $glb_detailTitle=''; $glb_flags ='';
$webTemplate ='';//��վģ��·��
$glb_url=''; $glb_filePath ='';//��ǰ������ַ,���ļ�·��
$glb_isonhtml ='';//�Ƿ����ɾ�̬��ҳ
$glb_locationType ='';//λ������

$glb_bodyContent ='';//��������
$glb_artitleAuthor ='';//��������
$glb_artitleAdddatetime ='';//�������ʱ��
$glb_upArticle ='';//��һƪ����
$glb_downArticle ='';//��һƪ����
$glb_aritcleRelatedTags ='';//���±�ǩ��
$glb_aritcleSmallImage=''; $glb_aritcleBigImage ='';//����Сͼ�����´�ͼ
$glb_searchKeyWord ='';//�����ؼ���

$isMakeHtml ='';//�Ƿ�������ҳ
//������   ReplaceValueParamΪ�����ַ���ʾ��ʽ
function handleAction($content){
    $startStr=''; $endStr=''; $ActionList=''; $splStr=''; $action=''; $s=''; $HandYes ='';
    $startStr= '{$' ; $endStr= '$}';
    $ActionList= getArray($content, $startStr, $endStr, true, true);
    //Call echo("ActionList ", ActionList)
    $splStr= aspSplit($ActionList, '$Array$');
    foreach( $splStr as $key=>$s){
        $action= AspTrim($s);
        $action= handleInModule($action, 'start'); //����\'�滻��
        if( $action <> '' ){
            $action= AspTrim(mid($action, 3, Len($action) - 4)) . ' ';
            //call echo("s",s)
            $HandYes= true; //����Ϊ��
            //{VB #} �����Ƿ���ͼƬ·���Ŀ����Ϊ����VB�ﲻ�������·��
            if( checkFunValue($action, '# ')== true ){
                $action= '';
                //����
            }else if( checkFunValue($action, 'GetLableValue ')== true ){
                $action= XY_getLableValue($action);
                //�����������������б�
            }else if( checkFunValue($action, 'TitleInSearchEngineList ')== true ){
                $action= XY_TitleInSearchEngineList($action);

                //�����ļ�
            }else if( checkFunValue($action, 'Include ')== true ){
                $action= XY_Include($action);
                //��Ŀ�б�
            }else if( checkFunValue($action, 'ColumnList ')== true ){
                $action= XY_AP_ColumnList($action);
                //�����б�
            }else if( checkFunValue($action, 'ArticleList ')== true || checkFunValue($action, 'CustomInfoList ')== true ){
                $action= XY_AP_ArticleList($action);
                //�����б�
            }else if( checkFunValue($action, 'CommentList ')== true ){
                $action= XY_AP_CommentList($action);
                //����ͳ���б�
            }else if( checkFunValue($action, 'SearchStatList ')== true ){
                $action= XY_AP_SearchStatList($action);
                //���������б�
            }else if( checkFunValue($action, 'Links ')== true ){
                $action= XY_AP_Links($action);

                //��ʾ��ҳ����
            }else if( checkFunValue($action, 'GetOnePageBody ')== true || checkFunValue($action, 'MainInfo ')== true ){
                $action= XY_AP_GetOnePageBody($action);
                //��ʾ��������
            }else if( checkFunValue($action, 'GetArticleBody ')== true ){
                $action= XY_AP_GetArticleBody($action);
                //��ʾ��Ŀ����
            }else if( checkFunValue($action, 'GetColumnBody ')== true ){
                $action= XY_AP_GetColumnBody($action);

                //�����ĿURL
            }else if( checkFunValue($action, 'GetColumnUrl ')== true ){
                $action= XY_GetColumnUrl($action);
                //�������URL
            }else if( checkFunValue($action, 'GetArticleUrl ')== true ){
                $action= XY_GetArticleUrl($action);
                //��õ�ҳURL
            }else if( checkFunValue($action, 'GetOnePageUrl ')== true ){
                $action= XY_GetOnePageUrl($action);

                //��ʾ������ ���ò���
            }else if( checkFunValue($action, 'DisplayWrap ')== true ){
                $action= XY_DisplayWrap($action);
                //��ʾ����
            }else if( checkFunValue($action, 'Layout ')== true ){
                $action= XY_Layout($action);
                //��ʾģ��
            }else if( checkFunValue($action, 'Module ')== true ){
                $action= XY_Module($action);
                //�������ģ�� 20150108
            }else if( checkFunValue($action, 'GetContentModule ')== true ){
                $action= XY_ReadTemplateModule($action);
                //��ģ����ʽ�����ñ���������   ������и���ĿStyle��������
            }else if( checkFunValue($action, 'ReadColumeSetTitle ')== true ){
                $action= XY_ReadColumeSetTitle($action);

                //��ʾJS��ȾASP/PHP/VB�ȳ���ı༭��
            }else if( checkFunValue($action, 'displayEditor ')== true ){
                $action= displayEditor($action);

                //Js����վͳ��
            }else if( checkFunValue($action, 'JsWebStat ')== true ){
                $action= XY_JsWebStat($action);

                //------------------- ������ -----------------------
                //��ͨ����A
            }else if( checkFunValue($action, 'HrefA ')== true ){
                $action= XY_HrefA($action);

                //��Ŀ�˵�(���ú�̨��Ŀ����)
            }else if( checkFunValue($action, 'ColumnMenu ')== true ){
                $action= XY_AP_ColumnMenu($action);

                //��վ�ײ�
            }else if( checkFunValue($action, 'WebSiteBottom ')== true ){
                $action= XY_AP_WebSiteBottom($action);
                //��ʾ��վ��Ŀ 20160331
            }else if( checkFunValue($action, 'DisplayWebColumn ')== true ){
                $action= XY_DisplayWebColumn($action);
                //URL����
            }else if( checkFunValue($action, 'escape ')== true ){
                $action= XY_escape($action);
                //URL����
            }else if( checkFunValue($action, 'unescape ')== true ){
                $action= XY_unescape($action);
                //asp��php�汾
            }else if( checkFunValue($action, 'EDITORTYPE ')== true ){
                $action= XY_EDITORTYPE($action);


                //��ʱ������
            }else if( checkFunValue($action, 'copyTemplateMaterial ')== true ){
                $action= '';
            }else if( checkFunValue($action, 'clearCache ')== true ){
                $action= '';
            }else{
                $HandYes= false; //����Ϊ��
            }
            //ע���������е�����ʾ �� And IsNul(action)=False
            if( isNul($action)== true ){ $action= '' ;}
            if( $HandYes== true ){
                $content= Replace($content, $s, $action);
            }
        }
    }
    $handleAction= $content;
    return @$handleAction;
}

//��ʾ��վ��Ŀ �°� ��֮ǰ��վ ��������Ľ�������
function XY_DisplayWebColumn($action){
    $i=''; $c=''; $s=''; $url=''; $sql=''; $dropDownMenu=''; $focusType=''; $addSql ='';
    $isConcise ='';//�����ʾ20150212
    $styleId=''; $styleValue ='';//��ʽID����ʽ����
    $cssNameAddId ='';
    $shopnavidwrap ='';//�Ƿ���ʾ��ĿID��

    $styleId= PHPTrim(RParam($action, 'styleID'));
    $styleValue= PHPTrim(RParam($action, 'styleValue'));
    $addSql= PHPTrim(RParam($action, 'addSql'));
    $shopnavidwrap= PHPTrim(RParam($action, 'shopnavidwrap'));
    //If styleId <> "" Then
    //Call ReadNavCSS(styleId, styleValue)
    //End If

    //Ϊ�������� ���Զ���ȡ��ʽ����  20150615
    if( checkStrIsNumberType($styleValue) ){
        $cssNameAddId= '_' . $styleValue; //Css����׷��Id���
    }
    $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn';
    //׷��sql
    if( $addSql <> '' ){
        $sql= getWhereAnd($sql, $addSql);
    }
    if( checkSql($sql)== false ){ eerr('Sql', $sql) ;}
    $rsObj=$GLOBALS['conn']->query( $sql);
    $dropDownMenu= strtolower(RParam($action, 'DropDownMenu'));
    $focusType= strtolower(RParam($action, 'FocusType'));
    $isConcise= strtolower(RParam($action, 'isConcise'));
    if( $isConcise== 'true' ){
        $isConcise= false;
    }else{
        $isConcise= true;
    }
    if( $isConcise== true ){ $c= $c . copyStr(' ', 4) . '<li class=left></li>' . vbCrlf() ;}
    for( $i= 1 ; $i<= @mysql_num_rows($rsObj); $i++){

        $rs=mysql_fetch_array($rsObj); //��PHP�ã���Ϊ�� asptophpת��������
        $url= getColumnUrl($rs['columnname'], 'name');
        if( $rs['columnname']== $GLOBALS['glb_columnName'] ){
            if( $focusType== 'a' ){
                $s= copyStr(' ', 8) . '<li class=focus><a href="' . $url . '">' . $rs['columname'] . '</a>';
            }else{
                $s= copyStr(' ', 8) . '<li class=focus>' . $rs['columnname'];
            }
        }else{
            $s= copyStr(' ', 8) . '<li><a href="' . $url . '">' . $rs['columnname'] . '</a>';
        }
        $url= WEB_ADMINURL . '?act=addEditHandle&actionType=WebColumn&lableTitle=��վ��Ŀ&nPageSize=10&page=&id=' . $rs['id'] . '&n=' . getRnd(11);
        $s= handleDisplayOnlineEditDialog($url, $s, '', 'div|li|span'); //�����Ƿ���������޸Ĺ�����

        $c= $c . $s;

        //С��

        $c= $c . copyStr(' ', 8) . '</li>' . vbCrlf();

        if( $isConcise== true ){ $c= $c . copyStr(' ', 8) . '<li class=line></li>' . vbCrlf() ;}
    }
    if( $isConcise== true ){ $c= $c . copyStr(' ', 8) . '<li class=right></li>' . vbCrlf() ;}

    if( $styleId <> '' ){
        $c= '<ul class=\'nav' . $styleId . $cssNameAddId . '\'>' . vbCrlf() . $c . vbCrlf() . '</ul>' . vbCrlf();
    }
    if( $shopnavidwrap== '1' || $shopnavidwrap== 'true' ){
        $c= '<div id=\'nav' . $styleId . $cssNameAddId . '\'>' . vbCrlf() . $c . vbCrlf() . '</div>' . vbCrlf();
    }

    $XY_DisplayWebColumn= $c;
    return @$XY_DisplayWebColumn;
}

//�滻ȫ�ֱ��� {$cfg_websiteurl$}
function replaceGlobleVariable( $content){
    $content= handleRGV($content, '{$cfg_webSiteUrl$}', $GLOBALS['cfg_webSiteUrl']); //��ַ
    $content= handleRGV($content, '{$cfg_webTemplate$}', $GLOBALS['cfg_webTemplate']); //ģ��
    $content= handleRGV($content, '{$cfg_webImages$}', $GLOBALS['cfg_webImages']); //ͼƬ·��
    $content= handleRGV($content, '{$cfg_webCss$}', $GLOBALS['cfg_webCss']); //css·��
    $content= handleRGV($content, '{$cfg_webJs$}', $GLOBALS['cfg_webJs']); //js·��
    $content= handleRGV($content, '{$cfg_webTitle$}', $GLOBALS['cfg_webTitle']); //��վ����
    $content= handleRGV($content, '{$cfg_webKeywords$}', $GLOBALS['cfg_webKeywords']); //��վ�ؼ���
    $content= handleRGV($content, '{$cfg_webDescription$}', $GLOBALS['cfg_webDescription']); //��վ����
    $content= handleRGV($content, '{$cfg_webSiteBottom$}', $GLOBALS['cfg_webSiteBottom']); //��վ����

    $content= handleRGV($content, '{$glb_columnId$}', $GLOBALS['glb_columnId']); //��ĿId
    $content= handleRGV($content, '{$glb_columnName$}', $GLOBALS['glb_columnName']); //��Ŀ����
    $content= handleRGV($content, '{$glb_columnType$}', $GLOBALS['glb_columnType']); //��Ŀ����
    $content= handleRGV($content, '{$glb_columnENType$}', $GLOBALS['glb_columnENType']); //��ĿӢ������

    $content= handleRGV($content, '{$glb_Table$}', $GLOBALS['glb_table']); //��
    $content= handleRGV($content, '{$glb_Id$}', $GLOBALS['glb_id']); //id


    //���ݾɰ汾 ��������ȥ��
    $content= handleRGV($content, '{$WebImages$}', $GLOBALS['cfg_webImages']); //ͼƬ·��
    $content= handleRGV($content, '{$WebCss$}', $GLOBALS['cfg_webCss']); //css·��
    $content= handleRGV($content, '{$WebJs$}', $GLOBALS['cfg_webJs']); //js·��
    $content= handleRGV($content, '{$Web_Title$}', $GLOBALS['cfg_webTitle']);
    $content= handleRGV($content, '{$Web_KeyWords$}', $GLOBALS['cfg_webKeywords']);
    $content= handleRGV($content, '{$Web_Description$}', $GLOBALS['cfg_webDescription']);


    $content= handleRGV($content, '{$EDITORTYPE$}', EDITORTYPE); //��׺
    $content= handleRGV($content, '{$WEB_VIEWURL$}', WEB_VIEWURL); //��ҳ��ʾ��ַ
    //�����õ�
    $content= handleRGV($content, '{$glb_artitleAuthor$}', $GLOBALS['glb_artitleAuthor']); //��������
    $content= handleRGV($content, '{$glb_artitleAdddatetime$}', $GLOBALS['glb_artitleAdddatetime']); //�������ʱ��
    $content= handleRGV($content, '{$glb_upArticle$}', $GLOBALS['glb_upArticle']); //��һƪ����
    $content= handleRGV($content, '{$glb_downArticle$}', $GLOBALS['glb_downArticle']); //��һƪ����
    $content= handleRGV($content, '{$glb_aritcleRelatedTags$}', $GLOBALS['glb_aritcleRelatedTags']); //���±�ǩ��
    $content= handleRGV($content, '{$glb_aritcleBigImage$}', $GLOBALS['glb_aritcleBigImage']); //���´�ͼ
    $content= handleRGV($content, '{$glb_aritcleSmallImage$}', $GLOBALS['glb_aritcleSmallImage']); //����Сͼ
    $content= handleRGV($content, '{$glb_searchKeyWord$}', $GLOBALS['glb_searchKeyWord']); //��ҳ��ʾ��ַ


    $replaceGlobleVariable= $content;
    return @$replaceGlobleVariable;
}

//�����滻
function handleRGV( $content, $findStr, $replaceStr){
    $lableName ='';
    //��[$$]����
    $lableName= mid($findStr, 3, Len($findStr) - 4) . ' ';
    $lableName= mid($lableName, 1, instr($lableName, ' ') - 1);
    $content= replaceValueParam($content, $lableName, $replaceStr);
    $content= replaceValueParam($content, strtolower($lableName), $replaceStr);
    //ֱ���滻{$$}���ַ�ʽ������֮ǰ��վ
    $content= Replace($content, $findStr, $replaceStr);
    $content= Replace($content, strtolower($findStr), $replaceStr);
    $handleRGV= $content;
    return @$handleRGV;
}

//������ַ����
function loadWebConfig(){
    $templatedir ='';
    $GLOBALS['conn=']=OpenConn();
    $rsObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'website');
    $rs=mysql_fetch_array($rsObj);
    if( @mysql_num_rows($rsObj)!=0 ){
        $GLOBALS['cfg_webSiteUrl']= phptrim($rs['websiteurl']); //��ַ
        $GLOBALS['cfg_webTemplate']= $GLOBALS['webDir'] . phptrim($rs['webtemplate']); //ģ��·��
        $GLOBALS['cfg_webImages']= $GLOBALS['webDir'] . phptrim($rs['webimages']); //ͼƬ·��
        $GLOBALS['cfg_webCss']= $GLOBALS['webDir'] . phptrim($rs['webcss']); //css·��
        $GLOBALS['cfg_webJs']= $GLOBALS['webDir'] . phptrim($rs['webjs']); //js·��
        $GLOBALS['cfg_webTitle']= $rs['webtitle']; //��ַ����
        $GLOBALS['cfg_webKeywords']= $rs['webkeywords']; //��վ�ؼ���
        $GLOBALS['cfg_webDescription']= $rs['webdescription']; //��վ����
        $GLOBALS['cfg_webSiteBottom']= $rs['websitebottom']; //��վ�ص�
        $GLOBALS['cfg_flags']= $rs['flags']; //��

        //�Ļ�ģ��20160202
        if( @$_REQUEST['templatedir'] <> '' ){
            //ɾ������Ŀ¼ǰ���Ŀ¼������Ҫ�Ǹ�����20160414
            $templatedir= replace(handlePath(@$_REQUEST['templatedir']),handlePath('/'),'/');
            //call eerr("templatedir",templatedir)

            if((instr($templatedir, ':') > 0 || instr($templatedir, '..') > 0) && getIP() <> '127.0.0.1' ){
                eerr('��ʾ', 'ģ��Ŀ¼�зǷ��ַ�');
            }
            $templatedir= handlehttpurl(Replace($templatedir, handlePath('/'), '/'));

            $GLOBALS['cfg_webImages']= Replace($GLOBALS['cfg_webImages'], $GLOBALS['cfg_webTemplate'], $templatedir);
            $GLOBALS['cfg_webCss']= Replace($GLOBALS['cfg_webCss'], $GLOBALS['cfg_webTemplate'], $templatedir);
            $GLOBALS['cfg_webJs']= Replace($GLOBALS['cfg_webJs'], $GLOBALS['cfg_webTemplate'], $templatedir);
            $GLOBALS['cfg_webTemplate']= $templatedir;
        }
        $GLOBALS['webTemplate']= $GLOBALS['cfg_webTemplate'];
    }
}

//��վλ�� ������
function thisPosition($content){
    $c ='';
    $c= '<a href="/">��ҳ</a>';
    if( $GLOBALS['glb_columnName'] <> '' ){
        $c= $c . ' >> <a href="' . getColumnUrl($GLOBALS['glb_columnName'], 'name') . '">' . $GLOBALS['glb_columnName'] . '</a>';
    }
    //20160330
    if( $GLOBALS['glb_locationType']== 'detail' ){
        $c= $c . ' >> �鿴����';
    }
    //call echo("glb_locationType",glb_locationType)

    $content= Replace($content, '[$detailPosition$]', $c);
    $content= Replace($content, '[$detailTitle$]', $GLOBALS['glb_detailTitle']);
    $content= Replace($content, '[$detailContent$]', $GLOBALS['glb_bodyContent']);

    $thisPosition= $content;
    return @$thisPosition;
}

//��ʾ�����б�
function getDetailList($action, $content, $actionName, $lableTitle, $fieldNameList, $nPageSize, $nPage, $addSql){
    $GLOBALS['conn=']=OpenConn();
    $defaultList=''; $i=''; $s=''; $c=''; $tableName=''; $j=''; $splxx=''; $sql ='';
    $x=''; $url=''; $nCount ='';
    $pageInfo ='';

    $fieldName ='';//�ֶ�����
    $splFieldName ='';//�ָ��ֶ�

    $replaceStr ='';//�滻�ַ�
    $tableName= strtolower($actionName); //������
    $listFileName ='';//�б��ļ�����
    $listFileName= RParam($action, 'listFileName');
    $abcolorStr ='';//A�Ӵֺ���ɫ
    $atargetStr ='';//A���Ӵ򿪷�ʽ
    $atitleStr ='';//A���ӵ�title20160407
    $anofollowStr ='';//A���ӵ�nofollow

    $id ='';
    $id= rq('id');
    checkIDSQL(@$_REQUEST['id']);

    if( $fieldNameList== '*' ){
        $fieldNameList= getHandleFieldList($GLOBALS['db_PREFIX'] . $tableName, '�ֶ��б�');
    }

    $fieldNameList= specialStrReplace($fieldNameList); //�����ַ�����
    $splFieldName= aspSplit($fieldNameList, ','); //�ֶηָ������


    $defaultList= getStrCut($content, '[list]', '[/list]', 2);
    $pageInfo= getStrCut($content, '[page]', '[/page]', 1);
    if( $pageInfo <> '' ){
        $content= Replace($content, $pageInfo, '');
    }

    $sql= 'select * from ' . $GLOBALS['db_PREFIX'] . $tableName . ' ' . $addSql;
    //���SQL
    if( checksql($sql)== false ){
        errorLog('������ʾ��<br>sql=' . $sql . '<br>');
        return '';
    }
    $rsObj=$GLOBALS['conn']->query( $sql);
    $nCount= @mysql_num_rows($rsObj);

    //Ϊ��̬��ҳ��ַ
    if( $GLOBALS['isMakeHtml']== true ){
        $url= '';
        if( Len($listFileName) > 5 ){
            $url= mid($listFileName, 1, Len($listFileName) - 5) . '[id].html';
            $url= urlAddHttpUrl($GLOBALS['cfg_webSiteUrl'], $url);
        }
    }else{
        $url= getUrlAddToParam(getUrl(), '?page=[id]', 'replace');
    }
    $content= Replace($content, '[$pageInfo$]', webPageControl($nCount, $nPageSize, $nPage, $url, $pageInfo));

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

        $s= $defaultList;
        $s= Replace($s, '[$id$]', $rs['id']);
        for( $j= 0 ; $j<= UBound($splFieldName); $j++){
            if( $splFieldName[$j] <> '' ){
                $splxx= aspSplit($splFieldName[$j] . '|||', '|');
                $fieldName= $splxx[0];
                $replaceStr= $rs[$fieldName] . '';
                $s= replaceValueParam($s, $fieldName, $replaceStr);
            }

            if( $GLOBALS['isMakeHtml']== true ){
                $url= getHandleRsUrl($rs['filename'], $rs['customaurl'], '/detail/detail' . $rs['id']);
            }else{
                $url= handleWebUrl('?act=detail&id=' . $rs['id']);
                if( $rs['customaurl'] <> '' ){
                    $url= $rs['customaurl'];
                }
            }

            //A���������ɫ
            $abcolorStr= '';
            if( instr($fieldNameList, ',titlecolor,') > 0 ){
                //A������ɫ
                if( $rs['titlecolor'] <> '' ){
                    $abcolorStr= 'color:' . $rs['titlecolor'] . ';';
                }
            }
            if( instr($fieldNameList, ',flags,') > 0 ){
                //A���ӼӴ�
                if( instr($rs['flags'], '|b|') > 0 ){
                    $abcolorStr= $abcolorStr . 'font-weight:bold;';
                }
            }
            if( $abcolorStr <> '' ){
                $abcolorStr= ' style="' . $abcolorStr . '"';
            }

            //�򿪷�ʽ2016
            if( instr($fieldNameList, ',target,') > 0 ){
                $atargetStr= IIF($rs['target'] <> '', ' target="' . $rs['target'] . '"', '');
            }

            //A��title
            if( instr($fieldNameList, ',title,') > 0 ){
                $atitleStr= IIF($rs['title'] <> '', ' title="' . $rs['title'] . '"', '');
            }

            //A��nofollow
            if( instr($fieldNameList, ',nofollow,') > 0 ){
                $anofollowStr= IIF($rs['nofollow'] <> 0, ' rel="nofollow"', '');
            }



            $s= replaceValueParam($s, 'url', $url);
            $s= replaceValueParam($s, 'abcolor', $abcolorStr); //A���Ӽ���ɫ��Ӵ�
            $s= replaceValueParam($s, 'atitle', $atitleStr); //A����title
            $s= replaceValueParam($s, 'anofollow', $anofollowStr); //A����nofollow
            $s= replaceValueParam($s, 'atarget', $atargetStr); //A���Ӵ򿪷�ʽ


        }
        //�����б�����߱༭
        $url= WEB_ADMINURL . '?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page=&parentid=&id=' . $rs['id'] . '&n=' . getRnd(11);
        $s= handleDisplayOnlineEditDialog($url, $s, '', 'div|li|span');

        $c= $c . $s;
    }
    $content= Replace($content, '[list]' . $defaultList . '[/list]', $c);

    if( $GLOBALS['isMakeHtml']== true ){
        $url= '';
        if( Len($listFileName) > 5 ){
            $url= mid($listFileName, 1, Len($listFileName) - 5) . '[id].html';
            $url= urlAddHttpUrl($GLOBALS['cfg_webSiteUrl'], $url);
        }
    }else{
        $url= getUrlAddToParam(getUrl(), '?page=[id]', 'replace');
    }

    $getDetailList= $content;
    return @$getDetailList;
}


//****************************************************
//Ĭ���б�ģ��
function defaultListTemplate(){
    $c=''; $templateHtml=''; $listTemplate=''; $lableName=''; $startStr=''; $endStr ='';

    $templateHtml= getFText($GLOBALS['cfg_webTemplate'] . '/' . $GLOBALS['templateName']);

    $lableName= 'list';
    $startStr= '<!--#' . $lableName . ' start#-->';
    $endStr= '<!--#' . $lableName . ' end#-->';
    //call rwend(templateHtml)
    if( instr($templateHtml, $startStr) > 0 && instr($templateHtml, $endStr) > 0 ){
        $listTemplate= strCut($templateHtml, $startStr, $endStr, 2);
    }else{
        $startStr= '<!--#' . $lableName;
        $endStr= '#-->';
        if( instr($templateHtml, $startStr) > 0 && instr($templateHtml, $endStr) > 0 ){
            $listTemplate= strCut($templateHtml, $startStr, $endStr, 2);
        }
    }
    if( $listTemplate== '' ){
        $c= '<ul class="list">' . vbCrlf();
        $c= $c . '[list]    <li><a href="[$url$]"[$atitle$][$atarget$][$abcolor$][$anofollow$]>[$title$]</a><span class="time">[$adddatetime format_time=\'7\'$]</span></li>' . vbCrlf();
        $c= $c . '[/list]' . vbCrlf();
        $c= $c . '</ul>' . vbCrlf();
        $c= $c . '<div class="clear10"></div>' . vbCrlf();
        $c= $c . '<div>[$pageInfo$]</div>' . vbCrlf();
        $listTemplate= $c;
    }

    $defaultListTemplate= $listTemplate;
    return @$defaultListTemplate;
}


//��¼��ǰ׺
if( @$_REQUEST['db_PREFIX'] <> '' ){
    $db_PREFIX= @$_REQUEST['db_PREFIX'];
}else if( @$_SESSION['db_PREFIX'] <> '' ){
    $db_PREFIX= @$_SESSION['db_PREFIX'];
}
//������ַ����
loadWebConfig();
$isMakeHtml= false; //Ĭ������HTMLΪ�ر�
if( @$_REQUEST['isMakeHtml']== '1' || @$_REQUEST['isMakeHtml']== 'true' ){
    $isMakeHtml= true;
}
$templateName= @$_REQUEST['templateName']; //ģ������

//�������ݴ���ҳ
switch ( @$_REQUEST['act'] ){
    case 'savedata' ; saveData(@$_REQUEST['stype']) ; die(); //��������
    break;//'վ��ͳ�� | ����IP[653] | ����PV[9865] | ��ǰ����[65]')
    case 'webstat' ; webStat($adminDir . '/Data/Stat/');die(); //��վͳ��
    break;
    case 'saveSiteMap' ; $isMakeHtml=true;saveSiteMap() ;die(); //����sitemap.xml
    break;
    case 'handleAction';
    if( @$_REQUEST['ishtml']=='1' ){
        $isMakeHtml= true;
    }
    rwend(handleAction(@$_REQUEST['content']));		//������

}

//����html
if( @$_REQUEST['act']== 'makehtml' ){
    ASPEcho('makehtml', 'makehtml');
    $isMakeHtml= true;
    makeWebHtml(' action actionType=\'' . @$_REQUEST['act'] . '\' columnName=\'' . @$_REQUEST['columnName'] . '\' id=\'' . @$_REQUEST['id'] . '\' ');
    createFileGBK('index.html', $code);

    //����Html����վ
}else if( @$_REQUEST['act']== 'copyHtmlToWeb' ){
    copyHtmlToWeb();
    //ȫ������
}else if( @$_REQUEST['act']== 'makeallhtml' ){
    makeAllHtml('', '', @$_REQUEST['id']);

    //���ɵ�ǰҳ��
}else if( @$_REQUEST['isMakeHtml'] <> '' && @$_REQUEST['isSave'] <> '' ){

    handlePower('���ɵ�ǰHTMLҳ��'); //����Ȩ�޴���
    writeSystemLog('', '���ɵ�ǰHTMLҳ��'); //ϵͳ��־

    $isMakeHtml= true;


    checkIDSQL(@$_REQUEST['id']);
    rw(makeWebHtml(' action actionType=\'' . @$_REQUEST['act'] . '\' columnName=\'' . @$_REQUEST['columnName'] . '\' columnType=\'' . @$_REQUEST['columnType'] . '\' id=\'' . @$_REQUEST['id'] . '\' npage=\'' . @$_REQUEST['page'] . '\' '));
    $glb_filePath= Replace($glb_url, $cfg_webSiteUrl, '');
    if( Right($glb_filePath, 1)== '/' ){
        $glb_filePath= $glb_filePath . 'index.html';
    }else if( $glb_filePath== '' && $glb_columnType== '��ҳ' ){
        $glb_filePath= 'index.html';
    }
    //�ļ���Ϊ��  ���ҿ�������html
    if( $glb_filePath <> '' && $glb_isonhtml== true ){
        createDirFolder(getFileAttr($glb_filePath, '1'));
        createFileGBK($glb_filePath, $code);
        if( @$_REQUEST['act']== 'detail' ){
            connExecute('update ' . $db_PREFIX . 'ArticleDetail set ishtml=true where id=' . @$_REQUEST['id']);
        }else if( @$_REQUEST['act']== 'nav' ){
            if( @$_REQUEST['id'] <> '' ){
                connExecute('update ' . $db_PREFIX . 'WebColumn set ishtml=true where id=' . @$_REQUEST['id']);
            }else{
                connExecute('update ' . $db_PREFIX . 'WebColumn set ishtml=true where columnname=\'' . @$_REQUEST['columnName'] . '\'');
            }
        }
        ASPEcho('�����ļ�·��', '<a href="' . $glb_filePath . '" target=\'_blank\'>' . $glb_filePath . '</a>');

        //�������������� 20160216
        if( $glb_columnType== '����' ){
            makeAllHtml('', '', $glb_columnId);
        }

    }

    //ȫ������
}else if( @$_REQUEST['act']== 'Search' ){
    rw(makeWebHtml('actionType=\'Search\' npage=\'1\' '));
}else{
    if( strtolower(@$_REQUEST['issave'])== '1' ){
        makeAllHtml(@$_REQUEST['columnType'], @$_REQUEST['columnName'], @$_REQUEST['columnId']);
    }else{
        checkIDSQL(@$_REQUEST['id']);
        rw(makeWebHtml(' action actionType=\'' . @$_REQUEST['act'] . '\' columnName=\'' . @$_REQUEST['columnName'] . '\' columnType=\'' . @$_REQUEST['columnType'] . '\' id=\'' . @$_REQUEST['id'] . '\' npage=\'' . @$_REQUEST['page'] . '\' '));
    }
}
//���ID�Ƿ�SQL��ȫ
function checkIDSQL($id){
    if( checkNumber($id)== false && $id <> '' ){
        eerr('��ʾ', 'id���зǷ��ַ�');
    }
}




//http://127.0.0.1/aspweb.asp?act=nav&columnName=ASP
//http://127.0.0.1/aspweb.asp?act=detail&id=75
//����html��̬ҳ
function makeWebHtml($action){
    $actionType=''; $npagesize=''; $npage=''; $url=''; $addSql=''; $sortSql ='';
    $actionType= RParam($action, 'actionType');
    $npage= RParam($action, 'npage');
    $npage= getnumber($npage);
    if( $npage== '' ){
        $npage= 1;
    }else{
        $npage= intval($npage);
    }
    //����
    if( $actionType== 'nav' ){
        $GLOBALS['glb_columnType']= RParam($action, 'columnType');
        $GLOBALS['glb_columnName']= RParam($action, 'columnName');
        $GLOBALS['glb_columnId']= RParam($action, 'columnId');
        if( $GLOBALS['glb_columnId']== '' ){
            $GLOBALS['glb_columnId']= RParam($action, 'id');
        }
        if( $GLOBALS['glb_columnType'] <> '' ){
            $addSql= 'where columnType=\'' . $GLOBALS['glb_columnType'] . '\'';
        }
        if( $GLOBALS['glb_columnName'] <> '' ){
            $addSql= getWhereAnd($addSql, 'where columnName=\'' . $GLOBALS['glb_columnName'] . '\'');
        }
        if( $GLOBALS['glb_columnId'] <> '' ){
            $addSql= getWhereAnd($addSql, 'where id=' . $GLOBALS['glb_columnId'] . '');
        }
        //call echo("addsql",addsql)
        $rsObj=$GLOBALS['conn']->query( 'Select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn ' . $addSql);
        $rs=mysql_fetch_array($rsObj);
        if( @mysql_num_rows($rsObj)!=0 ){
            $GLOBALS['glb_columnId']= $rs['id'];
            $GLOBALS['glb_columnName']= $rs['columnname'];
            $GLOBALS['glb_columnType']= $rs['columntype'];
            $GLOBALS['glb_bodyContent']= $rs['bodycontent'];
            $GLOBALS['glb_detailTitle']= $GLOBALS['glb_columnName'];
            $GLOBALS['glb_flags']= $rs['flags'];
            $npagesize= $rs['npagesize']; //ÿҳ��ʾ����
            $GLOBALS['glb_isonhtml']= $rs['isonhtml']; //�Ƿ����ɾ�̬��ҳ
            $sortSql= ' ' . $rs['sortsql']; //����SQL

            if( $rs['webtitle'] <> '' ){
                $GLOBALS['cfg_webTitle']= $rs['webtitle']; //��ַ����
            }
            if( $rs['webkeywords'] <> '' ){
                $GLOBALS['cfg_webKeywords']= $rs['webkeywords']; //��վ�ؼ���
            }
            if( $rs['webdescription'] <> '' ){
                $GLOBALS['cfg_webDescription']= $rs['webdescription']; //��վ����
            }
            if( $GLOBALS['templateName']== '' ){
                if( AspTrim($rs['templatepath']) <> '' ){
                    $GLOBALS['templateName']= $rs['templatepath'];
                }else if( $rs['columntype'] <> '��ҳ' ){
                    $GLOBALS['templateName']= getDateilTemplate($rs['id'], 'List');
                }
            }
        }
        $GLOBALS['glb_columnENType']= handleColumnType($GLOBALS['glb_columnType']);
        $GLOBALS['glb_url']= getColumnUrl($GLOBALS['glb_columnName'], 'name');

        //�������б�
        if( instr('|��Ʒ|����|��Ƶ|����|����|', '|' . $GLOBALS['glb_columnType'] . '|') > 0 ){
            $GLOBALS['glb_bodyContent']= getDetailList($action, defaultListTemplate(), 'ArticleDetail', '��Ŀ�б�', '*', $npagesize, $npage, 'where parentid=' . $GLOBALS['glb_columnId'] . $sortSql);
            //�������б�
        }else if( instr('|����|', '|' . $GLOBALS['glb_columnType'] . '|') > 0 ){
            $GLOBALS['glb_bodyContent']= getDetailList($action, defaultListTemplate(), 'GuestBook', '�����б�', '*', $npagesize, $npage, ' where isthrough<>0 ' . $sortSql);
        }else if( $GLOBALS['glb_columnType']== '�ı�' ){
            //������Ŀ�ӹ���
            if( @$_REQUEST['gl']== 'edit' ){
                $GLOBALS['glb_bodyContent']= '<span>' . $GLOBALS['glb_bodyContent'] . '</span>';
            }
            $url= WEB_ADMINURL . '?act=addEditHandle&actionType=WebColumn&lableTitle=��վ��Ŀ&nPageSize=10&page=&id=' . $GLOBALS['glb_columnId'] . '&n=' . getRnd(11);
            $GLOBALS['glb_bodyContent']= handleDisplayOnlineEditDialog($url, $GLOBALS['glb_bodyContent'], '', 'span');

        }
        //ϸ��
    }else if( $actionType== 'detail' ){
        $GLOBALS['glb_locationType']= 'detail';
        $rsObj=$GLOBALS['conn']->query( 'Select * from ' . $GLOBALS['db_PREFIX'] . 'articledetail where id=' . RParam($action, 'id'));
        $rs=mysql_fetch_array($rsObj);
        if( @mysql_num_rows($rsObj)!=0 ){
            $GLOBALS['glb_columnName']= getColumnName($rs['parentid']);
            $GLOBALS['glb_detailTitle']= $rs['title'];
            $GLOBALS['glb_flags']= $rs['flags'];
            $GLOBALS['glb_isonhtml']= $rs['isonhtml']; //�Ƿ����ɾ�̬��ҳ
            $GLOBALS['glb_id']= $rs['id']; //����ID
            if( $GLOBALS['isMakeHtml']== true ){
                $GLOBALS['glb_url']= getHandleRsUrl($rs['filename'], $rs['customaurl'], '/detail/detail' . $rs['id']);
            }else{
                $GLOBALS['glb_url']= handleWebUrl('?act=detail&id=' . $rs['id']);
            }

            if( $rs['webtitle'] <> '' ){
                $GLOBALS['cfg_webTitle']= $rs['webtitle']; //��ַ����
            }
            if( $rs['webkeywords'] <> '' ){
                $GLOBALS['cfg_webKeywords']= $rs['webkeywords']; //��վ�ؼ���
            }
            if( $rs['webdescription'] <> '' ){
                $GLOBALS['cfg_webDescription']= $rs['webdescription']; //��վ����
            }

            $GLOBALS['glb_artitleAuthor']= $rs['author'];
            $GLOBALS['glb_artitleAdddatetime']= $rs['adddatetime'];
            $GLOBALS['glb_upArticle']= upArticle($rs['parentid'], 'sortrank', $rs['sortrank']);
            $GLOBALS['glb_downArticle']= downArticle($rs['parentid'], 'sortrank', $rs['sortrank']);
            $GLOBALS['glb_aritcleRelatedTags']= aritcleRelatedTags($rs['relatedtags']);
            $GLOBALS['glb_aritcleSmallImage']= $rs['smallimage'];
            $GLOBALS['glb_aritcleBigImage']= $rs['bigimage'];

            //��������
            //glb_bodyContent = "<div class=""articleinfowrap"">[$articleinfowrap$]</div>" & rs("bodycontent") & "[$relatedtags$]<ul class=""updownarticlewrap"">[$updownArticle$]</ul>"
            //��һƪ���£���һƪ����
            //glb_bodyContent = Replace(glb_bodyContent, "[$updownArticle$]", upArticle(rs("parentid"), "sortrank", rs("sortrank")) & downArticle(rs("parentid"), "sortrank", rs("sortrank")))
            //glb_bodyContent = Replace(glb_bodyContent, "[$articleinfowrap$]", "��Դ��" & rs("author") & " &nbsp; ����ʱ�䣺" & format_Time(rs("adddatetime"), 1))
            //glb_bodyContent = Replace(glb_bodyContent, "[$relatedtags$]", aritcleRelatedTags(rs("relatedtags")))

            $GLOBALS['glb_bodyContent']= $rs['bodycontent'];

            //������ϸ�ӿ���
            if( @$_REQUEST['gl']== 'edit' ){
                $GLOBALS['glb_bodyContent']= '<span>' . $GLOBALS['glb_bodyContent'] . '</span>';
            }
            $url= WEB_ADMINURL . '?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page=&parentid=&id=' . RParam($action, 'id') . '&n=' . getRnd(11);
            $GLOBALS['glb_bodyContent']= handleDisplayOnlineEditDialog($url, $GLOBALS['glb_bodyContent'], '', 'span');

            if( $GLOBALS['templateName']== '' ){
                if( AspTrim($rs['templatepath']) <> '' ){
                    $GLOBALS['templateName']= $rs['templatepath'];
                }else{
                    $GLOBALS['templateName']= getDateilTemplate($rs['parentid'], 'Detail');
                }
            }

        }

        //��ҳ
    }else if( $actionType== 'onepage' ){
        $rsObj=$GLOBALS['conn']->query( 'Select * from ' . $GLOBALS['db_PREFIX'] . 'onepage where id=' . RParam($action, 'id'));
        $rs=mysql_fetch_array($rsObj);
        if( @mysql_num_rows($rsObj)!=0 ){
            $GLOBALS['glb_detailTitle']= $rs['title'];
            $GLOBALS['glb_isonhtml']= $rs['isonhtml']; //�Ƿ����ɾ�̬��ҳ
            if( $GLOBALS['isMakeHtml']== true ){
                $GLOBALS['glb_url']= getHandleRsUrl($rs['filename'], $rs['customaurl'], '/page/page' . $rs['id']);
            }else{
                $GLOBALS['glb_url']= handleWebUrl('?act=detail&id=' . $rs['id']);
            }

            if( $rs['webtitle'] <> '' ){
                $GLOBALS['cfg_webTitle']= $rs['webtitle']; //��ַ����
            }
            if( $rs['webkeywords'] <> '' ){
                $GLOBALS['cfg_webKeywords']= $rs['webkeywords']; //��վ�ؼ���
            }
            if( $rs['webdescription'] <> '' ){
                $GLOBALS['cfg_webDescription']= $rs['webdescription']; //��վ����
            }
            //����
            $GLOBALS['glb_bodyContent']= $rs['bodycontent'];


            //������ϸ�ӿ���
            if( @$_REQUEST['gl']== 'edit' ){
                $GLOBALS['glb_bodyContent']= '<span>' . $GLOBALS['glb_bodyContent'] . '</span>';
            }
            $url= WEB_ADMINURL . '?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page=&parentid=&id=' . RParam($action, 'id') . '&n=' . getRnd(11);
            $GLOBALS['glb_bodyContent']= handleDisplayOnlineEditDialog($url, $GLOBALS['glb_bodyContent'], '', 'span');


            if( $GLOBALS['templateName']== '' ){
                if( AspTrim($rs['templatepath']) <> '' ){
                    $GLOBALS['templateName']= $rs['templatepath'];
                }else{
                    $GLOBALS['templateName']= 'Main_Model.html';
                    //call echo(templateName,"templateName")
                }
            }

        }

        //����
    }else if( $actionType== 'Search' ){
        $GLOBALS['templateName']= 'Main_Model.html';
        $GLOBALS['glb_searchKeyWord']= @$_REQUEST['wd'];
        $addSql= ' where title like \'%' . $GLOBALS['glb_searchKeyWord'] . '%\'';
        $npagesize= 20;
        //call echo(npagesize, npage)
        $GLOBALS['glb_bodyContent']= getDetailList($action, defaultListTemplate(), 'ArticleDetail', '��վ��Ŀ', '*', $npagesize, $npage, $addSql);

        //���صȴ�
    }else if( $actionType== 'loading' ){
        rwend('ҳ�����ڼ����С�����');
    }
    //ģ��Ϊ�գ�����Ĭ����ҳģ��
    if( $GLOBALS['templateName']== '' ){
        $GLOBALS['templateName']= 'Index_Model.html'; //Ĭ��ģ��
    }
    //��⵱ǰ·���Ƿ���ģ��
    if( instr($GLOBALS['templateName'], '/')== false ){
        $GLOBALS['templateName']= $GLOBALS['cfg_webTemplate'] . '/' . $GLOBALS['templateName'];
    }
    //call echo("templateName",templateName)
    $GLOBALS['code']= getftext($GLOBALS['templateName']);


    $GLOBALS['code']= handleAction($GLOBALS['code']); //������
    $GLOBALS['code']= thisPosition($GLOBALS['code']); //λ��
    $GLOBALS['code']= replaceGlobleVariable($GLOBALS['code']); //�滻ȫ�ֱ�ǩ
    $GLOBALS['code']= handleAction($GLOBALS['code']); //������    '����һ�Σ��������������ﶯ��

    $GLOBALS['code']= handleAction($GLOBALS['code']); //������
    $GLOBALS['code']= handleAction($GLOBALS['code']); //������
    $GLOBALS['code']= thisPosition($GLOBALS['code']); //λ��
    $GLOBALS['code']= replaceGlobleVariable($GLOBALS['code']); //�滻ȫ�ֱ�ǩ
    $GLOBALS['code']= delTemplateMyNote($GLOBALS['code']); //ɾ����������

    //��ʽ��HTML
    if( instr($GLOBALS['cfg_flags'], '|formattinghtml|') > 0 ){
        //code = HtmlFormatting(code)        '��
        $GLOBALS['code']= handleHtmlFormatting($GLOBALS['code'], false, 0, 'ɾ������'); //�Զ���
        //��ʽ��HTML�ڶ���
    }else if( instr($GLOBALS['cfg_flags'], '|formattinghtmltow|') > 0 ){
        $GLOBALS['code']= htmlFormatting($GLOBALS['code']); //��
        $GLOBALS['code']= handleHtmlFormatting($GLOBALS['code'], false, 0, 'ɾ������'); //�Զ���
        //ѹ��HTML
    }else if( instr($GLOBALS['cfg_flags'], '|ziphtml|') > 0 ){
        $GLOBALS['code']= ziphtml($GLOBALS['code']);

    }
    //�պϱ�ǩ
    if( instr($GLOBALS['cfg_flags'], '|labelclose|') > 0 ){
        $GLOBALS['code']= handleCloseHtml($GLOBALS['code'], true, ''); //ͼƬ�Զ���alt  "|*|",
    }

    //���߱༭20160127
    if( rq('gl')== 'edit' ){
        if( instr($GLOBALS['code'], '</head>') > 0 ){
            if( instr($GLOBALS['code'], 'jquery.Min.js')== false ){
                $GLOBALS['code']= Replace($GLOBALS['code'], '</head>', '<script src="/Jquery/jquery.Min.js"></script></head>');
            }
            $GLOBALS['code']= Replace($GLOBALS['code'], '</head>', '<script src="/Jquery/Callcontext_menu.js"></script></head>');
        }
        if( instr($GLOBALS['code'], '<body>') > 0 ){
            //Code = Replace(Code,"<body>", "<body onLoad=""ContextMenu.intializeContextMenu()"">")
        }
    }
    //call echo(templateName,templateName)
    $makeWebHtml= $GLOBALS['code'];
    return @$makeWebHtml;
}

//���Ĭ��ϸ��ģ��ҳ
function getDateilTemplate($parentid, $templateType){
    $templateName ='';
    $templateName= 'Main_Model.html';
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn where id=' . $parentid);
    $rsx=mysql_fetch_array($rsxObj);
    if( @mysql_num_rows($rsxObj)!=0 ){
        //call echo("columntype",rsx("columntype"))
        if( $rsx['columntype']== '����' ){
            //����ϸ��ҳ
            if( checkFile($GLOBALS['cfg_webTemplate'] . '/News_' . $templateType . '.html')== true ){
                $templateName= 'News_' . $templateType . '.html';
            }
        }else if( $rsx['columntype']== '��Ʒ' ){
            //��Ʒϸ��ҳ
            if( checkFile($GLOBALS['cfg_webTemplate'] . '/Product_' . $templateType . '.html')== true ){
                $templateName= 'Product_' . $templateType . '.html';
            }
        }else if( $rsx['columntype']== '����' ){
            //����ϸ��ҳ
            if( checkFile($GLOBALS['cfg_webTemplate'] . '/Down_' . $templateType . '.html')== true ){
                $templateName= 'Down_' . $templateType . '.html';
            }

        }else if( $rsx['columntype']== '��Ƶ' ){
            //��Ƶϸ��ҳ
            if( checkFile($GLOBALS['cfg_webTemplate'] . '/Video_' . $templateType . '.html')== true ){
                $templateName= 'Video_' . $templateType . '.html';
            }
        }else if( $rsx['columntype']== '����' ){
            //��Ƶϸ��ҳ
            if( checkFile($GLOBALS['cfg_webTemplate'] . '/GuestBook_' . $templateType . '.html')== true ){
                $templateName= 'Video_' . $templateType . '.html';
            }
        }else if( $rsx['columntype']== '�ı�' ){
            //��Ƶϸ��ҳ
            if( checkFile($GLOBALS['cfg_webTemplate'] . '/Page_' . $templateType . '.html')== true ){
                $templateName= 'Page_' . $templateType . '.html';
            }
        }
    }
    //call echo(templateType,templateName)
    $getDateilTemplate= $templateName;

    return @$getDateilTemplate;
}


//����ȫ��htmlҳ��
function makeAllHtml($columnType, $columnName, $columnId){
    $action=''; $s=''; $i=''; $nPageSize=''; $nCountSize=''; $nPage=''; $addSql=''; $url=''; $articleSql ='';
    handlePower('����ȫ��HTMLҳ��'); //����Ȩ�޴���
    writeSystemLog('', '����ȫ��HTMLҳ��'); //ϵͳ��־

    $GLOBALS['isMakeHtml']= true;
    //��Ŀ
    ASPEcho('��Ŀ', '');
    if( $columnType <> '' ){
        $addSql= 'where columnType=\'' . $columnType . '\'';
    }
    if( $columnName <> '' ){
        $addSql= getWhereAnd($addSql, 'where columnName=\'' . $columnName . '\'');
    }
    if( $columnId <> '' ){
        $addSql= getWhereAnd($addSql, 'where id in(' . $columnId . ')');
    }
    $rssObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn ' . $addSql . ' order by sortrank asc');
    while( $rss= $GLOBALS['conn']->fetch_array($rssObj)){
        $GLOBALS['glb_columnName']= '';
        //��������html
        if( $rss['isonhtml']== true ){
            if( instr('|��Ʒ|����|��Ƶ|����|����|����|����|��Ƹ|����|', '|' . $rss['columntype'] . '|') > 0 ){
                if( $rss['columntype']== '����' ){
                    $nCountSize= getRecordCount($GLOBALS['db_PREFIX'] . 'guestbook', ''); //��¼��
                }else{
                    $nCountSize= getRecordCount($GLOBALS['db_PREFIX'] . 'articledetail', ' where parentid=' . $rss['id']); //��¼��
                }
                $nPageSize= $rss['npagesize'];
                $nPage= getPageNumb(intval($nCountSize), intval($nPageSize));
                if( $nPage <= 0 ){
                    $nPage= 1;
                }
                for( $i= 1 ; $i<= $nPage; $i++){
                    $url= getHandleRsUrl($rss['filename'], $rss['customaurl'], '/nav' . $rss['id']);
                    $GLOBALS['glb_filePath']= Replace($url, $GLOBALS['cfg_webSiteUrl'], '');
                    if( Right($GLOBALS['glb_filePath'], 1)== '/' || $GLOBALS['glb_filePath']== '' ){
                        $GLOBALS['glb_filePath']= $GLOBALS['glb_filePath'] . 'index.html';
                    }
                    //call echo("glb_filePath",glb_filePath)
                    $action= ' action actionType=\'nav\' columnName=\'' . $rss['columnname'] . '\' npage=\'' . $i . '\' listfilename=\'' . $GLOBALS['glb_filePath'] . '\' ';
                    //call echo("action",action)
                    makeWebHtml($action);
                    if( $i > 1 ){
                        $GLOBALS['glb_filePath']= mid($GLOBALS['glb_filePath'], 1, Len($GLOBALS['glb_filePath']) - 5) . $i . '.html';
                    }
                    $s= '<a href="' . $GLOBALS['glb_filePath'] . '" target=\'_blank\'>' . $GLOBALS['glb_filePath'] . '</a>(' . $rss['isonhtml'] . ')';
                    ASPEcho($action, $s);
                    if( $GLOBALS['glb_filePath'] <> '' ){
                        createDirFolder(getFileAttr($GLOBALS['glb_filePath'], '1'));
                        createFileGBK($GLOBALS['glb_filePath'], $GLOBALS['code']);
                    }
                    doevents();
                    $GLOBALS['templateName']= ''; //���ģ���ļ�����
                }
            }else{
                $action= ' action actionType=\'nav\' columnName=\'' . $rss['columnname'] . '\'';
                makeWebHtml($action);
                $GLOBALS['glb_filePath']= Replace(getColumnUrl($rss['columnname'], 'name'), $GLOBALS['cfg_webSiteUrl'], '');
                if( Right($GLOBALS['glb_filePath'], 1)== '/' || $GLOBALS['glb_filePath']== '' ){
                    $GLOBALS['glb_filePath']= $GLOBALS['glb_filePath'] . 'index.html';
                }
                $s= '<a href="' . $GLOBALS['glb_filePath'] . '" target=\'_blank\'>' . $GLOBALS['glb_filePath'] . '</a>(' . $rss['isonhtml'] . ')';
                ASPEcho($action, $s);
                if( $GLOBALS['glb_filePath'] <> '' ){
                    createDirFolder(getFileAttr($GLOBALS['glb_filePath'], '1'));
                    createFileGBK($GLOBALS['glb_filePath'], $GLOBALS['code']);
                }
                doevents();
                $GLOBALS['templateName']= '';
            }
            connExecute('update ' . $GLOBALS['db_PREFIX'] . 'WebColumn set ishtml=true where id=' . $rss['id']); //���µ���Ϊ����״̬
        }
    }

    //��������ָ����Ŀ��Ӧ����
    if( $columnId <> '' ){
        $articleSql= 'select * from ' . $GLOBALS['db_PREFIX'] . 'articledetail where parentid=' . $columnId . ' order by sortrank asc';
        //������������
    }else if( $addSql== '' ){
        $articleSql= 'select * from ' . $GLOBALS['db_PREFIX'] . 'articledetail order by sortrank asc';
    }
    if( $articleSql <> '' ){
        //����
        ASPEcho('����', '');
        $rssObj=$GLOBALS['conn']->query( $articleSql);
        while( $rss= $GLOBALS['conn']->fetch_array($rssObj)){
            $GLOBALS['glb_columnName']= '';
            $action= ' action actionType=\'detail\' columnName=\'' . $rss['parentid'] . '\' id=\'' . $rss['id'] . '\'';
            //call echo("action",action)
            makeWebHtml($action);
            $GLOBALS['glb_filePath']= Replace($GLOBALS['glb_url'], $GLOBALS['cfg_webSiteUrl'], '');
            if( Right($GLOBALS['glb_filePath'], 1)== '/' ){
                $GLOBALS['glb_filePath']= $GLOBALS['glb_filePath'] . 'index.html';
            }
            $s= '<a href="' . $GLOBALS['glb_filePath'] . '" target=\'_blank\'>' . $GLOBALS['glb_filePath'] . '</a>(' . $rss['isonhtml'] . ')';
            ASPEcho($action, $s);
            //�ļ���Ϊ��  ���ҿ�������html
            if( $GLOBALS['glb_filePath'] <> '' && $rss['isonhtml']== true ){
                createDirFolder(getFileAttr($GLOBALS['glb_filePath'], '1'));
                createFileGBK($GLOBALS['glb_filePath'], $GLOBALS['code']);
                connExecute('update ' . $GLOBALS['db_PREFIX'] . 'ArticleDetail set ishtml=true where id=' . $rss['id']); //��������Ϊ����״̬
            }
            $GLOBALS['templateName']= ''; //���ģ���ļ�����
        }
    }

    if( $addSql== '' ){
        //��ҳ
        ASPEcho('��ҳ', '');
        $rssObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'onepage order by sortrank asc');
        while( $rss= $GLOBALS['conn']->fetch_array($rssObj)){
            $GLOBALS['glb_columnName']= '';
            $action= ' action actionType=\'onepage\' id=\'' . $rss['id'] . '\'';
            //call echo("action",action)
            makeWebHtml($action);
            $GLOBALS['glb_filePath']= Replace($GLOBALS['glb_url'], $GLOBALS['cfg_webSiteUrl'], '');
            if( Right($GLOBALS['glb_filePath'], 1)== '/' ){
                $GLOBALS['glb_filePath']= $GLOBALS['glb_filePath'] . 'index.html';
            }
            $s= '<a href="' . $GLOBALS['glb_filePath'] . '" target=\'_blank\'>' . $GLOBALS['glb_filePath'] . '</a>(' . $rss['isonhtml'] . ')';
            ASPEcho($action, $s);
            //�ļ���Ϊ��  ���ҿ�������html
            if( $GLOBALS['glb_filePath'] <> '' && $rss['isonhtml']== true ){
                createDirFolder(getFileAttr($GLOBALS['glb_filePath'], '1'));
                createFileGBK($GLOBALS['glb_filePath'], $GLOBALS['code']);
                connExecute('update ' . $GLOBALS['db_PREFIX'] . 'onepage set ishtml=true where id=' . $rss['id']); //���µ�ҳΪ����״̬
            }
            $GLOBALS['templateName']= ''; //���ģ���ļ�����
        }

    }


}

//����html����վ
function copyHtmlToWeb(){
    $webDir='';$toWebDir=''; $toFilePath=''; $filePath=''; $fileName=''; $fileList=''; $splStr=''; $content=''; $s=''; $s1=''; $c=''; $webImages=''; $webCss=''; $webJs=''; $splJs ='';
    $webFolderName=''; $jsFileList=''; $setFileCode=''; $nErrLevel=''; $jsFilePath='';

    $setFileCode= @$_REQUEST['setcode']; //�����ļ��������

    handlePower('��������HTMLҳ��'); //����Ȩ�޴���
    writeSystemLog('', '��������HTMLҳ��'); //ϵͳ��־

    $webFolderName= $GLOBALS['cfg_webTemplate'];
    if( substr($webFolderName, 0 , 1)== '/' ){
        $webFolderName= mid($webFolderName, 2,-1);
    }
    if( Right($webFolderName, 1)== '/' ){
        $webFolderName= mid($webFolderName, 1, Len($webFolderName) - 1);
    }
    if( instr($webFolderName, '/') > 0 ){
        $webFolderName= mid($webFolderName, instr($webFolderName, '/') + 1,-1);
    }
    $webDir= '/htmladmin/' . $webFolderName . '/';
    $toWebDir='/htmlw' . 'eb/viewweb/';
    createDirFolder($toWebDir);
    $toWebDir= $toWebDir . pinYin2($webFolderName) . '/';

    deleteFolder($toWebDir);				//ɾ��
    createFolder('/htmlweb/web');		//�����ļ��� ��ֹweb�ļ��в�����20160504
    deleteFolder($webDir);
    createDirFolder($webDir);
    $webImages= $webDir . 'Images/';
    $webCss= $webDir . 'Css/';
    $webJs= $webDir . 'Js/';
    copyFolder($GLOBALS['cfg_webImages'], $webImages);
    copyFolder($GLOBALS['cfg_webCss'], $webCss);
    createFolder($webJs); //����Js�ļ���


    //����Js�ļ���
    $splJs= aspSplit(getDirJsList($webJs), vbCrlf());
    foreach( $splJs as $key=>$filePath){
        if( $filePath <> '' ){
            $toFilePath= $webJs . getFileName($filePath);
            ASPEcho('js', $filePath);
            moveFile($filePath, $toFilePath);
        }
    }
    //����Css�ļ���
    $splStr= aspSplit(getDirCssList($webCss), vbCrlf());
    foreach( $splStr as $key=>$filePath){
        if( $filePath <> '' ){
            $content= getftext($filePath);
            $content= Replace($content, $GLOBALS['cfg_webImages'], '../images/');

            $content= deleteCssNote($content);
            $content=phptrim($content);
            //����Ϊutf-8���� 20160527
            if( strtolower($setFileCode)=='utf-8' ){
                $content=replace($content,'gb2312','utf-8');
            }
            writeToFile($filePath, $content, $setFileCode);
            ASPEcho('css', $GLOBALS['cfg_webImages']);
        }
    }
    //������ĿHTML
    $GLOBALS['isMakeHtml']= true;
    $rssObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn where isonhtml=true');
    while( $rss= $GLOBALS['conn']->fetch_array($rssObj)){
        $GLOBALS['glb_filePath']= Replace(getColumnUrl($rss['columnname'], 'name'), $GLOBALS['cfg_webSiteUrl'], '');
        if( Right($GLOBALS['glb_filePath'], 1)== '/' || Right($GLOBALS['glb_filePath'], 1)== '' ){
            $GLOBALS['glb_filePath']= $GLOBALS['glb_filePath'] . 'index.html';
        }
        if( Right($GLOBALS['glb_filePath'], 5)== '.html' ){
            if( Right($GLOBALS['glb_filePath'], 11)== '/index.html' ){
                $fileList= $fileList . $GLOBALS['glb_filePath'] . vbCrlf();
            }else{
                $fileList= $GLOBALS['glb_filePath'] . vbCrlf() . $fileList;
            }
            $fileName= Replace($GLOBALS['glb_filePath'], '/', '_');
            $toFilePath= $webDir . $fileName;
            copyfile($GLOBALS['glb_filePath'], $toFilePath);
            ASPEcho('����', $GLOBALS['glb_filePath']);
        }
    }
    //��������HTML
    $rssObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'articledetail where isonhtml=true');
    while( $rss= $GLOBALS['conn']->fetch_array($rssObj)){
        $GLOBALS['glb_url']= getHandleRsUrl($rss['filename'], $rss['customaurl'], '/detail/detail' . $rss['id']);
        $GLOBALS['glb_filePath']= Replace($GLOBALS['glb_url'], $GLOBALS['cfg_webSiteUrl'], '');
        if( Right($GLOBALS['glb_filePath'], 1)== '/' || Right($GLOBALS['glb_filePath'], 1)== '' ){
            $GLOBALS['glb_filePath']= $GLOBALS['glb_filePath'] . 'index.html';
        }
        if( Right($GLOBALS['glb_filePath'], 5)== '.html' ){
            if( Right($GLOBALS['glb_filePath'], 11)== '/index.html' ){
                $fileList= $fileList . $GLOBALS['glb_filePath'] . vbCrlf();
            }else{
                $fileList= $GLOBALS['glb_filePath'] . vbCrlf() . $fileList;
            }
            $fileName= Replace($GLOBALS['glb_filePath'], '/', '_');
            $toFilePath= $webDir . $fileName;
            copyfile($GLOBALS['glb_filePath'], $toFilePath);
            ASPEcho('����' . $rss['title'], $GLOBALS['glb_filePath']);
        }
    }
    //���Ƶ���HTML
    $rssObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'onepage where isonhtml=true');
    while( $rss= $GLOBALS['conn']->fetch_array($rssObj)){
        $GLOBALS['glb_url']= getHandleRsUrl($rss['filename'], $rss['customaurl'], '/page/page' . $rss['id']);
        $GLOBALS['glb_filePath']= Replace($GLOBALS['glb_url'], $GLOBALS['cfg_webSiteUrl'], '');
        if( Right($GLOBALS['glb_filePath'], 1)== '/' || Right($GLOBALS['glb_filePath'], 1)== '' ){
            $GLOBALS['glb_filePath']= $GLOBALS['glb_filePath'] . 'index.html';
        }
        if( Right($GLOBALS['glb_filePath'], 5)== '.html' ){
            if( Right($GLOBALS['glb_filePath'], 11)== '/index.html' ){
                $fileList= $fileList . $GLOBALS['glb_filePath'] . vbCrlf();
            }else{
                $fileList= $GLOBALS['glb_filePath'] . vbCrlf() . $fileList;
            }
            $fileName= Replace($GLOBALS['glb_filePath'], '/', '_');
            $toFilePath= $webDir . $fileName;
            copyfile($GLOBALS['glb_filePath'], $toFilePath);
            ASPEcho('��ҳ' . $rss['title'], $GLOBALS['glb_filePath']);
        }
    }
    //��������html�ļ��б�
    //call echo(cfg_webSiteUrl,cfg_webTemplate)
    //call rwend(fileList)
    $sourceUrl=''; $replaceUrl ='';
    $splStr= aspSplit($fileList, vbCrlf());
    foreach( $splStr as $key=>$filePath){
        if( $filePath <> '' ){
            $filePath= $webDir . Replace($filePath, '/', '_');
            ASPEcho('filePath', $filePath);
            $content= getftext($filePath);

            foreach( $splStr as $key=>$s){
                $s1= $s;
                if( Right($s1, 11)== '/index.html' ){
                    $s1= substr($s1, 0 , Len($s1) - 11) . '/';
                }
                $sourceUrl= $GLOBALS['cfg_webSiteUrl'] . $s1;
                $replaceUrl= $GLOBALS['cfg_webSiteUrl'] . Replace($s, '/', '_');
                //Call echo(sourceUrl, replaceUrl) 							'����  ���������ʾ20160613
                $content= Replace($content, $sourceUrl, $replaceUrl);
            }
            $content= Replace($content, $GLOBALS['cfg_webSiteUrl'], ''); //ɾ����ַ
            $content= Replace($content, $GLOBALS['cfg_webTemplate'], ''); //ɾ��ģ��·��
            //content=nullLinkAddDefaultName(content)
            foreach( $splJs as $key=>$s){
                if( $s <> '' ){
                    $fileName= getFileName($s);
                    $content= Replace($content, 'Images/' . $fileName, 'js/' . $fileName);
                }
            }
            if( instr($content, '/Jquery/Jquery.Min.js') > 0 ){
                $content= Replace($content, '/Jquery/Jquery.Min.js', 'js/Jquery.Min.js');
                copyfile('/Jquery/Jquery.Min.js', $webJs . '/Jquery.Min.js');
            }
            $content= Replace($content, '<a href="" ', '<a href="index.html" '); //����ҳ��index.html
            createFileGBK($filePath, $content);
        }
    }

    //�Ѹ�����վ���µ�images/�ļ����µ�js�Ƶ�js/�ļ�����  20160315
    $htmlFileList=''; $splHtmlFile=''; $splJsFile=''; $htmlFilePath=''; $jsFileName ='';
    $jsFileList= getDirJsNameList($webImages);
    $htmlFileList= getDirHtmlList($webDir);
    $splHtmlFile= aspSplit($htmlFileList, vbCrlf());
    $splJsFile= aspSplit($jsFileList, vbCrlf());
    foreach( $splHtmlFile as $key=>$htmlFilePath){
        $content= getftext($htmlFilePath);
        foreach( $splJsFile as $key=>$jsFileName){
            $content= regExp_Replace($content, 'Images/' . $jsFileName, 'js/' . $jsFileName);
        }

        $nErrLevel= 0;
        $content= handleHtmlFormatting($content, false, $nErrLevel, '|ɾ������|'); //|ɾ������|
        $content= handleCloseHtml($content, true, ''); //�պϱ�ǩ
        $nErrLevel= checkHtmlFormatting($content);
        if( checkHtmlFormatting($content)== false ){
            eerr($htmlFilePath . '(��ʽ������)', $nErrLevel);		 //ע��
        }
        //����Ϊutf-8����
        if( strtolower($setFileCode)=='utf-8' ){
            $content=replace($content,'<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />','<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />');
        }
        $content=phptrim($content);
        writeToFile($htmlFilePath, $content, $setFileCode);
    }
    //images��js�ƶ���js��
    foreach( $splJsFile as $key=>$jsFileName){
        $jsFilePath=$webImages . $jsFileName;
        $content= getftext($jsFilePath);
        $content=phptrim($content);
        writeToFile($webJs . $jsFileName, $content, $setFileCode);
        deleteFile($jsFilePath);
    }

    copyFolder($webDir,$toWebDir);
    //ʹhtmlWeb�ļ�����phpѹ��
    if( @$_REQUEST['isMakeZip']=='1' ){
        makeHtmlWebToZip($webDir);
    }
    //ʹ��վ��xml���20160612
    if( @$_REQUEST['isMakeXml']=='1' ){
        makeHtmlWebToXmlZip('/htmladmin/', $webFolderName);
    }
}

//ʹhtmlWeb�ļ�����phpѹ��
function makeHtmlWebToZip($webDir){
    $content=''; $splStr=''; $filePath=''; $c=''; $fileArray=''; $fileName=''; $fileType=''; $isTrue ='';
    $webFolderName ='';
    $cleanFileList ='';
    $splStr= aspSplit($webDir, '/');
    $webFolderName= $splStr[2];
    //call eerr(webFolderName,webDir)
    $content= getFileFolderList($webDir, true, 'ȫ��', '', 'ȫ���ļ���', '', '');
    $splStr= aspSplit($content, vbCrlf());
    foreach( $splStr as $key=>$filePath){
        if( checkfolder($filePath)== false ){
            $fileArray= handleFilePathArray($filePath);
            $fileName= strtolower($fileArray[2]);
            $fileType= strtolower($fileArray[4]);
            $fileName= remoteNumber($fileName);
            $isTrue= true;

            if( instr('|' . $cleanFileList . '|', '|' . $fileName . '|') > 0 && $fileType== 'html' ){
                $isTrue= false;
            }
            if( $isTrue== true ){
                //call echo(fileType,fileName)
                if( $c <> '' ){ $c= $c . '|' ;}
                $c= $c . Replace($filePath, handlePath('/'), '');
                $cleanFileList= $cleanFileList . $fileName . '|';
            }
        }
    }
    rw($c);
    $c= $c . '|||||';
    createFileGBK('htmlweb/1.txt', $c);
    ASPEcho('<hr>cccccccccccc', $c);
    //���ж�����ļ�����20160309
    if( checkFile('/myZIP.php')== true ){
        ASPEcho('', XMLPost(getHost() . '/myZIP.php?webFolderName=' . $webFolderName , 'content=' . escape($c)));
    }

}
//ʹ��վ��xml���20160612
function makeHtmlWebToXmlZip($newWebDir,$rootDir){
    $xmlFileName='';$xmlSize='';
    $xmlFileName= getIP() . '_update.xml';

    //newWebDir="\Templates2015\"
    //rootDir="\sharembweb\"

    $objXmlZIP =''; $objXmlZIP= new xmlZIP();
    $objXmlZIPrun(handlePath($newWebDir), handlePath($newWebDir . $rootDir), false, $xmlFileName);
    ASPEcho(handlePath($newWebDir), handlePath($newWebDir . $rootDir));
    $objXmlZIP= $GLOBALS['Nothing'];
    doevents();
    $xmlSize=getFSize($xmlFileName);
    $xmlSize=printSpaceValue($xmlSize);
    ASPEcho('����xml����ļ�', '<a href=/tools/downfile.asp?act=download&downfile=' . xorEnc('/' . $xmlFileName, 31380) . ' title=\'�������\'>�������' . $xmlFileName . '('. $xmlSize .')</a>');
}


//���ɸ���sitemap.xml 20160118
function saveSiteMap(){
    $isWebRunHtml ='';//�Ƿ�Ϊhtml��ʽ��ʾ��վ
    $changefreg ='';//����Ƶ��
    $priority ='';//���ȼ�
    $c=''; $url ='';
    handlePower('�޸�����SiteMap'); //����Ȩ�޴���

    $changefreg= @$_REQUEST['changefreg'];
    $priority= @$_REQUEST['priority'];
    loadWebConfig(); //��������
    //call eerr("cfg_flags",cfg_flags)
    if( instr($GLOBALS['cfg_flags'], '|htmlrun|') > 0 ){
        $isWebRunHtml= true;
    }else{
        $isWebRunHtml= false;
    }

    $c= $c . '<?xml version="1.0" encoding="UTF-8"?>' . vbCrlf();
    $c= $c . "\t" . '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">' . vbCrlf();

    //��Ŀ
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn order by sortrank asc');
    while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
        if( $rsx['nofollow']== false ){
            $c= $c . copystr("\t", 2) . '<url>' . vbCrlf();
            if( $isWebRunHtml== true ){
                $url= getRsUrl($rsx['filename'], $rsx['customaurl'], '/nav' . $rsx['id']);
                $url=handleAction($url);
            }else{
                $url= escape('?act=nav&columnName=' . $rsx['columnname']);
            }
            $url= urlAddHttpUrl($GLOBALS['cfg_webSiteUrl'], $url);
            //call echo(cfg_webSiteUrl,url)

            $c= $c . copystr("\t", 3) . '<loc>' . $url . '</loc>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<lastmod>' . format_Time($rsx['updatetime'], 2) . '</lastmod>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<changefreq>' . $changefreg . '</changefreq>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<priority>' . $priority . '</priority>' . vbCrlf();
            $c= $c . copystr("\t", 2) . '</url>' . vbCrlf();
            ASPEcho('��Ŀ', '<a href="' . $url . '" target=\'_blank\'>' . $url . '</a>');
        }
    }

    //����
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'articledetail order by sortrank asc');
    while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
        if( $rsx['nofollow']== false ){
            $c= $c . copystr("\t", 2) . '<url>' . vbCrlf();
            if( $isWebRunHtml== true ){
                $url= getRsUrl($rsx['filename'], $rsx['customaurl'], '/detail/detail' . $rsx['id']);
                $url=handleAction($url);
            }else{
                $url= '?act=detail&id=' . $rsx['id'];
            }
            $url= urlAddHttpUrl($GLOBALS['cfg_webSiteUrl'], $url);
            //call echo(cfg_webSiteUrl,url)

            $c= $c . copystr("\t", 3) . '<loc>' . $url . '</loc>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<lastmod>' . format_Time($rsx['updatetime'], 2) . '</lastmod>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<changefreq>' . $changefreg . '</changefreq>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<priority>' . $priority . '</priority>' . vbCrlf();
            $c= $c . copystr("\t", 2) . '</url>' . vbCrlf();
            ASPEcho('����', '<a href="' . $url . '" target=\'_blank\'>' . $url . '</a>');
        }
    }

    //��ҳ
    $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'onepage order by sortrank asc');
    while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
        if( $rsx['nofollow']== false ){
            $c= $c . copystr("\t", 2) . '<url>' . vbCrlf();
            if( $isWebRunHtml== true ){
                $url= getRsUrl($rsx['filename'], $rsx['customaurl'], '/page/detail' . $rsx['id']);
                $url=handleAction($url);
            }else{
                $url= '?act=onepage&id=' . $rsx['id'];
            }
            $url= urlAddHttpUrl($GLOBALS['cfg_webSiteUrl'], $url);
            //call echo(cfg_webSiteUrl,url)

            $c= $c . copystr("\t", 3) . '<loc>' . $url . '</loc>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<lastmod>' . format_Time($rsx['updatetime'], 2) . '</lastmod>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<changefreq>' . $changefreg . '</changefreq>' . vbCrlf();
            $c= $c . copystr("\t", 3) . '<priority>' . $priority . '</priority>' . vbCrlf();
            $c= $c . copystr("\t", 2) . '</url>' . vbCrlf();
            ASPEcho('��ҳ', '<a href="' . $url . '" target=\'_blank\'>' . $url . '</a>');
        }
    }


    $c= $c . "\t" . '</urlset>' . vbCrlf();

    loadWebConfig();

    createfile('sitemap.xml', $c);
    ASPEcho('����sitemap.xml�ļ��ɹ�', '<a href=\'/sitemap.xml\' target=\'_blank\'>���Ԥ��sitemap.xml</a>');

    //�ж��Ƿ�����sitemap.html
    if( @$_REQUEST['issitemaphtml']== '1' ){
        $c= '';
        //�ڶ���
        //��Ŀ
        $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'webcolumn order by sortrank asc');
        while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
            if( $rsx['nofollow']== false ){
                if( $isWebRunHtml== true ){
                    $url= getRsUrl($rsx['filename'], $rsx['customaurl'], '/nav' . $rsx['id']);
                    $url=handleAction($url);
                }else{
                    $url= escape('?act=nav&columnName=' . $rsx['columnname']);
                }
                $url= urlAddHttpUrl($GLOBALS['cfg_webSiteUrl'], $url);

                $c= $c . '<li style="width:20%;"><a href="' . $url . '">' . $rsx['columnname'] . '</a><ul>' . vbCrlf();

                //����
                $rssObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'articledetail where parentId=' . $rsx['id'] . ' order by sortrank asc');
                while( $rss= $GLOBALS['conn']->fetch_array($rssObj)){
                    if( $rss['nofollow']== false ){
                        if( $isWebRunHtml== true ){
                            $url= getRsUrl($rss['filename'], $rss['customaurl'], '/detail/detail' . $rss['id']);
                            $url=handleAction($url);
                        }else{
                            $url= '?act=detail&id=' . $rss['id'];
                        }
                        $url= urlAddHttpUrl($GLOBALS['cfg_webSiteUrl'], $url);
                        $c= $c . '<li style="width:20%;"><a href="' . $url . '" target="_blank">' . $rss['title'] . '</a>' . vbCrlf();
                    }
                }
                $c= $c . '</ul></li>' . vbCrlf();


            }
        }

        //����
        $c= $c . '<li style="width:20%;"><a href="javascript:;">�����б�</a><ul>' . vbCrlf();
        $rsxObj=$GLOBALS['conn']->query( 'select * from ' . $GLOBALS['db_PREFIX'] . 'onepage order by sortrank asc');
        while( $rsx= $GLOBALS['conn']->fetch_array($rsxObj)){
            if( $rsx['nofollow']== false ){
                $c= $c . copystr("\t", 2) . '<url>' . vbCrlf();
                if( $isWebRunHtml== true ){
                    $url= getRsUrl($rsx['filename'], $rsx['customaurl'], '/page/detail' . $rsx['id']);
                    $url=handleAction($url);
                }else{
                    $url= '?act=onepage&id=' . $rsx['id'];
                }
                $c= $c . '<li style="width:20%;"><a href="' . $url . '" target="_blank">' . $rsx['title'] . '</a>' . vbCrlf();
            }
        }
        $c= $c . '</ul></li>' . vbCrlf();

        $templateContent ='';
        $templateContent= getftext($GLOBALS['adminDir'] . '/template_SiteMap.html');


        $templateContent= Replace($templateContent, '{$content$}', $c);
        $templateContent= Replace($templateContent, '{$Web_Title$}', $GLOBALS['cfg_webTitle']);


        createfile('sitemap.html', $templateContent);
        ASPEcho('����sitemap.html�ļ��ɹ�', '<a href=\'/sitemap.html\' target=\'_blank\'>���Ԥ��sitemap.html</a>');
    }
    writeSystemLog('', '����sitemap.xml'); //ϵͳ��־
}
?>


