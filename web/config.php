<?php 


define('WEBCOLUMNTYPE','��ҳ|�ı�|��Ʒ|����|��Ƶ|����|����|����|����|��Ƹ|����'); 		//��վ��Ŀ�����б�
define('EDITORTYPE','php'); 		//�༭�����ͣ���ASP,��PHP,��jSP,��.NET
define('WEB_VIEWURL','../phpweb.php'); 		//��վ��ʾURL
define('WEB_ADMINURL','/web/1.php'); 		//�����վ�����߱༭�õ�20160216
$webDir='';				//��վĿ¼
//=========


$ReadBlockList='';

$SysStyle=array(9);
$SysStyle[0]= '#999999';
$makeHtmlFileToLCase	 =''; $makeHtmlFileToLCase=true;		//����HTML�ļ�תСд
$isWebLabelClose =''; $isWebLabelClose=true;					//�պϱ�ǩ(20150831)

$HandleisCache =''; $HandleisCache=false;						//�����Ƿ�����
$db_PREFIX =''; $db_PREFIX= 'xy_'; 		 //��ǰ׺
$adminDir ='';$adminDir='/web/';							//��̨Ŀ¼


$openErrorLog =''; $openErrorLog= true; //����������־
$openWriteSystemLog =''; $openWriteSystemLog= '|txt|database|'; //����дϵͳ��־ txtд���ı� databaseд�����ݿ�
$isTestEcho=''; $isTestEcho=true;											//���ز��Ի���
$webVersion =''; $webVersion='ASPPHPCMS v1.5';												//��վ�汾


$WEB_CACHEFile =''; $WEB_CACHEFile=$webDir . '/web/'. EDITORTYPE .'cachedata.txt';								//�����ļ�
$WEB_CACHEContent =''; $WEB_CACHEContent='';								//�����ļ�����
$isCacheTip =''; $isCacheTip=false;			//�Ƿ���������ʾ

$language =''; $language='#en-us';			//en-us  | zh-cn | zh-tw

//�����滻����
function batchReplaceL($content,$str){
    $splstr='';$s='';$i='';
    $splstr=aspSplit($str,'|*|');
    for( $i= 0 ; $i<= ubound($splstr); $i++){
        $s=$splstr[$i];
        if( $s <>'' ){
            $content=replaceL($content,$s);
        }
    }
    $batchReplaceL=$content;
    return @$batchReplaceL;
}
//�滻����
function replaceL($content,$str){
    $replaceL=replace($content,$str, setL($str));
    return @$replaceL;
}
//����
function setL($str){
    $c='';
    $c=$str;
    if( $GLOBALS['language']=='en-us' ){
        $c=languageEN($c);
    }
    $setL=$c;
    return @$setL;
}
//����������  c=handleDisplayLanguage(c,"loginok")
function handleDisplayLanguage($c,$sType){
    //�����ֱ��ת���ˣ���Ҫ��һ��һ��ת�ˣ�
    if( $GLOBALS['language']=='zh-tw' ){
        $handleDisplayLanguage=simplifiedTransfer($c);
        return @$handleDisplayLanguage;
    }
    if( $sType=='login' ){

        $c= batchReplaceL($c, '�벻Ҫ���������ַ�|*|������ȷ|*|�û�����������ĸ|*|�û�����������ĸ|*|�����û���Ϊ��|*|�����������ĸ|*|��������Ϊ��');
        $c= batchReplaceL($c, '��¼��̨|*|����Ա��¼|*|��������ǹ���Ա|*|������ֹͣ���ĵ�½��Ϊ|*|�û���|*|��');
        $c= batchReplaceL($c, '��&nbsp;��|*|����|*|������|*|�� ¼|*|��¼|*|�� ��|*|����');


    }else if( $sType=='loginok' ){
        $c= batchReplaceL($c, '��̨��ͼ|*|�������|*|��������Ա|*|��ǰλ��|*|����Ա��Ϣ|*|�޸�����|*|���·ÿ���Ϣ|*|�����Ŷ�|*|��Ȩ����|*|������֧���Ŷ�');
        $c= batchReplaceL($c, '���������޸�ģʽ|*|ϵͳ��Ϣ|*|��ѿ�Դ��|*|��Ȩ��Ϣ|*|����������|*|�������汾|*|����Ⱥ|*|�������|*|��¼��̨');
        $c= batchReplaceL($c, '�û���|*|��ǰ׺|*|����|*|�˳�|*|����|*|��ҳ|*|Ȩ��|*|�˿�|*|����|*|����|*|��|*|�ƶ�');
        $c= batchReplaceL($c, 'ϵͳ����|*|�ҵ����|*|��Ŀ����|*|ģ�����|*|��Ա����|*|���ɹ���|*|��������');

        $c= batchReplaceL($c, 'վ������|*|��վͳ��|*|����|*|��̨������־|*|��̨����Ա|*|��վ��Ŀ|*|������Ϣ|*|����|*|����ͳ��|*|��ҳ����|*|��������|*|��Ƹ����');
        $c= batchReplaceL($c, '��������|*|���Թ���|*|��Ա����|*|���۴�|*|��վ����|*|��վģ��|*|��̨�˵�|*|ִ��|*|��վ');


    }
    $handleDisplayLanguage=$c;
    return @$handleDisplayLanguage;
}

//ΪӢ��
function languageEN($str){
    $c='';
    if( $str=='��¼�ɹ������ڽ����̨...' ){
        $c='Login successfully, is entering the background...';
    }else if( $str=='�˺��������<br>��¼����Ϊ ' ){
        $c='Account password error <br> login ';
    }else if( $str=='��¼�ɹ������ڽ����̨...' ){
        $c='Login successfully, is entering the background...';
    }else if( $str=='�˳��ɹ�' ){
        $c='Exit success';
    }else if( $str=='�˳��ɹ������ڽ����¼����...' ){
        $c='Quit successfully, is entering the login screen...';
    }else if( $str=='���������ɣ����ڽ����̨����...' ){
        $c='Clear buffer finish, is entering the background interface...';
    }else if( $str=='��ʾ��Ϣ' ){
        $c='Prompt info';
    }else if( $str=='������������û���Զ���ת����������' ){
        $c='If your browser does not automatically jump, please click here';
    }else if( $str=='����ʱ'	 ){
        $c='Countdown ';
    }else if( $str=='��̨��ͼ'	 ){
        $c='Admin map';
    }else if( $str=='�������'	 ){
        $c='Clear buffer';
    }else if( $str=='��������Ա'	 ){
        $c='Super administrator';
    }else if( $str=='��ǰλ��'	 ){
        $c='current location';
    }else if( $str=='����Ա��Ϣ'	 ){
        $c='Admin info';
    }else if( $str=='�޸�����'	 ){
        $c='Modify password';
    }else if( $str=='�û���'	 ){
        $c='username';
    }else if( $str=='��ǰ׺'	 ){
        $c='Table Prefix';
    }else if( $str=='���������޸�ģʽ'	 ){
        $c='online modification';
    }else if( $str=='ϵͳ��Ϣ'	 ){
        $c='system info';
    }else if( $str=='��Ȩ��Ϣ'	 ){
        $c='Authorization information';
    }else if( $str=='��ѿ�Դ��'	 ){
        $c='Free open source';
    }else if( $str=='����������'	 ){
        $c='Server name';
    }else if( $str=='�������汾'	 ){
        $c='Server version';
    }else if( $str=='���·ÿ���Ϣ'	 ){
        $c='visitor info';
    }else if( $str=='�����Ŷ�'	 ){
        $c='team info';
    }else if( $str=='��Ȩ����'	 ){
        $c='copyright';
    }else if( $str=='������֧���Ŷ�'	 ){
        $c='Develop and support team';
    }else if( $str=='����Ⱥ'	 ){
        $c='Exchange group';
    }else if( $str=='�������'	 ){
        $c='Related links';
    }else if( $str=='ϵͳ����'	 ){
        $c='System';
    }else if( $str=='�ҵ����'	 ){
        $c='My panel';
    }else if( $str=='��Ŀ����'	 ){
        $c='Column';
    }else if( $str=='ģ�����'	 ){
        $c='Template';
    }else if( $str=='��Ա����'	 ){
        $c='Member';
    }else if( $str=='���ɹ���'	 ){
        $c='Generation';
    }else if( $str=='��������'	 ){
        $c='More settings';


    }else if( $str=='��¼��̨'	 ){
        $c='Login management background';
    }else if( $str=='����Ա��¼'	 ){
        $c='Administrator login ';
    }else if( $str=='��������ǹ���Ա'	 ){
        $c='If you are not an administrator';
    }else if( $str=='������ֹͣ���ĵ�½��Ϊ'	 ){
        $c='Please stop your login immediately';
    }else if( $str=='��&nbsp;��'	|| $str=='����' ){
        $c='password';
    }else if( $str=='������'	 ){
        $c='Please input';
    }else if( $str=='�� ¼'	|| $str=='��¼' ){
        $c='login';
    }else if( $str=='�� ��' || $str=='����' ){
        $c='reset';
    }else if( $str=='�벻Ҫ���������ַ�'	 ){
        $c='Please do not enter special characters';
    }else if( $str=='������ȷ'	 ){
        $c='Input correct';
    }else if( $str=='�û�����������ĸ'	 ){
        $c='Username use ';
    }else if( $str=='�����û���Ϊ��'	 ){
        $c='Your username is empty';
    }else if( $str=='�����������ĸ'	 ){
        $c='Passwords use ';
    }else if( $str=='��������Ϊ��'	 ){
        $c='Your password is empty';
    }else if( $str=='վ������'	 ){
        $c='Site configuration';
    }else if( $str=='��վͳ��'	 ){
        $c='Website statistics';
    }else if( $str=='��̨������־'	 ){
        $c='Admin log ';
    }else if( $str=='��̨����Ա'	 ){
        $c='Background manager';
    }else if( $str=='��վ��Ŀ'	 ){
        $c='Website column';
    }else if( $str=='������Ϣ'	 ){
        $c='Classified information';
    }else if( $str=='����ͳ��'	 ){
        $c='Search statistics';
    }else if( $str=='��ҳ����'	 ){
        $c='Single page';
    }else if( $str=='��������'	 ){
        $c='Friendship link';
    }else if( $str=='��Ƹ����'	 ){
        $c='Recruitment management';
    }else if( $str=='��������'	 ){
        $c='Feedback management';
    }else if( $str=='���Թ���'	 ){
        $c='message management';
    }else if( $str=='��Ա����'	 ){
        $c='Member allocation';
    }else if( $str=='���۴�'	 ){
        $c='Bidding words';
    }else if( $str=='��վ����'	 ){
        $c='Website layout';
    }else if( $str=='��վģ��'	 ){
        $c='Website module';
    }else if( $str=='��̨�˵�'	 ){
        $c='Background menu';
    }else if( $str=='��վ'	 ){
        $c='Template website ';

    }else if( $str=='11111'	 ){
        $c='1111111';



    }else if( $str=='ִ��'	 ){
        $c='implement ';
    }else if( $str=='����'	 ){
        $c='comment ';
    }else if( $str=='����'	 ){
        $c='generate ';
    }else if( $str=='Ȩ��'	 ){
        $c='jurisdiction';
    }else if( $str=='����'	 ){
        $c='Help';
    }else if( $str=='�˳�'	 ){
        $c='sign out';
    }else if( $str=='����'	 ){
        $c='hello';
    }else if( $str=='��ҳ'	 ){
        $c='home';
    }else if( $str=='�˿�'	 ){
        $c='port';
    }else if( $str=='����'	 ){
        $c='official website';
    }else if( $str=='����'	 ){
        $c='Emai';
    }else if( $str=='�ƶ�'	 ){
        $c='Cloud';

    }else if( $str=='��'	 ){
        $c=' edition';




    }else{
        $c=$str;
    }
    $languageEN=$c;
    return @$languageEN;
}
?>