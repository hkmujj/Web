<?php 


define('WEBCOLUMNTYPE','首页|文本|产品|新闻|视频|下载|案例|留言|反馈|招聘|订单'); 		//网站栏目类型列表
define('EDITORTYPE','php'); 		//编辑器类型，是ASP,或PHP,或jSP,或.NET
define('WEB_VIEWURL','../phpweb.php'); 		//网站显示URL
define('WEB_ADMINURL','/web/1.php'); 		//后端网站，在线编辑用到20160216
$webDir='';				//网站目录
//=========


$ReadBlockList='';

$SysStyle=array(9);
$SysStyle[0]= '#999999';
$makeHtmlFileToLCase	 =''; $makeHtmlFileToLCase=true;		//生成HTML文件转小写
$isWebLabelClose =''; $isWebLabelClose=true;					//闭合标签(20150831)

$HandleisCache =''; $HandleisCache=false;						//缓冲是否处理了
$db_PREFIX =''; $db_PREFIX= 'xy_'; 		 //表前缀
$adminDir ='';$adminDir='/web/';							//后台目录


$openErrorLog =''; $openErrorLog= true; //开启错误日志
$openWriteSystemLog =''; $openWriteSystemLog= '|txt|database|'; //开启写系统日志 txt写入文本 database写入数据库
$isTestEcho=''; $isTestEcho=true;											//开关测试回显
$webVersion =''; $webVersion='ASPPHPCMS v1.5';												//网站版本


$WEB_CACHEFile =''; $WEB_CACHEFile=$webDir . '/web/'. EDITORTYPE .'cachedata.txt';								//缓冲文件
$WEB_CACHEContent =''; $WEB_CACHEContent='';								//缓冲文件内容
$isCacheTip =''; $isCacheTip=false;			//是否开启缓冲提示

$language =''; $language='#en-us';			//en-us  | zh-cn | zh-tw

//批量替换语言
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
//替换语言
function replaceL($content,$str){
    $replaceL=replace($content,$str, setL($str));
    return @$replaceL;
}
//语言
function setL($str){
    $c='';
    $c=$str;
    if( $GLOBALS['language']=='en-us' ){
        $c=languageEN($c);
    }
    $setL=$c;
    return @$setL;
}
//处理显语言  c=handleDisplayLanguage(c,"loginok")
function handleDisplayLanguage($c,$sType){
    //繁体就直接转换了，不要再一个一个转了，
    if( $GLOBALS['language']=='zh-tw' ){
        $handleDisplayLanguage=simplifiedTransfer($c);
        return @$handleDisplayLanguage;
    }
    if( $sType=='login' ){

        $c= batchReplaceL($c, '请不要输入特殊字符|*|输入正确|*|用户名可以用字母|*|用户名可以用字母|*|您的用户名为空|*|密码可以用字母|*|您的密码为空');
        $c= batchReplaceL($c, '登录后台|*|管理员登录|*|如果您不是管理员|*|请立即停止您的登陆行为|*|用户名|*|版');
        $c= batchReplaceL($c, '密&nbsp;码|*|密码|*|请输入|*|登 录|*|登录|*|重 置|*|重置');


    }else if( $sType=='loginok' ){
        $c= batchReplaceL($c, '后台地图|*|清除缓冲|*|超级管理员|*|当前位置|*|管理员信息|*|修改密码|*|最新访客信息|*|开发团队|*|版权所有|*|开发与支持团队');
        $c= batchReplaceL($c, '进入在线修改模式|*|系统信息|*|免费开源版|*|授权信息|*|服务器名称|*|服务器版本|*|交流群|*|相关链接|*|登录后台');
        $c= batchReplaceL($c, '用户名|*|表前缀|*|帮助|*|退出|*|您好|*|首页|*|权限|*|端口|*|邮箱|*|官网|*|版|*|云端');
        $c= batchReplaceL($c, '系统管理|*|我的面板|*|栏目管理|*|模板管理|*|会员管理|*|生成管理|*|更多设置');

        $c= batchReplaceL($c, '站点配置|*|网站统计|*|生成|*|后台操作日志|*|后台管理员|*|网站栏目|*|分类信息|*|评论|*|搜索统计|*|单页管理|*|友情链接|*|招聘管理');
        $c= batchReplaceL($c, '反馈管理|*|留言管理|*|会员配置|*|竞价词|*|网站布局|*|网站模块|*|后台菜单|*|执行|*|仿站');


    }
    $handleDisplayLanguage=$c;
    return @$handleDisplayLanguage;
}

//为英文
function languageEN($str){
    $c='';
    if( $str=='登录成功，正在进入后台...' ){
        $c='Login successfully, is entering the background...';
    }else if( $str=='账号密码错误<br>登录次数为 ' ){
        $c='Account password error <br> login ';
    }else if( $str=='登录成功，正在进入后台...' ){
        $c='Login successfully, is entering the background...';
    }else if( $str=='退出成功' ){
        $c='Exit success';
    }else if( $str=='退出成功，正在进入登录界面...' ){
        $c='Quit successfully, is entering the login screen...';
    }else if( $str=='清除缓冲完成，正在进入后台界面...' ){
        $c='Clear buffer finish, is entering the background interface...';
    }else if( $str=='提示信息' ){
        $c='Prompt info';
    }else if( $str=='如果您的浏览器没有自动跳转，请点击这里' ){
        $c='If your browser does not automatically jump, please click here';
    }else if( $str=='倒计时'	 ){
        $c='Countdown ';
    }else if( $str=='后台地图'	 ){
        $c='Admin map';
    }else if( $str=='清除缓冲'	 ){
        $c='Clear buffer';
    }else if( $str=='超级管理员'	 ){
        $c='Super administrator';
    }else if( $str=='当前位置'	 ){
        $c='current location';
    }else if( $str=='管理员信息'	 ){
        $c='Admin info';
    }else if( $str=='修改密码'	 ){
        $c='Modify password';
    }else if( $str=='用户名'	 ){
        $c='username';
    }else if( $str=='表前缀'	 ){
        $c='Table Prefix';
    }else if( $str=='进入在线修改模式'	 ){
        $c='online modification';
    }else if( $str=='系统信息'	 ){
        $c='system info';
    }else if( $str=='授权信息'	 ){
        $c='Authorization information';
    }else if( $str=='免费开源版'	 ){
        $c='Free open source';
    }else if( $str=='服务器名称'	 ){
        $c='Server name';
    }else if( $str=='服务器版本'	 ){
        $c='Server version';
    }else if( $str=='最新访客信息'	 ){
        $c='visitor info';
    }else if( $str=='开发团队'	 ){
        $c='team info';
    }else if( $str=='版权所有'	 ){
        $c='copyright';
    }else if( $str=='开发与支持团队'	 ){
        $c='Develop and support team';
    }else if( $str=='交流群'	 ){
        $c='Exchange group';
    }else if( $str=='相关链接'	 ){
        $c='Related links';
    }else if( $str=='系统管理'	 ){
        $c='System';
    }else if( $str=='我的面板'	 ){
        $c='My panel';
    }else if( $str=='栏目管理'	 ){
        $c='Column';
    }else if( $str=='模板管理'	 ){
        $c='Template';
    }else if( $str=='会员管理'	 ){
        $c='Member';
    }else if( $str=='生成管理'	 ){
        $c='Generation';
    }else if( $str=='更多设置'	 ){
        $c='More settings';


    }else if( $str=='登录后台'	 ){
        $c='Login management background';
    }else if( $str=='管理员登录'	 ){
        $c='Administrator login ';
    }else if( $str=='如果您不是管理员'	 ){
        $c='If you are not an administrator';
    }else if( $str=='请立即停止您的登陆行为'	 ){
        $c='Please stop your login immediately';
    }else if( $str=='密&nbsp;码'	|| $str=='密码' ){
        $c='password';
    }else if( $str=='请输入'	 ){
        $c='Please input';
    }else if( $str=='登 录'	|| $str=='登录' ){
        $c='login';
    }else if( $str=='重 置' || $str=='重置' ){
        $c='reset';
    }else if( $str=='请不要输入特殊字符'	 ){
        $c='Please do not enter special characters';
    }else if( $str=='输入正确'	 ){
        $c='Input correct';
    }else if( $str=='用户名可以用字母'	 ){
        $c='Username use ';
    }else if( $str=='您的用户名为空'	 ){
        $c='Your username is empty';
    }else if( $str=='密码可以用字母'	 ){
        $c='Passwords use ';
    }else if( $str=='您的密码为空'	 ){
        $c='Your password is empty';
    }else if( $str=='站点配置'	 ){
        $c='Site configuration';
    }else if( $str=='网站统计'	 ){
        $c='Website statistics';
    }else if( $str=='后台操作日志'	 ){
        $c='Admin log ';
    }else if( $str=='后台管理员'	 ){
        $c='Background manager';
    }else if( $str=='网站栏目'	 ){
        $c='Website column';
    }else if( $str=='分类信息'	 ){
        $c='Classified information';
    }else if( $str=='搜索统计'	 ){
        $c='Search statistics';
    }else if( $str=='单页管理'	 ){
        $c='Single page';
    }else if( $str=='友情链接'	 ){
        $c='Friendship link';
    }else if( $str=='招聘管理'	 ){
        $c='Recruitment management';
    }else if( $str=='反馈管理'	 ){
        $c='Feedback management';
    }else if( $str=='留言管理'	 ){
        $c='message management';
    }else if( $str=='会员配置'	 ){
        $c='Member allocation';
    }else if( $str=='竞价词'	 ){
        $c='Bidding words';
    }else if( $str=='网站布局'	 ){
        $c='Website layout';
    }else if( $str=='网站模块'	 ){
        $c='Website module';
    }else if( $str=='后台菜单'	 ){
        $c='Background menu';
    }else if( $str=='仿站'	 ){
        $c='Template website ';

    }else if( $str=='11111'	 ){
        $c='1111111';



    }else if( $str=='执行'	 ){
        $c='implement ';
    }else if( $str=='评论'	 ){
        $c='comment ';
    }else if( $str=='生成'	 ){
        $c='generate ';
    }else if( $str=='权限'	 ){
        $c='jurisdiction';
    }else if( $str=='帮助'	 ){
        $c='Help';
    }else if( $str=='退出'	 ){
        $c='sign out';
    }else if( $str=='您好'	 ){
        $c='hello';
    }else if( $str=='首页'	 ){
        $c='home';
    }else if( $str=='端口'	 ){
        $c='port';
    }else if( $str=='官网'	 ){
        $c='official website';
    }else if( $str=='邮箱'	 ){
        $c='Emai';
    }else if( $str=='云端'	 ){
        $c='Cloud';

    }else if( $str=='版'	 ){
        $c=' edition';




    }else{
        $c=$str;
    }
    $languageEN=$c;
    return @$languageEN;
}
?>