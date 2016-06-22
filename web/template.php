<?php 
//系统
require_once './../PHP2/ImageWaterMark/Include/ASP.php';
require_once './../PHP2/ImageWaterMark/Include/sys_FSO.php';
require_once './../PHP2/ImageWaterMark/Include/Conn.php';
require_once './../PHP2/ImageWaterMark/Include/MySqlClass.php'; 
require_once './../PHP2/ImageWaterMark/Include/sys_Url.php'; 

require_once './../PHP2/Web/Inc/Common.php';  

require_once './config.php'; 
require_once './setAccess.php';
require_once './function.php';

?><!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>模板文件管理</title>
</head>
<body>
<style type="text/css">
<!--
body {
    margin-left: 0px;
    margin-top: 0px;
    margin-right: 0px;
    margin-bottom: 0px;
}
a:link,a:visited,a:active {
    color: #000000;
    text-decoration: none;
}
a:hover {
    color: #666666;
    text-decoration: none;
}
.tableline{
    border: 1px solid #999999;
}
body,td,th {
    font-size: 12px;
}
a {
    font-size: 12px;
}
-->
</style>
<script language="javascript">
function checkDel()
    {
    if(confirm("确认要删除吗？删除后将不可恢复！"))
    return true;
    else
    return false;
}
</script>
<?PHP

if( @$_SESSION['adminusername']== '' ){
    eerr('提示', '未登录，请先登录');
}


switch ( @$_REQUEST['act'] ){
    case 'templateFileList' ; displayTemplateDirDialog(@$_REQUEST['dir']) ; templateFileList(@$_REQUEST['dir']);break;//模板列表
    case 'delTemplateFile' ; delTemplateFile(@$_REQUEST['dir'], @$_REQUEST['fileName']) ; displayTemplateDirDialog(@$_REQUEST['dir']) ; templateFileList(@$_REQUEST['dir'])		;break;//删除模板文件
    case 'addEditFile' ; displayTemplateDirDialog(@$_REQUEST['dir']) ; addEditFile(@$_REQUEST['dir'], @$_REQUEST['fileName']);break;//显示添加修改文件
    default ; displayTemplateDirDialog(@$_REQUEST['dir']); //显示模板目录面板
}

//模板文件列表
function templateFileList($dir){
    $content=''; $splStr=''; $fileName=''; $s ='';
    if( @$_SESSION['adminusername']== 'ASPPHPCMS' ){
        $content= getDirFileNameList($dir,'');
    }else{
        $content= getDirHtmlNameList($dir);
    }
    $splStr= aspSplit($content, vbCrlf());
    foreach( $splStr as $key=>$fileName){
        if( $fileName<>'' ){
            $s= '<a href="../phpweb.php?templatedir=' . escape($dir) . '&templateName=' . $fileName . '" target=\'_blank\'>预览</a> ';
            ASPEcho('<img src=\'Images/Icon/2/htm.gif\'>' . $fileName, $s . '| <a href=\'?act=addEditFile&dir=' . $dir . '&fileName=' . $fileName . '\'>修改</a> | <a href=\'?act=delTemplateFile&dir=' . @$_REQUEST['dir'] . '&fileName=' . $fileName . '\' onclick=\'return checkDel()\'>删除</a>');
        }
    }
}

//删除模板文件
function delTemplateFile($dir, $fileName){
    $filePath ='';

    handlePower('删除模板文件');						//管理权限处理

    $filePath= $dir . '/' . $fileName;
    deleteFile($filePath);
    ASPEcho('删除文件', $filePath);
}

//显示面板样式列表
function displayPanelList($dir){
    $content='';$splstr='';$s='';$c='';
    $content=getDirFolderNameList($dir);
    $splstr=aspSplit($content,vbCrlf());
    $c='<select name=\'selectLeftStyle\'>';
    foreach( $splstr as $key=>$s){
        $s='<option value=\'\'>'. $s .'</option>';
        $c=$c . $s . vbCrlf();
    }
    $displayPanelList= $c . '</select>';
    return @$displayPanelList;
}


//添加修改文件
function addEditFile($dir, $fileName){
    $filePath='';$promptMsg='';

    if( Right(strtolower($fileName), 5) <> '.html' && @$_SESSION['adminusername'] <> 'ASPPHPCMS' ){
        $fileName= $fileName . '.html';
    }
    $filePath= $dir . '/' . $fileName;

    if( checkFile($filePath)==false ){
        handlePower('添加模板文件');						//管理权限处理
    }else{
        handlePower('修改模板文件');						//管理权限处理
    }

    //保存内容
    if( @$_REQUEST['issave']== 'true' ){
        createfile($filePath, @$_REQUEST['content']);
        $promptMsg='保存成功';
    }
    ?>
    <form name="form1" method="post" action="?act=addEditFile&issave=true">
    <table width="99%" border="0" cellspacing="0" cellpadding="0" class="tableline">
    <tr>
    <td height="30">目录<?=$dir?><br>
    <input name="dir" type="hidden" id="dir" value="<?=$dir?>" /></td>
    </tr>
    <tr>
    <td>文件名称
    <input name="fileName" type="text" id="fileName" value="<?=$fileName?>" size="40"><?=$promptMsg?>
    <br>
    <textarea name="content" style="width:99%;height:480px;"id="content"><? rw(getFText($filePath))?></textarea></td>
    </tr>
    <tr>
    <td height="40" align="center"><input type="submit" name="button" id="button" value=" 保存 " /></td>
    </tr>
    </table>
    </form>
    <?PHP }
    //文件夹搜索
    function displayTemplateDirDialog($dir){
        $folderPath='';
        ?>
        <form name="form2" method="post" action="?act=templateFileList">
        <table width="99%" border="0" cellspacing="0" cellpadding="0" class="tableline">
        <tr>
        <td height="30"><input name="dir" type="text" id="dir" value="<?=$dir?>" size="60" />
        <input type="submit" name="button2" id="button2" value=" 进入 " /><?PHP
        $folderPath=$dir . '/images/column/';
        if( checkFolder($folderPath) ){
            rw('面板样式' . displayPanelList($folderPath));
        }
        $folderPath=$dir . '/images/nav/';
        if( checkFolder($folderPath) ){
            rw('导航样式' . displayPanelList($folderPath));
        }
        ?></td>
        </tr>
        </table>
        </form>
        <? }?>


