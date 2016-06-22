<?php
//系统
require_once './../PHP2/ImageWaterMark/Include/ASP.php';
require_once './../PHP2/ImageWaterMark/Include/sys_FSO.php';
require_once './../PHP2/ImageWaterMark/Include/Conn.php';
require_once './../PHP2/ImageWaterMark/Include/MySqlClass.php';
require_once './../PHP2/ImageWaterMark/Include/sys_System.php'; 



// 更新引用部分    http://127.0.0.1/php2/web/%E8%8E%B7%E5%BE%97inc%E5%BC%95%E7%94%A8%E5%86%85%E5%AE%B9.asp

//生成
/*
*/
//引用
require_once './../PHP2/Web/Inc/2014_Array.php';
require_once './../PHP2/Web/Inc/2014_Author.php';
require_once './../PHP2/Web/Inc/2014_Css.php';
require_once './../PHP2/Web/Inc/2014_Js.php';
require_once './../PHP2/Web/Inc/2014_Nav.php';
require_once './../PHP2/Web/Inc/2015_APGeneral.php';
require_once './../PHP2/Web/Inc/2015_Color.php';
require_once './../PHP2/Web/Inc/2015_Formatting.php';
require_once './../PHP2/Web/Inc/2015_Param.php';
//require_once './PHP2/Web/Inc/2015_ToMyPHP.php';
//require_once './PHP2/Web/Inc/2015_ToPhpCms.php';
require_once './../PHP2/Web/Inc/Cai.php';
require_once './../PHP2/Web/Inc/Check.php';
require_once './../PHP2/Web/Inc/Common.php';
require_once './../PHP2/Web/Inc/Config.php';
require_once './../PHP2/Web/Inc/Incpage.php';
require_once './../PHP2/Web/Inc/Print.php';
require_once './../PHP2/Web/Inc/StringNumber.php';
require_once './../PHP2/Web/Inc/Time.php';
require_once './../PHP2/Web/Inc/URL.php';
require_once './../PHP2/Web/Inc/FunHTML.php'; 

header("Content-Type:text/html;charset=gb2312");

 
$dirName="./../UploadFiles/image";

$upFileName='file';
if(isset($_FILES[$upFileName])==false){
	$upFileName='filedata';
}

$fileName=$_FILES[$upFileName]["name"];
$fileType=handleFilePathArray($fileName)[4];
//aspecho('$fileType',$fileType);
$displayFileName=format_Time(now(),6) . '.' . $fileType;
$newFileName=$dirName . '/' . $displayFileName;
//aspecho('$newFileName',$newFileName);
//aspecho('$_FILES[$upFileName]["type"]',$_FILES[$upFileName]["type"]);
//文件大小，与类型以后判断  200000大概2M
// file代表上传文件，提交过来的上传file的name是什么名称不重要

if ((($_FILES[$upFileName]["type"] == "image/gif") || ($_FILES[$upFileName]["type"] == "image/jpeg") || ($_FILES[$upFileName]["type"] == "image/png") || ($_FILES[$upFileName]["type"] == "image/pjpeg")) && ($_FILES[$upFileName]["size"] < 200000)){
    if ($_FILES[$upFileName]["error"] > 0){
        echo "Return Code: " . $_FILES[$upFileName]["error"] . "<br />";
    }else{
        //echo "Upload: " . $_FILES[$upFileName]["name"] . "<br />";
        //echo "Type: " . $_FILES[$upFileName]["type"] . "<br />";
        //echo "Size: " . ($_FILES[$upFileName]["size"] / 1024) . " Kb<br />";
        //echo "Temp file: " . $_FILES[$upFileName]["tmp_name"] . "<br />";

        if (file_exists($newFileName))
            {
            echo$newFileName . " already exists. ";
        }else{
            move_uploaded_file($_FILES[$upFileName]["tmp_name"],$newFileName);
            //echo "Stored in: " . "upload/" . $_FILES[$upFileName]["name"];
            echo('<SCRIPT language=javascript>parent.document.all.'. @$_GET['returnInputName'] .'.value=\'/UploadFiles/image/'. $displayFileName .'\';</script>');
            echo('上传文件成功');
        }
    }
}else{
    echo "Invalid file" .$_FILES[$upFileName]["type"] . ' size' . $_FILES[$upFileName]["size"];
}

?>
