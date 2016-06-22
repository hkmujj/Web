<?PHP
header("Content-Type: text/html; charset=gb2312");
?><!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>PHP上传文件</title>
<style type="text/css">
body{padding:0;margin:0;}
</style>
</head>
<body> 

<form action="upload_file.php?returnInputName=<?=@$_GET['returnInputName']?>" method="post"
enctype="multipart/form-data">
<label for="file"></label>
<input type="file" name="file" id="file" />
<input type="submit" name="submit" value="上传" />
</form>

</body>
</html>