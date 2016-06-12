<!--#Include virtual = "/Inc/_Config.Asp"--> 
<%
dim httpurl,refererUrl, cookiesStr, postStr, sSetChar
httpurl="http://aa/4.asp"
refererUrl="www.baidu.com"
call rw( getWinHttp(httpUrl, refererUrl, cookiesStr, postStr, sSetChar))
%>
 
<script>

function createAjax() { 
	var _xmlhttp;
	try {
		_xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");	//IE的创建方式
	}catch (e) {
		try {
			_xmlhttp=new XMLHttpRequest();	//FF等浏览器的创建方式
		}catch (e) {
			_xmlhttp=false;		//如果创建失败，将返回false
		}
	}
	return _xmlhttp;	//返回xmlhttp对象实例
}
//Ajax("/Ajax.Asp?act=ShowArticleNumb","ArticleViewNumb")
//Ajax
function Ajax(URL,ShowID) {  
	var xmlhttp=createAjax();
	if (xmlhttp) {
		URL+= "&n="+Math.random() 
		xmlhttp.open('post', URL, true);//基本方法
		xmlhttp.setRequestHeader("cache-control","no-cache"); 
		xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded"); 		
		xmlhttp.onreadystatechange=function() {		
			if (xmlhttp.readyState==4 && xmlhttp.status==200) {	
				var str=unescape(xmlhttp.responseText)
				if(ShowID!="undefined" && ShowID!=""){
					document.getElementById(ShowID).innerHTML = str; 
				}else if(str!=""){
					alert(str)	
				}else{
					alert("显示"+str+URL)
				}
			}
			else {
				if(ShowID!="undefined" && ShowID!=""){
					document.getElementById(ShowID).innerHTML = "<img src=/Images/lodin.gif>正在加载..."
				}
			}
		}
		xmlhttp.send(null);	
		//alert("网络错误");
	}
}
function testcall(){
 Ajax("http://aa/4.asp?a=a","msg")
 return false;
}


</script>
<a href="javascript:;" onclick="testcall()">testcall</a>
<div id="msg">msgstr</div>