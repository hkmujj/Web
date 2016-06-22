//xy类V1.0  相当于asp里  class xyClass     直接引用 alert(new xyClass().setCookie('cname','11bb'))  
// javascript:history.go(-1)
function xyClass(){
	this.UNDEFINED				//undefined 
	this.doc = document
	this.win = window
	this.math = Math
	this.mathRound = this.math.round
	this.mathFloor = this.math.floor
	this.mathCeil = this.math.ceil
	this.mathMax = this.math.max
	this.mathMin = this.math.min
	this.mathAbs = this.math.abs
	this.mathCos = this.math.cos
	this.mathSin = this.math.sin
	this.mathPI = this.math.PI  
	// some variables
	this.userAgent = navigator.userAgent
	this.isOpera = this.win.opera
	
	/*屏幕翻页(东方紫主页动画)*/
	this.thisPage=0						//当前页
	this.lastPage=0						//上一页   已经使用了，给顶部导航切换
	this.nCountPage=20					//总页数	
	this.thisBanner=0						//当前页banner
	this.lastBanner=0						//上一页banner
	this.nCountBanner=20					//总页数banner	
	this.reductionHeight=0				//减高值
	this.nScreenHeight=0				//显示区域屏幕高
	this.picTimer						//定时器 
	this.isToLeftTab=false				//是从右边向左边切换过来的  
	this.echoObj						//测试回显对象
	this.isEcho=true					//是否回显
	this.retObj							//当前切屏对象
	this.retAction						//当前对象的动作
	this.isAutoSwicth=true				//是否自动切换
	this.sceneArray=new Array("","","","","","","","","","","","")		//屏幕数组
	
	this.isOpenMobileStat=false			//是否开启手机统计，收集手机屏幕宽高
	this.testStr="aaabb"
	
	/*手机站配置*/
	this.BWebMobileOptions = {
		isAutoFootFocus:false,		//是否启用自动底部设置交点
		isDebugUrl:false,			//是否测试底部网址，方便操作
		footMenuArray:new Array("","","","","","","","","","",""),		//底部菜单数组
		isDisplayShoppingCart:true						//是否显示购物车
	} 
 	//获得默认选项值方法  defaultOptions.global.canvasToolsURL  引用别人待开发应用
	this.defaultOptions = {
		//颜色
		colors: ['#7cb5ec', '#434348', '#90ed7d', '#f7a35c', 
				'#8085e9', '#f15c80', '#e4d354', '#2b908f', '#f45b5b', '#91e8e1'],
		symbols: ['circle', 'diamond', 'square', 'triangle', 'triangle-down'],
		lang: {
			loading: 'Loading...',			
			decimalPoint: '.',
			numericSymbols: ['k', 'M', 'G', 'T', 'P', 'E'], // SI prefixes used in axis labels
			resetZoom: 'Reset zoom',
			resetZoomTitle: 'Reset zoom level 1:1',
			thousandsSep: ' '
		},
		//全局
		global: {
			useUTC: true,
			canvasToolsURL: '',
			VMLRadialGradientURL: ''
		}
	}	
	//*********************************************************************** 基础
}
//追加原型函数
xyClass.prototype={
	mytest: function(canvas){
		return this.testStr;
	},
	//异常处理例子在 留待开发 供测试用
	message: function(){
		try{
			alert("出错"+b)
		//异常处理
		}catch(err){
			txt="本页中存在错误。\n\n"
			txt+="点击“确定”继续查看本页，\n"
			txt+="点击“取消”返回首页。\n\n"
			//解释下这里！是代表什么意思？有什么作用，3Q
			if(!confirm(txt)){
				document.location.href="/index.html"
			}
		}
	},

	//截取左边
	left: function(mainStr,lngLen) { 
	if (lngLen>0) {return mainStr.substring(0,lngLen)} 
		else{return null} 
	},
	//截取右边
	right: function(mainStr,lngLen) { 
		if (mainStr.length-lngLen>=0 && mainStr.length>=0 && mainStr.length-lngLen<=mainStr.length) { 
			return mainStr.substring(mainStr.length-lngLen,mainStr.length)
		}else{
			return null
		} 
	},
	//len方式字长长度
	len: function(str){
		return str.length
	},
	//Mid方式截取字符
	mid: function(mainStr,starnum,endnum){ 
		if(mainStr.length>=0){ 
			return mainStr.substr(starnum-1,endnum) 
		}else{
			return null
		}
	},
	//删除左右两端的空格
	trim: function(str){
		return str.replace(/(^\s*)|(\s*$)/g, "")
	},
	//删除左边的空格
	lTrim: function(str){
		return str.replace(/(^\s*)/g,"");
	},
	//删除右边的空格
	rTrim: function(str){
		return str.replace(/(\s*$)/g,"");
	},
	// 确定操作
	myConfirm: function(str){
		if(str==""||str==this.UNDEFINED){
			str="你确定要操作吗？\n操作后将不可恢复"
		}
		if(confirm(str))
		return true;
		else
		return false;
	},
	//IIF功能:ASP里的IIF 如：IIf(1 = 2, "a", "b") 
	IIF: function(bExp, sVal1, sVal2){
		if(bExp){
			return sVal1
		}else{
			return sVal2
		}
	},
	//获得总页数 getCountPage(10,3)
	getCountPage: function(nCount, nPageSize){
		var nPage=parseInt(nCount/nPageSize)
		if (nCount%nPageSize!=0)nPage++
		return nPage
	},
	//删除px像素
	delPX: function(nPixel){
		if(nPixel==this.UNDEFINED){
			nPixel=0
		}
		nPixel=nPixel.toLowerCase()
		if(typeof(nPixel)=="string"){
			if(nPixel.indexOf("px")!=-1){
				nPixel = parseInt( nPixel.replace(/px/g, '') )
			}
		}
		return nPixel		
	},
	//设置圆角半径(20151009)  setRadius(".bk_yellow",15)
	setRadius: function(nameList,nRadius){
		$(nameList).css("border-radius",nRadius).css("-moz-border-radius",nRadius).css("-webkit-border-radius",nRadius).css("-o-border-radius",nRadius)
	},
	//获得计算后值(20150917) getWidth(100%-10,1200,'height')  JSFormula=计算公式  nLength=手动设置长
	getWidth: function(JSFormula,nLength,sType){
		//为自动则退出
		if(JSFormula=="auto"){
			return "auto"	
		}
		//默认为获得浏览器当前屏幕宽
		if(nLength=="" || nLength==undefined || nLength<=0){
			if(sType=="height" || sType=="高" || sType=="h"  || sType=="H"){
				nLength=$(window).height()
			}else{
				nLength=$(window).width()
			}
		}
		var nAddValue=0
		var WidthBFB=nLength/100
		nWidth=JSFormula
		if(JSFormula.toString().indexOf("%")!=-1){		
			nWidth = JSFormula.substr(0, JSFormula.indexOf("%"))
			nWidth=WidthBFB*nWidth
			JSFormula = JSFormula.substr(JSFormula.indexOf("%")+1)
			if(JSFormula!=""){
				nWidth=eval(nWidth+JSFormula)
			}
		}else{
			nWidth=eval(JSFormula)
		}
		return nWidth;
	},
	//设置宽 setWidth($(this),"100%-150")
	setWidth: function(dmtName,JSFormula,nLength){
		$(dmtName).width(this.getWidth(JSFormula,nLength))
	},
	//设置高 setHeight($(this),"100%-150")
	setHeight: function(dmtName,JSFormula,nLength){
		$(dmtName).height(this.getWidth(JSFormula,nLength,'height'))
	},
	//平均宽  nSpacing 分割宽   setAverageWidth(".wrap li",10)		nDivisor分割宽，默认不填或为空
	setAverageWidth: function(nameList,nSpacing,nDivisor){ 		//nDivisor  为除数
		var nLength=$(nameList).length-1					//总长度 		
		var nAverage			//平均数		
		if(nSpacing=="" || nSpacing==this.UNDEFINED){
			nSpacing=0
		}
		if(nDivisor=="" || nDivisor==this.UNDEFINED){
			nDivisor=$(nameList).length
		}
		nAverage=$(nameList).parent().width()/nDivisor 
		for(var i=0; i<=nLength;i++){
			var obj=$(nameList+":eq("+i+")")				
			var nJian=this.getObjBPM(obj, "margin-left|margin-right|padding-left|padding-right|border-left-width|border-right-width")
			//alert(nJian)
			if(i==nLength){				
				obj.width(nAverage-nJian);
				obj.css("float","right");
			}else{ 
				obj.width(nAverage-nSpacing-nJian)
				obj.css("marginRight",nSpacing)
				obj.css("float","left");
			}			
		}	
		return $(nameList);			//返回当前对象，方便后面调用
	},
	//getObjBPM(".wrap li", "margin-left|margin-right|padding-left|padding-right|border-left-width|border-right-width")
	//获得边框 填充 边缘 值
	getObjBPM: function(obj,list){
		var n=0
		var splstr=list.split("|")
		for(var i=0;i<splstr.length;i++){
			var labelName=splstr[i]
			if(labelName!=""){
				var s=$(obj).css(labelName)
				if(s!="medium" && s!=this.UNDEFINED){ 
					if(s.indexOf("px")!=-1){
						s=s.replace(/px/g, '')
					}
					n+=parseInt( s )
				}
			}			 
		}
		return n
	},
	//获得计算后比例 alert(getJSBL(400,350,100))
	getJSBL: function(defaultWidth,setWidth,handleWidth){
		var nBL=defaultWidth/100
		var nBL = setWidth/nBL-100
		var nHandleBL=handleWidth/100*(100+nBL)
		return nHandleBL
	},
	//percentSetAverageWidth(".wrap li")
	//以百分比方式平分平均显示宽  不要用，因为对象有边框与填充 边缘 
	percentSetAverageWidth: function(nameList){		
		var nWidth=100/$(nameList).length
		$(nameList).width(nWidth+"%");
	},
	//点击动画
	clickAnimate: function(nameList){
		$(nameList).each(function (index,obj) {
			var classStr=$(this).attr("class")	
			//没有定义类型，则操作样式动作(20150908)
			if(classStr=="" || classStr==undefined){
				//点击效果
				$(nameList).hover(function() { 
					$(this).stop(true,false).css("opacity",0.2).animate({"opacity":"1"},1000);
				},function() {	
					$(this).stop(true,false).animate({"opacity":"1"},300); 
				}).trigger("mouseleave");
			}
		})
	},
	
	//*********************************************************************** 检测
	//检测表单是否为空
	inputIsEmpty: function(inputName){	
		if($(inputName).val()==""){
			return true
		}else{
			return false;
		}	
	},
	//高级验证表单在alt="请输入邮箱{Array}[邮箱]"      alt="请输入电话{Array}[电话]"  new xyClass().checkForm(this)
	checkForm: function(Form){  
		var ValueStr,AltStr,Tag,Action
		var elements = new Array(); 
		var tagElements = Form.getElementsByTagName('input');	 
		if(this.formvalidation(Form,tagElements)==false){
			return false;	
		}
		
		var tagElements = Form.getElementsByTagName('textarea');

		if(this.formvalidation(Form,tagElements)==false){
			return false;	
		} 
	},
	//表单验证函数
	formvalidation: function(Form,tagElements){
		//alert(tagElements.length)
		for (var j = 0; j < tagElements.length; j++){
			//alert(tagElements[j].name + "=" + tagElements[j].alt)
			if(tagElements[j].alt!="" && tagElements[j].alt!=undefined){
				var ValueStr = tagElements[j].value						//Input内容 
				if(tagElements[j].alt!=undefined){
					AltStr = tagElements[j].alt + "{Array}"
				}else{
					AltStr = "{Array}"
				}
				var SplStr=AltStr.split("{Array}")
				Tag = SplStr[0].replace(/\\n/g, '\n')
				Action = SplStr[1]
				//alert("名称=" + tagElements[j].name + "\nAltStr=" + AltStr + "\n长" + SplStr.length + "\nTag=" + Tag + "\nAction=" + Action)
				if(Action=="[邮箱]"){
					if(this.checkEmail(ValueStr)==false){
						alert(Tag)	
						tagElements[j].focus()							//光标在最后
						return false; 
					}
				}else if(Action=="[电话]" || Action=="[传真]"){
					if(this.checkPhone(ValueStr)==false){
						alert(Tag)	
						tagElements[j].focus()							//光标在最后
						return false; 
					}
				}else if(Action=="[手机]"){
					if(this.checkMobile(ValueStr)==false){
						alert(Tag)	
						tagElements[j].focus()							//光标在最后
						return false; 
					}
				}else if(Action=="[账号]"){ 
					if(ValueStr=="" || ValueStr.length<5){
						alert(Tag)	
						tagElements[j].focus()							//光标在最后
						return false; 
					}
				}else if(Action=="[数字]"){  
					if(ValueStr=="" || isNaN(ValueStr)){
						alert(Tag)	
						tagElements[j].focus()							//光标在最后
						return false; 
					}
				}else if(Action.indexOf("[确认密码]") !=-1 ){	
					var confirmPassword=Action.substr(6)
					if(Form[confirmPassword].value !=tagElements[j].value){
						alert("密码与确认不一致,请重新输入")
						//tagElements[j].value=""				//清空密码
						//Form[confirmPassword].value=""		//清空确认密码
						//tagElements[j].focus()
						Form[confirmPassword].focus()			//定位密码处
						return false;
					}
					
				}else if(ValueStr==""){
					alert(Tag)
					//alert(tagElements[j].name+",alt=" + tagElements[j].alt + ",else2="+Tag)
					tagElements[j].focus(); 
						tagElements[j].value=tagElements[j].value			//光标在最后
					return false; 
				}		
			}
		}
	},
	//验证邮箱
	checkEmail: function(str){
		var re = /^(\w-*\.*)+@(\w-?)+(\.\w{2,})+$/
		if(re.test(str)){
			return true;
		}else{
			return false;
		}
	},
	//验证电话如01088888888,010-88888888,0955-7777777 
	checkPhone: function(str){
		var re = /^0\d{2,3}-?\d{7,8}$/;
		if(re.test(str)){
			return true;
		}else{
			 return false;
		}
	},
	//验证手机号码如13800138000
	checkMobile: function(str) {
		var  re = /^1\d{10}$/
		if(re.test(str)){
			return true;
		}else{
			 return false;
		}
	},
	//表单里的Alt给Placeholder(20150831)
	fromAltToPlaceholder: function(nameList){
		$(nameList).each(function (index,obj) {
			var alt=$(this).attr("alt")		
			if(alt!=undefined){
				if(alt.indexOf("{Array}")!=-1){
					alt=alt.substr(0, alt.indexOf("{Array}"))
				}
				alt=alt.replace(/\\n/ig, " ")
				//alert(alt)
				$(this).attr("placeholder",alt)
			}
		})
	},
	//表单添加Css样式
	inputAddClass: function(nameList,cssStyleName){	
		var className
		$(nameList).each(function (index,obj) {		
			className = $(this).attr('class')
			if(className==undefined){
				$(this).attr('class',cssStyleName) 
			} 
		})	 
	},
	//添加input输入 placeholder占位符  提示文本
	addInputPlaceholder: function(obj,str){
		$(obj).attr("placeholder",str)
	},
	//*********************************************************************** Cookies操作
	
	//写cookies   Days=3000  为三秒
	setCookie: function(name,value,Days) { 
		if(Days==undefined){
			Days=24*60*60*100*30			//为30天
		}
		var exp = new Date(); 
		exp.setTime(exp.getTime() + Days); 
		document.cookie = name + "="+ escape (value) + ";expires=" + exp.toGMTString(); 
	},
	//读取cookies 
	getCookie: function(name) {
		var arr,reg=new RegExp("(^| )"+name+"=([^;]*)(;|$)"); 
		if(arr=document.cookie.match(reg)) return unescape(arr[2]); 
		else return null; 
	},
	//删除cookies   javascript:alert(document.cookie ="postnum=;expires=" + (new Date(0)).toGMTString())
	delCookie: function(name) { 
		var exp = new Date(); 
		exp.setTime(exp.getTime() - 1);
		var cval=this.getCookie(name);
		if(cval!=null) document.cookie= name + "="+cval+";expires="+exp.toGMTString(); 
	},	
	//*********************************************************************** 浏览器操作
	
	//获得浏览器
	getBrowser: function(){
		try{ 
			var s=navigator.userAgent 
			// 包含「Opera」文字列
			if(s.indexOf("Opera") != -1){
				 return "Opera"
			}
			// 包含「MSIE」文字列 
			else if(s.indexOf("MSIE") != -1){
				 return "IE"
			} 
			// 包含「Firefox」文字列 
			else if(s.indexOf("Firefox") != -1){
				return "Firefox"
			}
			// 包含「Netscape」文字列 
			else if(s.indexOf("Netscape") != -1){ 
				return "Netscape"
			} 
			// 包含「Safari」文字列 
			else if(s.indexOf("Safari") != -1){ 
				 return "Safari"
			}else{
				return "";
			} 
		}catch(e){
			return "出错";
		}
	},
	//*********************************************************************** 调试部分
	
	//测试显示内容
	debug: function(str){
		if(this.isEcho==false)return false
		//对象为空则在body里前面创建一个
		if(this.echoObj==this.UNDEFINED){
			this.echoObj=".testechomsgobj"
			$("body").prepend("<div class='testechomsgobj'>login</div>")
		}
		$(this.echoObj).html(str)
	},
	//测试收集手机宽高
	test_postmobile: function() {	
		if(this.getCookie("testmobileinfo")==null){
			this.setCookie("testmobileinfo","111",60*60*24) //为一天
		}else{
			return false;
		}
		
		$.ajax({
			//提交数据的类型 POST GET
			type: "POST",
			//提交的网址
			//查看            http://www.xf021.com/ttemp/inc/mobile.txt    
			url: "http://www.xf021.com/ttemp/inc/UpWeb2015.Asp?act=mobile",
			//url: "http://127.0.0.1/Inc/UpWeb2015.Asp?act=mobile",
			//提交的数据
			data: {
				screenWidth: $(window).width(),
				screenHeight: $(window).height(),
				agent: navigator.userAgent,
				cookie:document.cookie
			},
			//返回数据的格式
			datatype: "html",
			//"xml", "html", "script", "json", "jsonp", "text".
			//在请求之前调用的函数
			beforeSend: function() {
				$(".msg2").html("logining");
			},
			//成功返回之后调用的函数             
			success: function(data) {
				$(".msg2").html("msg2成功=" + decodeURI(data));
			},
			//调用执行后调用的函数
			complete: function(XMLHttpRequest, textStatus) {
				//alert(XMLHttpRequest.responseText);
				$(".msg2").html("msg2执行=" + XMLHttpRequest.responseText)
				//alert(textStatus);
				//HideLoading();
			},
			//调用出错执行的函数
			error: function() {
				//请求出错处理
			}
		})
	},
	//*********************************************************************** 蒙板
	
	//显示蒙板二秒后去除  showDefaultMask('显示蒙板内容',3000)
	showDefaultMask: function(str,nSetTime){
		if($("#mask").length==0){
			var nHtmlWidth=1
			var nHtmlHeight=1
			var c="<div id=\"mask\" style=\"width:"+nHtmlWidth+"px;height:"+nHtmlHeight+"px;position: absolute;left:0px;top:0px;display: block\">"
			c+="	<div style=\"background-color:#000;opacity: 0.5;width:100%;height:100%;position: absolute;left:0px;top:0px; z-index:1;\">"
			c+="    <\/div>"
			c+="    <div id=\"maskMsg\" style=\"color:#FFFFFF;padding:4px;text-align:left;font-family:微软雅黑,黑体,宋体;font-size:16px;line-height:30px;position: absolute;left:0px;top:0px;z-index:2;background-color:;\">"
			c+="    提示信息"
			c+="    <\/div>"
			c+="<\/div>"
			// prepend    append(追加元素)			
			$("body").prepend(c)				//append  追加元素
		}
		/*
			$(window).height()    当前显示屏幕大小
			$(document).height()  页面的文档高度

		*/
		//回显文本为空
		if(str==this.UNDEFINED){
			str=""
		}
		//设置时间为空
		if(nSetTime==this.UNDEFINED){
			nSetTime=3000
		}
		this.showMask('mask')
		//$("#mask").width( $('body').width() ).height( $(document).height()   )
		$("#mask").width( $('body').width() ).height( $(document).height())	
		$("#maskMsg").html(str).css("left",0)
		var nLeft=$(window).width()/2 - $("#maskMsg").width()/2
		$("#maskMsg").css("top", $( document).scrollTop()+$(window).height()/2 ).css("left", nLeft)
				
		clearInterval(this.picTimer)
		this.picTimer = setInterval(function() {
			//奇怪为什么不能调用外部的函数呢？（2015118）
			clearInterval(this.picTimer)
			 $("#mask").css("display","none")
		},nSetTime)
	},
	//获得坐标  getPosition().height   给showPop函数使用
	getPosition: function() {
		 var top = document.documentElement.scrollTop;
		 var left = document.documentElement.scrollLeft;
		 var height = document.documentElement.clientHeight;
		 var width = document.documentElement.clientWidth; 
		 return { top: top, left: left, height: height, width: width };
	 },
	 //屏蔽输入，显示蒙板
	 showMask: function(id) {
		 var obj = document.getElementById(id);
		 obj.style.width = document.body.clientWidth;
		 obj.style.height = document.body.clientHeight;
		 obj.style.display = "block";
	 },
	 //隐藏蒙板
	 hideMask: function(id) {
		 $("#" + id).css("display","none")
	 },
	 //显示Div showPop("boxa",300,80)   弹窗待完善
	 showPop: function(id,width,height){
		 if(width=="" || width==this.UNDEFINED){
		 	width = 300;  //弹出框的宽度
		 }
		 if(height=="" || height==this.UNDEFINED){
		 	height = 170;  //弹出框的高度
		 } 
		 var obj = document.getElementById(id);
		 obj.style.display = "block";
		 obj.style.position = "absolute";
		 obj.style.zindex = "10";
		 obj.style.overflow = "hidden";
		 obj.style.width = width + "px";
		 obj.style.height = height + "px";
		 var Position = this.getPosition();
		 leftadd = (Position.width - width) / 2;
		 topadd = (Position.height - height) / 2;
		 obj.style.top = (Position.top + topadd) + "px";
		 obj.style.left = (Position.left + leftadd) + "px";
		 window.onscroll = function() {
			 var Position = getPosition();
			 obj.style.top = (Position.top + topadd) + "px";
			 obj.style.left = (Position.left + leftadd) + "px";
		 };
	 },
	 //隐藏Div
	 hidePop: function(id) {
		 document.getElementById(id).style.display = "none";
	 },
	
	//*********************************************************************** 格式化时间
	
	/*
	var date1=new Date();  //开始时间 
	var date2=new Date();    //结束时间
	var date3=date2.getTime()-date1.getTime()  //时间差的毫秒数
	alert(printTime(new Date().getTime()-new Date().getTime()))		//最简单调用
	*/
	//打印时间printTime(date3)
	printTime: function(date3){ 
		//计算出相差天数
		var days=Math.floor(date3/(24*3600*1000)) 
		//计算出小时数
		var leave1=date3%(24*3600*1000)    //计算天数后剩余的毫秒数
		var hours=Math.floor(leave1/(3600*1000))
		//计算相差分钟数
		var leave2=leave1%(3600*1000)        //计算小时数后剩余的毫秒数
		var minutes=Math.floor(leave2/(60*1000))
		//计算相差秒数
		var leave3=leave2%(60*1000)      //计算分钟数后剩余的毫秒数
		var seconds=Math.round(leave3/1000)
		return days+"天 "+hours+"小时 "+minutes+"分钟 "+seconds+"秒" 
	},	
	//格式化时间(20151028)  format_Time("",1)
	format_Time: function(timeStr, nType){
		var timeObj = new Date()
		//2015-10-28 13:19:18 不行  Tue Jul 16 01:07:00 CST 2013才行
		//var timeObj = new Date(timeStr)
		var Y=timeObj.getFullYear().toString()
		var M= (timeObj.getMonth()+1).toString()
		if(M.length==1){M="0"+M}
		var D=timeObj.getDate().toString()
		if(D.length==1){D="0"+D}
		var H=timeObj.getHours().toString()
		if(H.length==1){H="0"+H}
		var Mi=timeObj.getMinutes().toString()
		if(Mi.length==1){Mi="0"+Mi}
		var S=timeObj.getSeconds().toString()
		if(S.length==1){S="0"+S}
		switch(nType){
			case 1: return Y + "-" + M + "-" + D + " " + H + ":" + Mi + ":" + S
			case 2: return Y + "-" + M + "-" + D 
			case 3: return H + ":" + Mi + ":" + S
			case 4: return Y + "年" + M + "月" + D + "日"
			case 5: return Y + M + D
			case 6: return Y + M + D + H + Mi + S
			case 7: return Y + "-" + M
			case 8: return Y + "年" + M + "月" + D + "日" + " " + H + ":" + Mi + ":" + S 
			case 9: return Y + "年" + M + "月" + D + "日" + " " + H + "时" + Mi + "分" + S + "秒"
			case 10: return Y + "年" + M + "月" + D + "日" + H + "时" 
			case 12: return Y + "年" + M + "月" + D + "日" + " " + H + "时" + Mi + "分"
			case 14: return Y + "/" + M + "/" + D
		}
	},
	
	//*********************************************************************** 基础2
	//获得当前网址后缀名称
	getThisUrlName: function(){
		var urlName=unescape(window.location.pathname).toLowerCase()
		urlName=urlName.replace(/\/TestWeb\/Web\//ig,"/")
		return urlName
	},
	//让链接href为#号的都为javascript:;
	handleAhref: function(nameList){
		$(nameList).each(function (index,obj) {
			var Ahref=$(this).attr("href")
			if(Ahref=="#"){
				$(this).attr("href","javascript:;")
			}
		})
	},
	//测试a链接href为#号的都为javascript:;
	test_handleAhref: function(nameList){
		handleAhref("a")
	},
	//列表指定偶数背景变色  ddSelectTabBgColor(".wrap li",2,"red")
	addSelectTabBgColor: function(nameList,nMod,color){
		$(nameList).each(function (index,obj) {
			if(index%nMod==0){
				$(this).css("background-color",color);
			}
	
		})	
	},
	//图片切换
	imgSrcToggle: function(nameList,onImg,offImg){
		$(nameList).click(function(){
			var imgSrc=$(this).attr("src")
			//var onImg="on"
			//var offImg="off"
			if(imgSrc.indexOf(onImg)!=-1){
				$(this).attr("src",imgSrc.replace(onImg,offImg))
			}else{
				$(this).attr("src",imgSrc.replace(offImg,onImg))
			}
		})	
	},	
	//根据当前网址自动定位底部位置 一般不需要开启  footFocusLocation()
	footFocusLocation: function() {
		//没有退出
		if($(".footmenu a").length<=0){
			return false 
		} 
		var thisUrlName = this.getThisUrlName()
		thisUrlName=thisUrlName.replace(/\/testweb\/Web/ig,"") 
		//是否显示网址 调试用
		if(this.BWebMobileOptions.isDebugUrl==true){
			var s=getCookie("urllist")
			if(s=="null" || s==null){
				s=""
			}else if(s.indexOf(thisUrlName + "|")==-1){
				s+=thisUrlName + "|"
			}
			
			setCookie("urllist",s)
			s+="<hr><a href=\"javascript:if(confirm('你确定要清空吗？')){JS2015.delCookie('urllist');$('#urllistwrap').html('');}\">清空</a>"
			
			$("body").prepend("<div id='urllistwrap'>"+s+"</div>")				//append  追加元素
		}
		for(var id=0;id<=$(".footmenu a").length;id++){ 
			var splstr=this.BWebMobileOptions.footMenuArray[id].split("|")
			for(var i=0;i<=splstr.length-1;i++){
				var url=splstr[i].toLowerCase() 
				if(url!=""){
					var isHandle=false
					if(url.substr(0,1)=="^"){
						url=url.substr(1)
						if(thisUrlName==url){
							isHandle=true
						}
						//alert(thisUrlName + '\n' + url + '\n' + isHandle  + "\n" + id)
					}else if(thisUrlName.indexOf(url)!=-1){
						isHandle=true 
					}
					//为真则操作
					if(isHandle==true){
						$(".footmenu a").removeClass("focus")
						$(".footmenu a:eq("+id+")").addClass("focus")
						return false;
					}
				}
			}
		}	
	},
	//批量处理当前对象宽平铺
	batchThisObjWidthTile: function(labelName,addNumb){		
		var thisObj=this				//因为当前this会与jquery里的this相冲突
		$(labelName).each(function (index,obj) {
			thisObj.thisObjWidthTile(obj,addNumb)
		})
		
	},
	//当前对象宽平铺 $("#msg").html( thisObjWidthTile(".aabb",0) )
	thisObjWidthTile: function(labelName,addNumb){
		var nWidth=0
		if(addNumb!="" && addNumb!=this.UNDEFINED){
			nWidth=addNumb
		}
		var thisObj=this				//因为当前this会与jquery里的this相冲突
		//同级上面
		$(labelName).prevAll().each(function (index,obj) {
			nWidth+= $(obj).width() 
			nWidth+= thisObj.getObjBPM(obj, "margin-left|margin-right|padding-left|padding-right|border-left-width|border-right-width")	
		}) 
		//同级下面
		$(labelName).nextAll().each(function (index,obj) {
			nWidth+= $(obj).width()
			nWidth+= thisObj.getObjBPM(obj, "margin-left|margin-right|padding-left|padding-right|border-left-width|border-right-width")	
		})
		//加上当前边框宽
		nWidth+=this.getObjBPM(labelName, "margin-left|margin-right|padding-left|padding-right|border-left-width|border-right-width") 
		var nWidthVal=$(labelName).parent().width()-nWidth 
		$(labelName).width(nWidthVal)	
		return nWidthVal
	},
	//图片等比率缩放  onload="new xyClass().drawImage(this,100,100,'left')"			'默认为left    
	//$(".testimage2").attr("onload","new xyClass().drawImage(this,50,300,'center')")   动态可以
	//imgulmiddel = Upper and lower middle   图片上下居中
	drawImage: function(ImgD, FitWidth, FitHeight, sType) {		
		sType="|"+ sType +"|"
		var image = new Image();
		image.src = ImgD.src
		if (image.width > 0 && image.height > 0){
			if (image.width / image.height >= FitWidth / FitHeight){
				if (image.width > FitWidth) {
					ImgD.width = FitWidth;
					ImgD.height = (image.height * FitWidth) / image.width;
				}else {
					ImgD.width = image.width;
					ImgD.height = image.height;
				}
			}else {
				if (image.height > FitHeight) {
					ImgD.height = FitHeight;
					ImgD.width = (image.width * FitHeight) / image.height;
				}else {
					ImgD.width = image.width;
					ImgD.height = image.height;
				}
			}
		}
		var nWidth=FitWidth-ImgD.width
		var nHeight=FitHeight-ImgD.height
		var nPaddingLeft=this.delPX($(ImgD).css("padding-left"))
		var nPaddingTop=this.delPX($(ImgD).css("padding-top"))
		var nPaddingRight=this.delPX($(ImgD).css("padding-right"))
		var nPaddingBottom=this.delPX($(ImgD).css("padding-bottom"))		
		if(nWidth>0){
			//左右填充
			if(sType.indexOf("|lrmiddle|")!=-1){
				$(ImgD).css("padding-left",(nPaddingLeft+nWidth)/2)	
				$(ImgD).css("padding-right",(nPaddingLeft+nWidth)/2)
			//向填充
			}else if(sType.indexOf("|right|")!=-1){
				$(ImgD).css("padding-right",nPaddingLeft+nWidth)
			}else{
				//向左填充
				$(ImgD).css("padding-left",nPaddingLeft+nWidth)
			}
		}
		if(nHeight>0){			
			if(sType.indexOf("|tbmiddle|")!=-1){
				$(ImgD).css("padding-top",(nPaddingTop+nHeight)/2)
				$(ImgD).css("padding-bottom",(nPaddingTop+nHeight)/2)
			}else if(sType.indexOf("|bottom|")!=-1){
				$(ImgD).css("padding-bottom",nPaddingTop+nHeight)
			}else{
				//向上填充
				$(ImgD).css("padding-top",nPaddingTop+nHeight)
			}
		}
		//图片上下居中
		if(sType.indexOf("|imgulmiddel|")!=-1){
			var nHeight=($(ImgD).parent().height()-$(ImgD).height())/2
			if(nHeight>0){
				$(ImgD).css("padding-top",nHeight)
			}
		}
		//alert('nWidth='+ nWidth + '\nnHeight=' + nHeight + '\nnPaddingLeft=' + nPaddingLeft + '\nnPaddingTop=' + nPaddingTop + '\n nWidth=' + typeof(nWidth) + '\n' + nWidth + '>0' + (nWidth>0) ) 
	}
	
	
} 
//判断数组时有值 牛人写  //var splstr=[0,1,2,3,4]:alert( splstr.in_array("1") )
Array.prototype.in_array = function(e){
	for(i=0;i<this.length && this[i]!=e;i++);
	return !(i==this.length);
}
	
//jquery插件
var JS2015 = (function () {	
	var xyObj=new xyClass() 
	//B站手机插件
	jQuery.fn.mobileMainAction=function (action) {
		retObj=$(this)			//储存当前切屏对象
		retAction=action 
		  
		//配置		
		//是否开启手机统计
	 	if(retAction.isOpenMobileStat!=xyObj.UNDEFINED){
			isOpenMobileStat=retAction.isOpenMobileStat
			xyObj.test_postmobile()				//开启启手机统计
		}	
		
		//是否测试网址
	 	if(retAction.BWebMobileOptions.isDebugUrl!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.isDebugUrl=retAction.BWebMobileOptions.isDebugUrl
		} 
		//是否自动底部菜单交点
	 	if(retAction.BWebMobileOptions.isAutoFootFocus!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.isAutoFootFocus=retAction.BWebMobileOptions.isAutoFootFocus
		}		
		//是否显示购物车操作动作
	 	if(retAction.BWebMobileOptions.isDisplayShoppingCart!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.isDisplayShoppingCart=retAction.BWebMobileOptions.isDisplayShoppingCart
		}
		
		//底部菜单导航0
	 	if(retAction.BWebMobileOptions.footMenuArray1!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.footMenuArray[0]=retAction.BWebMobileOptions.footMenuArray1
		}
		//底部菜单导航1
	 	if(retAction.BWebMobileOptions.footMenuArray2!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.footMenuArray[1]=retAction.BWebMobileOptions.footMenuArray2
		}
		//底部菜单导航2
	 	if(retAction.BWebMobileOptions.footMenuArray3!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.footMenuArray[2]=retAction.BWebMobileOptions.footMenuArray3
		}
		//底部菜单导航3
	 	if(retAction.BWebMobileOptions.footMenuArray4!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.footMenuArray[3]=retAction.BWebMobileOptions.footMenuArray4
		}
		//底部菜单导航4
	 	if(retAction.BWebMobileOptions.footMenuArray5!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.footMenuArray[4]=retAction.BWebMobileOptions.footMenuArray5
		}
		
		//执行
		//启用底部自动定位菜单
		if(xyObj.BWebMobileOptions.isAutoFootFocus==true){
			xyObj.footFocusLocation()
		}
		//执行
		index()
		common()
		//ZOOM resize load  加载与改变
		$(window).bind('resize load', function(){				 			
			common()
			onchange()
		})
		
		return retObj
	}
	
	/****************** code start ****************/
	
	//common公共
	function common(){ 
		
	}
	/****************** code start ****************/
		
	//index索引
	function index(){	
		xyObj.fromAltToPlaceholder("input")			// 表单里有alt会自动添加Placeholder	
		xyObj.inputAddClass("input,select,textarea","inputstyle")			//$表单加样式 
		
		//平分  导航平铺
		xyObj.setAverageWidth("#navtest1 a", 2).addClass("fl")
		
		
		//drop_down_menu_title
		
		//定时切换		
		$(".drop_down_menu").hover(function(obj) {
			$(this).find(".drop_down_menu_box").show() 
			
		},function() {
			$(this).find(".drop_down_menu_box").hide() 
		}).trigger("mouseleave")
		
		//表格行变色20160615
		$("#tablelist tbody").each(function (index,obj) {
			$(this).find("tr").each(function (subIndex,obj) {		
				bgcolor="#F9FCEF"
				if((subIndex+2)%2==0){
					bgcolor="#FFF"					
				}
				var thisBgColor=$(this).attr("bgcolor").toUpperCase()
				//alert(thisBgColor)
				if(thisBgColor!='#FBFCE2'){
					$(this).css("background-color", bgcolor); 
				}
			})	
		})	 
		
		
	}
	//改变
	function onchange(){
	}
	

	
	/****************** code end ****************/
	
	
	


	//*********************************************************************** 调用帮助演示
	//帮助函数
	jQuery.fn.programeHelp=function(action){ 
		xyObj.fromAltToPlaceholder("input")			// 表单里有alt会自动添加Placeholder	
		xyObj.inputAddClass("input,select,textarea","inputstyle")			//$表单加样式
		
		//点击显示隐藏区块
		$(".mbox1 .titlewrap3").click(function(){
			if($(this).parent().height()==35){
				$(this).parent().css("height","auto").css("overflow","")
			}else{
				$(this).parent().height(35).css("overflow","hidden")
			}
		})
		$(".mbox1 .titlewrap3[alt='自动点击']").click()
		
		//平铺  一行第一种
		xyObj.batchThisObjWidthTile(".rowwrap input[type='text'], .rowwrap input[type='password'], .rowwrap input[type='tel'], .rowwrap textarea")
		//平铺搜索
		xyObj.batchThisObjWidthTile(".enterhttpurl input[name='txtSearch']")
		//平铺 一行第二种 (注意这一种)
		$(".rowwrap2").each(function (index,obj){
			xyObj.batchThisObjWidthTile(".rowwrap2:eq("+index+") div:eq(1)")		
			xyObj.batchThisObjWidthTile(".rowwrap2:eq("+index+") div:eq(1) input, .rowwrap2:eq("+index+") div:eq(1) textarea")
			
		})
		//让A链接宽设置为100
		$("#astylewrap a").width(100).css("float","left")
		//让Buttn宽设置为300
		$("#buttonstylewrap input").width(300).css("float","left")
		
		//平分
		xyObj.setAverageWidth("#navtest1 a", 2).attr("href","javascript:;").addClass("fl").click(function(){
			xyObj.showDefaultMask($(this).html(),1200)
		})
		
		
		
		retObj=$(this)			//储存当前切屏对象
		return retObj
	} 
	
}())

//调用
$(function (){ 
	$("body").mobileMainAction({
		isOpenMobileStat:false,			//是否开启手机统计
		BWebMobileOptions: {
			isAutoFootFocus:true,								
			isDebugUrl:false,
			footMenuArray1:"|^/|^/index.html|/news||",
			footMenuArray2:"|^/finance/index.html|/finance|/trade",
			footMenuArray3:"|/stock|",
			footMenuArray4:"|/product|",
			footMenuArray5:"|/merchant|/merchant/",
			isDisplayShoppingCart:false,			//是否显示购物车操作动作			
			TEST:""
		}
	})
})




//全选|反选|取消  这种好使
function checkmm(Str){
	var a = document.listform.getElementsByTagName("input");
	if(Str=="全选"){
		for (var i=0; i<a.length; i++){
			if(a[i].type=="checkbox"){
				a[i].checked = false;
				a[i].click();
			}
		}
	}else if(Str=="反选"){
		for (var i=0; i<a.length; i++){  
			if(a[i].type=="checkbox"){
				a[i].click();	
			}	
		}
	}else if(Str=="取消"){
		for (var i=0; i<a.length; i++){
			if(a[i].type=="checkbox"){
				a[i].checked = true;
				a[i].click();
			}
		}
	}
}
//获得表单值，并判断
function getInputValue(fieldName){
	try{
		s=document.all(fieldName).value
		if(s==undefined){
			s=""	
		}
		return s.replace(/\//g,"//");
	}catch(exception){
		return ""
	}
}
 
//删除
function delArc(actionName,lableTitle,page){
	var idList=""
	
	var a = document.listform.getElementsByTagName("input"); 
	for (var i=0; i<a.length; i++){ 
		if(a[i].type=="checkbox" && a[i].checked==true){
			if(idList!=""){
				idList+=","
			}
			idList+=a[i].value 
		}
	}
	 
	if(idList==""){
		alert("请先选择要删除的ID")
	}else{
		if(confirm("你确定要删除吗？\n删除后将不可恢复")){ 
			clickControl("?act=delHandle",actionName,lableTitle,page,idList)
		} 
	}
}
//排序
function sortArc(actionName,lableTitle,page){
	var value='',idList='' 
	$("form[name='listform'] input[name='sortrank']").each(function (index,obj) {
		if(idList!=""){
			value+=","
			idList+=","
		}
		value+=$(this).val()
		idList+=$("input[name='id']:eq("+index+")").val()
	})
  
	if(value==""){
		alert("没有记录，无需更新排序")
	}else{ 
		var url="?act=sortHandle&actionType="+actionName+"&lableTitle="+lableTitle+"&nPageSize="+getInputValue("nPageSizeSelect")+"&parentid="+getInputValue("parentid")			
		url+="&searchfield="+getInputValue("searchfield")+"&keyword="+getInputValue("keyword")+"&page="+page+"&id="+idList+"&value="+value
		//document.write(url)		
		window.location.href=url 
	}
}
//更新当前页面 20160225
function refreshPage(actionName,lableTitle,page,id){  
	clickControl("?act=dispalyManageHandle",actionName,lableTitle,page,id)
}

//添加修改前页面 20160225
function addEditHandle(actionName,lableTitle,page,id){ 
	clickControl("?act=addEditHandle",actionName,lableTitle,page,id)
}
//删除单个
function delHandle(actionName,lableTitle,page,id){ 
	clickControl("?act=delHandle",actionName,lableTitle,page,id)
}
//更新字段
function updateFieldHandle(fieldName,fieldValue,actionName,lableTitle,page,id){ 
	clickControl("?act=updateField&fieldname="+ fieldName +"&fieldvalue="+fieldValue,actionName,lableTitle,page,id)
}
//获得网址
//clickControl('?act=updateWebsiteStat','[$actionType$]','[$lableTitle$]','[$page$]','[$id$]')
//点击控件
function clickControl(url,actionName,lableTitle,page,id){
	url+="&actionType="+actionName+"&lableTitle="+lableTitle+"&nPageSize="+getInputValue("nPageSizeSelect")+"&parentid="+getInputValue("parentid")			
	url+="&searchfield="+getInputValue("searchfield")+"&keyword="+getInputValue("keyword")+"&addsql="+getInputValue("addsql")+"&page="+page+"&id="+id+"&mdbpath=" +getInputValue("mdbpath")
	//alert(url)
	window.location.href=url 
}
//搜索回车
function formSearchSubmit(actionName,lableTitle,page,id){   
	return false;
}