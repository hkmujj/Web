//xy��V1.0  �൱��asp��  class xyClass     ֱ������ alert(new xyClass().setCookie('cname','11bb'))  
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
	
	/*��Ļ��ҳ(��������ҳ����)*/
	this.thisPage=0						//��ǰҳ
	this.lastPage=0						//��һҳ   �Ѿ�ʹ���ˣ������������л�
	this.nCountPage=20					//��ҳ��	
	this.thisBanner=0						//��ǰҳbanner
	this.lastBanner=0						//��һҳbanner
	this.nCountBanner=20					//��ҳ��banner	
	this.reductionHeight=0				//����ֵ
	this.nScreenHeight=0				//��ʾ������Ļ��
	this.picTimer						//��ʱ�� 
	this.isToLeftTab=false				//�Ǵ��ұ�������л�������  
	this.echoObj						//���Ի��Զ���
	this.isEcho=true					//�Ƿ����
	this.retObj							//��ǰ��������
	this.retAction						//��ǰ����Ķ���
	this.isAutoSwicth=true				//�Ƿ��Զ��л�
	this.sceneArray=new Array("","","","","","","","","","","","")		//��Ļ����
	
	this.isOpenMobileStat=false			//�Ƿ����ֻ�ͳ�ƣ��ռ��ֻ���Ļ���
	this.testStr="aaabb"
	
	/*�ֻ�վ����*/
	this.BWebMobileOptions = {
		isAutoFootFocus:false,		//�Ƿ������Զ��ײ����ý���
		isDebugUrl:false,			//�Ƿ���Եײ���ַ���������
		footMenuArray:new Array("","","","","","","","","","",""),		//�ײ��˵�����
		isDisplayShoppingCart:true						//�Ƿ���ʾ���ﳵ
	} 
 	//���Ĭ��ѡ��ֵ����  defaultOptions.global.canvasToolsURL  ���ñ��˴�����Ӧ��
	this.defaultOptions = {
		//��ɫ
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
		//ȫ��
		global: {
			useUTC: true,
			canvasToolsURL: '',
			VMLRadialGradientURL: ''
		}
	}	
	//*********************************************************************** ����
}
//׷��ԭ�ͺ���
xyClass.prototype={
	mytest: function(canvas){
		return this.testStr;
	},
	//�쳣���������� �������� ��������
	message: function(){
		try{
			alert("����"+b)
		//�쳣����
		}catch(err){
			txt="��ҳ�д��ڴ���\n\n"
			txt+="�����ȷ���������鿴��ҳ��\n"
			txt+="�����ȡ����������ҳ��\n\n"
			//����������Ǵ���ʲô��˼����ʲô���ã�3Q
			if(!confirm(txt)){
				document.location.href="/index.html"
			}
		}
	},

	//��ȡ���
	left: function(mainStr,lngLen) { 
	if (lngLen>0) {return mainStr.substring(0,lngLen)} 
		else{return null} 
	},
	//��ȡ�ұ�
	right: function(mainStr,lngLen) { 
		if (mainStr.length-lngLen>=0 && mainStr.length>=0 && mainStr.length-lngLen<=mainStr.length) { 
			return mainStr.substring(mainStr.length-lngLen,mainStr.length)
		}else{
			return null
		} 
	},
	//len��ʽ�ֳ�����
	len: function(str){
		return str.length
	},
	//Mid��ʽ��ȡ�ַ�
	mid: function(mainStr,starnum,endnum){ 
		if(mainStr.length>=0){ 
			return mainStr.substr(starnum-1,endnum) 
		}else{
			return null
		}
	},
	//ɾ���������˵Ŀո�
	trim: function(str){
		return str.replace(/(^\s*)|(\s*$)/g, "")
	},
	//ɾ����ߵĿո�
	lTrim: function(str){
		return str.replace(/(^\s*)/g,"");
	},
	//ɾ���ұߵĿո�
	rTrim: function(str){
		return str.replace(/(\s*$)/g,"");
	},
	// ȷ������
	myConfirm: function(str){
		if(str==""||str==this.UNDEFINED){
			str="��ȷ��Ҫ������\n�����󽫲��ɻָ�"
		}
		if(confirm(str))
		return true;
		else
		return false;
	},
	//IIF����:ASP���IIF �磺IIf(1 = 2, "a", "b") 
	IIF: function(bExp, sVal1, sVal2){
		if(bExp){
			return sVal1
		}else{
			return sVal2
		}
	},
	//�����ҳ�� getCountPage(10,3)
	getCountPage: function(nCount, nPageSize){
		var nPage=parseInt(nCount/nPageSize)
		if (nCount%nPageSize!=0)nPage++
		return nPage
	},
	//ɾ��px����
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
	//����Բ�ǰ뾶(20151009)  setRadius(".bk_yellow",15)
	setRadius: function(nameList,nRadius){
		$(nameList).css("border-radius",nRadius).css("-moz-border-radius",nRadius).css("-webkit-border-radius",nRadius).css("-o-border-radius",nRadius)
	},
	//��ü����ֵ(20150917) getWidth(100%-10,1200,'height')  JSFormula=���㹫ʽ  nLength=�ֶ����ó�
	getWidth: function(JSFormula,nLength,sType){
		//Ϊ�Զ����˳�
		if(JSFormula=="auto"){
			return "auto"	
		}
		//Ĭ��Ϊ����������ǰ��Ļ��
		if(nLength=="" || nLength==undefined || nLength<=0){
			if(sType=="height" || sType=="��" || sType=="h"  || sType=="H"){
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
	//���ÿ� setWidth($(this),"100%-150")
	setWidth: function(dmtName,JSFormula,nLength){
		$(dmtName).width(this.getWidth(JSFormula,nLength))
	},
	//���ø� setHeight($(this),"100%-150")
	setHeight: function(dmtName,JSFormula,nLength){
		$(dmtName).height(this.getWidth(JSFormula,nLength,'height'))
	},
	//ƽ����  nSpacing �ָ��   setAverageWidth(".wrap li",10)		nDivisor�ָ��Ĭ�ϲ����Ϊ��
	setAverageWidth: function(nameList,nSpacing,nDivisor){ 		//nDivisor  Ϊ����
		var nLength=$(nameList).length-1					//�ܳ��� 		
		var nAverage			//ƽ����		
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
		return $(nameList);			//���ص�ǰ���󣬷���������
	},
	//getObjBPM(".wrap li", "margin-left|margin-right|padding-left|padding-right|border-left-width|border-right-width")
	//��ñ߿� ��� ��Ե ֵ
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
	//��ü������� alert(getJSBL(400,350,100))
	getJSBL: function(defaultWidth,setWidth,handleWidth){
		var nBL=defaultWidth/100
		var nBL = setWidth/nBL-100
		var nHandleBL=handleWidth/100*(100+nBL)
		return nHandleBL
	},
	//percentSetAverageWidth(".wrap li")
	//�԰ٷֱȷ�ʽƽ��ƽ����ʾ��  ��Ҫ�ã���Ϊ�����б߿������ ��Ե 
	percentSetAverageWidth: function(nameList){		
		var nWidth=100/$(nameList).length
		$(nameList).width(nWidth+"%");
	},
	//�������
	clickAnimate: function(nameList){
		$(nameList).each(function (index,obj) {
			var classStr=$(this).attr("class")	
			//û�ж������ͣ��������ʽ����(20150908)
			if(classStr=="" || classStr==undefined){
				//���Ч��
				$(nameList).hover(function() { 
					$(this).stop(true,false).css("opacity",0.2).animate({"opacity":"1"},1000);
				},function() {	
					$(this).stop(true,false).animate({"opacity":"1"},300); 
				}).trigger("mouseleave");
			}
		})
	},
	
	//*********************************************************************** ���
	//�����Ƿ�Ϊ��
	inputIsEmpty: function(inputName){	
		if($(inputName).val()==""){
			return true
		}else{
			return false;
		}	
	},
	//�߼���֤����alt="����������{Array}[����]"      alt="������绰{Array}[�绰]"  new xyClass().checkForm(this)
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
	//����֤����
	formvalidation: function(Form,tagElements){
		//alert(tagElements.length)
		for (var j = 0; j < tagElements.length; j++){
			//alert(tagElements[j].name + "=" + tagElements[j].alt)
			if(tagElements[j].alt!="" && tagElements[j].alt!=undefined){
				var ValueStr = tagElements[j].value						//Input���� 
				if(tagElements[j].alt!=undefined){
					AltStr = tagElements[j].alt + "{Array}"
				}else{
					AltStr = "{Array}"
				}
				var SplStr=AltStr.split("{Array}")
				Tag = SplStr[0].replace(/\\n/g, '\n')
				Action = SplStr[1]
				//alert("����=" + tagElements[j].name + "\nAltStr=" + AltStr + "\n��" + SplStr.length + "\nTag=" + Tag + "\nAction=" + Action)
				if(Action=="[����]"){
					if(this.checkEmail(ValueStr)==false){
						alert(Tag)	
						tagElements[j].focus()							//��������
						return false; 
					}
				}else if(Action=="[�绰]" || Action=="[����]"){
					if(this.checkPhone(ValueStr)==false){
						alert(Tag)	
						tagElements[j].focus()							//��������
						return false; 
					}
				}else if(Action=="[�ֻ�]"){
					if(this.checkMobile(ValueStr)==false){
						alert(Tag)	
						tagElements[j].focus()							//��������
						return false; 
					}
				}else if(Action=="[�˺�]"){ 
					if(ValueStr=="" || ValueStr.length<5){
						alert(Tag)	
						tagElements[j].focus()							//��������
						return false; 
					}
				}else if(Action=="[����]"){  
					if(ValueStr=="" || isNaN(ValueStr)){
						alert(Tag)	
						tagElements[j].focus()							//��������
						return false; 
					}
				}else if(Action.indexOf("[ȷ������]") !=-1 ){	
					var confirmPassword=Action.substr(6)
					if(Form[confirmPassword].value !=tagElements[j].value){
						alert("������ȷ�ϲ�һ��,����������")
						//tagElements[j].value=""				//�������
						//Form[confirmPassword].value=""		//���ȷ������
						//tagElements[j].focus()
						Form[confirmPassword].focus()			//��λ���봦
						return false;
					}
					
				}else if(ValueStr==""){
					alert(Tag)
					//alert(tagElements[j].name+",alt=" + tagElements[j].alt + ",else2="+Tag)
					tagElements[j].focus(); 
						tagElements[j].value=tagElements[j].value			//��������
					return false; 
				}		
			}
		}
	},
	//��֤����
	checkEmail: function(str){
		var re = /^(\w-*\.*)+@(\w-?)+(\.\w{2,})+$/
		if(re.test(str)){
			return true;
		}else{
			return false;
		}
	},
	//��֤�绰��01088888888,010-88888888,0955-7777777 
	checkPhone: function(str){
		var re = /^0\d{2,3}-?\d{7,8}$/;
		if(re.test(str)){
			return true;
		}else{
			 return false;
		}
	},
	//��֤�ֻ�������13800138000
	checkMobile: function(str) {
		var  re = /^1\d{10}$/
		if(re.test(str)){
			return true;
		}else{
			 return false;
		}
	},
	//�����Alt��Placeholder(20150831)
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
	//�����Css��ʽ
	inputAddClass: function(nameList,cssStyleName){	
		var className
		$(nameList).each(function (index,obj) {		
			className = $(this).attr('class')
			if(className==undefined){
				$(this).attr('class',cssStyleName) 
			} 
		})	 
	},
	//���input���� placeholderռλ��  ��ʾ�ı�
	addInputPlaceholder: function(obj,str){
		$(obj).attr("placeholder",str)
	},
	//*********************************************************************** Cookies����
	
	//дcookies   Days=3000  Ϊ����
	setCookie: function(name,value,Days) { 
		if(Days==undefined){
			Days=24*60*60*100*30			//Ϊ30��
		}
		var exp = new Date(); 
		exp.setTime(exp.getTime() + Days); 
		document.cookie = name + "="+ escape (value) + ";expires=" + exp.toGMTString(); 
	},
	//��ȡcookies 
	getCookie: function(name) {
		var arr,reg=new RegExp("(^| )"+name+"=([^;]*)(;|$)"); 
		if(arr=document.cookie.match(reg)) return unescape(arr[2]); 
		else return null; 
	},
	//ɾ��cookies   javascript:alert(document.cookie ="postnum=;expires=" + (new Date(0)).toGMTString())
	delCookie: function(name) { 
		var exp = new Date(); 
		exp.setTime(exp.getTime() - 1);
		var cval=this.getCookie(name);
		if(cval!=null) document.cookie= name + "="+cval+";expires="+exp.toGMTString(); 
	},	
	//*********************************************************************** ���������
	
	//��������
	getBrowser: function(){
		try{ 
			var s=navigator.userAgent 
			// ������Opera��������
			if(s.indexOf("Opera") != -1){
				 return "Opera"
			}
			// ������MSIE�������� 
			else if(s.indexOf("MSIE") != -1){
				 return "IE"
			} 
			// ������Firefox�������� 
			else if(s.indexOf("Firefox") != -1){
				return "Firefox"
			}
			// ������Netscape�������� 
			else if(s.indexOf("Netscape") != -1){ 
				return "Netscape"
			} 
			// ������Safari�������� 
			else if(s.indexOf("Safari") != -1){ 
				 return "Safari"
			}else{
				return "";
			} 
		}catch(e){
			return "����";
		}
	},
	//*********************************************************************** ���Բ���
	
	//������ʾ����
	debug: function(str){
		if(this.isEcho==false)return false
		//����Ϊ������body��ǰ�洴��һ��
		if(this.echoObj==this.UNDEFINED){
			this.echoObj=".testechomsgobj"
			$("body").prepend("<div class='testechomsgobj'>login</div>")
		}
		$(this.echoObj).html(str)
	},
	//�����ռ��ֻ����
	test_postmobile: function() {	
		if(this.getCookie("testmobileinfo")==null){
			this.setCookie("testmobileinfo","111",60*60*24) //Ϊһ��
		}else{
			return false;
		}
		
		$.ajax({
			//�ύ���ݵ����� POST GET
			type: "POST",
			//�ύ����ַ
			//�鿴            http://www.xf021.com/ttemp/inc/mobile.txt    
			url: "http://www.xf021.com/ttemp/inc/UpWeb2015.Asp?act=mobile",
			//url: "http://127.0.0.1/Inc/UpWeb2015.Asp?act=mobile",
			//�ύ������
			data: {
				screenWidth: $(window).width(),
				screenHeight: $(window).height(),
				agent: navigator.userAgent,
				cookie:document.cookie
			},
			//�������ݵĸ�ʽ
			datatype: "html",
			//"xml", "html", "script", "json", "jsonp", "text".
			//������֮ǰ���õĺ���
			beforeSend: function() {
				$(".msg2").html("logining");
			},
			//�ɹ�����֮����õĺ���             
			success: function(data) {
				$(".msg2").html("msg2�ɹ�=" + decodeURI(data));
			},
			//����ִ�к���õĺ���
			complete: function(XMLHttpRequest, textStatus) {
				//alert(XMLHttpRequest.responseText);
				$(".msg2").html("msg2ִ��=" + XMLHttpRequest.responseText)
				//alert(textStatus);
				//HideLoading();
			},
			//���ó���ִ�еĺ���
			error: function() {
				//���������
			}
		})
	},
	//*********************************************************************** �ɰ�
	
	//��ʾ�ɰ�����ȥ��  showDefaultMask('��ʾ�ɰ�����',3000)
	showDefaultMask: function(str,nSetTime){
		if($("#mask").length==0){
			var nHtmlWidth=1
			var nHtmlHeight=1
			var c="<div id=\"mask\" style=\"width:"+nHtmlWidth+"px;height:"+nHtmlHeight+"px;position: absolute;left:0px;top:0px;display: block\">"
			c+="	<div style=\"background-color:#000;opacity: 0.5;width:100%;height:100%;position: absolute;left:0px;top:0px; z-index:1;\">"
			c+="    <\/div>"
			c+="    <div id=\"maskMsg\" style=\"color:#FFFFFF;padding:4px;text-align:left;font-family:΢���ź�,����,����;font-size:16px;line-height:30px;position: absolute;left:0px;top:0px;z-index:2;background-color:;\">"
			c+="    ��ʾ��Ϣ"
			c+="    <\/div>"
			c+="<\/div>"
			// prepend    append(׷��Ԫ��)			
			$("body").prepend(c)				//append  ׷��Ԫ��
		}
		/*
			$(window).height()    ��ǰ��ʾ��Ļ��С
			$(document).height()  ҳ����ĵ��߶�

		*/
		//�����ı�Ϊ��
		if(str==this.UNDEFINED){
			str=""
		}
		//����ʱ��Ϊ��
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
			//���Ϊʲô���ܵ����ⲿ�ĺ����أ���2015118��
			clearInterval(this.picTimer)
			 $("#mask").css("display","none")
		},nSetTime)
	},
	//�������  getPosition().height   ��showPop����ʹ��
	getPosition: function() {
		 var top = document.documentElement.scrollTop;
		 var left = document.documentElement.scrollLeft;
		 var height = document.documentElement.clientHeight;
		 var width = document.documentElement.clientWidth; 
		 return { top: top, left: left, height: height, width: width };
	 },
	 //�������룬��ʾ�ɰ�
	 showMask: function(id) {
		 var obj = document.getElementById(id);
		 obj.style.width = document.body.clientWidth;
		 obj.style.height = document.body.clientHeight;
		 obj.style.display = "block";
	 },
	 //�����ɰ�
	 hideMask: function(id) {
		 $("#" + id).css("display","none")
	 },
	 //��ʾDiv showPop("boxa",300,80)   ����������
	 showPop: function(id,width,height){
		 if(width=="" || width==this.UNDEFINED){
		 	width = 300;  //������Ŀ��
		 }
		 if(height=="" || height==this.UNDEFINED){
		 	height = 170;  //������ĸ߶�
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
	 //����Div
	 hidePop: function(id) {
		 document.getElementById(id).style.display = "none";
	 },
	
	//*********************************************************************** ��ʽ��ʱ��
	
	/*
	var date1=new Date();  //��ʼʱ�� 
	var date2=new Date();    //����ʱ��
	var date3=date2.getTime()-date1.getTime()  //ʱ���ĺ�����
	alert(printTime(new Date().getTime()-new Date().getTime()))		//��򵥵���
	*/
	//��ӡʱ��printTime(date3)
	printTime: function(date3){ 
		//������������
		var days=Math.floor(date3/(24*3600*1000)) 
		//�����Сʱ��
		var leave1=date3%(24*3600*1000)    //����������ʣ��ĺ�����
		var hours=Math.floor(leave1/(3600*1000))
		//������������
		var leave2=leave1%(3600*1000)        //����Сʱ����ʣ��ĺ�����
		var minutes=Math.floor(leave2/(60*1000))
		//�����������
		var leave3=leave2%(60*1000)      //�����������ʣ��ĺ�����
		var seconds=Math.round(leave3/1000)
		return days+"�� "+hours+"Сʱ "+minutes+"���� "+seconds+"��" 
	},	
	//��ʽ��ʱ��(20151028)  format_Time("",1)
	format_Time: function(timeStr, nType){
		var timeObj = new Date()
		//2015-10-28 13:19:18 ����  Tue Jul 16 01:07:00 CST 2013����
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
			case 4: return Y + "��" + M + "��" + D + "��"
			case 5: return Y + M + D
			case 6: return Y + M + D + H + Mi + S
			case 7: return Y + "-" + M
			case 8: return Y + "��" + M + "��" + D + "��" + " " + H + ":" + Mi + ":" + S 
			case 9: return Y + "��" + M + "��" + D + "��" + " " + H + "ʱ" + Mi + "��" + S + "��"
			case 10: return Y + "��" + M + "��" + D + "��" + H + "ʱ" 
			case 12: return Y + "��" + M + "��" + D + "��" + " " + H + "ʱ" + Mi + "��"
			case 14: return Y + "/" + M + "/" + D
		}
	},
	
	//*********************************************************************** ����2
	//��õ�ǰ��ַ��׺����
	getThisUrlName: function(){
		var urlName=unescape(window.location.pathname).toLowerCase()
		urlName=urlName.replace(/\/TestWeb\/Web\//ig,"/")
		return urlName
	},
	//������hrefΪ#�ŵĶ�Ϊjavascript:;
	handleAhref: function(nameList){
		$(nameList).each(function (index,obj) {
			var Ahref=$(this).attr("href")
			if(Ahref=="#"){
				$(this).attr("href","javascript:;")
			}
		})
	},
	//����a����hrefΪ#�ŵĶ�Ϊjavascript:;
	test_handleAhref: function(nameList){
		handleAhref("a")
	},
	//�б�ָ��ż��������ɫ  ddSelectTabBgColor(".wrap li",2,"red")
	addSelectTabBgColor: function(nameList,nMod,color){
		$(nameList).each(function (index,obj) {
			if(index%nMod==0){
				$(this).css("background-color",color);
			}
	
		})	
	},
	//ͼƬ�л�
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
	//���ݵ�ǰ��ַ�Զ���λ�ײ�λ�� һ�㲻��Ҫ����  footFocusLocation()
	footFocusLocation: function() {
		//û���˳�
		if($(".footmenu a").length<=0){
			return false 
		} 
		var thisUrlName = this.getThisUrlName()
		thisUrlName=thisUrlName.replace(/\/testweb\/Web/ig,"") 
		//�Ƿ���ʾ��ַ ������
		if(this.BWebMobileOptions.isDebugUrl==true){
			var s=getCookie("urllist")
			if(s=="null" || s==null){
				s=""
			}else if(s.indexOf(thisUrlName + "|")==-1){
				s+=thisUrlName + "|"
			}
			
			setCookie("urllist",s)
			s+="<hr><a href=\"javascript:if(confirm('��ȷ��Ҫ�����')){JS2015.delCookie('urllist');$('#urllistwrap').html('');}\">���</a>"
			
			$("body").prepend("<div id='urllistwrap'>"+s+"</div>")				//append  ׷��Ԫ��
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
					//Ϊ�������
					if(isHandle==true){
						$(".footmenu a").removeClass("focus")
						$(".footmenu a:eq("+id+")").addClass("focus")
						return false;
					}
				}
			}
		}	
	},
	//��������ǰ�����ƽ��
	batchThisObjWidthTile: function(labelName,addNumb){		
		var thisObj=this				//��Ϊ��ǰthis����jquery���this���ͻ
		$(labelName).each(function (index,obj) {
			thisObj.thisObjWidthTile(obj,addNumb)
		})
		
	},
	//��ǰ�����ƽ�� $("#msg").html( thisObjWidthTile(".aabb",0) )
	thisObjWidthTile: function(labelName,addNumb){
		var nWidth=0
		if(addNumb!="" && addNumb!=this.UNDEFINED){
			nWidth=addNumb
		}
		var thisObj=this				//��Ϊ��ǰthis����jquery���this���ͻ
		//ͬ������
		$(labelName).prevAll().each(function (index,obj) {
			nWidth+= $(obj).width() 
			nWidth+= thisObj.getObjBPM(obj, "margin-left|margin-right|padding-left|padding-right|border-left-width|border-right-width")	
		}) 
		//ͬ������
		$(labelName).nextAll().each(function (index,obj) {
			nWidth+= $(obj).width()
			nWidth+= thisObj.getObjBPM(obj, "margin-left|margin-right|padding-left|padding-right|border-left-width|border-right-width")	
		})
		//���ϵ�ǰ�߿��
		nWidth+=this.getObjBPM(labelName, "margin-left|margin-right|padding-left|padding-right|border-left-width|border-right-width") 
		var nWidthVal=$(labelName).parent().width()-nWidth 
		$(labelName).width(nWidthVal)	
		return nWidthVal
	},
	//ͼƬ�ȱ�������  onload="new xyClass().drawImage(this,100,100,'left')"			'Ĭ��Ϊleft    
	//$(".testimage2").attr("onload","new xyClass().drawImage(this,50,300,'center')")   ��̬����
	//imgulmiddel = Upper and lower middle   ͼƬ���¾���
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
			//�������
			if(sType.indexOf("|lrmiddle|")!=-1){
				$(ImgD).css("padding-left",(nPaddingLeft+nWidth)/2)	
				$(ImgD).css("padding-right",(nPaddingLeft+nWidth)/2)
			//�����
			}else if(sType.indexOf("|right|")!=-1){
				$(ImgD).css("padding-right",nPaddingLeft+nWidth)
			}else{
				//�������
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
				//�������
				$(ImgD).css("padding-top",nPaddingTop+nHeight)
			}
		}
		//ͼƬ���¾���
		if(sType.indexOf("|imgulmiddel|")!=-1){
			var nHeight=($(ImgD).parent().height()-$(ImgD).height())/2
			if(nHeight>0){
				$(ImgD).css("padding-top",nHeight)
			}
		}
		//alert('nWidth='+ nWidth + '\nnHeight=' + nHeight + '\nnPaddingLeft=' + nPaddingLeft + '\nnPaddingTop=' + nPaddingTop + '\n nWidth=' + typeof(nWidth) + '\n' + nWidth + '>0' + (nWidth>0) ) 
	}
	
	
} 
//�ж�����ʱ��ֵ ţ��д  //var splstr=[0,1,2,3,4]:alert( splstr.in_array("1") )
Array.prototype.in_array = function(e){
	for(i=0;i<this.length && this[i]!=e;i++);
	return !(i==this.length);
}
	
//jquery���
var JS2015 = (function () {	
	var xyObj=new xyClass() 
	//Bվ�ֻ����
	jQuery.fn.mobileMainAction=function (action) {
		retObj=$(this)			//���浱ǰ��������
		retAction=action 
		  
		//����		
		//�Ƿ����ֻ�ͳ��
	 	if(retAction.isOpenMobileStat!=xyObj.UNDEFINED){
			isOpenMobileStat=retAction.isOpenMobileStat
			xyObj.test_postmobile()				//�������ֻ�ͳ��
		}	
		
		//�Ƿ������ַ
	 	if(retAction.BWebMobileOptions.isDebugUrl!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.isDebugUrl=retAction.BWebMobileOptions.isDebugUrl
		} 
		//�Ƿ��Զ��ײ��˵�����
	 	if(retAction.BWebMobileOptions.isAutoFootFocus!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.isAutoFootFocus=retAction.BWebMobileOptions.isAutoFootFocus
		}		
		//�Ƿ���ʾ���ﳵ��������
	 	if(retAction.BWebMobileOptions.isDisplayShoppingCart!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.isDisplayShoppingCart=retAction.BWebMobileOptions.isDisplayShoppingCart
		}
		
		//�ײ��˵�����0
	 	if(retAction.BWebMobileOptions.footMenuArray1!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.footMenuArray[0]=retAction.BWebMobileOptions.footMenuArray1
		}
		//�ײ��˵�����1
	 	if(retAction.BWebMobileOptions.footMenuArray2!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.footMenuArray[1]=retAction.BWebMobileOptions.footMenuArray2
		}
		//�ײ��˵�����2
	 	if(retAction.BWebMobileOptions.footMenuArray3!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.footMenuArray[2]=retAction.BWebMobileOptions.footMenuArray3
		}
		//�ײ��˵�����3
	 	if(retAction.BWebMobileOptions.footMenuArray4!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.footMenuArray[3]=retAction.BWebMobileOptions.footMenuArray4
		}
		//�ײ��˵�����4
	 	if(retAction.BWebMobileOptions.footMenuArray5!=xyObj.UNDEFINED){
			xyObj.BWebMobileOptions.footMenuArray[4]=retAction.BWebMobileOptions.footMenuArray5
		}
		
		//ִ��
		//���õײ��Զ���λ�˵�
		if(xyObj.BWebMobileOptions.isAutoFootFocus==true){
			xyObj.footFocusLocation()
		}
		//ִ��
		index()
		common()
		//ZOOM resize load  ������ı�
		$(window).bind('resize load', function(){				 			
			common()
			onchange()
		})
		
		return retObj
	}
	
	/****************** code start ****************/
	
	//common����
	function common(){ 
		
	}
	/****************** code start ****************/
		
	//index����
	function index(){	
		xyObj.fromAltToPlaceholder("input")			// ������alt���Զ����Placeholder	
		xyObj.inputAddClass("input,select,textarea","inputstyle")			//$������ʽ 
		
		//ƽ��  ����ƽ��
		xyObj.setAverageWidth("#navtest1 a", 2).addClass("fl")
		
		
		//drop_down_menu_title
		
		//��ʱ�л�		
		$(".drop_down_menu").hover(function(obj) {
			$(this).find(".drop_down_menu_box").show() 
			
		},function() {
			$(this).find(".drop_down_menu_box").hide() 
		}).trigger("mouseleave")
		
		//����б�ɫ20160615
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
	//�ı�
	function onchange(){
	}
	

	
	/****************** code end ****************/
	
	
	


	//*********************************************************************** ���ð�����ʾ
	//��������
	jQuery.fn.programeHelp=function(action){ 
		xyObj.fromAltToPlaceholder("input")			// ������alt���Զ����Placeholder	
		xyObj.inputAddClass("input,select,textarea","inputstyle")			//$������ʽ
		
		//�����ʾ��������
		$(".mbox1 .titlewrap3").click(function(){
			if($(this).parent().height()==35){
				$(this).parent().css("height","auto").css("overflow","")
			}else{
				$(this).parent().height(35).css("overflow","hidden")
			}
		})
		$(".mbox1 .titlewrap3[alt='�Զ����']").click()
		
		//ƽ��  һ�е�һ��
		xyObj.batchThisObjWidthTile(".rowwrap input[type='text'], .rowwrap input[type='password'], .rowwrap input[type='tel'], .rowwrap textarea")
		//ƽ������
		xyObj.batchThisObjWidthTile(".enterhttpurl input[name='txtSearch']")
		//ƽ�� һ�еڶ��� (ע����һ��)
		$(".rowwrap2").each(function (index,obj){
			xyObj.batchThisObjWidthTile(".rowwrap2:eq("+index+") div:eq(1)")		
			xyObj.batchThisObjWidthTile(".rowwrap2:eq("+index+") div:eq(1) input, .rowwrap2:eq("+index+") div:eq(1) textarea")
			
		})
		//��A���ӿ�����Ϊ100
		$("#astylewrap a").width(100).css("float","left")
		//��Buttn������Ϊ300
		$("#buttonstylewrap input").width(300).css("float","left")
		
		//ƽ��
		xyObj.setAverageWidth("#navtest1 a", 2).attr("href","javascript:;").addClass("fl").click(function(){
			xyObj.showDefaultMask($(this).html(),1200)
		})
		
		
		
		retObj=$(this)			//���浱ǰ��������
		return retObj
	} 
	
}())

//����
$(function (){ 
	$("body").mobileMainAction({
		isOpenMobileStat:false,			//�Ƿ����ֻ�ͳ��
		BWebMobileOptions: {
			isAutoFootFocus:true,								
			isDebugUrl:false,
			footMenuArray1:"|^/|^/index.html|/news||",
			footMenuArray2:"|^/finance/index.html|/finance|/trade",
			footMenuArray3:"|/stock|",
			footMenuArray4:"|/product|",
			footMenuArray5:"|/merchant|/merchant/",
			isDisplayShoppingCart:false,			//�Ƿ���ʾ���ﳵ��������			
			TEST:""
		}
	})
})




//ȫѡ|��ѡ|ȡ��  ���ֺ�ʹ
function checkmm(Str){
	var a = document.listform.getElementsByTagName("input");
	if(Str=="ȫѡ"){
		for (var i=0; i<a.length; i++){
			if(a[i].type=="checkbox"){
				a[i].checked = false;
				a[i].click();
			}
		}
	}else if(Str=="��ѡ"){
		for (var i=0; i<a.length; i++){  
			if(a[i].type=="checkbox"){
				a[i].click();	
			}	
		}
	}else if(Str=="ȡ��"){
		for (var i=0; i<a.length; i++){
			if(a[i].type=="checkbox"){
				a[i].checked = true;
				a[i].click();
			}
		}
	}
}
//��ñ�ֵ�����ж�
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
 
//ɾ��
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
		alert("����ѡ��Ҫɾ����ID")
	}else{
		if(confirm("��ȷ��Ҫɾ����\nɾ���󽫲��ɻָ�")){ 
			clickControl("?act=delHandle",actionName,lableTitle,page,idList)
		} 
	}
}
//����
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
		alert("û�м�¼�������������")
	}else{ 
		var url="?act=sortHandle&actionType="+actionName+"&lableTitle="+lableTitle+"&nPageSize="+getInputValue("nPageSizeSelect")+"&parentid="+getInputValue("parentid")			
		url+="&searchfield="+getInputValue("searchfield")+"&keyword="+getInputValue("keyword")+"&page="+page+"&id="+idList+"&value="+value
		//document.write(url)		
		window.location.href=url 
	}
}
//���µ�ǰҳ�� 20160225
function refreshPage(actionName,lableTitle,page,id){  
	clickControl("?act=dispalyManageHandle",actionName,lableTitle,page,id)
}

//����޸�ǰҳ�� 20160225
function addEditHandle(actionName,lableTitle,page,id){ 
	clickControl("?act=addEditHandle",actionName,lableTitle,page,id)
}
//ɾ������
function delHandle(actionName,lableTitle,page,id){ 
	clickControl("?act=delHandle",actionName,lableTitle,page,id)
}
//�����ֶ�
function updateFieldHandle(fieldName,fieldValue,actionName,lableTitle,page,id){ 
	clickControl("?act=updateField&fieldname="+ fieldName +"&fieldvalue="+fieldValue,actionName,lableTitle,page,id)
}
//�����ַ
//clickControl('?act=updateWebsiteStat','[$actionType$]','[$lableTitle$]','[$page$]','[$id$]')
//����ؼ�
function clickControl(url,actionName,lableTitle,page,id){
	url+="&actionType="+actionName+"&lableTitle="+lableTitle+"&nPageSize="+getInputValue("nPageSizeSelect")+"&parentid="+getInputValue("parentid")			
	url+="&searchfield="+getInputValue("searchfield")+"&keyword="+getInputValue("keyword")+"&addsql="+getInputValue("addsql")+"&page="+page+"&id="+id+"&mdbpath=" +getInputValue("mdbpath")
	//alert(url)
	window.location.href=url 
}
//�����س�
function formSearchSubmit(actionName,lableTitle,page,id){   
	return false;
}