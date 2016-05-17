<%
Class ClassYouKou

	'Div版网站导航
	Function DivNavigation(NavDid,AddSql)
		Dim I, C, S, Url, FileName, BigFolder, Did, ShowDid, Target 
		Rs.Open "Select * From [NavBigClass] " & AddSql, Conn, 1, 1 
		C = C & "<li class=left></li>" & vbCrlf
		For I = 1 To Rs.RecordCount
			FileName = Rs("FileName") 
			If FileName <> "" And Right(FileName,1)<>"/" Then FileName = FileName & ".html"  
		
			BigFolder = "/" & Rs("FolderName") & "/" 
			Did = Rs("BigClassName") 
			ShowDid = FontColorFontB(Did, Rs("FontB"), Rs("FontColor")) 
			Target = Rs("Target") 
			
			Url = Rs("FileName")
			If CheckRemoteUrl(Url)=False Then
				Url = BigFolder & FileName 
				Url = Replace(Replace(Url, "//", "/"), "//", "/")
				Url = FullHttpUrl(Host,Url)
				Call WebDebug(Url,"act=Nav&NavDid=" & Did)'调试
			End If
			
			
			If NavDid=Rs("BigClassName") Then
				'Call Echo(NavDid,Rs("BigClassName") & "," & ShowDid)
				C = C & "<li class=focus>"& ShowDid &"</li>" &vbCrlf
			Else
				C = C & "<li><a " & AHref(Url, "", Target) & ">" & ShowDid & "</a></li>" &vbCrlf
			End If
			C = C & "<li class=line></li>" & vbCrlf
		Rs.MoveNext : Next : Rs.Close 
		C = C & "<li class=right></li>" & vbCrlf
		DivNavigation = C 
	End Function

	'显示处理后的会员
	Function ShowHandleUserList(TopNumb)
		Dim SplStr,Splxx,SplSex,I,S,C
		SplStr=Split("王|李|张|刘|陈|杨|黄|吴|赵|周|徐|孙|马|朱|胡|林|郭|何|高|罗|郑|梁|谢|宋|唐|许|邓|冯|韩|曹|曾|彭|萧|蔡|潘|田|董|袁|于|余|叶|蒋|杜|苏|魏|程|吕|丁|沈|任|姚|卢|傅|钟|姜|崔|谭|廖|范|汪|陆|金|石|戴|贾|韦|夏|邱|方|侯|邹|熊|孟|秦|白|江|阎|薛|尹|段|雷|黎|史|龙|陶|贺|顾|毛|郝|龚|邵|万|钱|严|赖|覃|洪","|")
		Splxx=Split("武北京市|天津市|辽宁省沈阳市|吉林省长春市|黑龙江省哈尔滨市|上海市|江苏省南京市|湖北省武汉市|广东省广州市|重庆市|四川省成都市|陕西省西安市|河北省石家庄市|河北省唐山市|山西省太原市|内蒙包头市|辽宁省大连市|辽宁省鞍山市|辽宁省抚顺市|吉林省吉林市|黑龙江省齐齐哈尔市|江苏省徐州市|浙江省杭州市|福建省福州市|江西省南昌市|山东省济南市|山东省青岛市|山东省淄博市|河南省郑州市|湖南省长沙市|贵州省贵阳市|云南省昆明市|甘肃省兰州市|新疆乌鲁木齐市|河北省邯郸市|河北省保定市|河北省张家口市|山西省大同市|内蒙呼和浩特市|辽宁省本溪市|辽宁省丹东市|辽宁省锦州市|辽宁省阜新市|辽宁省辽阳市|黑龙江省鸡西市|黑龙江省鹤岗市|黑龙江省大庆市|黑龙江省伊春市|黑龙江省佳木斯市|黑龙江省牡丹|江市江苏省无锡市|江苏省常州市|江苏省苏州市|浙江省宁波市|安徽省合肥市|安徽省淮南市|安徽省淮北市|福建省厦门市|山东省枣庄市|山东省烟台市|山东省潍坊市|山东省泰安市|山东省临沂市|河南省开封市|河南省洛阳市|河南省平顶山市|河南省安阳市|河南省新乡市|河南省焦作市|湖北省黄石市|湖北省襄樊市|湖北省荆州市|湖南省株洲市|湖南省湘潭市|湖南省衡阳市|广东省深圳市|","|")
		Randomize
		For I=0 To TopNumb
			S = "来自 " &  Splxx(Int(Rnd*Ubound(Splxx))) & " "
			S = S & SplStr(Int(Rnd*Ubound(SplStr))) 
			S = S & IIF( Int(Rnd*2)=1, "先生", "女士" )
			C = C & "<li>" & S & "</li>"		
		Next
		ShowHandleUserList = C
	End Function
	
	'详细页显示Banner
	Function MainBanner(NavDid,NavSid,ProDid,ProSid,ProTid)
		Dim Sql, ImgPath
		If ProTid <> "" Then
			Sql="Select * From [ThreeClass] Where BigClassName='"& ProDid &"' And SmallClassName='"& ProSid &"' And ThreeClassName='" & ProTid & "'"
		ElseIf ProSid <> "" Then
			Sql="Select * From [SmallClass] Where BigClassName='"& ProDid &"' And SmallClassName='"& ProSid &"'"
		ElseIf ProDid <> "" Then
			Sql="Select * From [BigClass] Where BigClassName='"& ProDid &"'"
		ElseIf NavSid <> "" Then	
			Sql="Select * From [NavSmallClass] Where BigClassName='"& NavDid &"' And SmallClassName='"& NavSid &"'"
		ElseIf NavDid <> "" Then	
			Sql="Select * From [NavBigClass] Where BigClassName='"& NavDid &"'"
		End If
		If Sql <> "" Then
			Rs.Open Sql,Conn,1,1
			If Not Rs.Eof Then
				ImgPath = Rs("ImgPath")
			End If : Rs.Close			
		End If
		'图片为空时为默认
		If ImgPath = "" Then
			ImgPath = DefaultWebBanner
		End If
		'获得绝对图片地址
		If InStr(Lcase(ImgPath),"http://") = False And InStr(Lcase(ImgPath),"http://") = False Then
			ImgPath = "http://" & Request.ServerVariables("HTTP_HOST") & ImgPath
		End If
		MainBanner = ImgPath 
	End Function 


	'网站导航（简单类型 首页 | 关于我们 | 产品展示 | 联系方式）
	'顶部 WebNavigation("Where NavTop=True  Order By Sort Asc")
	'底部 WebNavigation("Where NavButtom=True  Order By Sort Asc")
	Function NewWebNavigation(NavDid,AddSql)
		Dim I, C, Url, FileName, BigFolder, Did, ShowDid, Target
		Rs.Open "Select * From [NavBigClass] " & AddSql, Conn, 1, 1 
		For I = 1 To Rs.RecordCount
			FileName = Rs("FileName") 
			If FileName <> "" And Right(FileName,1)<>"/" Then FileName = FileName & ".html"  
		
			BigFolder = "/" & Rs("FolderName") & "/" 
			Did = Rs("BigClassName") 
			ShowDid = FontColorFontB(Did, Rs("FontB"), Rs("FontColor")) 
			Target = Rs("Target")

			Url = Rs("FileName")
			If CheckRemoteUrl(Url)=False Then
				Url = BigFolder & FileName 
				Url = Replace(Replace(Url, "//", "/"), "//", "/")
				Call WebDebug(Url,"act=Nav&NavDid=" & Did)'调试
			End If
			
			If NavDid=Rs("BigClassName") Then
				C = C & "<span class='NavValue'>" & ShowDid & "</span>" 
			Else
				C = C & "<a " & AHref(Url, "", Target) & " class='nav'>" & ShowDid & "</a>" 
			End If 
			If I <> Rs.RecordCount Then C = C & "<span class=""NavLine""></span>" 
		Rs.MoveNext : Next : Rs.Close 
		NewWebNavigation = C 
	End Function
	
	'网站导航（简单类型 首页 | 关于我们 | 产品展示 | 联系方式）
	'顶部 WebNavigation("Where NavTop=True  Order By Sort Asc")
	'底部 WebNavigation("Where NavButtom=True  Order By Sort Asc")
	Function WebNavigation(AddSql)
		Dim I, C, Url, FileName, BigFolder, Did, ShowDid, Target 
		Rs.Open "Select * From [NavBigClass] " & AddSql, Conn, 1, 1 
		For I = 1 To Rs.RecordCount
			FileName = Rs("FileName") 
			If FileName <> "" And Right(FileName,1)<>"/" Then FileName = FileName & ".html"  
		
			BigFolder = "/" & Rs("FolderName") & "/" 
			Did = Rs("BigClassName") 
			ShowDid = FontColorFontB(Did, Rs("FontB"), Rs("FontColor")) 
			Target = Rs("Target") 
			
			Url = Rs("FileName")
			If CheckRemoteUrl(Url)=False Then
				Url = BigFolder & FileName 
				Url = Replace(Replace(Url, "//", "/"), "//", "/")
				Call WebDebug(Url,"act=Nav&NavDid=" & Did)'调试
			End If
			
			C = C & "<a " & AHref(Url, "", Target) & ">" & ShowDid & "</a>" 
			If I <> Rs.RecordCount Then C = C & " &nbsp;| &nbsp; " 
		Rs.MoveNext : Next : Rs.Close 
		WebNavigation = C 
	End Function
	
	
	'友情链接
	Function FriendLink()
		Dim Url,Title,Target,Rel,ImgPath,I,C
		Rsx.Open"Select * From [FriendLink] Where Through=True Order By Sort",Conn,1,1
		For I=1 To Rsx.RecordCount
			Url =  Rsx("Url")
			Target = ATarget(Rsx("Target"))
			Rel = ARel(Rsx("NoFollow"))
			Title = Rsx("Title")
			ImgPath = Rsx("ImgPath")
			C=C & "<div class='LinkBlock'><a href='"& Url &"'"& Target &""&Rel&"><img src="& ImgPath &"></a><br>"
			C=C & "<a href='"& Url &"'"& Target &""&Rel&">"& FontColorFontB(Title,Rsx("FontB"),Rsx("FontColor")) &"</a></div>" & vbCrlf		
		Rsx.MoveNext:Next:Rsx.Close
		FriendLink = C
	End Function
	
	'Banner切换 BannerTab("Banner",6,"")
	Function BannerTab(Did, TopNumb, PhotoWidth,PhotoHeight,AddSql)
		Dim SplStr, I, S, C
		If PhotoWidth="" Then PhotoWidth = 980
		If PhotoHeight="" Then PhotoHeight=230
		C=C & "<style type=""text/css""> " & vbCrlf
		C=C & ".container, .container *{margin:0; padding:0;}" & vbCrlf
		C=C & ".container{width:"& PhotoWidth &"px; height:"& PhotoHeight &"px; overflow:hidden;position:relative;}" & vbCrlf
		C=C & ".slider{position:absolute;}" & vbCrlf
		C=C & ".slider li{ list-style:none;display:inline;}" & vbCrlf
		C=C & ".slider img{ width:"& PhotoWidth &"px; height:"& PhotoHeight &"px; display:block;}" & vbCrlf
		C=C & ".slider2{width:8000px;}" & vbCrlf
		C=C & ".slider2 li{float:left;}" & vbCrlf
		C=C & ".num{ position:absolute; right:5px; bottom:5px;}" & vbCrlf
		C=C & ".num li{" & vbCrlf
		C=C & "	float: left;" & vbCrlf
		C=C & "	color: #FF7300;" & vbCrlf
		C=C & "	text-align: center;" & vbCrlf
		C=C & "	line-height: 16px;" & vbCrlf
		C=C & "	width: 16px;" & vbCrlf
		C=C & "	height: 16px;" & vbCrlf
		C=C & "	font-family: Arial;" & vbCrlf
		C=C & "	font-size: 12px;" & vbCrlf
		C=C & "	cursor: pointer;" & vbCrlf
		C=C & "	overflow: hidden;" & vbCrlf
		C=C & "	margin: 3px 1px;" & vbCrlf
		C=C & "	border: 1px solid #FF7300;" & vbCrlf
		C=C & "	background-color: #fff;" & vbCrlf
		C=C & "}" & vbCrlf
		C=C & ".num li.on{" & vbCrlf
		C=C & "	color: #fff;" & vbCrlf
		C=C & "	line-height: 21px;" & vbCrlf
		C=C & "	width: 21px;" & vbCrlf
		C=C & "	height: 21px;" & vbCrlf
		C=C & "	font-size: 16px;" & vbCrlf
		C=C & "	margin: 0 1px;" & vbCrlf
		C=C & "	border: 0;" & vbCrlf
		C=C & "	background-color: #FF7300;" & vbCrlf
		C=C & "	font-weight: bold;" & vbCrlf
		C=C & "}" & vbCrlf
		C=C & "</style>" & vbCrlf
		
		
	    Rs.Open "Select * From[Product] Where BigClassName='" & Did & "'", Conn, 1, 1 
		
		C=C & "<div class=""container"" id=""idTransformView2"">" & vbCrlf
		C=C & "  <ul class=""slider slider2"" id=""idSlider2"">" & vbCrlf
		For I=1 To TopNumb
			If Rs.Eof Then TopNumb=I-1 : Exit For
			C=C & "    <li><img src="""& Rs("BigFiles") &"""/></li>" & vbCrlf
 		Rs.MoveNext:Next : Rs.Close 
		C=C & "  </ul>" & vbCrlf
		C=C & "  <ul class=""num"" id=""idNum2"">" & vbCrlf
		 
		For I=1 To TopNumb
			C=C & "    <li>"& I &"</li>" & vbCrlf
		Next
		C=C & "  </ul>" & vbCrlf
		C=C & "</div>" & vbCrlf
		C=C & "<script type=""text/javascript"">" & vbCrlf
		C=C & "var $Mid = function (id) {" & vbCrlf
		C=C & "	return ""string"" == typeof id ? document.getElementById(id) : id;" & vbCrlf
		C=C & "};" & vbCrlf
		C=C & "var Class = {" & vbCrlf
		C=C & "  create: function() {" & vbCrlf
		C=C & "	return function() {" & vbCrlf
		C=C & "	  this.initialize.apply(this, arguments);" & vbCrlf
		C=C & "	}" & vbCrlf
		C=C & "  }" & vbCrlf
		C=C & "}" & vbCrlf
		C=C & "Object.extend = function(destination, source) {" & vbCrlf
		C=C & "	for (var property in source) {" & vbCrlf
		C=C & "		destination[property] = source[property];" & vbCrlf
		C=C & "	}" & vbCrlf
		C=C & "	return destination;" & vbCrlf
		C=C & "}" & vbCrlf
		C=C & "var TransformView = Class.create();" & vbCrlf
		C=C & "TransformView.prototype = {" & vbCrlf
		C=C & "  //容器对象,滑动对象,切换参数,切换数量" & vbCrlf
		C=C & "  initialize: function(container, slider, parameter, count, options) {" & vbCrlf
		C=C & "	if(parameter <= 0 || count <= 0) return;" & vbCrlf
		C=C & "	var oContainer = $Mid(container), oSlider = $Mid(slider), oThis = this;" & vbCrlf
		C=C & "	this.Index = 0;//当前索引" & vbCrlf
		C=C & "	" & vbCrlf
		C=C & "	this._timer = null;//定时器" & vbCrlf
		C=C & "	this._slider = oSlider;//滑动对象" & vbCrlf
		C=C & "	this._parameter = parameter;//切换参数" & vbCrlf
		C=C & "	this._count = count || 0;//切换数量" & vbCrlf
		C=C & "	this._target = 0;//目标参数" & vbCrlf
		C=C & "	" & vbCrlf
		C=C & "	this.SetOptions(options);" & vbCrlf
		C=C & "	" & vbCrlf
		C=C & "	this.Up = !!this.options.Up;" & vbCrlf
		C=C & "	this.Step = Math.abs(this.options.Step);" & vbCrlf
		C=C & "	this.Time = Math.abs(this.options.Time);" & vbCrlf
		C=C & "	this.Auto = !!this.options.Auto;" & vbCrlf
		C=C & "	this.Pause = Math.abs(this.options.Pause);" & vbCrlf
		C=C & "	this.onStart = this.options.onStart;" & vbCrlf
		C=C & "	this.onFinish = this.options.onFinish;" & vbCrlf
		C=C & "	oContainer.style.overflow = ""hidden"";" & vbCrlf
		C=C & "	oContainer.style.position = ""relative"";" & vbCrlf
		C=C & "	oSlider.style.position = ""absolute"";" & vbCrlf
		C=C & "	oSlider.style.top = oSlider.style.left = 0;" & vbCrlf
		C=C & "  }," & vbCrlf
		C=C & "  //设置默认属性" & vbCrlf
		C=C & "  SetOptions: function(options) {" & vbCrlf
		C=C & "	this.options = {//默认值" & vbCrlf
		C=C & "		Up:		true,//是否向上(否则向左)" & vbCrlf
		C=C & "		Step:		5,//滑动变化率" & vbCrlf
		C=C & "		Time:		10,//滑动延时" & vbCrlf
		C=C & "		Auto:		true,//是否自动转换" & vbCrlf
		C=C & "		Pause:		2000,//停顿时间(Auto为true时有效)" & vbCrlf
		C=C & "		onStart:	function(){},//开始转换时执行" & vbCrlf
		C=C & "		onFinish:	function(){}//完成转换时执行" & vbCrlf
		C=C & "	};" & vbCrlf
		C=C & "	Object.extend(this.options, options || {});" & vbCrlf
		C=C & "  }," & vbCrlf
		C=C & "  //开始切换设置" & vbCrlf
		C=C & "  Start: function() {" & vbCrlf
		C=C & "	if(this.Index < 0){" & vbCrlf
		C=C & "		this.Index = this._count - 1;" & vbCrlf
		C=C & "	} else if (this.Index >= this._count){ this.Index = 0; }" & vbCrlf
		C=C & "	" & vbCrlf
		C=C & "	this._target = -1 * this._parameter * this.Index;" & vbCrlf
		C=C & "	this.onStart();" & vbCrlf
		C=C & "	this.Move();" & vbCrlf
		C=C & "  }," & vbCrlf
		C=C & "  //移动" & vbCrlf
		C=C & "  Move: function() {" & vbCrlf
		C=C & "	clearTimeout(this._timer);" & vbCrlf
		C=C & "	var oThis = this, style = this.Up ? ""top"" : ""left"", iNow = parseInt(this._slider.style[style]) || 0, iStep = this.GetStep(this._target, iNow);" & vbCrlf
		C=C & "	" & vbCrlf
		C=C & "	if (iStep != 0) {" & vbCrlf
		C=C & "		this._slider.style[style] = (iNow + iStep) + ""px"";" & vbCrlf
		C=C & "		this._timer = setTimeout(function(){ oThis.Move(); }, this.Time);" & vbCrlf
		C=C & "	} else {" & vbCrlf
		C=C & "		this._slider.style[style] = this._target + ""px"";" & vbCrlf
		C=C & "		this.onFinish();" & vbCrlf
		C=C & "		if (this.Auto) { this._timer = setTimeout(function(){ oThis.Index++; oThis.Start(); }, this.Pause); }" & vbCrlf
		C=C & "	}" & vbCrlf
		C=C & "  }," & vbCrlf
		C=C & "  //获取步长" & vbCrlf
		C=C & "  GetStep: function(iTarget, iNow) {" & vbCrlf
		C=C & "	var iStep = (iTarget - iNow) / this.Step;" & vbCrlf
		C=C & "	if (iStep == 0) return 0;" & vbCrlf
		C=C & "	if (Math.abs(iStep) < 1) return (iStep > 0 ? 1 : -1);" & vbCrlf
		C=C & "	return iStep;" & vbCrlf
		C=C & "  }," & vbCrlf
		C=C & "  //停止" & vbCrlf
		C=C & "  Stop: function(iTarget, iNow) {" & vbCrlf
		C=C & "	clearTimeout(this._timer);" & vbCrlf
		C=C & "	this._slider.style[this.Up ? ""top"" : ""left""] = this._target + ""px"";" & vbCrlf
		C=C & "  }" & vbCrlf
		C=C & "};" & vbCrlf
		C=C & "window.onload=function(){" & vbCrlf
		C=C & "	function Each(list, fun){" & vbCrlf
		C=C & "		for (var i = 0, len = list.length; i < len; i++) { fun(list[i], i); }" & vbCrlf
		C=C & "	};" & vbCrlf
		C=C & "	var objs = $Mid(""idNum2"").getElementsByTagName(""li"");" & vbCrlf
		C=C & "	var tv = new TransformView(""idTransformView2"", ""idSlider2"", 980, "& TopNumb &", {" & vbCrlf
		C=C & "		onStart: function(){ Each(objs, function(o, i){ o.className = tv.Index == i ? ""on"" : """"; }) },//按钮样式" & vbCrlf
		C=C & "		Up: false" & vbCrlf
		C=C & "	});" & vbCrlf
		C=C & "	tv.Start();" & vbCrlf
		C=C & "	Each(objs, function(o, i){" & vbCrlf
		C=C & "		o.onmouseover = function(){" & vbCrlf
		C=C & "			o.className = ""on"";" & vbCrlf
		C=C & "			tv.Auto = false;" & vbCrlf
		C=C & "			tv.Index = i;" & vbCrlf
		C=C & "			tv.Start();" & vbCrlf
		C=C & "		}" & vbCrlf
		C=C & "		o.onmouseout = function(){" & vbCrlf
		C=C & "			o.className = """";" & vbCrlf
		C=C & "			tv.Auto = true;" & vbCrlf
		C=C & "			tv.Start();" & vbCrlf
		C=C & "		}" & vbCrlf
		C=C & "	})" & vbCrlf
		C=C & "}" & vbCrlf
		C=C & "</script>" & vbCrlf
	
		BannerTab=C
	End Function 

	
	

End Class
%>