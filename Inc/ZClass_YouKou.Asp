<%
Class ClassYouKou

	'Div����վ����
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
				Call WebDebug(Url,"act=Nav&NavDid=" & Did)'����
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

	'��ʾ������Ļ�Ա
	Function ShowHandleUserList(TopNumb)
		Dim SplStr,Splxx,SplSex,I,S,C
		SplStr=Split("��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|֣|��|л|��|��|��|��|��|��|��|��|��|��|��|��|��|��|Ԭ|��|��|Ҷ|��|��|��|κ|��|��|��|��|��|Ҧ|¬|��|��|��|��|̷|��|��|��|½|��|ʯ|��|��|Τ|��|��|��|��|��|��|��|��|��|��|��|Ѧ|��|��|��|��|ʷ|��|��|��|��|ë|��|��|��|��|Ǯ|��|��|��|��","|")
		Splxx=Split("�䱱����|�����|����ʡ������|����ʡ������|������ʡ��������|�Ϻ���|����ʡ�Ͼ���|����ʡ�人��|�㶫ʡ������|������|�Ĵ�ʡ�ɶ���|����ʡ������|�ӱ�ʡʯ��ׯ��|�ӱ�ʡ��ɽ��|ɽ��ʡ̫ԭ��|���ɰ�ͷ��|����ʡ������|����ʡ��ɽ��|����ʡ��˳��|����ʡ������|������ʡ���������|����ʡ������|�㽭ʡ������|����ʡ������|����ʡ�ϲ���|ɽ��ʡ������|ɽ��ʡ�ൺ��|ɽ��ʡ�Ͳ���|����ʡ֣����|����ʡ��ɳ��|����ʡ������|����ʡ������|����ʡ������|�½���³ľ����|�ӱ�ʡ������|�ӱ�ʡ������|�ӱ�ʡ�żҿ���|ɽ��ʡ��ͬ��|���ɺ��ͺ�����|����ʡ��Ϫ��|����ʡ������|����ʡ������|����ʡ������|����ʡ������|������ʡ������|������ʡ�׸���|������ʡ������|������ʡ������|������ʡ��ľ˹��|������ʡĵ��|���н���ʡ������|����ʡ������|����ʡ������|�㽭ʡ������|����ʡ�Ϸ���|����ʡ������|����ʡ������|����ʡ������|ɽ��ʡ��ׯ��|ɽ��ʡ��̨��|ɽ��ʡΫ����|ɽ��ʡ̩����|ɽ��ʡ������|����ʡ������|����ʡ������|����ʡƽ��ɽ��|����ʡ������|����ʡ������|����ʡ������|����ʡ��ʯ��|����ʡ�差��|����ʡ������|����ʡ������|����ʡ��̶��|����ʡ������|�㶫ʡ������|","|")
		Randomize
		For I=0 To TopNumb
			S = "���� " &  Splxx(Int(Rnd*Ubound(Splxx))) & " "
			S = S & SplStr(Int(Rnd*Ubound(SplStr))) 
			S = S & IIF( Int(Rnd*2)=1, "����", "Ůʿ" )
			C = C & "<li>" & S & "</li>"		
		Next
		ShowHandleUserList = C
	End Function
	
	'��ϸҳ��ʾBanner
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
		'ͼƬΪ��ʱΪĬ��
		If ImgPath = "" Then
			ImgPath = DefaultWebBanner
		End If
		'��þ���ͼƬ��ַ
		If InStr(Lcase(ImgPath),"http://") = False And InStr(Lcase(ImgPath),"http://") = False Then
			ImgPath = "http://" & Request.ServerVariables("HTTP_HOST") & ImgPath
		End If
		MainBanner = ImgPath 
	End Function 


	'��վ������������ ��ҳ | �������� | ��Ʒչʾ | ��ϵ��ʽ��
	'���� WebNavigation("Where NavTop=True  Order By Sort Asc")
	'�ײ� WebNavigation("Where NavButtom=True  Order By Sort Asc")
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
				Call WebDebug(Url,"act=Nav&NavDid=" & Did)'����
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
	
	'��վ������������ ��ҳ | �������� | ��Ʒչʾ | ��ϵ��ʽ��
	'���� WebNavigation("Where NavTop=True  Order By Sort Asc")
	'�ײ� WebNavigation("Where NavButtom=True  Order By Sort Asc")
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
				Call WebDebug(Url,"act=Nav&NavDid=" & Did)'����
			End If
			
			C = C & "<a " & AHref(Url, "", Target) & ">" & ShowDid & "</a>" 
			If I <> Rs.RecordCount Then C = C & " &nbsp;| &nbsp; " 
		Rs.MoveNext : Next : Rs.Close 
		WebNavigation = C 
	End Function
	
	
	'��������
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
	
	'Banner�л� BannerTab("Banner",6,"")
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
		C=C & "  //��������,��������,�л�����,�л�����" & vbCrlf
		C=C & "  initialize: function(container, slider, parameter, count, options) {" & vbCrlf
		C=C & "	if(parameter <= 0 || count <= 0) return;" & vbCrlf
		C=C & "	var oContainer = $Mid(container), oSlider = $Mid(slider), oThis = this;" & vbCrlf
		C=C & "	this.Index = 0;//��ǰ����" & vbCrlf
		C=C & "	" & vbCrlf
		C=C & "	this._timer = null;//��ʱ��" & vbCrlf
		C=C & "	this._slider = oSlider;//��������" & vbCrlf
		C=C & "	this._parameter = parameter;//�л�����" & vbCrlf
		C=C & "	this._count = count || 0;//�л�����" & vbCrlf
		C=C & "	this._target = 0;//Ŀ�����" & vbCrlf
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
		C=C & "  //����Ĭ������" & vbCrlf
		C=C & "  SetOptions: function(options) {" & vbCrlf
		C=C & "	this.options = {//Ĭ��ֵ" & vbCrlf
		C=C & "		Up:		true,//�Ƿ�����(��������)" & vbCrlf
		C=C & "		Step:		5,//�����仯��" & vbCrlf
		C=C & "		Time:		10,//������ʱ" & vbCrlf
		C=C & "		Auto:		true,//�Ƿ��Զ�ת��" & vbCrlf
		C=C & "		Pause:		2000,//ͣ��ʱ��(AutoΪtrueʱ��Ч)" & vbCrlf
		C=C & "		onStart:	function(){},//��ʼת��ʱִ��" & vbCrlf
		C=C & "		onFinish:	function(){}//���ת��ʱִ��" & vbCrlf
		C=C & "	};" & vbCrlf
		C=C & "	Object.extend(this.options, options || {});" & vbCrlf
		C=C & "  }," & vbCrlf
		C=C & "  //��ʼ�л�����" & vbCrlf
		C=C & "  Start: function() {" & vbCrlf
		C=C & "	if(this.Index < 0){" & vbCrlf
		C=C & "		this.Index = this._count - 1;" & vbCrlf
		C=C & "	} else if (this.Index >= this._count){ this.Index = 0; }" & vbCrlf
		C=C & "	" & vbCrlf
		C=C & "	this._target = -1 * this._parameter * this.Index;" & vbCrlf
		C=C & "	this.onStart();" & vbCrlf
		C=C & "	this.Move();" & vbCrlf
		C=C & "  }," & vbCrlf
		C=C & "  //�ƶ�" & vbCrlf
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
		C=C & "  //��ȡ����" & vbCrlf
		C=C & "  GetStep: function(iTarget, iNow) {" & vbCrlf
		C=C & "	var iStep = (iTarget - iNow) / this.Step;" & vbCrlf
		C=C & "	if (iStep == 0) return 0;" & vbCrlf
		C=C & "	if (Math.abs(iStep) < 1) return (iStep > 0 ? 1 : -1);" & vbCrlf
		C=C & "	return iStep;" & vbCrlf
		C=C & "  }," & vbCrlf
		C=C & "  //ֹͣ" & vbCrlf
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
		C=C & "		onStart: function(){ Each(objs, function(o, i){ o.className = tv.Index == i ? ""on"" : """"; }) },//��ť��ʽ" & vbCrlf
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