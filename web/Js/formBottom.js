 
//loadJs("hotkeys.js");			    //热键
//loadJs("cookie.js");			    //cookies配合热键

$(function(){
	//jQuery(document).bind('keydown', 'Alt+s', function (){  $("form[name='formsubmit'] input[type='submit']").click()  })	
	if($(".tabbut li").length>0){
		$(".tabbut li").click(function(){
			$(".tabbut li").attr("class","atitleleftclick mgr6").css("cursor","pointer")
			$(this).attr("class","atitlefocus mgr6").css("cursor","default")
			$(".inputtable .tbodybox").css("display","none")
			var index = $(".tabbut li").index(this)		//获得索引
			$(".inputtable .tbodybox:eq("+ index +")").css("display","")		
			$("#switchId").val(index)
		})
		var swtichId=$("#switchId").val()
		if(swtichId==''){
			swtichId='0'
		}
		$(".tabbut li:eq("+ swtichId +")").click()
	}
	$("select").addClass("inputstyle") 
})



//加载Js
function loadJs(name) {    
	document.write('<script src="'+name+'" type="text/javascript"></script>');
}
 
 
 
 
 
//颜色
function set_title_color(color) {
	$('#title').css('color',color);
	$('#titlecolor').val(color);
}
//清除标题颜色
function clear_title() {
	$('#title').css('color','');
	$('#titlecolor').val('');
	$('#title_colorpanel').html(' ');
}
//加粗
function input_font_bold() {
	if($('#title').css('font-weight') == '700' || $('#title').css('font-weight')=='bold') {
		$('#title').css('font-weight','normal');
		//$("input[value='b']").attr("checked",false);							//这咱方法会出错，在IE8
		var ctxB=false
	}else{
		$('#title').css('font-weight','bold')
		//$("input[value='b']").attr("checked",true);
		var ctxB=true
	}
	//这样做是为了兼容
	var a = document.getElementsByTagName("input"); 
	for (var i=0; i<a.length; i++){
		if(a[i].type=="checkbox"){
			if(a[i].value=="b"){
				a[i].checked = ctxB;
				break;
			} 
		}
	} 

}

