 
//loadJs("hotkeys.js");			    //�ȼ�
//loadJs("cookie.js");			    //cookies����ȼ�

$(function(){
	//jQuery(document).bind('keydown', 'Alt+s', function (){  $("form[name='formsubmit'] input[type='submit']").click()  })	
	if($(".tabbut li").length>0){
		$(".tabbut li").click(function(){
			$(".tabbut li").attr("class","atitleleftclick mgr6").css("cursor","pointer")
			$(this).attr("class","atitlefocus mgr6").css("cursor","default")
			$(".inputtable .tbodybox").css("display","none")
			var index = $(".tabbut li").index(this)		//�������
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



//����Js
function loadJs(name) {    
	document.write('<script src="'+name+'" type="text/javascript"></script>');
}
 
 
 
 
 
//��ɫ
function set_title_color(color) {
	$('#title').css('color',color);
	$('#titlecolor').val(color);
}
//���������ɫ
function clear_title() {
	$('#title').css('color','');
	$('#titlecolor').val('');
	$('#title_colorpanel').html(' ');
}
//�Ӵ�
function input_font_bold() {
	if($('#title').css('font-weight') == '700' || $('#title').css('font-weight')=='bold') {
		$('#title').css('font-weight','normal');
		//$("input[value='b']").attr("checked",false);							//���۷����������IE8
		var ctxB=false
	}else{
		$('#title').css('font-weight','bold')
		//$("input[value='b']").attr("checked",true);
		var ctxB=true
	}
	//��������Ϊ�˼���
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

