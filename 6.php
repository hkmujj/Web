<?PHP
 
/*
  ASP Translator Messages (These can be removed later)
  ----------------------------------------------------

  1. All "Dim" statements have been changed to comments for reference.
  2. Script uses Left(), Right(), or Mid(). See the Documentation for equivalent PHP functions

*/

Class $whileclass;
function myfun($nNumb) {
        if ($nNumb == 1) {
            echo "hello world<hr>"; 
        } else {
            echo "no numb<hr>";
        } 
    } 
function nfor($n) {
        //dim $i; 
        for ($i=1; $i<=$n; $i++) {
            echo $i."for、<hr>"; 
        } 
    } 
function nwhile($n) {
        While $n > 1;
            $n = $n - 1; 
            echo $n."while、<hr>"; 
        } 
    } 
function ndoloop($n) {
        Do While $n > 1;
            $n = $n - 1; 
            echo $n."doloop、<hr>"; 
        Loop; 
    } 
function nforeach() {
        //dim $splStr, $s; 
        $splStr = Array("aa", "bb", "cc"); 
        For Each $s In $splStr;
            echo "s=".$s."<hr>"; 
        } 
    } 
End Class; 
 
//判断类
Class $ifclass;
function testif($n) {
        if ($n > 10) {
            echo "n大于10<br>"; 
        } elseif $n > 5 Then;
            echo "n大于5<br>"; 
        } else {
            echo "n为默认<br>".$n; 
        } 
    } 
function testif2($a) {
        echo "testif2<hr>"; 
    } 
 
 
End Class; 
 
//字典类
Class $zdclass;
function testzd() {
        //dim $aspD, $title, $items, $i; 
        //dim $aA, $bB : Set $aspD = Server.CreateObject("Scripting.Dictionary");
            $aspD.$add "Abs", "返回数字的绝对值11111111"; 
            $aspD.$add "Sqr", "返回数值表达式的平方根aaaaaaaaaaaaaaaaaaaaaaaa"; 
            $aspD.$add "Sgn", "返回表示数字符号的整数22222222"; 
            $aspD.$add "Rnd", "返回一个随机生成的数字33333333333333"; 
            $aspD.$add "Log", "返回指定数值的自然对数ssssssssssssssss"; 
 
 
            echo "Abs=".$aspD["Abs"]."<hr>"; 
            echo "Rnd=".$aspD["Rnd"]."<hr>"; 
    }
End Class; 
 
//测试循环
function $testwhile() {
    //dim $obj : Set $obj = $new $whileclass;
        Call $obj.myfun(1); 
        echo "<br>33333333<br>"; 
        Call $obj.myfun(2); 
        Call $obj.nfor(6); 
        Call $obj.nwhile(6); 
        Call $obj.ndoloop(6); 
        Call $obj.nforeach(); 
 
}
//测试判断
function testif() {
    //dim $obj : Set $obj = $new $ifclass;
        Call $obj.testif(11); 
        Call $obj.testif(6); 
        Call $obj.testif(3); 
        $obj.testif2 3 : $obj.testif2 3; 
}
//测试字典
function testzd() {
    //dim $obj : Set $obj = $new $zdclass;
        Call $obj.testzd(); 
 
}
 
 
 
//截取字符串 更新20160114
//c=作者小云[A]sharembweb.com[/A]QQ313801120
//0=sharembweb.com
//1=[A]sharembweb.com[/A]
//3=[A]sharembweb.com
//4=sharembweb.com[/A]
function $strCutTest[$ByVal $content, $ByVal $startStr, $ByVal $endStr, $ByVal $cutType] {
    //On Error Resume Next
    //dim $s1, $s1Str, $s2, $s3, $c; 
    if (instr($content, $startStr) == False || instr($content, $endStr) == False) {
        $c = ""; 
        break; 
    } 
    Select Case $cutType;
        //完善于20150923
        Case 1;
            $s1 = instr($content, $startStr); 
            $s1Str = mid($content, $s1 + strlen($startStr)); 
            $s2 = $s1 + instr($s1Str, $endStr) + strlen($startStr) + strlen($endStr) - 1; //为什么要减1
 
        Case } else {
            $s1 = instr($content, $startStr) + strlen($startStr); 
            $s1Str = mid($content, $s1); 
            //S2 = instr(S1, Content, EndStr)
            $s2 = $s1 + instr($s1Str, $endStr) - 1; 
        //call echo("s2",s2)
    End Select;
    $s3 = $s2 - $s1; 
    if ($s3 >= 0) {
        $c = mid($content, $s1, $s3); 
    } else {
        $c = ""; 
    } 
    if ($cutType == 3) {
        $c = $startStr.$c; 
    } 
    if ($cutType == 4) {
        $c = $c.$endStr; 
    } 
    $return  $c; 
    //If Err.Number <> 0 Then Call eerr(startStr, content)
//doError Err.Description, "strCutTest 截取字符串 函数出错，StartStr=" & EchoHTML(StartStr) & "<hr>EndStr=" & EchoHTML(EndStr)
}
 
//测试实例
function testcase() {
 
    //dim $c; 
    $c = "作者小云[A]sharembweb.com[/A]QQ313801120"; 
    
    echo "c=".$c."<br>"; 
    
    echo "0=".$strCutTest[$c, "[A]", "[/A]", 0]."<br>"."\n"; 
    echo "1=".$strCutTest[$c, "[A]", "[/A]", 1]."<br>"."\n"; 
    //response.Write("2=" & strCutTest(c,"[A]","[/A]",2) & "<br>" & "\n")
    echo "3=".$strCutTest[$c, "[A]", "[/A]", 3]."<br>"."\n"; 
    echo "4=".$strCutTest[$c, "[A]", "[/A]", 4]."<br>"."\n"; 
 
}
 
 
//选择
Select Case Request("act");
    Case "testwhile" : $testwhile();                                        //测试循环
    Case "testif" : testif();                                              //测试判断
    Case "testzd" : testzd();                                              //测试字典
    Case "testcase" : testcase();                                              //测试实例
    
 
    Case } else { : $displayDefault();                                          //显示默认
End Select;
 
 
 
 
//显示默认
function $displayDefault() {
    echo "<a href='?act=testwhile'>测试循环</a> <br>"; 
    echo "<a href='?act=testif'>测试判断</a> <br>"; 
    echo "<a href='?act=testzd'>测试字典</a> <br>"; 
    echo "<a href='?act=testcase'>测试实例</a> <br>"; 
}  
