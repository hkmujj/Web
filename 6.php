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
            echo $i."for��<hr>"; 
        } 
    } 
function nwhile($n) {
        While $n > 1;
            $n = $n - 1; 
            echo $n."while��<hr>"; 
        } 
    } 
function ndoloop($n) {
        Do While $n > 1;
            $n = $n - 1; 
            echo $n."doloop��<hr>"; 
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
 
//�ж���
Class $ifclass;
function testif($n) {
        if ($n > 10) {
            echo "n����10<br>"; 
        } elseif $n > 5 Then;
            echo "n����5<br>"; 
        } else {
            echo "nΪĬ��<br>".$n; 
        } 
    } 
function testif2($a) {
        echo "testif2<hr>"; 
    } 
 
 
End Class; 
 
//�ֵ���
Class $zdclass;
function testzd() {
        //dim $aspD, $title, $items, $i; 
        //dim $aA, $bB : Set $aspD = Server.CreateObject("Scripting.Dictionary");
            $aspD.$add "Abs", "�������ֵľ���ֵ11111111"; 
            $aspD.$add "Sqr", "������ֵ���ʽ��ƽ����aaaaaaaaaaaaaaaaaaaaaaaa"; 
            $aspD.$add "Sgn", "���ر�ʾ���ַ��ŵ�����22222222"; 
            $aspD.$add "Rnd", "����һ��������ɵ�����33333333333333"; 
            $aspD.$add "Log", "����ָ����ֵ����Ȼ����ssssssssssssssss"; 
 
 
            echo "Abs=".$aspD["Abs"]."<hr>"; 
            echo "Rnd=".$aspD["Rnd"]."<hr>"; 
    }
End Class; 
 
//����ѭ��
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
//�����ж�
function testif() {
    //dim $obj : Set $obj = $new $ifclass;
        Call $obj.testif(11); 
        Call $obj.testif(6); 
        Call $obj.testif(3); 
        $obj.testif2 3 : $obj.testif2 3; 
}
//�����ֵ�
function testzd() {
    //dim $obj : Set $obj = $new $zdclass;
        Call $obj.testzd(); 
 
}
 
 
 
//��ȡ�ַ��� ����20160114
//c=����С��[A]sharembweb.com[/A]QQ313801120
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
        //������20150923
        Case 1;
            $s1 = instr($content, $startStr); 
            $s1Str = mid($content, $s1 + strlen($startStr)); 
            $s2 = $s1 + instr($s1Str, $endStr) + strlen($startStr) + strlen($endStr) - 1; //ΪʲôҪ��1
 
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
//doError Err.Description, "strCutTest ��ȡ�ַ��� ��������StartStr=" & EchoHTML(StartStr) & "<hr>EndStr=" & EchoHTML(EndStr)
}
 
//����ʵ��
function testcase() {
 
    //dim $c; 
    $c = "����С��[A]sharembweb.com[/A]QQ313801120"; 
    
    echo "c=".$c."<br>"; 
    
    echo "0=".$strCutTest[$c, "[A]", "[/A]", 0]."<br>"."\n"; 
    echo "1=".$strCutTest[$c, "[A]", "[/A]", 1]."<br>"."\n"; 
    //response.Write("2=" & strCutTest(c,"[A]","[/A]",2) & "<br>" & "\n")
    echo "3=".$strCutTest[$c, "[A]", "[/A]", 3]."<br>"."\n"; 
    echo "4=".$strCutTest[$c, "[A]", "[/A]", 4]."<br>"."\n"; 
 
}
 
 
//ѡ��
Select Case Request("act");
    Case "testwhile" : $testwhile();                                        //����ѭ��
    Case "testif" : testif();                                              //�����ж�
    Case "testzd" : testzd();                                              //�����ֵ�
    Case "testcase" : testcase();                                              //����ʵ��
    
 
    Case } else { : $displayDefault();                                          //��ʾĬ��
End Select;
 
 
 
 
//��ʾĬ��
function $displayDefault() {
    echo "<a href='?act=testwhile'>����ѭ��</a> <br>"; 
    echo "<a href='?act=testif'>�����ж�</a> <br>"; 
    echo "<a href='?act=testzd'>�����ֵ�</a> <br>"; 
    echo "<a href='?act=testcase'>����ʵ��</a> <br>"; 
}  
