<?php  
$picture = new Imagick('18.gif');
 $i=0;
 foreach($picture as $frame){
     $f='frame-'.$i.'.gif';
     file_put_contents($f,$frame);
     $i++;
 } 
?>   
