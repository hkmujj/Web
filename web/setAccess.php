<?PHP

//新的截取字符20160216
function newGetStrCut($content, $title){
    $s ='';
    //这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409
    if( instr($content,vbCrlf())==false ){
        $content=replace($content,chr(10),vbCrlf());
    }
    if( instr($content, '【/' . $title . '】') > 0 ){
        $s= ADSql(phptrim(getStrCut($content, '【' . $title . '】', '【/' . $title . '】', 0)));
    }else{
        $s= ADSql(phptrim(getStrCut($content, '【' . $title . '】', vbCrlf(), 0)));
    }
    $newGetStrCut= $s;
    return @$newGetStrCut;
}


//重置数据库数据
function resetAccessData(){

    handlePower('恢复模板数据');						//管理权限处理

    $GLOBALS['conn=']=OpenConn();
    $splStr=''; $i=''; $s=''; $columnname=''; $title=''; $nCount=''; $webdataDir ='';
    $webdataDir= @$_REQUEST['webdataDir'];
    if( $webdataDir <> '' ){
        if( checkFolder($webdataDir)== false ){
            eerr('网站数据目录不存在，恢复默认数据未成功', $webdataDir);
        }
    }else{
        $webdataDir= '/Data/WebData/';
    }

    ASPEcho('提示', '恢复数据完成');
    rw('<hr><a href=\'../' . EDITORTYPE . 'web.' . EDITORTYPE . '\' target=\'_blank\'>进入首页</a> | <a href="?" target=\'_blank\'>进入后台</a>');

    $content=''; $filePath=''; $parentid=''; $author=''; $adddatetime=''; $fileName=''; $bodycontent=''; $webtitle=''; $webkeywords=''; $webdescription=''; $sortrank=''; $labletitle=''; $target ='';
    $websitebottom=''; $webtemplate=''; $webimages=''; $webcss=''; $webjs=''; $flags=''; $websiteurl=''; $splxx=''; $columntype=''; $relatedtags=''; $npagesize=''; $customaurl=''; $nofollow ='';
    $templatepath=''; $isthrough='';$titlecolor='';
    $showreason=''; $ncomputersearch=''; $nmobliesearch=''; $ncountsearch=''; $ndegree ='';//竞价表
    $displaytitle=''; $aboutcontent=''; $isonhtml ='';//单页表
    $columnenname ='';//导航表
    $smallimage=''; $bigImage=''; $bannerimage ='';//文章表
    $httpurl='';

    //网站配置
    $content= getftext($webdataDir . '/website.txt');
    //这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409
    if( instr($content,vbCrlf())==false ){
        $content=replace($content,chr(10),vbCrlf());
    }
    if( $content <> '' ){
        $webtitle= newGetStrCut($content, 'webtitle');
        $webkeywords= newGetStrCut($content, 'webkeywords');
        $webdescription= newGetStrCut($content, 'webdescription');
        $websitebottom= newGetStrCut($content, 'websitebottom');
        $webtemplate= newGetStrCut($content, 'webtemplate');
        $webimages= newGetStrCut($content, 'webimages');
        $webcss= newGetStrCut($content, 'webcss');
        $webjs= newGetStrCut($content, 'webjs');
        $flags= newGetStrCut($content, 'flags');
        $websiteurl= newGetStrCut($content, 'websiteurl');

        if( getRecordCount($GLOBALS['db_PREFIX'] . 'website', '')== 0 ){
            connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'website(webtitle) values(\'测试\')');
        }

        connExecute('update ' . $GLOBALS['db_PREFIX'] . 'website  set webtitle=\'' . $webtitle . '\',webkeywords=\'' . $webkeywords . '\',webdescription=\'' . $webdescription . '\',websitebottom=\'' . $websitebottom . '\',webtemplate=\'' . $webtemplate . '\',webimages=\'' . $webimages . '\',webcss=\'' . $webcss . '\',webjs=\'' . $webjs . '\',flags=\'' . $flags . '\',websiteurl=\'' . $websiteurl . '\'');
    }

    //导航
    connExecute('delete from ' . $GLOBALS['db_PREFIX'] . 'webcolumn');
    $content= getDirTxtList($webdataDir . '/webcolumn/');
    $splStr= aspSplit($content, vbCrlf());
    hr();
    foreach( $splStr as $filePath){
        $fileName= getfilename($filePath);
        if( $filePath <> '' && instr('_#', substr($fileName, 0 , 1))== false ){
            ASPEcho('导航', $filePath);
            $content= getftext($filePath);
            //这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409
            if( instr($content,vbCrlf())==false ){
                $content=replace($content,chr(10),vbCrlf());
            }
            $splxx= aspSplit($content, vbCrlf() . '-------------------------------');
            foreach( $splxx as $s){
                if( instr($s, '【webtitle】') > 0 ){
                    $s=$s . vbCrlf();
                    $webtitle= newGetStrCut($s, 'webtitle');
                    $webkeywords= newGetStrCut($s, 'webkeywords');
                    $webdescription= newGetStrCut($s, 'webdescription');

                    $sortrank= newGetStrCut($s, 'sortrank');
                    if( $sortrank== '' ){ $sortrank= 0 ;}
                    $fileName= newGetStrCut($s, 'filename');
                    $columnname= newGetStrCut($s, 'columnname');
                    $columnenname= newGetStrCut($s, 'columnenname');
                    $columntype= newGetStrCut($s, 'columntype');
                    $flags= newGetStrCut($s, 'flags');
                    $parentid= newGetStrCut($s, 'parentid');

                    $parentid= phptrim(getColumnId($parentid));										 //可根据栏目名称找到对应ID   不存在为-1
                    //call echo("parentid",parentid)
                    $labletitle= newGetStrCut($s, 'labletitle');
                    //每页显示条数
                    $npagesize= newGetStrCut($s, 'npagesize');
                    if( $npagesize== '' ){ $npagesize= 10 ;}//默认分页数为10条

                    $target= newGetStrCut($s, 'target');

                    $smallimage= newGetStrCut($s, 'smallimage');
                    $bigImage= newGetStrCut($s, 'bigImage');
                    $bannerimage= newGetStrCut($s, 'bannerimage');

                    $templatepath= newGetStrCut($s, 'templatepath');


                    $bodycontent= newGetStrCut($s, 'bodycontent');
                    $bodycontent= contentTranscoding($bodycontent);
                    //是否启用生成html
                    $isonhtml= newGetStrCut($s, 'isonhtml');
                    if( $isonhtml== '0' || strtolower($isonhtml)== 'false' ){
                        $isonhtml= 0;
                    }else{
                        $isonhtml= 1;
                    }
                    //是否为nofollow
                    $nofollow= newGetStrCut($s, 'nofollow');
                    if( $nofollow== '1' || strtolower($nofollow)== 'true' ){
                        $nofollow= 1;
                    }else{
                        $nofollow= 0;
                    }
                    //call echo(columnname,nofollow)


                    $aboutcontent= newGetStrCut($s, 'aboutcontent');
                    $aboutcontent= contentTranscoding($aboutcontent);

                    $bodycontent= newGetStrCut($s, 'bodycontent');
                    $bodycontent= contentTranscoding($bodycontent);

                    connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'webcolumn (webtitle,webkeywords,webdescription,columnname,columnenname,columntype,sortrank,filename,flags,parentid,labletitle,aboutcontent,bodycontent,npagesize,isonhtml,nofollow,target,smallimage,bigImage,bannerimage,templatepath) values(\'' . $webtitle . '\',\'' . $webkeywords . '\',\'' . $webdescription . '\',\'' . $columnname . '\',\'' . $columnenname . '\',\'' . $columntype . '\',' . $sortrank . ',\'' . $fileName . '\',\'' . $flags . '\',' . $parentid . ',\'' . $labletitle . '\',\'' . $aboutcontent . '\',\'' . $bodycontent . '\',' . $npagesize . ',' . $isonhtml . ',' . $nofollow . ',\'' . $target . '\',\'' . $smallimage . '\',\'' . $bigImage . '\',\'' . $bannerimage . '\',\'' . $templatepath . '\')');
                }
            }
        }
    }

    //文章
    connExecute('delete from ' . $GLOBALS['db_PREFIX'] . 'articledetail');
    $content= getDirAllFileList($webdataDir . '/articledetail/','txt');
    $splStr= aspSplit($content, vbCrlf());
    hr();
    foreach( $splStr as $filePath){
        $fileName= getfilename($filePath);
        if( $filePath <> '' && instr('_#', substr($fileName, 0 , 1))== false ){
            ASPEcho('文章', $filePath);
            $content= getftext($filePath);
            //这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409
            if( instr($content,vbCrlf())==false ){
                $content=replace($content,chr(10),vbCrlf());
            }
            $splxx= aspSplit($content, vbCrlf() . '-------------------------------');
            foreach( $splxx as $s){
                if( instr($s, '【title】') > 0 ){
                    $s= $s . vbCrlf();
                    $parentid= newGetStrCut($s, 'parentid');
                    $parentid= getColumnId($parentid);
                    $title= newGetStrCut($s, 'title');
                    $titlecolor= newGetStrCut($s, 'titlecolor');
                    $webtitle= newGetStrCut($s, 'webtitle');
                    $webkeywords= newGetStrCut($s, 'webkeywords');
                    $webdescription= newGetStrCut($s, 'webdescription');


                    $author= newGetStrCut($s, 'author');
                    $sortrank= newGetStrCut($s, 'sortrank');
                    if( $sortrank== '' ){ $sortrank= 0 ;}
                    $adddatetime= newGetStrCut($s, 'adddatetime');
                    $fileName= newGetStrCut($s, 'filename');
                    $templatepath= newGetStrCut($s, 'templatepath');
                    $flags= newGetStrCut($s, 'flags');
                    $relatedtags= newGetStrCut($s, 'relatedtags');

                    $customaurl= newGetStrCut($s, 'customaurl');
                    $target= newGetStrCut($s, 'target');


                    $smallimage= newGetStrCut($s, 'smallimage');
                    $bigImage= newGetStrCut($s, 'bigImage');
                    $bannerimage= newGetStrCut($s, 'bannerimage');


                    $aboutcontent= newGetStrCut($s, 'aboutcontent');
                    $aboutcontent= contentTranscoding($aboutcontent);

                    $bodycontent= newGetStrCut($s, 'bodycontent');
                    $bodycontent= contentTranscoding($bodycontent);
                    //是否启用生成html
                    $isonhtml= newGetStrCut($s, 'isonhtml');
                    if( $isonhtml== '0' || strtolower($isonhtml)== 'false' ){
                        $isonhtml= 0;
                    }else{
                        $isonhtml= 1;
                    }
                    //是否为nofollow
                    $nofollow= newGetStrCut($s, 'nofollow');
                    if( $nofollow== '1' || strtolower($nofollow)== 'true' ){
                        $nofollow= 1;
                    }else{
                        $nofollow= 0;
                    }
                    connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'articledetail (parentid,title,titlecolor,webtitle,webkeywords,webdescription,author,sortrank,adddatetime,filename,flags,relatedtags,aboutcontent,bodycontent,updatetime,isonhtml,customaurl,nofollow,target,smallimage,bigImage,bannerimage,templatepath) values(' . $parentid . ',\'' . $title . '\',\''. $titlecolor .'\',\'' . $webtitle . '\',\'' . $webkeywords . '\',\'' . $webdescription . '\',\'' . $author . '\',' . $sortrank . ',\'' . $adddatetime . '\',\'' . $fileName . '\',\'' . $flags . '\',\'' . $relatedtags . '\',\''. $aboutcontent .'\',\'' . $bodycontent . '\',\'' . Now() . '\',' . $isonhtml . ',\'' . $customaurl . '\',' . $nofollow . ',\'' . $target . '\',\'' . $smallimage . '\',\'' . $bigImage . '\',\'' . $bannerimage . '\',\'' . $templatepath . '\')');
                }
            }
        }
    }

    //单页
    connExecute('delete from ' . $GLOBALS['db_PREFIX'] . 'OnePage');
    $content= getDirTxtList($webdataDir . '/OnePage/');
    $splStr= aspSplit($content, vbCrlf());
    hr();
    foreach( $splStr as $filePath){
        $fileName= getfilename($filePath);
        if( $filePath <> '' && instr('_#', substr($fileName, 0 , 1))== false ){
            ASPEcho('单页', $filePath);
            $content= getftext($filePath);
            //这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409
            if( instr($content,vbCrlf())==false ){
                $content=replace($content,chr(10),vbCrlf());
            }
            $splxx= aspSplit($content, vbCrlf() . '-------------------------------');
            foreach( $splxx as $s){
                if( instr($s, '【webkeywords】') > 0 ){
                    $s= $s . vbCrlf();
                    $title= newGetStrCut($s, 'title');
                    $displaytitle= newGetStrCut($s, 'displaytitle');
                    $webtitle= newGetStrCut($s, 'webtitle');
                    $webkeywords= newGetStrCut($s, 'webkeywords');
                    $webdescription= newGetStrCut($s, 'webdescription');



                    $adddatetime= newGetStrCut($s, 'adddatetime');
                    $fileName= newGetStrCut($s, 'filename');

                    $aboutcontent= newGetStrCut($s, 'aboutcontent');

                    $aboutcontent= contentTranscoding($aboutcontent);
                    $target= newGetStrCut($s, 'target');
                    $templatepath= newGetStrCut($s, 'templatepath');

                    $bodycontent= newGetStrCut($s, 'bodycontent');
                    $bodycontent= contentTranscoding($bodycontent);
                    //是否启用生成html
                    $isonhtml= newGetStrCut($s, 'isonhtml');
                    if( $isonhtml== '0' || strtolower($isonhtml)== 'false' ){
                        $isonhtml= 0;
                    }else{
                        $isonhtml= 1;
                    }
                    //是否为nofollow
                    $nofollow= newGetStrCut($s, 'nofollow');
                    if( $nofollow== '1' || strtolower($nofollow)== 'true' ){
                        $nofollow= 1;
                    }else{
                        $nofollow= 0;
                    }


                    connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'onepage (title,displaytitle,webtitle,webkeywords,webdescription,adddatetime,filename,isonhtml,aboutcontent,bodycontent,nofollow,target,templatepath) values(\'' . $title . '\',\'' . $displaytitle . '\',\'' . $webtitle . '\',\'' . $webkeywords . '\',\'' . $webdescription . '\',\'' . $adddatetime . '\',\'' . $fileName . '\',' . $isonhtml . ',\'' . $aboutcontent . '\',\'' . $bodycontent . '\',' . $nofollow . ',\'' . $target . '\',\'' . $templatepath . '\')');
                }
            }
        }
    }

    //竞价
    connExecute('delete from ' . $GLOBALS['db_PREFIX'] . 'Bidding');
    $content= getDirTxtList($webdataDir . '/Bidding/');
    $splStr= aspSplit($content, vbCrlf());
    hr();
    foreach( $splStr as $filePath){
        $fileName= getfilename($filePath);
        if( $filePath <> '' && instr('_#', substr($fileName, 0 , 1))== false ){
            ASPEcho('竞价', $filePath);
            $content= getftext($filePath);
            //这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409
            if( instr($content,vbCrlf())==false ){
                $content=replace($content,chr(10),vbCrlf());
            }
            $splxx= aspSplit($content, vbCrlf() . '-------------------------------');
            foreach( $splxx as $s){
                if( instr($s, '【webkeywords】') > 0 ){
                    $s=$s . vbCrlf();
                    $webkeywords= newGetStrCut($s, 'webkeywords');
                    $showreason= newGetStrCut($s, 'showreason');
                    $ncomputersearch= newGetStrCut($s, 'ncomputersearch');
                    $nmobliesearch= newGetStrCut($s, 'nmobliesearch');
                    $ncountsearch= newGetStrCut($s, 'ncountsearch');
                    $ndegree= newGetStrCut($s, 'ndegree');
                    $ndegree= getnumber($ndegree);
                    if( $ndegree== '' ){
                        $ndegree= 0;
                    }
                    connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'Bidding (webkeywords,showreason,ncomputersearch,nmobliesearch,ndegree) values(\'' . $webkeywords . '\',\'' . $showreason . '\',' . $ncomputersearch . ',' . $nmobliesearch . ',' . $ndegree . ')');
                }
            }
        }
    }

    //搜索统计
    connExecute('delete from ' . $GLOBALS['db_PREFIX'] . 'SearchStat');
    $content= getDirTxtList($webdataDir . '/SearchStat/');
    $splStr= aspSplit($content, vbCrlf());
    hr();
    foreach( $splStr as $filePath){
        $fileName= getfilename($filePath);
        if( $filePath <> '' && instr('_#', substr($fileName, 0 , 1))== false ){
            ASPEcho('搜索统计', $filePath);
            $content= getftext($filePath);
            //这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409
            if( instr($content,vbCrlf())==false ){
                $content=replace($content,chr(10),vbCrlf());
            }
            $splxx= aspSplit($content, vbCrlf() . '-------------------------------');
            foreach( $splxx as $s){
                if( instr($s, '【title】') > 0 ){
                    $s=$s . vbCrlf();
                    $title= newGetStrCut($s, 'title');
                    $webtitle= newGetStrCut($s, 'webtitle');
                    $webkeywords= newGetStrCut($s, 'webkeywords');
                    $webdescription= newGetStrCut($s, 'webdescription');

                    $customaurl= newGetStrCut($s, 'customaurl');
                    $target= newGetStrCut($s, 'target');
                    $isthrough= newGetStrCut($s, 'isthrough');
                    if( $isthrough== '0' || strtolower($isthrough)== 'false' ){
                        $isthrough= 0;
                    }else{
                        $isthrough= 1;
                    }
                    $sortrank= newGetStrCut($s, 'sortrank');
                    if( $sortrank== '' ){ $sortrank= 0 ;}
                    //是否启用生成html
                    $isonhtml= newGetStrCut($s, 'isonhtml');
                    if( $isonhtml== '0' || strtolower($isonhtml)== 'false' ){
                        $isonhtml= 0;
                    }else{
                        $isonhtml= 1;
                    }
                    //是否为nofollow
                    $nofollow= newGetStrCut($s, 'nofollow');
                    if( $nofollow== '1' || strtolower($nofollow)== 'true' ){
                        $nofollow= 1;
                    }else{
                        $nofollow= 0;
                    }
                    //call echo("title",title)
                    connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'SearchStat (title,webtitle,webkeywords,webdescription,customaurl,target,isthrough,sortrank,isonhtml,nofollow) values(\'' . $title . '\',\'' . $webtitle . '\',\'' . $webkeywords . '\',\'' . $webdescription . '\',\'' . $customaurl . '\',\'' . $target . '\',' . $isthrough . ',' . $sortrank . ',' . $isonhtml . ',' . $nofollow . ')');

                }
            }
        }
    }
    $itemid='';$username='';$ip='';$reply='';$tablename			='';//评论
    //评论
    connExecute('delete from ' . $GLOBALS['db_PREFIX'] . 'TableComment');
    $content= getDirTxtList($webdataDir . '/TableComment/');
    $splStr= aspSplit($content, vbCrlf());
    hr();
    foreach( $splStr as $filePath){
        $fileName= getfilename($filePath);
        if( $filePath <> '' && instr('_#', substr($fileName, 0 , 1))== false ){
            ASPEcho('评论', $filePath);
            $content= getftext($filePath);
            //这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409
            if( instr($content,vbCrlf())==false ){
                $content=replace($content,chr(10),vbCrlf());
            }
            $splxx= aspSplit($content, vbCrlf() . '-------------------------------');
            foreach( $splxx as $s){
                if( instr($s, '【title】') > 0 ){
                    $s=$s . vbCrlf();

                    $tablename= newGetStrCut($s, 'tablename');
                    $title= newGetStrCut($s, 'title');
                    $itemid= getArticleId(newGetStrCut($s, 'itemid'));
                    if( $itemid=='' ){ $itemid=0;}
                    //call echo("itemID",itemID)
                    $adddatetime= newGetStrCut($s, 'adddatetime');
                    $username= newGetStrCut($s, 'username');
                    $ip= newGetStrCut($s, 'ip');
                    $bodycontent= newGetStrCut($s, 'bodycontent');
                    $reply= newGetStrCut($s, 'reply');



                    $isthrough= newGetStrCut($s, 'isthrough');
                    if( $isthrough== '0' || strtolower($isthrough)== 'false' ){
                        $isthrough= 0;
                    }else{
                        $isthrough= 1;
                    }



                    //call echo("title",title)
                    connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'TableComment (tablename,title,itemid,adddatetime,username,ip,bodycontent,reply,isthrough) values(\''. $tablename .'\',\'' . $title . '\','. $itemid .',\''. $adddatetime .'\',\''. $username .'\',\''. $ip .'\',\''. $bodycontent .'\',\''. $reply .'\','. $isthrough .')');

                }
            }
        }
    }

    //友情链接
    connExecute('delete from ' . $GLOBALS['db_PREFIX'] . 'FriendLink');
    $content= getDirTxtList($webdataDir . '/FriendLink/');
    $splStr= aspSplit($content, vbCrlf());
    hr();
    foreach( $splStr as $filePath){
        $fileName= getfilename($filePath);
        if( $filePath <> '' && instr('_#', substr($fileName, 0 , 1))== false ){
            ASPEcho('评论', $filePath);
            $content= getftext($filePath);
            //这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409
            if( instr($content,vbCrlf())==false ){
                $content=replace($content,chr(10),vbCrlf());
            }
            $splxx= aspSplit($content, vbCrlf() . '-------------------------------');
            foreach( $splxx as $s){
                if( instr($s, '【title】') > 0 ){
                    $s=$s . vbCrlf();

                    $title= newGetStrCut($s, 'title');
                    $httpurl= newGetStrCut($s, 'httpurl');
                    $smallimage= newGetStrCut($s, 'smallimage');
                    $flags= newGetStrCut($s, 'flags');
                    $target= newGetStrCut($s, 'target');


                    $sortrank= newGetStrCut($s, 'sortrank');
                    if( $sortrank== '0' || strtolower($sortrank)== 'false' ){
                        $sortrank= 0;
                    }else{
                        $sortrank= 1;
                    }
                    $isthrough= newGetStrCut($s, 'isthrough');
                    if( $isthrough== '0' || strtolower($isthrough)== 'false' ){
                        $isthrough= 0;
                    }else{
                        $isthrough= 1;
                    }
                    //call echo("title",title)
                    connExecute('insert into ' . $GLOBALS['db_PREFIX'] . 'FriendLink (title,httpurl,smallimage,flags,sortrank,isthrough,target) values(\''. $title .'\',\'' . $httpurl . '\',\''. $smallimage .'\',\''. $flags .'\','. $sortrank .','. $isthrough .',\''. $target .'\')');

                }
            }
        }
    }


    writeSystemLog('','恢复默认数据' . $GLOBALS['db_PREFIX']);			//系统日志

}

//内容转码
function contentTranscoding( $content){
    $content= Replace(Replace(Replace(Replace($content, '<?', '&lt;?'), '?>', '?&gt;'), '<' . '%', '&lt;%'), '?>', '%&gt;');


    $splStr=''; $i=''; $s=''; $c=''; $isTranscoding=''; $isBR ='';
    $isTranscoding= false;
    $isBR= false;
    $splStr= aspSplit($content, vbCrlf());
    foreach( $splStr as $s){
        if( instr($s, '[&html转码&]') > 0 ){
            $isTranscoding= true;
        }
        if( instr($s, '[&html转码end&]') > 0 ){
            $isTranscoding= false;
        }
        if( instr($s, '[&全部换行&]') > 0 ){
            $isBR= true;
        }
        if( instr($s, '[&全部换行end&]') > 0 ){
            $isBR= false;
        }

        if( $isTranscoding== true ){
            $s= Replace(Replace($s, '[&html转码&]', ''), '<', '&lt;');
        }else{
            $s= Replace($s, '[&html转码end&]', '');
        }
        if( $isBR== true ){
            $s= Replace($s, '[&全部换行&]', '') . '<br>';
        }else{
            $s= Replace($s, '[&全部换行end&]', '');
        }
        $c= $c . $s . vbCrlf();
    }
    $c=replace(replace($c,'【b】','<b>'),'【/b】','</b>');
    $c=replace(replace($c,'【strong】','<strong>'),'【/strong】','</strong>');
    $contentTranscoding= $c;
    return @$contentTranscoding;
}
?>