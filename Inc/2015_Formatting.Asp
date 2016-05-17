<% 
'��ʽ��20150212


'Html��ʽ�� �򵥰� ��ǿ��2014 12 09��20150709
Function htmlFormatting(content)
    htmlFormatting = handleHtmlFormatting(content, False, 0, "") 
End Function
 
'������ʽ��
Function handleHtmlFormatting(ByVal content, isMsgBox, nErrLevel, action)
    Dim splStr, s, tempS, lCaseS, c, left4Str, left5Str, left6Str, left7Str, left8Str 
    Dim nLevel                                                                      '����
    Dim elseS, elseLable 

    Dim levelArray(299), keyWord 
    Dim lableName                                                                   '��ǩ����
    Dim isJavascript                                                                'Ϊjavascript
    Dim isTextarea                                                                  'Ϊ�����ı���<textarea
    Dim isPre                                                                       'Ϊpre
    isJavascript = False                                                            'Ĭ��javascriptΪ��
    isTextarea = False                                                              '�����ļ���Ϊ��
    isPre = False                                                                   'Ĭ��preΪ��
    nLevel = 0                                                                      '������

    action = "|" & action & "|"                                                     '����
    splStr = Split(content, vbCrLf) 
    For Each s In splStr
        tempS = s : elseS = s 
        s = trimVbCrlf(s) : lCaseS = LCase(s) 
        '�ж���20150710
        If(Left(lCaseS, 8) = "<script " Or Left(lCaseS, 8) = "<script>") And InStr(s, "</script>") = False And isJavascript = False Then
            isJavascript = True 
            c = c & phptrim(tempS) & vbCrLf 
        ElseIf isJavascript = True Then

            If Left(lCaseS, 9) = "</script>" Then
                isJavascript = False 
                c = c & phptrim(tempS) & vbCrLf                                                 '���������߿ո�
            Else
                c = c & tempS & vbCrLf                                                          'Ϊjs����ʾԭ�ı�  ������������߿ո�phptrim(tempS)
            End If 

            '�����ı����ж���20151019
        ElseIf(Left(lCaseS, 10) = "<textarea " Or Left(lCaseS, 10) = "<textarea>") And InStr(s, "</textarea>") = False And isTextarea = False Then
            isTextarea = True 
            c = c & phptrim(tempS) & vbCrLf 
        ElseIf isTextarea = True Then
            c = c & phptrim(tempS) & vbCrLf 
            If Left(lCaseS, 11) = "</textarea>" Then
                isTextarea = False 
            End If 
            '�����ı����ж���20151019
        ElseIf(Left(lCaseS, 5) = "<pre " Or Left(lCaseS, 5) = "<pre>") And InStr(s, "</pre>") = False And isPre = False Then
            isPre = True 
            c = c & phptrim(tempS) & vbCrLf 
        ElseIf isPre = True Then
            c = c & tempS & vbCrLf 
            If Left(lCaseS, 6) = "</pre>" Then
                isPre = False 
            End If 


        ElseIf s <> "" And isJavascript = False And isTextarea = False Then
            left4Str = "|" & Left(lCaseS, 4) & "|" : left5Str = "|" & Left(lCaseS, 5) & "|" : left6Str = "|" & Left(lCaseS, 6) & "|" 
            left7Str = "|" & Left(lCaseS, 7) & "|" : left8Str = "|" & Left(lCaseS, 8) & "|" 

            keyWord = ""                                                                    '�ؼ��ʳ�ʼ���
            lableName = ""                                                                  '��ǩ����
            If InStr("|<ul>|<ul |<li>|<li |<dt>|<dt |<dl>|<dl |<dd>|<dd |<tr>|<tr |<td>|<td |", left4Str) > 0 Then
                keyWord = left4Str 
                lableName = Mid(left4Str, 3, 2) 
            ElseIf InStr("|<div>|<div |", left5Str) > 0 Then
                keyWord = left5Str 
                lableName = Mid(left5Str, 3, 3) 
            ElseIf InStr("|<span>|<span |<form>|<form |", left6Str) > 0 Then
                keyWord = left6Str 
                lableName = Mid(left6Str, 3, 4) 

            ElseIf InStr("|<table>|<table |<tbody>|<tbody |", left7Str) > 0 Then
                keyWord = left7Str 
                lableName = Mid(left7Str, 3, 5) 

            ElseIf InStr("|<center>|<center |", left8Str) > 0 Then
                keyWord = left8Str 
                lableName = Mid(left8Str, 3, 6) 
            End If 
            keyWord = Trim(Replace(Replace(keyWord, "<", ""), ">", "")) 
            'call echo(KeyWord,lableName)
            '��ʼ
            If keyWord <> "" Then
                s = copyStr("    ", nLevel) & s 
                If Right(lCaseS, 3 + Len(lableName)) <> "</" & lableName & ">" And InStr(lCaseS, "</" & lableName & ">") = False Then
                    nLevel = nLevel + 1 
                    If nLevel >= 0 Then
                        levelArray(nLevel) = keyWord 
                    End If 
                End If 
            ElseIf InStr("|</ul>|</li>|</dl>|</dt>|</dd>|</tr>|</td>|", "|" & Left(lCaseS, 5) & "|") > 0 Or InStr("|</div>|", "|" & Left(lCaseS, 6) & "|") > 0 Or InStr("|</span>|</form>|", "|" & Left(lCaseS, 7) & "|") > 0 Or InStr("|</table>|</tbody>|", "|" & Left(lCaseS, 8) & "|") > 0 Or InStr("|</center>|", "|" & Left(lCaseS, 9) & "|") > 0 Then
                nLevel = nLevel - 1 
                s = copyStr("    ", nLevel) & s 
            Else
                s = copyStr("    ", nLevel) & s 
                '����ǽ�����ǩ���һ��
                If Right(lCaseS, 6) = "</div>" Then
                    If checkHtmlFormatting(lCaseS) = False Then
                        s = Left(s, Len(s) - 6) 
                        nLevel = nLevel - 1 
                        s = s & vbCrLf & copyStr("    ", nLevel) & "</div>" 
                    End If 
                ElseIf Right(lCaseS, 7) = "</span>" Then
                    If checkHtmlFormatting(lCaseS) = False Then
                        s = Left(s, Len(s) - 7) 
                        nLevel = nLevel - 1 
                        s = s & vbCrLf & copyStr("    ", nLevel) & "</span>" 
                    End If 
                ElseIf InStr("|</ul>|</dt>|<dl>|<dd>|", left5Str) > 0 Then
                    s = Left(s, Len(s) - 5) 
                    nLevel = nLevel - 1 
                    s = s & vbCrLf & copyStr("    ", nLevel) & Right(lCaseS, 5) 
                End If 


                '��   aaa</li>   ���ֽ�����   20160106
                elseS = phptrim(LCase(elseS)) 
                If InStr(elseS, "</") > 0 Then
                    elseLable = Mid(elseS, InStr(elseS, "</")) 
                    If InStr("|</ul>|</li>|</dl>|</dt>|</dd>|</tr>|</td>|</div>|</span>|<form>|", "|" & elseLable & "|") > 0 And nLevel > 0 Then
                        nLevel = nLevel - 1 
                    End If 
                End If 
                'call echo("s",replace(s,"<","&lt;"))


            End If 
            'call echo("",ShowHtml(temps)
            c = c & s & vbCrLf 
        ElseIf s = "" Then
            If InStr(action, "|delblankline|") = False And InStr(action, "|ɾ������|") = False Then'ɾ������
                c = c & vbCrLf 
            End If 
        End If 
    Next 
    handleHtmlFormatting = c 
    nErrLevel = nLevel                                                              '��ô��󼶱�
    If nLevel <> 0 And(LCase(isMsgBox) = "1" Or LCase(isMsgBox) = "true") Then
        Call echo("HTML��ǩ�д���", nLevel) 
    End If 
'Call Echo("nLevel",nLevel & "," & levelArray(nLevel))                '��ʾ�������20150212
End Function
 
'�����պ�HTML��ǩ(20150902)  ������ĸ����� �������
Function formatting(content, action)
    Dim i, endStr, s, c, labelName, startLabel, endLabel, endLabelStr, nLevel, isYes, parentLableName 
    Dim nextLableName                                                               '��һ����������
    Dim isA                                                                         '�Ƿ�ΪA����
    Dim isTextarea                                                                  '�Ƿ�Ϊ���������ı���
    Dim isScript                                                                    '�ű�����
    Dim isStyle                                                                     'Css�����ʽ��
    Dim isPre                                                                       '�Ƿ�Ϊpre
    startLabel = "<" 
    endLabel = ">" 
    nLevel = 0 
    action = "|" & action & "|"                                                     '�㼶
    isA = False : isTextarea = False : isScript = False : isStyle = False : isPre = False 
    content = Replace(Replace(content, vbCrLf, Chr(10)), vbTab, "    ") 

    For i = 1 To Len(content)
        s = Mid(content, i, 1) 
        endStr = Mid(content, i) 
        If s = "<" Then
            If InStr(endStr, ">") > 0 Then
                s = Mid(endStr, 1, InStr(endStr, ">")) 
                i = i + Len(s) - 1 
                s = Mid(s, 2, Len(s) - 2) 
                If Right(s, 1) = "/" Then
                    s = phptrim(Left(s, Len(s) - 1)) 
                End If 
                endStr = Right(endStr, Len(endStr) - Len(s) - 2)                                '����ַ���ȥ��ǰ��ǩ  -2����Ϊ����<>�����ַ�
                'ע��֮ǰ����labelName����
                labelName = Mid(s, 1, InStr(s & " ", " ") - 1) 
                labelName = LCase(labelName) 


                'call echo("labelName",labelName)
                If labelName = "a" Then
                    isA = True 
                ElseIf labelName = "/a" Then
                    isA = False 
                ElseIf labelName = "textarea" Then
                    isTextarea = True 
                ElseIf labelName = "/textarea" Then
                    isTextarea = False 
                ElseIf labelName = "script" Then
                    isScript = True 
                ElseIf labelName = "/script" Then
                    isScript = False 
                ElseIf labelName = "style" Then
                    isStyle = True 
                ElseIf labelName = "/style" Then
                    isStyle = False 
                ElseIf labelName = "pre" Then
                    isPre = True 
                ElseIf labelName = "/pre" Then
                    isPre = False 

                End If 

                endLabelStr = endLabel 
                nextLableName = getHtmlLableName(endStr, 0) 
                '��Ϊѹ��HTML
                If InStr(action, "|ziphtml|") = False And isPre = False Then
                    If isA = False Then
                        If InStr("|a|strong|u|i|s|script|", "|" & labelName & "|") = False And "/" & labelName <> nextLableName And InStr("|/a|/strong|/u|/i|/s|/script|", "|" & nextLableName & "|") = False Then
                            endLabelStr = endLabelStr & Chr(10) 
                        End If 
                    End If 
                End If 
                s = startLabel & s & endLabelStr 
                '��Ϊѹ��HTML
                If InStr(action, "|ziphtml|") = False And isPre = False Then
                    '�������            aaaaa</span>
                    If isA = False And isYes = False And Left(labelName, 1) = "/" And labelName <> "/script" And labelName <> "/a" Then
                        '�ų�����    <span>���췢��</span>     �����ж���һ���ֶβ�����vbcrlf����
                        If "/" & parentLableName <> labelName And Right(Trim(c), 1) <> Chr(10) Then
                            s = Chr(10) & s 
                        End If 
                    End If 
                End If 
                parentLableName = labelName 
                isYes = True 
            End If 
        ElseIf s <> "" Then
            isYes = False 
            'call echo("isPre",isPre)
            If isPre = False Then
                If s = Chr(10) Then
                    If isTextarea = False And isScript = False And isStyle = False Then
                        s = "" 
                    ElseIf isScript = True Then
                        If InStr(action, "|zipscripthtml|") > 0 Then
                            s = " " 
                        End If 
                    ElseIf isStyle = True Then
                        If InStr(action, "|zipstylehtml|") > 0 Then
                            s = "" 
                        End If 
                    ElseIf isTextarea = True Then
                        If InStr(action, "|ziptextareahtml|") > 0 Then
                            s = "" 
                        End If 
                    Else
                        s = Chr(10) 
                    End If
					' Right(Trim(c), 1) = ">")   Ϊ��ѹ��ʱ�õ�
                ElseIf (Right(Trim(c), 1) = Chr(10) or Right(Trim(c), 1) = ">") And phptrim(s) = "" And isTextarea = False And isScript = False Then
                    s = "" 
                End If 
            End If 
        End If 
        c = c & s 
    Next 
    c = Replace(c, Chr(10), vbCrLf)
    formatting = c 
End Function
 
Function zipHTML(c)
    zipHTML = formatting(c, "ziphtml|zipscripthtml|zipstylehtml")                   'ziphtml|zipscripthtml|zipstylehtml|ziptextareahtml
End Function
 
'���HTML��ǩ�Ƿ�ɶԳ��� �磨<div><ul><li>aa</li></ul></div></div>��
Function checkHtmlFormatting(ByVal content)
    Dim splStr, s, c, splxx, nLable, lableStr 
    content = LCase(content) 
    splStr = Split("ul|li|dt|dd|dl|div|span", "|") 
    For Each s In splStr
        s = phptrim(s) 
        If s <> "" Then
            nLable = 0 
            lableStr = "<" & s & " " 
            If InStr(content, lableStr) > 0 Then
                splxx = Split(content, lableStr) 
                nLable = nLable + UBound(splxx) 
            End If 
            lableStr = "<" & s & ">" 
            If InStr(content, lableStr) > 0 Then
                splxx = Split(content, lableStr) 
                nLable = nLable + UBound(splxx) 
            End If 
            lableStr = "</" & s & ">" 
            If InStr(content, lableStr) > 0 Then
                splxx = Split(content, lableStr) 
                nLable = nLable - UBound(splxx) 
            End If 
            'call echo(ShowHtml(lableStr),nLable)
            If nLable <> 0 Then
                checkHtmlFormatting = False 
                Exit Function 
            End If 
        End If 
    Next 
    checkHtmlFormatting = True 
End Function
 

%>    