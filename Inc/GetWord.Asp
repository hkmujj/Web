<%
dim WordReadContent
'��ùؼ��� 
Function GetKeyWord(ByVal Content)
	GetKeyWord = GetHandleKeyWord(Content,"")
End Function 
'��ô�����ؼ���  20150224
Function GetHandleKeyWord(ByVal Content, SType)
    Dim nOK, SplStr, I, J, TempContent, C, S 
	If WordReadContent = "" Then      
		WordReadContent = vbCrlf & GetFText("/DataDir/keyWord.txt")  & vbCrlf 
	End If
    'Call Die(WordReadContent)
    nOK = 0 
    If CStr(SType)="1" Then 
        TempContent = Content & "    " 
        For I = 1 To Len(Content)
            For J = 2 To 4
                S = Mid(TempContent, I, J) 
                If InStr(" " & C & " ", " " & S & " ") = False Then
                    If InStr(WordReadContent, vbCrlf & S & vbCrlf) > 0 Then
                        C = C & S & " " 
                        nOK = nOK + 1 
                    End If 
                End If 
            Next 
        Next 
	Else 	'���ַ��������淽����N��
        SplStr = Split(WordReadContent, vbCrlf) 
        For Each S In SplStr
			'Call Echo("S",S)
            If S <> "" Then
                If InStr(" " & C & " ", " " & S & " ") = False Then
                    If InStr(Content, S) > 0 Then
                        C = C & S & " " 
                        nOK = nOK + 1 
                    End If 
                End If 
            End If 
        Next
    End If 
	GetHandleKeyWord = C 
End Function 


'������ĺ������� 
Function GetChinaStr(Content)
    Dim I, C, J, S
    For I = 1 To Len(Content)
		J = Asc(Mid(Content, I, 1))
		S = Mid(Content, I, 1)
		 '�Ǻ����ۼ�
        If J < 0 Then
			If (J <= - 22033 And J >= - 24158) = False And InStr ("�? ", S) = False Then
            	C = C & S
			End If
		'���Ǻ��� �ж�ǰһ���ַ��Ƿ�Ϊ�ո�
        ElseIf Right(C,1)<>" " Then
            C = C & " " 
        End If 
    Next
	
    GetChinaStr = C 
End Function 
%> 