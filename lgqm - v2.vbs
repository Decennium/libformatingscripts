'��ݽű���
'1 ���ػһ�wiki�ϵ��ٸ�����ͬ��Ŀ¼
'2 ����Ŀ¼��������ͬ��С˵
'3 ȥ������Ҫ��html��ǲ�ת���gb2312
'4 �ϲ���һ���ı��ļ�
'��ʹ��cscript.exeִ��
'ʹ��wscript.exeִ�нű���᲻ͣ���ȷ����ťֱ������
'��ν��֮��Ԥ

startTime = Now()

currentFolder = "E:\Temp\lgqm-v2\"
'����ļ��п����������
url = "http://lgqm.huiji.wiki/wiki/%E5%90%8C%E4%BA%BA%E4%BD%9C%E5%93%81%E7%AE%80%E8%A6%81%E4%BF%A1%E6%81%AF%E4%B8%80%E8%A7%88"

Set http = CreateObject("Msxml2.XMLHTTP")

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = currentFolder

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True

Wscript.echo "����index������Ŀ¼������ҳ����"
DownloadContent

EndTime = Now()
UsedTime = DateDiff("s",StartTime,EndTime)

wscript.echo vbCrLf & "���ȫ����������ʱ " & UsedTime & " �롣"
objShell.exec "explorer.exe /e, " & currentFolder

Sub DownloadContent()
	Set objall = CreateObject("ADODB.Stream")
	objall.Charset = "gb2312"
	objall.Type = 2
	objall.Open
	objall.LineSeparator = -1      'CRLF

	http.open "GET", url, False
	http.setRequestHeader "Accept-Encoding", "gzip"
	http.send
	strIndex = http.responseText
	aIndex = Split(strIndex, Chr(10) )
	i = 1
'	For Each strLine in aIndex
	For x = 300 to UBound(aIndex)
		strLine = aIndex(x)
		If instr(strLine,"printfooter") > 0 Then Exit For

		objRegEx.Pattern = "^<td> *<a href="&chr(34)&"(.*?)"&chr(34)&" title="&chr(34)&"(.*?)"&chr(34)&".+$"
		Set myMatches = objRegEx.Execute(strline)
		If myMatches.count > 0 Then
			url = "http://lgqm.huiji.wiki" & myMatches(0).Submatches(0)
			objRegEx.Pattern = "[\/:?*<>"&chr(34)&"|]"
			title = objRegEx.Replace(myMatches(0).Submatches(1)," ")

			n = Right("000" & i,4)
			i = i + 1
			http.open "GET", url, False
			http.setRequestHeader "Accept-Encoding", "gzip"
			http.send
			strContent = http.responseText
			aContent = Split(strContent, Chr(10) )
			'For Each ContentLine in aContent
			For y = 280 to UBound(aContent)
				ContentLine = aContent(y)
				'�������������һ���Ĵ���
				If instr(ContentLine,"printfooter") > 0 Then Exit For
				
				objRegEx.Pattern = "(<h\d|<p|<th|</table>)"
				If True = objRegEx.Test(ContentLine) Then
					'wscript.echo strline
					objRegEx.Pattern = "</h1>"
					newline = objRegEx.Replace(ContentLine,vbCrLf)
					objRegEx.Pattern = "<th[^>]*>"
					newline = objRegEx.Replace(newline,vbCrLf)
					objRegEx.Pattern = "</th>"
					newline = objRegEx.Replace(newline,": ")
					objRegEx.Pattern = "</table>"
					newline = objRegEx.Replace(newline,vbCrLf)
					objRegEx.Pattern = "<[^>]+>"
					newline = objRegEx.Replace(newline,"")
					objRegEx.Pattern = "^ +"
					newline = objRegEx.Replace(newline,"")
					objRegEx.Pattern = "^.*��Ŀ¼.*$"
					newline = objRegEx.Replace(newline,"")
					If Len(newline) > 0 Then
						objall.WriteText newline,1
					End If
					'wscript.echo newline
				End If
			Next
			objall.WriteText vbCrLf & "//==" & vbCrLf,1
			wscript.echo "������� " & n & " - " & title
		End If
	Next
	objall.SaveToFile currentFolder & "�ٸ�����wiki.txt",2
	objall.Close
	wscript.echo "���ٸ�����wiki.txt���ϲ����"
End Sub
