'��ݽű���
'1 ���ػһ�wiki�ϵ��ٸ�����ͬ��Ŀ¼
'2 ����Ŀ¼��������ͬ��С˵
'3 ȥ������Ҫ��html��ǲ�ת���gb2312
'4 �ϲ���һ���ı��ļ�
'��ʹ��cscript.exeִ��
'ʹ��wscript.exeִ�нű���᲻ͣ���ȷ����ťֱ������
'��ν��֮��Ԥ

Set objFS = CreateObject("Scripting.FileSystemObject")

currentFolder = "E:\Temp\lgqm\"
'����ļ��п����������
If Not objFS.FolderExists(currentFolder) Then 
	objFS.CreateFolder currentFolder
End If

url = "http://lgqm.huiji.wiki/wiki/%E5%90%8C%E4%BA%BA%E4%BD%9C%E5%93%81%E7%AE%80%E8%A6%81%E4%BF%A1%E6%81%AF%E4%B8%80%E8%A7%88"
tongren = "E:\lgqm-best.txt"

Set objall = CreateObject("ADODB.Stream")
objall.Charset = "gb2312"
objall.Type = 2
objall.LineSeparator = -1      'CRLF

Set http = CreateObject("Msxml2.XMLHTTP")

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = currentFolder

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True
URLPattern = "^<td> *<a href="&chr(34)&"(.*?)"&chr(34)&" title="&chr(34)&"(.*?)"&chr(34)&">[^<]+<\/a>.*$"

UserAgent = "Mozilla/5.0 (Windows NT 6.3; Trident/7.0; rv:11.0) like Gecko"

URLHead = "http://lgqm.huiji.wiki"
MsgHead = "����Ŀ¼������Ŀ¼������ҳ���ݣ����� "
MsgTime = " ���ز�������ɡ���ʱ "
CopyTail = " �Ѿ����Ƶ�ͼ��⡣"
ServerAddress = "\\192.168.3.5\NewAdded\"

EndKey = "printfooter"
'===
startTime = Now()
lgqm_File_Name = "�ٸ�����wiki������.txt"
Wscript.echo MsgHead & lgqm_File_Name
DownloadAll
EndTime = Now()
UsedTime = DateDiff("s",StartTime,EndTime)

wscript.echo lgqm_File_Name & MsgTime & UsedTime & " �롣"
objFS.CopyFile currentFolder & lgqm_File_Name , ServerAddress ,True
wscript.echo vbCrLf & lgqm_File_Name & CopyTail
'===
startTime = Now()
lgqm_File_Name = "�ٸ�����wiki������ת���汾.txt"
Wscript.echo MsgHead & lgqm_File_Name
DownloadEndedPart
EndTime = Now()
UsedTime = DateDiff("s",StartTime,EndTime)

wscript.echo lgqm_File_Name & MsgTime & UsedTime & " �롣"
objFS.CopyFile currentFolder & lgqm_File_Name , ServerAddress ,True
wscript.echo vbCrLf & lgqm_File_Name & CopyTail
'===
startTime = Now()
lgqm_File_Name = "�ٸ�����wiki�����.txt"
Wscript.echo MsgHead & lgqm_File_Name
DownloadBest
EndTime = Now()
UsedTime = DateDiff("s",StartTime,EndTime)

wscript.echo lgqm_File_Name & MsgTime & UsedTime & " �롣"
objFS.CopyFile currentFolder & lgqm_File_Name , ServerAddress ,True
wscript.echo vbCrLf & lgqm_File_Name & CopyTail
'===
objShell.Run "explorer.exe /e, " & currentFolder , 3 ,False

Set objall = Nothing
Set http = Nothing
Set objShell = Nothing
Set objFS = Nothing
Set objRegEx = Nothing

Sub DownloadAll()
	objall.Open

	http.open "GET", url, False
	http.setRequestHeader "Accept-Encoding", "gzip"
	http.setRequestHeader "User-Agent", UserAgent
	http.send
	strIndex = http.responseText
	aIndex = Split(strIndex, Chr(10) )
	i = 0
'	For Each strLine in aIndex
	For x = 300 to UBound(aIndex)
	'�ӵ�300�к�Զ�ſ�ʼ���Ĳ��֡�
		strLine = aIndex(x)
		If instr(strLine, EndKey) > 0 Then Exit For

		objRegEx.Pattern = URLPattern
		Set myMatches = objRegEx.Execute(strline)
		If myMatches.count > 0 Then
			url2 = URLHead & myMatches(0).Submatches(0)
			objRegEx.Pattern = "[\/:?*<>"&chr(34)&"|]"
			title = objRegEx.Replace(myMatches(0).Submatches(1)," ")

			DownloadURL url2

			i = i + 1
			wscript.echo "������� " & Right("000" & i,4) & " - " & titlex
			x = x + 20
		End If
	Next
	objall.WriteText "���� " & i & " ��ͬ�˹���"
	objall.WriteText "����ʱ�䣺" & Now()
	objall.SaveToFile currentFolder & lgqm_File_Name , 2

	objall.Close
End Sub

Sub DownloadEndedPart()
	objall.Open

	http.open "GET", url, False
	http.setRequestHeader "Accept-Encoding", "gzip"
	http.setRequestHeader "User-Agent", UserAgent
	http.send
	strIndex = http.responseText
	aIndex = Split(strIndex, Chr(10) )
	i = 0
'	For Each strLine in aIndex
	For x = 300 to UBound(aIndex)
	'�ӵ�300�к�Զ�ſ�ʼ���Ĳ��֡�
		strLine = aIndex(x)
		If instr(strLine, EndKey) > 0 Then Exit For

		objRegEx.Pattern = URLPattern
		Set myMatches = objRegEx.Execute(strline)
		If myMatches.count > 0 Then
			If aIndex(x+12) = "<td>���" OR aIndex(x+14) <> "<td>��ת��" Then
'			If aIndex(x+11) = "<p>���" OR aIndex(x+12) = "<td>���" _
'				OR aIndex(x+14) <> "<td>��ת��" Then
				url2 = URLHead & myMatches(0).Submatches(0)
				objRegEx.Pattern = "[\/:?*<>"&chr(34)&"|]"
				title = objRegEx.Replace(myMatches(0).Submatches(1)," ")

				DownloadURL url2

				i = i + 1
				wscript.echo "������� " & Right("000" & i,4) & " - " & title
				x = x + 20
				'ֱ������20�к󡣵�21����һ���µ���ҳ��ַ
				'���Ŀ¼ҳ���ַ����ı䣬�����ֵ������Ҫ�޸�
			End If
		End If
	Next
	objall.WriteText "���� " & i & " ��ͬ�˹���"
	objall.WriteText "����ʱ�䣺" & Now()
	objall.SaveToFile currentFolder & lgqm_File_Name , 2

	objall.Close
End Sub

Sub DownloadBest()
	objall.Open

	Const FOR_READING = 1
	Set objTS = objFS.OpenTextFile(tongren, FOR_READING)
	i = 1
	Do
		url = objTS.Readline
		
		DownloadURL url

		i = i + 1
		wscript.echo "������� - " & url
	Loop until objTS.AtEndOfStream

	objall.WriteText vbCrLf & "//==" & vbCrLf,1

	objall.WriteText "����ʱ�䣺" & Now()
	objall.SaveToFile currentFolder & lgqm_File_Name , 2

	objall.Close
End Sub

Sub DownloadURL(url)
	http.open "GET", url, False
	http.setRequestHeader "Accept-Encoding", "gzip"
	http.send
	strContent = http.responseText
	aContent = Split(strContent, Chr(10) )
	'For Each ContentLine in aContent
	For y = 280 to UBound(aContent)
	'�ӵ�280�к�Զ�ſ�ʼ���Ĳ��֡�
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
			objRegEx.Pattern = "^.*�ܷ���.*$"
			newline = objRegEx.Replace(newline,"")
			If Len(newline) > 0 Then
				objall.WriteText newline,1
			End If
			'wscript.echo newline
		End If
	Next
	objall.WriteText vbCrLf & "//==" & vbCrLf,1
End Sub