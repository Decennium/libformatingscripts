'��ݽű���
'1 ���ػһ�wiki�ϵ��ٸ�����ͬ��Ŀ¼
'2 ����Ŀ¼��������ͬ��С˵
'3 ȥ������Ҫ��html��ǲ�ת���gb2312
'4 �ϲ���һ���ı��ļ�
'��ʹ��cscript.exeִ��
'ʹ��wscript.exeִ�нű���᲻ͣ���ȷ����ťֱ������
'��ν��֮��Ԥ

startTime = Now()

Set objFS = CreateObject("Scripting.FileSystemObject")

currentFolder = "E:\Temp\lgqm\"
'����ļ��п����������
If Not objFS.FolderExists(currentFolder) Then 
	objFS.CreateFolder currentFolder
End If

url = "http://lgqm.huiji.wiki/wiki/%E5%90%8C%E4%BA%BA%E4%BD%9C%E5%93%81%E7%AE%80%E8%A6%81%E4%BF%A1%E6%81%AF%E4%B8%80%E8%A7%88"
AllStory = True
'�������ΪTrue����������ͬ�˹���
'����������ΪFalse�����������������ת����ͬ�˹���
'�������ط�ʽ�õ����ļ�����ͬ��
Set oArgs = WScript.Arguments
If oArgs.count >0 Then
	If UCase(oArgs(0)) = "PART" Then
		AllStory = False
	End If
End If
Set oArgs = Nothing

If AllStory Then
	lgqm_File_Name = "�ٸ�����wiki������.txt"
Else
	lgqm_File_Name = "�ٸ�����wiki������ת���汾.txt"
End If

Set http = CreateObject("Msxml2.XMLHTTP")

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = currentFolder

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True

Wscript.echo "����index������Ŀ¼������ҳ���ݣ����� " & lgqm_File_Name
DownloadContent
Wscript.echo lgqm_File_Name & " ���ز�������ɡ�"

EndTime = Now()
UsedTime = DateDiff("s",StartTime,EndTime)

wscript.echo vbCrLf & "���ȫ����������ʱ " & UsedTime & " �롣"
objShell.Run "explorer.exe /e, " & currentFolder , 3 ,False
objFS.CopyFile currentFolder & lgqm_File_Name , "\\192.168.3.5\NewAdded\" ,True
wscript.echo vbCrLf & lgqm_File_Name & " �Ѿ����Ƶ�ͼ��⡣"

Set http = Nothing
Set objShell = Nothing
Set objFS = Nothing
Set objRegEx = Nothing

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
	i = 0
'	For Each strLine in aIndex
	For x = 300 to UBound(aIndex)
	'�ӵ�300�к�Զ�ſ�ʼ���Ĳ��֡�
		strLine = aIndex(x)
		If instr(strLine,"printfooter") > 0 Then Exit For

		objRegEx.Pattern = "^<td> *<a href="&chr(34)&"(.*?)"&chr(34)&" title="&chr(34)&"(.*?)"&chr(34)&".+$"
		Set myMatches = objRegEx.Execute(strline)
		If myMatches.count > 0 Then
			If AllStory OR aIndex(x+11) = "<p>���" OR aIndex(x+12) = "<td>���" OR aIndex(x+14) = "<td>��ת��" OR aIndex(x+14) = "<td>����ת��" Then
				url = "http://lgqm.huiji.wiki" & myMatches(0).Submatches(0)
				objRegEx.Pattern = "[\/:?*<>"&chr(34)&"|]"
				title = objRegEx.Replace(myMatches(0).Submatches(1)," ")

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
						If Len(newline) > 0 Then
							objall.WriteText newline,1
						End If
						'wscript.echo newline
					End If
				Next
				objall.WriteText vbCrLf & "//==" & vbCrLf,1
				i = i + 1
				wscript.echo "������� " & Right("000" & i,4) & " - " & title
			Else
				x = x +20
				'ֱ������20�к󡣵�21����һ���µ���ҳ��ַ
				'���Ŀ¼ҳ���ַ����ı䣬�����ֵ������Ҫ�޸�
			End If
		End If
	Next
	objall.WriteText "���� " & i & " ��ͬ�˹���"
	objall.WriteText "����ʱ�䣺" & Now()
	objall.SaveToFile currentFolder & lgqm_File_Name , 2
	objall.Close
	Set objall = Nothing
End Sub
