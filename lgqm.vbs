'��ݽű���
'1 ���ػһ�wiki�ϵ��ٸ�����ͬ��Ŀ¼
'2 ����Ŀ¼��������ͬ��С˵
'3 ȥ������Ҫ��html��ǲ�ת���gb2312
'4 �ϲ���һ���ı��ļ�
'��ʹ��cscript.exeִ��
'ʹ��wscript.exeִ�нű���᲻ͣ���ȷ����ֱ������
'��ν��֮��Ԥ

startTime = Now()

currentFolder = "E:\Temp\lgqm-v2\"
url = "http://lgqm.huiji.wiki/wiki/%E5%90%8C%E4%BA%BA%E4%BD%9C%E5%93%81%E7%AE%80%E8%A6%81%E4%BF%A1%E6%81%AF%E4%B8%80%E8%A7%88"
indexfile = currentFolder & "index.html"
curl = "curl -H " & chr(34) & "Accept-Encoding: gzip" & chr(34) & " "

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = currentFolder

Set objFS = CreateObject("Scripting.FileSystemObject")
Set Folder = objFS.GetFolder(currentFolder)

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True

Wscript.echo "����index"
DownloadIndex

WScript.echo "����Ŀ¼������ҳ����"
DownloadContent

WScript.echo "��ѹ��gz�ļ�"
gunzipFiles

WScript.echo "�ϲ��ļ�"
EmergeFiles

EndTime = Now()
UsedTime = DateDiff("s",StartTime,EndTime)

wscript.echo "���ȫ���������õ����°桶�ٸ�����wiki.txt������ʱ " & UsedTime & " �롣"

Sub DownloadIndex()
	gzfile = indexfile & ".gz"
	s = curl & url & " -o " & chr(34) & gzfile & chr(34)
	objShell.Run s ,7, True
	'��ѹ��
	objShell.exec "gzip -df " & chr(34) & gzfile & chr(34)
'	objShell.Run chr(34) & "c:\Program Files\7-Zip\7z.exe" & chr(34) &" e -y " & gzfile & " -o " & currentFolder ,1, True
	If objFS.FileExists(gzfile) Then
	'��Щ�ļ�������gz��ʽ��������txt��ʽ
	'���Բ��ܱ���ѹ�������Ҫ�򵥵ĸ���
		objFS.CopyFile gzfile,Mid(gzfile,1,Len(gzfile)-3),True
		objFS.DeleteFile gzfile, True
	End If
	wscript.sleep 1000
End Sub

Sub DownloadContent()
	i = 1
	Set objStream = CreateObject("ADODB.Stream")
	objStream.Charset = "utf-8"
	objStream.Type = 2
	objStream.Open
	objStream.LoadFromFile indexfile
	objStream.LineSeparator = 10      'LF
	Do Until objStream.EOS
		strLine = objStream.ReadText(-2)
		objRegEx.Pattern = "^<td> *<a href="&chr(34)&"(.*?)"&chr(34)&" title="&chr(34)&"(.*?)"&chr(34)&".+$"
		Set myMatches = objRegEx.Execute(strline)
		If myMatches.count > 0 Then
			url = "http://lgqm.huiji.wiki" & myMatches(0).Submatches(0)
			objRegEx.Pattern = "[\/:?*<>"&chr(34)&"|]"
			title = objRegEx.Replace(myMatches(0).Submatches(1)," ")

			n = Right("000" & i,4)
			gzfile = currentFolder & n & " - " & title & ".txt.gz"
			s = curl & url & " -o " & chr(34) & gzfile & chr(34)
			objShell.Run s,7, True
			WScript.echo "������� " & n & " - " & title & ".txt.gz"
			i = i + 1
		End If
	Loop
	objStream.Close
	objFS.DeleteFile indexfile
End Sub

Sub gunzipFiles()
	For Each file in Folder.Files
		filename = file.name
		filepath = file.path
		extName = objFS.GetExtensionName(file.path)
		If (extName = "gz") Then
			objShell.Run "gzip -df " &chr(34) & file.path & chr(34) ,7, True
			If objFS.FileExists(filepath) Then
			'��Щ�ļ�������gz��ʽ��������txt��ʽ
			'���Բ��ܱ���ѹ�������Ҫ�򵥵ĸ���
				objFS.CopyFile file.path,Mid(file.path,1,Len(file.path)-3),True
				objFS.DeleteFile file.path, True
			End If
			WScript.echo "��ѹ��� " & filename
		End If
	Next
End Sub

Sub EmergeFiles()
	Set objStream = CreateObject("ADODB.Stream")
	objStream.Charset = "utf-8"
	objStream.Type = 2 'adTypeText
	objStream.LineSeparator = 10      'LF

	Set objall = CreateObject("ADODB.Stream")
	objall.Charset = "gb2312"
	objall.Type = 2
	objall.Open
	objall.LineSeparator = -1      'CRLF

	For Each file in Folder.Files
		filename = file.name
		extName = objFS.GetExtensionName(file.Path)
		If (extName = "txt") AND (left(file.name,1)="0") Then
			objStream.Open
			objStream.LoadFromFile file.Path
			Do
				'�������������һ���Ĵ���
				strLine = objStream.ReadText(-2)
				If instr(strline,"printfooter") > 0 Then Exit Do
				
				objRegEx.Pattern = "(<h\d|<p|<th|</table>)"
				If True = objRegEx.Test(strline) Then
					'wscript.echo strline
					objRegEx.Pattern = "</h1>"
					newline = objRegEx.Replace(strLine,Chr(13)&chr(10))
					objRegEx.Pattern = "<th[^>]*>"
					newline = objRegEx.Replace(newline,Chr(13)&chr(10))
					objRegEx.Pattern = "</th>"
					newline = objRegEx.Replace(newline,": ")
					objRegEx.Pattern = "</table>"
					newline = objRegEx.Replace(newline,Chr(13)&chr(10))
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
			Loop Until objStream.EOS
			objall.WriteText "",1
			objall.WriteText "//==",1
			objall.WriteText "",1
			objStream.Close
			objFS.DeleteFile file.path, True
			wscript.echo "�ϲ���� " & filename
		End If
	Next
	objall.SaveToFile currentFolder & "�ٸ�����wiki.txt",2
	objall.Close
End Sub
