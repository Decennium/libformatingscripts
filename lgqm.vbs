'这份脚本：
'1 下载灰机wiki上的临高启明同人目录
'2 根据目录下载所有同人小说
'3 去除不必要的html标记并转码成gb2312
'4 合并成一份文本文件
'请使用cscript.exe执行
'使用wscript.exe执行脚本你会不停点击确定键直到累死
'勿谓言之不预

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

Wscript.echo "下载index"
DownloadIndex

WScript.echo "依照目录下载网页内容"
DownloadContent

WScript.echo "解压缩gz文件"
gunzipFiles

WScript.echo "合并文件"
EmergeFiles

EndTime = Now()
UsedTime = DateDiff("s",StartTime,EndTime)

wscript.echo "完成全部工作，得到最新版《临高启明wiki.txt》。耗时 " & UsedTime & " 秒。"

Sub DownloadIndex()
	gzfile = indexfile & ".gz"
	s = curl & url & " -o " & chr(34) & gzfile & chr(34)
	objShell.Run s ,7, True
	'解压缩
	objShell.exec "gzip -df " & chr(34) & gzfile & chr(34)
'	objShell.Run chr(34) & "c:\Program Files\7-Zip\7z.exe" & chr(34) &" e -y " & gzfile & " -o " & currentFolder ,1, True
	If objFS.FileExists(gzfile) Then
	'有些文件并不是gz格式，而就是txt格式
	'所以不能被解压，这就需要简单的更名
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
			WScript.echo "下载完成 " & n & " - " & title & ".txt.gz"
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
			'有些文件并不是gz格式，而就是txt格式
			'所以不能被解压，这就需要简单的更名
				objFS.CopyFile file.path,Mid(file.path,1,Len(file.path)-3),True
				objFS.DeleteFile file.path, True
			End If
			WScript.echo "解压完成 " & filename
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
				'在这里可以做进一步的处理
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
					objRegEx.Pattern = "^.*总目录.*$"
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
			wscript.echo "合并完成 " & filename
		End If
	Next
	objall.SaveToFile currentFolder & "临高启明wiki.txt",2
	objall.Close
End Sub
