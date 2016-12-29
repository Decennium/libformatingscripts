'这份脚本：
'1 下载灰机wiki上的临高启明同人目录
'2 根据目录下载所有同人小说
'3 去除不必要的html标记并转码成gb2312
'4 合并成一份文本文件
'请使用cscript.exe执行
'使用wscript.exe执行脚本你会不停点击确定按钮直到累死
'勿谓言之不预

startTime = Now()

currentFolder = "E:\Temp\lgqm-v2\"
'这个文件夹可以随意更改
url = "http://lgqm.huiji.wiki/wiki/%E5%90%8C%E4%BA%BA%E4%BD%9C%E5%93%81%E7%AE%80%E8%A6%81%E4%BF%A1%E6%81%AF%E4%B8%80%E8%A7%88"

Set http = CreateObject("Msxml2.XMLHTTP")

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = currentFolder

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True

Wscript.echo "下载index并依照目录下载网页内容"
DownloadContent

EndTime = Now()
UsedTime = DateDiff("s",StartTime,EndTime)

wscript.echo vbCrLf & "完成全部工作。耗时 " & UsedTime & " 秒。"
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
				'在这里可以做进一步的处理
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
					objRegEx.Pattern = "^.*总目录.*$"
					newline = objRegEx.Replace(newline,"")
					If Len(newline) > 0 Then
						objall.WriteText newline,1
					End If
					'wscript.echo newline
				End If
			Next
			objall.WriteText vbCrLf & "//==" & vbCrLf,1
			wscript.echo "处理完成 " & n & " - " & title
		End If
	Next
	objall.SaveToFile currentFolder & "临高启明wiki.txt",2
	objall.Close
	wscript.echo "《临高启明wiki.txt》合并完成"
End Sub
