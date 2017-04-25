'这份脚本：
'1 下载灰机wiki上的临高启明同人目录
'2 根据目录下载所有同人小说
'3 去除不必要的html标记并转码成gb2312
'4 合并成一份文本文件
'请使用cscript.exe执行
'使用WScript.exe执行脚本你会不停点击确定按钮直到累死
'勿谓言之不预

Set objFS = CreateObject("Scripting.FileSystemObject")

currentFolder = "E:\Temp\lgqm\"
'这个文件夹可以随意更改
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

Const UserAgent = "Mozilla/5.0 (Windows NT 6.3; Trident/7.0; rv:11.0) like Gecko"

Const URLHead = "http://lgqm.huiji.wiki"
Const MsgHead = "下载目录并依照目录下载网页内容，生成 "
Const Msg2 = " 下载并生成完成。耗时 "
Const CopyTail = " 已经复制到图书库。"
Const ServerAddress = "\\192.168.3.5\NewAdded\"

Const EndKey = "printfooter"

Const dtAll = 0
Const dtEnded = 1
Const dtBest = 2
'===
DownloadAll dtAll
'===
DownloadAll dtEnded
'===
DownloadAll dtBest
'===
objShell.Run "explorer.exe /e, " & currentFolder , 3 ,False

Set objall = Nothing
Set http = Nothing
Set objShell = Nothing
Set objFS = Nothing
Set objRegEx = Nothing

Sub DownloadAll(DownloadType)
	startTime = Now()
	Select Case DownloadType
	Case dtEnded
		lgqm_File_Name = "临高启明wiki完结或已转正版本.txt"
	Case dtBest
		lgqm_File_Name = "临高启明wiki优秀版.txt"
	Case Else
		lgqm_File_Name = "临高启明wiki完整版.txt"
	End Select
	WScript.echo MsgHead & lgqm_File_Name

	objall.Open

	http.open "GET", url, False
	http.setRequestHeader "Accept-Encoding", "gzip"
	http.setRequestHeader "User-Agent", UserAgent
	http.send
	strIndex = http.responseText
	aIndex = Split(strIndex, Chr(10) )
	i = 1
'	For Each strLine in aIndex
	For x = 300 to UBound(aIndex)
	'从第300行后不远才开始正文部分。
		strLine = aIndex(x)
		If instr(strLine, EndKey) > 0 Then Exit For

		objRegEx.Pattern = URLPattern
		Set myMatches = objRegEx.Execute(strline)
		If myMatches.count > 0 Then
			url2 = URLHead & myMatches(0).Submatches(0)
			Select Case DownloadType
			Case dtEnded
				If (aIndex(x+12) = "<td>完结" OR aIndex(x+14) <> "<td>待转正") Then
					DownloadURL url2, i
					i = i + 1
				End If
			Case dtBest
				If (InStr(aIndex(x+18), "铁拳爆菊大出血杯")>0) Then
					DownloadURL url2, i
					i = i + 1
				End If
			Case Else
				DownloadURL url2, i
				i = i + 1
			End Select
			x = x + 20
			'直接跳到20行后。第21行是一行新的网页地址
			'如果目录页布局发生改变，这个数值可能需要修改
		End If
	Next

	objall.WriteText "共计 " & i & " 份同人故事"
	objall.WriteText "更新时间：" & Now()
	objall.SaveToFile currentFolder & lgqm_File_Name , 2

	objall.Close

	EndTime = Now()
	UsedTime = DateDiff("s",StartTime,EndTime)

	WScript.echo lgqm_File_Name & Msg2 & UsedTime & " 秒。"
	objFS.CopyFile currentFolder & lgqm_File_Name , ServerAddress ,True
	WScript.echo vbCrLf & lgqm_File_Name & CopyTail
End Sub

Sub DownloadURL(url, i)
	http.open "GET", url, False
	http.setRequestHeader "Accept-Encoding", "gzip"
	http.send
	strContent = http.responseText
	aContent = Split(strContent, Chr(10) )
	'For Each ContentLine in aContent
	For y = 270 to UBound(aContent)
	'从第270行后不远才开始正文部分。
		ContentLine = aContent(y)
		If instr(ContentLine,"printfooter") > 0 Then Exit For

		objRegEx.Pattern = "<h1>(.*?)<\/h1>"
		Set myMatches = objRegEx.Execute(ContentLine)
		If myMatches.count > 0 Then title = myMatches(0).Submatches(0)
		StoryTitle = "第" & Right("000" & i,4) & "篇" & " - " & title
		objall.WriteText StoryTitle
		objall.WriteText vbCrLf
		
		objRegEx.Pattern = "(<h\d|<p|<th|</table>)"
		If True = objRegEx.Test(ContentLine) Then
			'WScript.echo strline
			objRegEx.Pattern = "</h1>"
			newline = objRegEx.Replace(ContentLine,vbCrLf)
			objRegEx.Pattern = "<th[^>]*>"
			newline = objRegEx.Replace(newline,vbCrLf)
			objRegEx.Pattern = "</table>"
			newline = objRegEx.Replace(newline,vbCrLf)
			objRegEx.Pattern = "</th>"
			newline = objRegEx.Replace(newline,": ")
			objRegEx.Pattern = "<[^>]+>"
			newline = objRegEx.Replace(newline,"")
			objRegEx.Pattern = "^ +"
			newline = objRegEx.Replace(newline,"")
			objRegEx.Pattern = "^.*总目录.*$"
			newline = objRegEx.Replace(newline,"")
			objRegEx.Pattern = "^.*总分类.*$"
			newline = objRegEx.Replace(newline,"")
			If Len(newline) > 0 Then
				objall.WriteText newline,1
			End If
			'WScript.echo newline
		End If
	Next
	objall.WriteText vbCrLf & "==EOF==" & vbCrLf,1
	WScript.echo "处理完成 " & StoryTitle
End Sub