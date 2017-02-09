'这份脚本：
'1 下载灰机wiki上的临高启明同人中的铁拳爆菊入围作品
'2 去除不必要的html标记并转码成gb2312
'3 合并成一份文本文件
'请使用cscript.exe执行
'使用wscript.exe执行脚本你会不停点击确定按钮直到累死
'勿谓言之不预

startTime = Now()

Set objFS = CreateObject("Scripting.FileSystemObject")

currentFolder = "E:\Temp\lgqm\"
'这个文件夹可以随意更改
If Not objFS.FolderExists(currentFolder) Then 
	objFS.CreateFolder currentFolder
End If

tongren = "lgqm-tongren.txt"

lgqm_File_Name = "临高启明wiki优秀版.txt"

Set http = CreateObject("Msxml2.XMLHTTP")

Set objall = CreateObject("ADODB.Stream")
objall.Charset = "gb2312"
objall.Type = 2
objall.Open
objall.LineSeparator = -1      'CRLF

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = currentFolder

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True

Wscript.echo "读取index并依照目录下载网页内容，生成 " & lgqm_File_Name
DownloadContent
Wscript.echo lgqm_File_Name & " 下载并生成完成。"

EndTime = Now()
UsedTime = DateDiff("s",StartTime,EndTime)

wscript.echo vbCrLf & "完成全部工作。耗时 " & UsedTime & " 秒。"
objShell.Run "explorer.exe /e, " & currentFolder , 3 ,False
objFS.CopyFile currentFolder & lgqm_File_Name , "\\192.168.3.5\NewAdded\" ,True
wscript.echo vbCrLf & lgqm_File_Name & " 已经复制到图书库。"

Set http = Nothing
Set objShell = Nothing
Set objFS = Nothing
Set objRegEx = Nothing

Sub DownloadContent()
	Const FOR_READING = 1
	Set objTS = objFS.OpenTextFile(tongren, FOR_READING)
	i = 1
	Do
		url = objTS.Readline
		
		DownloadURL url

		i = i + 1
		wscript.echo "处理完成 - " & url
	Loop until objTS.AtEndOfStream

	objall.WriteText vbCrLf & "//==" & vbCrLf,1

	objall.WriteText "更新时间：" & Now()
	objall.SaveToFile currentFolder & lgqm_File_Name , 2
	objall.Close
	Set objall = Nothing
End Sub

Sub DownloadURL( url)
	http.open "GET", url, False
	http.setRequestHeader "Accept-Encoding", "gzip"
	http.send
	strContent = http.responseText
	aContent = Split(strContent, Chr(10) )
	'For Each ContentLine in aContent
	For y = 280 to UBound(aContent)
	'从第280行后不远才开始正文部分。
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
			objRegEx.Pattern = "^.*总分类.*$"
			newline = objRegEx.Replace(newline,"")
			If Len(newline) > 0 Then
				objall.WriteText newline,1
			End If
			'wscript.echo newline
		End If
	Next
	objall.WriteText vbCrLf & "//==" & vbCrLf,1
End Sub