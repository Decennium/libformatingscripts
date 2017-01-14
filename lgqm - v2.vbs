'这份脚本：
'1 下载灰机wiki上的临高启明同人目录
'2 根据目录下载所有同人小说
'3 去除不必要的html标记并转码成gb2312
'4 合并成一份文本文件
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

url = "http://lgqm.huiji.wiki/wiki/%E5%90%8C%E4%BA%BA%E4%BD%9C%E5%93%81%E7%AE%80%E8%A6%81%E4%BF%A1%E6%81%AF%E4%B8%80%E8%A7%88"
AllStory = True
'这个变量为True，下载所有同人故事
'这个变量如果为False，则仅下载已完结或已转正的同人故事
'两种下载方式得到的文件名不同。
Set oArgs = WScript.Arguments
If oArgs.count >0 Then
	If UCase(oArgs(0)) = "PART" Then
		AllStory = False
	End If
End If
Set oArgs = Nothing

If AllStory Then
	lgqm_File_Name = "临高启明wiki完整版.txt"
Else
	lgqm_File_Name = "临高启明wiki完结或已转正版本.txt"
End If

Set http = CreateObject("Msxml2.XMLHTTP")

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = currentFolder

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True

Wscript.echo "下载index并依照目录下载网页内容，生成 " & lgqm_File_Name
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
	'从第300行后不远才开始正文部分。
		strLine = aIndex(x)
		If instr(strLine,"printfooter") > 0 Then Exit For

		objRegEx.Pattern = "^<td> *<a href="&chr(34)&"(.*?)"&chr(34)&" title="&chr(34)&"(.*?)"&chr(34)&".+$"
		Set myMatches = objRegEx.Execute(strline)
		If myMatches.count > 0 Then
			If AllStory OR aIndex(x+11) = "<p>完结" OR aIndex(x+12) = "<td>完结" OR aIndex(x+14) = "<td>已转正" OR aIndex(x+14) = "<td>部分转正" Then
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
						If Len(newline) > 0 Then
							objall.WriteText newline,1
						End If
						'wscript.echo newline
					End If
				Next
				objall.WriteText vbCrLf & "//==" & vbCrLf,1
				i = i + 1
				wscript.echo "处理完成 " & Right("000" & i,4) & " - " & title
			Else
				x = x +20
				'直接跳到20行后。第21行是一行新的网页地址
				'如果目录页布局发生改变，这个数值可能需要修改
			End If
		End If
	Next
	objall.WriteText "共计 " & i & " 份同人故事"
	objall.WriteText "更新时间：" & Now()
	objall.SaveToFile currentFolder & lgqm_File_Name , 2
	objall.Close
	Set objall = Nothing
End Sub
