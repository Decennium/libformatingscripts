'这份脚本：
'1 下载燃文 http://www.ranwena.com/files/article/0/996/
'2 根据目录下载所有小说章节
'3 去除不必要的html标记并转码成gb2312
'4 合并成一份文本文件
'请使用cscript.exe执行
'使用WScript.exe执行脚本你会不停点击确定按钮直到累死
'勿谓言之不预

Set objFSO = CreateObject("Scripting.FileSystemObject")

BaseFolder = "D:\Temp\lgqm_Novel\"
TempFolder = BaseFolder & "txt\"
'这个文件夹可以随意更改
If Not objFSO.FolderExists(TempFolder) Then 
	CreateMultiLevelFolder TempFolder
End If
HaveNewFile = False

url = "http://www.ranwena.com/files/article/0/996/index.html"

'Set http = CreateObject("Msxml2.XMLHTTP")
Set http = CreateObject("Msxml2.XMLHttp.6.0")
'Set http = CreateObject("Msxml2.ServerXMLHttp.6.0")

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = BaseFolder

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True
objRegEx.IgnoreCase = True
URLPattern = "<dd><a href="&chr(34)&"(.*?)"&chr(34)&">(.*?)<\/a><\/dd>$"

Const UserAgent = "Mozilla/5.0 (Windows NT 6.3; Trident/7.0; rv:11.0) like Gecko"

Const URLHead = "http://www.ranwena.com/files/article/0/996/"
Const MsgHead = "下载目录并依照目录下载网页内容，生成 "
Const MsgDone = " 下载并生成完成。耗时 "
Const MsgFailed = " 下载失败。"
Const MsgCopy = " 已经复制到图书库。"
Const MsgNoUpdate = " 没有更新，不做合并，不做复制。"
Const ServerAddress = "\\192.168.3.5\NewAdded\"

Const EndKey = "footer"

'===
startTime = Now()

lgqm_File_Name = "临高启明小说完整版.txt"
WScript.echo MsgHead & lgqm_File_Name

DownloadAll

If HaveNewFile Then EmergeAll

EndTime = Now()
UsedTime = DateDiff("s",StartTime,EndTime)

If HaveNewFile Then
	WScript.echo lgqm_File_Name & MsgDone & UsedTime & " 秒。"
	objFSO.CopyFile BaseFolder & lgqm_File_Name , ServerAddress ,True
	WScript.echo vbCrLf & lgqm_File_Name & MsgCopy & vbCrLf
	objShell.Run "explorer.exe /e, " & BaseFolder , 3 ,False
Else
	WScript.echo vbCrLf & lgqm_File_Name & MsgNoUpdate & vbCrLf
End If
'===

Set http = Nothing
Set objShell = Nothing
Set objFS = Nothing
Set objRegEx = Nothing

Sub DownloadAll()

	http.open "GET", url, False
	http.setRequestHeader "Accept-Encoding", "gzip"
	http.setRequestHeader "User-Agent", UserAgent
	http.send
	strIndex = http.responseText
	
	aIndex = Split(strIndex, vbCrlf )
	If UBound(aIndex)<10 Then
		WScript.Echo lgqm_File_Name & MsgFailed
		objall.Close
		Exit Sub
	End If
	i = 1
'	For Each strLine in aIndex
	For x = 80 to UBound(aIndex)
	'从第80行后不远才开始正文部分。
		strLine = aIndex(x)
		If instr(strLine, EndKey) > 0 Then Exit For

		objRegEx.Pattern = URLPattern
		Set myMatches = objRegEx.Execute(strline)
		If myMatches.count > 0 Then
			url2 = myMatches(0).Submatches(0)
			StoryTitle = myMatches(0).Submatches(1)
			If Not(objFSO.FileExists(TempFolder & Right("000" & i,4) & ".TXT")) Then 
				DownloadURL url2, i
				HaveNewFile = True
				WScript.echo StoryTitle & " 处理完成"
				WScript.Sleep 2000
			Else
				WScript.echo StoryTitle & " 已经存在，跳过下载"
			End If
			i = i + 1
		End If
	Next
End Sub

Sub DownloadURL(url, i)
	Set objONE = CreateObject("ADODB.Stream")
	objONE.Charset = "gb2312"
	objONE.Type = 2
	objONE.LineSeparator = -1      'CRLF

	objONE.Open
	'on error resume next
	http.open "GET", url, False
	http.setRequestHeader "Accept-Encoding", "gzip"
	http.setRequestHeader "User-Agent", UserAgent
	http.setRequestHeader "referer", url
	http.send
	strContent = http.responseText
	aContent = Split(strContent, chr(9))
	'WScript.Echo UBound(aContent)
	If UBound(aContent)<10 Then
		WScript.Echo url & MsgFailed
		objONE.Close
		Exit Sub
	End If
	'For Each ContentLine in aContent
	For y = 50 to UBound(aContent)
	'从第50行后不远才开始正文部分。
		ContentLine = aContent(y)
		If instr(ContentLine, EndKey) > 0 Then Exit For

		objRegEx.Pattern = "<h1>(.*?)<\/h1>"
		Set myMatches = objRegEx.Execute(ContentLine)
		If myMatches.count > 0 Then
			title = myMatches(0).Submatches(0)
			StoryTitle = "第" & Right("000" & i,4) & "章" & " - " & title
			objONE.WriteText StoryTitle
			objONE.WriteText vbCrLf
		End If
		objRegEx.Pattern = "<div id="&chr(34)&"content"&chr(34)&">(.*)</div>"
		If True = objRegEx.Test(ContentLine) Then
			'WScript.echo strline
			objRegEx.Pattern = "</*div[^>]*>"
			newline = objRegEx.Replace(ContentLine,"")
			objRegEx.Pattern = "</*a[^>]*>"
			newline = objRegEx.Replace(newline,"")
			objRegEx.Pattern = "&nbsp;|www|\.com"
			newline = objRegEx.Replace(newline,"")
			objRegEx.Pattern = "<br /><br />"
			newline = objRegEx.Replace(newline,vbCrLf)
			If Len(newline) > 0 Then
				objONE.WriteText newline,1
			End If
			'WScript.echo newline
		End If
	Next
	If objONE.State = 1 Then
		objONE.SaveToFile TempFolder & Right("000" & i,4) & ".TXT" , 2
		objONE.Close
	End If

	'Set objall = Nothing
End Sub

Sub EmergeAll()
	Const ForReading = 1
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objOutputFile = objFSO.CreateTextFile(BaseFolder & lgqm_File_Name)

	Set objFolder = objFSO.GetFolder(TempFolder)
	'Wscript.Echo objFolder.Path

	Set colFiles = objFolder.Files

	For Each objFile in colFiles
		'WScript.echo objFile.Name
		Set objTextFile = objFSO.OpenTextFile(TempFolder & objFile.Name, ForReading)
		strText = objTextFile.ReadAll
		objTextFile.Close
		objOutputFile.WriteLine strText
	Next

	objOutputFile.Close
End Sub

Sub CreateMultiLevelFolder(strPath)
	If objFSO.FolderExists(strPath) Then Exit Sub
	If Not objFSO.FolderExists(objFSO.GetParentFolderName(strPath)) Then 
		CreateMultiLevelFolder objFSO.GetParentFolderName(strPath) 
	End If 
	objFSO.CreateFolder strPath
End Sub
