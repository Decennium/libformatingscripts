'这份脚本：
'1 下载灰机wiki上的临高启明同人目录
'2 根据目录下载所有同人小说
'3 去除不必要的html标记并转码成gb2312
'4 合并成一份文本文件
'请使用cscript.exe执行
'使用WScript.exe执行脚本你会不停点击确定按钮直到累死
'勿谓言之不预

Set objFSO = CreateObject("Scripting.FileSystemObject")

BaseFolder = "D:\Temp\lgqm_wiki\"
TempFolder = BaseFolder & "txt\"
'这个文件夹可以随意更改
If Not objFSO.FolderExists(TempFolder) Then 
	CreateMultiLevelFolder TempFolder
End If

url = "http://lgqm.huiji.wiki/wiki/%E5%90%8C%E4%BA%BA%E4%BD%9C%E5%93%81%E7%AE%80%E8%A6%81%E4%BF%A1%E6%81%AF%E4%B8%80%E8%A7%88"

'Set http = CreateObject("Msxml2.XMLHTTP")
Set http = CreateObject("Msxml2.XMLHttp.6.0")
'Set http = CreateObject("Msxml2.ServerXMLHttp.6.0")

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = BaseFolder

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True
URLPattern = "^<td> *<a href="&chr(34)&"(.*?)"&chr(34)&" title="&chr(34)&"(.*?)"&chr(34)&">[^<]+<\/a>.*$"

Const UserAgent = "Mozilla/5.0 (Windows NT 6.3; Trident/7.0; rv:11.0) like Gecko"

Const URLHead = "http://lgqm.huiji.wiki"
Const MsgStart = "开始下载目录并依照目录下载网页内容，生成 "
Const MsgEnd = " 下载并生成完成。耗时 "
Const MsgFailed = " 下载失败。"
Const MsgCopy = " 已经复制到图书库。"
Const ServerAddress = "\\192.168.3.5\NewAdded\"

Const EndKey = "printfooter"

Const dtAll = 0
Const dtBest = 1
Const dtClosed = 2
Dim FileHead(3)
FileHead(dtAll) = "ALL_"
FileHead(dtBest) = "BEST"
FileHead(dtClosed) = "CLOS"
Dim FileEmergeName(3)
FileEmergeName(dtAll) = "临高启明wiki完整版.txt"
FileEmergeName(dtBest) = "临高启明wiki优秀版.txt"
FileEmergeName(dtClosed) = "临高启明wiki完结或已转正版本.txt"
'===
For iDLDType = dtAll to dtClosed
	startTime = Now()

	lgqm_File_Name = FileEmergeName(iDLDType)
	WScript.echo MsgStart & lgqm_File_Name
	DownloadAll iDLDType
	EmergeAll iDLDType

	EndTime = Now()
	UsedTime = DateDiff("s",StartTime,EndTime)

	WScript.echo lgqm_File_Name & MsgEnd & UsedTime & " 秒。"
	On Error resume next
	objFSO.CopyFile BaseFolder & lgqm_File_Name , ServerAddress ,True
	If Err.Number Then
		WScript.Echo Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
		Err.Clear
	Else
		WScript.echo vbCrLf & lgqm_File_Name & MsgCopy & vbCrLf
	End If
	On Error goto 0
Next
objShell.Run "explorer.exe /e, " & BaseFolder , 3 ,False

Set http = Nothing
Set objShell = Nothing
Set objFSO = Nothing
Set objRegEx = Nothing

Sub DownloadAll(DownloadType)
	http.open "GET", url, False
	http.setRequestHeader "Accept-Encoding", "gzip"
	http.setRequestHeader "User-Agent", UserAgent
	http.send
	strIndex = http.responseText
	aIndex = Split(strIndex, Chr(10) )
	If UBound(aIndex)<10 Then
		WScript.Echo lgqm_File_Name & MsgFailed
		Exit Sub
	End If
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
			Case dtClosed
				If (aIndex(x+12) = "<td>完结" OR aIndex(x+14) <> "<td>待转正") Then
					DownloadURL url2, i, DownloadType
					i = i + 1
				End If
			Case dtBest
				If (InStr(aIndex(x+18), "铁拳爆菊大出血杯")>0) Then
					DownloadURL url2, i, DownloadType
					i = i + 1
				End If
			Case dtAll
				DownloadURL url2, i, DownloadType
				i = i + 1
			End Select
			x = x + 20
			'直接跳到20行后。第21行是一行新的网页地址
			'如果目录页布局发生改变，这个数值可能需要修改
		End If
	Next
End Sub

Sub DownloadURL(url, i, DownloadType)
	File_Head = FileHead(DownloadType)

	Set objONE = CreateObject("ADODB.Stream")
	objONE.Charset = "gb2312"
	objONE.Type = 2
	objONE.LineSeparator = -1      'CRLF

	objONE.Open
	on error resume next
	http.open "GET", url, False
	http.setRequestHeader "Accept-Encoding", "gzip"
	http.setRequestHeader "User-Agent", UserAgent
	http.send
	strContent = http.responseText
	aContent = Split(strContent, Chr(10) )
	If UBound(aContent)<10 Then
		WScript.Echo url & MsgFailed
		objONE.Close
		Exit Sub
	End If
	'For Each ContentLine in aContent
	For y = 270 to UBound(aContent)
	'从第270行后不远才开始正文部分。
		ContentLine = aContent(y)
		If instr(ContentLine, EndKey) > 0 Then Exit For

		objRegEx.Pattern = "<h1>(.*?)<\/h1>"
		Set myMatches = objRegEx.Execute(ContentLine)
		If myMatches.count > 0 Then
			title = myMatches(0).Submatches(0)
			StoryTitle = "第" & Right("000" & i,4) & "篇" & " - " & title
			objONE.WriteText StoryTitle
			objONE.WriteText vbCrLf
		End If
		objRegEx.Pattern = "(<h\d|<p|<th|</table>)"
		If True = objRegEx.Test(ContentLine) Then
			'WScript.echo strline
			objRegEx.Pattern = "</h1>"
			newline = objRegEx.Replace(ContentLine,vbCrLf)
			objRegEx.Pattern = "<th[^>]*>|</table>"
			newline = objRegEx.Replace(newline,vbCrLf)
			objRegEx.Pattern = "</th>"
			newline = objRegEx.Replace(newline,": ")
			objRegEx.Pattern = "<[^>]+>|^ +|^.*总目录.*$|^.*总分类.*$"
			newline = objRegEx.Replace(newline,"")
			If Len(newline) > 0 Then
				objONE.WriteText newline,1
			End If
			'WScript.echo newline
		End If
	Next
	'objONE.WriteText vbCrLf & "==EOF==" & vbCrLf,1
	If objONE.State = 1 Then
		objONE.SaveToFile TempFolder & File_Head & Right("000" & i,4) & ".TXT" , 2
		objONE.Close
	End If
	WScript.echo "处理完成 " & StoryTitle
End Sub

Sub EmergeAll(DownloadType)
	File_Head = FileHead(DownloadType)

	Const ForReading = 1
	Set objOutputFile = objFSO.CreateTextFile(BaseFolder & lgqm_File_Name)

	Set objFolder = objFSO.GetFolder(TempFolder)
	'Wscript.Echo objFolder.Path

	Set colFiles = objFolder.Files

	For Each objFile in colFiles
		If Instr(objFile.Name,File_Head)>0 Then
			Set objTextFile = objFSO.OpenTextFile(TempFolder & objFile.Name, ForReading)
			strText = objTextFile.ReadAll
			objTextFile.Close
			objOutputFile.WriteLine strText
		End If
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
