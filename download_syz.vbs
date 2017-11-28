Set objFS = CreateObject("Scripting.FileSystemObject")

currentFolder = "D:\Temp\syz\"

If Not objFS.FolderExists(currentFolder) Then 
	objFS.CreateFolder currentFolder
End If

baseURL = "http://www.dqjj.com/syz/1.asp"

Set objall = CreateObject("ADODB.Stream")
objall.Charset = "utf-8"
objall.Type = 2
objall.LineSeparator = -1      'CRLF

'Set http = CreateObject("Msxml2.XMLHTTP")
Set http = CreateObject("Msxml2.XMLHttp.6.0")
'Set http = CreateObject("Msxml2.ServerXMLHttp.6.0")

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = currentFolder

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True
URLPattern = "<script>"

Const UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36"

Const URLHead = "http://www.dqjj.com/"

Const EndKey = "<script>"
NextPattern = "<a href=" & Chr(34) & "(\/syz\/.*?)" & Chr(34) & " title=" & Chr(34) & ".*?" & Chr(34) & ">.*?<\/a>"

'===
StartDownload
'===
objShell.Run "explorer.exe /e, " & currentFolder , 3 ,False

Set objall = Nothing
Set http = Nothing
Set objShell = Nothing
Set objFS = Nothing
Set objRegEx = Nothing

Sub StartDownload()
	nextpage = baseURL
	Do
		nextpage = DownloadURL_ALL(nextpage)
		If Len(nextpage) < 1 Then Exit Do
	Loop
End Sub


Function DownloadURL_ALL(url)
	on error resume next
	DownloadURL = ""
	
	filename = Mid(url, InStrRev(url, "/" )+1)
	WScript.Echo url
	
	http.open "GET", url, False
	http.setRequestHeader "User-Agent", UserAgent
	http.setRequestHeader "Content-Type", "text/html;charset=utf-8"
	http.setRequestHeader "referer", url
	http.send
	
	strContent = http.responseText
	aContent = Split(strContent, vbCrlf)
	WScript.Echo aContent
	
	If UBound(aContent)<10 Then
		WScript.Echo url & " 下载失败。"
		Exit Function
	End If

	outFile = currentFolder & filename & ".html"
	Set objFile = objFS.CreateTextFile(outFile,True,True)

	For y = 80 to UBound(aContent)
	'从第80行后不远才开始正文部分。
		ContentLine = aContent(y)
		If instr(ContentLine, EndKey) > 0 Then Exit For

		objRegEx.Pattern = NextPattern
		Set myMatches = objRegEx.Execute(ContentLine)
		If myMatches.count > 0 Then
			url_next = URLHead & myMatches(1).Submatches(0)
			If Instr(url_next, "mulu.asp") > 0 AND Instr(url, "331.asp") > 0 Then
				url_next = ""
			End If
			DownloadURL = url_next
			'WScript.echo url_next
		End If
		
		' objRegEx.Pattern = "<table[^>]*>"
		' newline = objRegEx.Replace(ContentLine,"")
		' objRegEx.Pattern = "<tr[^>]*>"
		' newline = objRegEx.Replace(newline,"")
		' objRegEx.Pattern = "<td[^>]*>"
		' newline = objRegEx.Replace(newline,"")
		' objRegEx.Pattern = "</td>"
		' newline = objRegEx.Replace(newline,"")
		' objRegEx.Pattern = "</tr>"
		' newline = objRegEx.Replace(newline,"")
		' objRegEx.Pattern = "</table>"
		' newline = objRegEx.Replace(newline,"")
		
		'newline = Trim(newline)
		
		If Len(ContentLine) > 0 Then
			objFile.WriteLine ContentLine
		End If
	Next

	
	objFile.Close

	WScript.echo "处理完成 " & filename 'StoryTitle
End Function

Function DownloadURL(url)
	on error resume next
	DownloadURL = ""
	
	filename = Mid(url, InStrRev(url, "/" )+1)
	objall.Open

	http.open "GET", url, False
	http.setRequestHeader "User-Agent", UserAgent
	http.setRequestHeader "Content-Type", "text/html;charset=gb2312"
	http.setRequestHeader "Pragma", "no-cache"
	http.setRequestHeader "Accept-Encoding", "gzip, deflate"
	http.setRequestHeader "Accept-Language", "zh-CN,zh;q=0.8,en;q=0.6"
	http.setRequestHeader "Cache-Control", "no-cache"
	http.setRequestHeader "Host", "www.dqjj.com"
	http.setRequestHeader "Connection", "keep-alive"
	http.setRequestHeader "Accept", "*/*"
	http.setRequestHeader "referer", url
	http.send
	
	strContent = http.responseText
	aContent = Split(strContent, Chr(10) )
	If UBound(aContent)<10 Then
		WScript.Echo url & " 下载失败。"
		objall.Close
		Exit Function
	End If
	'For Each ContentLine in aContent
	For y = 80 to UBound(aContent)
	'从第80行后不远才开始正文部分。
		ContentLine = aContent(y)
		If instr(ContentLine, EndKey) > 0 Then Exit For

		objRegEx.Pattern = NextPattern
		Set myMatches = objRegEx.Execute(ContentLine)
		If myMatches.count > 0 Then
			url_next = URLHead & myMatches(1).Submatches(0)
			If Instr(url_next, "menu.asp") > 0 Then
				url_next = ""
			End If
			DownloadURL = url_next
			'WScript.echo url_next
		End If
		
		objRegEx.Pattern = "<table[^>]*>"
		newline = objRegEx.Replace(ContentLine,"")
		objRegEx.Pattern = "<tr[^>]*>"
		newline = objRegEx.Replace(newline,"")
		objRegEx.Pattern = "<td[^>]*>"
		newline = objRegEx.Replace(newline,"")
		objRegEx.Pattern = "</td>"
		newline = objRegEx.Replace(newline,"")
		objRegEx.Pattern = "</tr>"
		newline = objRegEx.Replace(newline,"")
		objRegEx.Pattern = "</table>"
		newline = objRegEx.Replace(newline,"")

		newline = Trim(newline)
		If Len(newline) > 0 Then
			objall.WriteText newline,1
		End If
	Next

	objall.SaveToFile currentFolder & filename & ".html", 2
	objall.Close

	WScript.echo "处理完成 " & filename 'StoryTitle
End Function