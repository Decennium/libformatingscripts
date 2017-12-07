Set objFSO = CreateObject("Scripting.FileSystemObject")

BaseFolder = "D:\Temp\WTJ_Blog\"

TempFolder = BaseFolder & "txt\"
'这个文件夹可以随意更改
' If Not objFSO.FolderExists(TempFolder) Then 
	' CreateMultiLevelFolder TempFolder
' End If
Blog_File_Name = "WTJ_Blogs.txt"

HaveNewFile = False

baseURL = "http://www.caogen.com/blog/infor_detail.aspx?id=95&articleId=2358"

Set objall = CreateObject("ADODB.Stream")
objall.Charset = "gb2312"
objall.Type = 2
objall.LineSeparator = -1      'CRLF

Set http = CreateObject("Msxml2.XMLHttp.6.0")
'Set http = CreateObject("Msxml2.ServerXMLHttp.6.0")

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = BaseFolder

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True
objRegEx.IgnoreCase = True

URLPattern = "<script>"

Const UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36"

Const URLHead = "http://www.caogen.com/blog/infor_detail.aspx?id=95&articleId="

Const EndKey = "Next_Down"
NextPattern = "<a id="&Chr(34)&"Next_Up"&Chr(34)&" title=[^>]+ href=" & Chr(34) & "\/blog\/infor_detail.aspx.+articleId=(.*?)" & Chr(34) & ">上一篇　(.*?)<\/a>"

'===
StartDownload

If HaveNewFile Then EmergeAll

'===
objShell.Run "explorer.exe /e, " & BaseFolder , 3 ,False

Set objall = Nothing
Set http = Nothing
Set objShell = Nothing
Set objFSO = Nothing
Set objRegEx = Nothing

Sub StartDownload()
	nextpage = baseURL
	Do
		filename = Right("0000" & Mid(nextpage, InStrRev(nextpage, "=" )+1),6)&".txt"
		nextpage = DownloadURL(nextpage)
		If Len(nextpage) < 1 Then Exit Do
	Loop
End Sub

Function DownloadURL(url)
	on error resume next
	http.open "GET", url, False
	http.setRequestHeader "User-Agent", UserAgent
	http.setRequestHeader "Content-Type", "text/html;charset=gb2312"
	http.setRequestHeader "Accept-Encoding", "gzip, deflate"
	http.setRequestHeader "Accept", "*/*"
	http.setRequestHeader "referer", url
	http.send
	If Err.Number > 0 Then
		WScript.Echo Err.Number & " Src: " & Err.Source & " Desc: " & Err.Description
		Err.Clear
		Exit Function
	End If
	On Error goto 0

	strContent = http.responseText
	aContent = Split(strContent, vbCrLf )
	If UBound(aContent)<10 Then
		WScript.Echo url & " 下载失败。"
		Exit Function
	End If

	DownloadURL = ""
	filename = Right("0000" & Mid(url, InStrRev(url, "=" )+1),6)&".txt"

If Not(objFSO.FileExists(TempFolder & filename)) Then
	objall.Open
	'For Each ContentLine in aContent
	For y = 280 to UBound(aContent)
	'从第280行后不远才开始正文部分。
		ContentLine = aContent(y)
		If instr(ContentLine, EndKey) > 0 Then Exit For

		objRegEx.Pattern = NextPattern
		Set myMatches = objRegEx.Execute(ContentLine)
		If myMatches.count > 0 Then
			url_next = URLHead & myMatches(0).Submatches(0)
			DownloadURL = url_next
			HaveNewFile = True
			'WScript.echo url_next
		Else
			url_next = ""
		End If

		objRegEx.Pattern = "<span"
		Set myMatches = objRegEx.Execute(ContentLine)
		If myMatches.count > 0 Then
			objRegEx.Pattern = "<span[^>]*Blog_Infor[^>]*>(.+)</span>"
			newline = objRegEx.Replace(ContentLine,"$1")
			objRegEx.Pattern = "<span[^>]*Intime[^>]*>(.+)</span>"
			newline = objRegEx.Replace(newline,"$1")
			objRegEx.Pattern = "^.*changeFontSize.*$|^.*bds_renren.*$"
			newline = objRegEx.Replace(newline,"")
			objRegEx.Pattern = "</*span[^>]*>|</*div[^>]*>|<p>"
			newline = objRegEx.Replace(newline,"")
			objRegEx.Pattern = "(&nbsp;)+"
			newline = objRegEx.Replace(newline,"　　")
			objRegEx.Pattern = "　+$| +$"
			newline = objRegEx.Replace(newline,"")
			objRegEx.Pattern = "</p>|(<br[^>]*>)+"
			newline = objRegEx.Replace(newline,vbCrLf)

			newline = Trim(newline)
			If Len(newline) > 0 Then
				objall.WriteText newline,1
			End If
		End If
	Next

	objall.SaveToFile TempFolder & filename, 2
	objall.Close

	WScript.echo "处理完成 " & filename 'StoryTitle
Else
	For y = 280 to UBound(aContent)
	'从第280行后不远才开始正文部分。
		ContentLine = aContent(y)
		If instr(ContentLine, EndKey) > 0 Then Exit For

		objRegEx.Pattern = NextPattern
		Set myMatches = objRegEx.Execute(ContentLine)
		If myMatches.count > 0 Then
			url_next = URLHead & myMatches(0).Submatches(0)
			DownloadURL = url_next
			HaveNewFile = True
			'WScript.echo url_next
		Else
			url_next = ""
		End If
	Next
	WScript.Echo filename & " Exist, Skip It..."
End If

End Function

Sub EmergeAll()
	Const ForReading = 1
	Set objOutputFile = objFSO.CreateTextFile(BaseFolder & Blog_File_Name)

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
