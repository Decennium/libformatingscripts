'��ݽű���
'1 ����ȼ�� http://www.ranwena.com/files/article/0/996/
'2 ����Ŀ¼��������С˵�½�
'3 ȥ������Ҫ��html��ǲ�ת���gb2312
'4 �ϲ���һ���ı��ļ�
'��ʹ��cscript.exeִ��
'ʹ��WScript.exeִ�нű���᲻ͣ���ȷ����ťֱ������
'��ν��֮��Ԥ

Set objFSO = CreateObject("Scripting.FileSystemObject")

BaseFolder = "D:\Temp\lgqm_Novel\"
TempFolder = BaseFolder & "txt\"
'����ļ��п����������
If Not objFSO.FolderExists(TempFolder) Then 
	CreateMultiLevelFolder TempFolder
End If
HaveNewFile = False

url = "https://www.ranwena.com/files/article/0/996/"

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

Const URLHead = "https://www.ranwena.com/files/article/0/996/"
Const MsgHead = "����Ŀ¼������Ŀ¼������ҳ���ݣ����� "
Const MsgDone = " ���ز�������ɡ���ʱ "
Const MsgFailed = " ����ʧ�ܡ�"
Const MsgCopy = " �Ѿ����Ƶ�ͼ��⡣"
Const MsgNoUpdate = " û�и��£������ϲ����������ơ�"
Const ServerAddress = "\\192.168.3.5\NewAdded\"

Const EndKey = "footer"

'===
startTime = Now()

lgqm_File_Name = "�ٸ�����С˵������.txt"
WScript.echo MsgHead & lgqm_File_Name

DownloadAll

If HaveNewFile Then EmergeAll

EndTime = Now()
UsedTime = DateDiff("s",StartTime,EndTime)

If HaveNewFile Then
	WScript.echo lgqm_File_Name & MsgDone & UsedTime & " �롣"
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
	On Error Resume Next
	http.open "GET", url, False
	http.setRequestHeader "Accept-Encoding", "gzip"
	http.setRequestHeader "User-Agent", UserAgent
	http.send
	If Err.Number > 0 Then
		WScript.Echo Err.Number & " Src: " & Err.Source & " Desc: " & Err.Description
		Err.Clear
		Exit Sub
	End If
	On Error goto 0
	
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
	'�ӵ�80�к�Զ�ſ�ʼ���Ĳ��֡�
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
				WScript.echo StoryTitle & " �������"
				WScript.Sleep 2000
			Else
				WScript.echo StoryTitle & " �Ѿ����ڣ���������"
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
	On Error Resume Next
	http.open "GET", url, False
	http.setRequestHeader "Accept-Encoding", "gzip"
	http.setRequestHeader "User-Agent", UserAgent
	http.setRequestHeader "referer", url
	http.send
	If Err.Number > 0 Then
		WScript.Echo Err.Number & " Src: " & Err.Source & " Desc: " & Err.Description
		Err.Clear
		objONE.Close
		Exit Sub
	End If
	On Error goto 0
	
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
	'�ӵ�50�к�Զ�ſ�ʼ���Ĳ��֡�
		ContentLine = aContent(y)
		If instr(ContentLine, EndKey) > 0 Then Exit For

		objRegEx.Pattern = "<h1>(.*?)<\/h1>"
		Set myMatches = objRegEx.Execute(ContentLine)
		If myMatches.count > 0 Then
			title = myMatches(0).Submatches(0)
			StoryTitle = "��" & Right("000" & i,4) & "��" & " - " & title
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
