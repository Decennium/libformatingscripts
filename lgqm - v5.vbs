'��ݽű���
'1 ���ػһ�wiki�ϵ��ٸ�����ͬ��Ŀ¼
'2 ����Ŀ¼��������ͬ��С˵
'3 ȥ������Ҫ��html��ǲ�ת���gb2312
'4 �ϲ���һ���ı��ļ�
'��ʹ��cscript.exeִ��
'ʹ��WScript.exeִ�нű���᲻ͣ���ȷ����ťֱ������
'��ν��֮��Ԥ
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")

BaseFolder = "D:\Temp\lgqm_wiki\"
TempFolder = BaseFolder & "txt\"
Const DownloadTampFile = "update.log"
'����ļ��п����������
If Not objFSO.FolderExists(TempFolder) Then 
	CreateMultiLevelFolder TempFolder
End If

Const url = "https://lgqm.huijiwiki.com/wiki/%E5%90%8C%E4%BA%BA%E4%BD%9C%E5%93%81%E5%88%97%E8%A1%A8"

Set http = CreateObject("Msxml2.XMLHttp.6.0")
'Set http = CreateObject("Msxml2.ServerXMLHttp.6.0")

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = BaseFolder

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True
URLPattern = "^<td> *<a href="&chr(34)&"(.*?)"&chr(34)&" title="&chr(34)&"(.*?)"&chr(34)&">[^<]+<\/a>.*$"

Const UserAgent = "Mozilla/5.0 (Windows NT 6.3; Trident/7.0; rv:11.0) like Gecko"

Const URLHead = "https://lgqm.huijiwiki.com"
Const MsgStart = "��ʼ����Ŀ¼������Ŀ¼������ҳ���ݣ����� "
Const MsgEnd = " ���ز�������ɡ���ʱ "
Const MsgFailed = " ����ʧ�ܡ�"
Const MsgCopy = " �Ѿ����Ƶ�ͼ��⡣"
Const MsgNoUpdate = " û�и��£������ϲ����������ơ�"
Const ServerAddress = "\\192.168.3.5\NewAdded\"

Const EndKey = "printfooter"

Const dtAll = 0
Const dtBest = 1
Const dtClosed = 2

Public FileHead(3)
FileHead(dtAll) = "ALL_"
FileHead(dtBest) = "BEST"
FileHead(dtClosed) = "CLOS"

Public FileEmergeName(3)
FileEmergeName(dtAll) = "�ٸ�����wiki������.txt"
FileEmergeName(dtBest) = "�ٸ�����wiki�����.txt"
FileEmergeName(dtClosed) = "�ٸ�����wiki������ת���汾.txt"

Public HaveUpdate(3)
HaveUpdate(dtAll) = False
HaveUpdate(dtBest) = False
HaveUpdate(dtClosed) = False
'===
For iDLDType = dtAll to dtClosed
	startTime = Now()

	lgqm_File_Name = FileEmergeName(iDLDType)
	WScript.echo MsgStart & lgqm_File_Name
	DownloadAll iDLDType
	If HaveUpdate(iDLDType) Then EmergeAll(iDLDType)

	EndTime = Now()
	UsedTime = DateDiff("s",StartTime,EndTime)

	WScript.echo lgqm_File_Name & MsgEnd & UsedTime & " �롣"
	If HaveUpdate(iDLDType) Then
		On Error resume next
		objFSO.CopyFile BaseFolder & lgqm_File_Name , ServerAddress ,True
		If Err.Number Then
			WScript.Echo Err.Number & " Srce: " & Err.Source & " Desc: " & Err.Description
			Err.Clear
		Else
			WScript.echo vbCrLf & lgqm_File_Name & MsgCopy & vbCrLf
		End If
		On Error goto 0
	Else
		WScript.echo vbCrLf & lgqm_File_Name & MsgNoUpdate & vbCrLf
	End If
Next
WriteDownloadTime
objShell.Run "explorer.exe /e, " & BaseFolder , 3 ,False

Set http = Nothing
Set objShell = Nothing
Set objFSO = Nothing
Set objRegEx = Nothing

WScript.Quit

Sub DownloadAll(DownloadType)
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
	aIndex = Split(strIndex, Chr(10) )
	If UBound(aIndex)<10 Then
		WScript.Echo lgqm_File_Name & MsgFailed
		Exit Sub
	End If
	i = 1
'	For Each strLine in aIndex
	For x = 300 to UBound(aIndex)
	'�ӵ�300�к�Զ�ſ�ʼ���Ĳ��֡�
		strLine = aIndex(x)
		If instr(strLine, EndKey) > 0 Then Exit For

		objRegEx.Pattern = URLPattern
		Set myMatches = objRegEx.Execute(strline)
		If myMatches.count > 0 Then
			url2 = URLHead & myMatches(0).Submatches(0)

			UpdateTime = CDate(Mid(aIndex(x+10),5,10))
			DownloadTamp = ReadDownloadTime()
			
			TargetFile = TempFolder & FileHead(dtAll) & Right("000" & i,4) & ".TXT"
			BestFile = TempFolder & FileHead(dtBest) & Right("000" & i,4) & ".TXT"
			ClosedFile = TempFolder & FileHead(dtClosed) & Right("000" & i,4) & ".TXT"
			
			If DateDiff("d",UpdateTime, DownloadTamp)<0 OR NOT objFSO.FileExists(TargetFile) Then
				DownloadedFile = DownloadURL(url2, i, dtAll)
				HaveUpdate(dtAll) = True
			Else
				WScript.Echo myMatches(0).Submatches(1) & " �Ѿ����£�������"
			End If
			If objFSO.FileExists(DownloadedFile) Then
				If (InStr(aIndex(x+18), "��ȭ���մ��Ѫ��")>0) Then
					objFSO.CopyFile DownloadedFile, BestFile
					HaveUpdate(dtBest) = True
				End If
				If (aIndex(x+12) = "<td>���" OR aIndex(x+14) = "<td>��ת��" OR aIndex(x+14) = "<td>����ת��") Then
					objFSO.CopyFile DownloadedFile, ClosedFile
					HaveUpdate(dtClosed) = True
				End If
			End If
			i = i + 1
			x = x + 20
			'ֱ������20�к󡣵�21����һ���µ���ҳ��ַ
			'���Ŀ¼ҳ���ַ����ı䣬�����ֵ������Ҫ�޸�
		End If
	Next
End Sub

Function DownloadURL(url, i, DownloadType)
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
	If Err.Number > 0 Then
		WScript.Echo Err.Number & " Src: " & Err.Source & " Desc: " & Err.Description
		Err.Clear
		objONE.Close
		Exit Function
	End If
	On Error goto 0

	strContent = http.responseText
	aContent = Split(strContent, Chr(10) )
	If UBound(aContent)<10 Then
		WScript.Echo url & MsgFailed
		objONE.Close
		Exit Function
	End If
	'For Each ContentLine in aContent
	For y = 270 to UBound(aContent)
	'�ӵ�270�к�Զ�ſ�ʼ���Ĳ��֡�
		ContentLine = aContent(y)
		If instr(ContentLine, EndKey) > 0 Then Exit For

		objRegEx.Pattern = "<h1>(.*?)<\/h1>"
		Set myMatches = objRegEx.Execute(ContentLine)
		If myMatches.count > 0 Then
			title = myMatches(0).Submatches(0)
			StoryTitle = "��" & Right("000" & i,4) & "ƪ" & " - " & title
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
			objRegEx.Pattern = "<[^>]+>|^ +|^.*��Ŀ¼.*$|^.*�ܷ���.*$"
			newline = objRegEx.Replace(newline,"")
			If Len(newline) > 0 Then
				objONE.WriteText newline,1
			End If
			'WScript.echo newline
		End If
	Next
	'objONE.WriteText vbCrLf & "==EOF==" & vbCrLf,1
	If objONE.State = 1 Then
		FullFileName = TempFolder & File_Head & Right("000" & i,4) & ".TXT"
		objONE.SaveToFile FullFileName , 2
		objONE.Close
	Else
		FullFileName= ""
	End If
	DownloadURL = FullFileName
	WScript.echo "������� " & StoryTitle
End Function

Sub EmergeAll(DownloadType)
	If Not HaveUpdate(DownloadType) Then Exit Sub

	File_Head = FileHead(DownloadType)

	Const ForReading = 1
	Const ForWriting = 2

	Set objOutputFile = objFSO.CreateTextFile(BaseFolder & lgqm_File_Name, True)

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

Sub WriteDownloadTime()
	Set objTampFile = objFSO.CreateTextFile(BaseFolder & DownloadTampFile)
	objTampFile.WriteLine Date()
	objTampFile.Close
End Sub

Function ReadDownloadTime()
	Set objTampFile = objFSO.OpenTextFile(BaseFolder & DownloadTampFile, ForReading, True)
	If Not objTampFile.AtEndOfStream Then
		strText = objTampFile.ReadLine
	End If
	objTampFile.Close
	If Trim(strText) = "" Then strText = "1970-01-01"
	ReadDownloadTime = CDate(strText)
End Function

