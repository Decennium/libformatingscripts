'��ݽű���
'1 ���ػһ�wiki�ϵ��ٸ�����ͬ���е���ȭ������Χ��Ʒ
'2 ȥ������Ҫ��html��ǲ�ת���gb2312
'3 �ϲ���һ���ı��ļ�
'��ʹ��cscript.exeִ��
'ʹ��wscript.exeִ�нű���᲻ͣ���ȷ����ťֱ������
'��ν��֮��Ԥ

startTime = Now()

Set objFS = CreateObject("Scripting.FileSystemObject")

currentFolder = "E:\Temp\lgqm\"
'����ļ��п����������
If Not objFS.FolderExists(currentFolder) Then 
	objFS.CreateFolder currentFolder
End If

tongren = "lgqm-tongren.txt"

lgqm_File_Name = "�ٸ�����wiki�����.txt"

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

Wscript.echo "��ȡindex������Ŀ¼������ҳ���ݣ����� " & lgqm_File_Name
DownloadContent
Wscript.echo lgqm_File_Name & " ���ز�������ɡ�"

EndTime = Now()
UsedTime = DateDiff("s",StartTime,EndTime)

wscript.echo vbCrLf & "���ȫ����������ʱ " & UsedTime & " �롣"
objShell.Run "explorer.exe /e, " & currentFolder , 3 ,False
objFS.CopyFile currentFolder & lgqm_File_Name , "\\192.168.3.5\NewAdded\" ,True
wscript.echo vbCrLf & lgqm_File_Name & " �Ѿ����Ƶ�ͼ��⡣"

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
		wscript.echo "������� - " & url
	Loop until objTS.AtEndOfStream

	objall.WriteText vbCrLf & "//==" & vbCrLf,1

	objall.WriteText "����ʱ�䣺" & Now()
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
	'�ӵ�280�к�Զ�ſ�ʼ���Ĳ��֡�
		ContentLine = aContent(y)
		'�������������һ���Ĵ���
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
			objRegEx.Pattern = "^.*��Ŀ¼.*$"
			newline = objRegEx.Replace(newline,"")
			objRegEx.Pattern = "^.*�ܷ���.*$"
			newline = objRegEx.Replace(newline,"")
			If Len(newline) > 0 Then
				objall.WriteText newline,1
			End If
			'wscript.echo newline
		End If
	Next
	objall.WriteText vbCrLf & "//==" & vbCrLf,1
End Sub