'��ݽű���
'1 ����http://www.mzitu.com�ϵ���ŮͼƬ
'2 ÿ����Ůһ���ļ���
'��ʹ��cscript.exeִ��
'ʹ��wscript.exeִ�нű���᲻ͣ���ȷ����ťֱ������
'��ν��֮��Ԥ

Set objFS = CreateObject("Scripting.FileSystemObject")
Set bStrm = createobject("Adodb.Stream")

currentFolder = "D:\Temp\mzitu\"
'����ļ��п����������
If Not objFS.FolderExists(currentFolder) Then 
	objFS.CreateFolder currentFolder
End If

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = currentFolder

Public PageInfo(2)
url = "http://www.mzitu.com/"

'Set http = CreateObject("Msxml2.XMLHTTP")
Set http = CreateObject("Msxml2.XMLHttp.6.0")

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True

HomeEndKey = "navigation pagination"
NextKey = "��һ"
EndURL = "/"
UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36"
cookie = "Hm_lvt_dbc355aef238b6c32b43eacbbf161c3c=1509168374; Hm_lpvt_dbc355aef238b6c32b43eacbbf161c3c=1509168521"
Head_Accept = "*/*"
PagePattern = "^.+<a href="&chr(34)&"(http.*?\/\d+)"&chr(34)&" target="&chr(34)&"_blank"&chr(34)&">(.*?)<\/a>"&".+$"
NextPagePattern = "^.*a href='"&"(http\:\/\/www.mzitu.com\/\d+\/\d+)"&"'><span>��һҳ.+$"
NextMeiziPattern = "^.*a href='(.*?)'><span>��һ��.+$"
ImagePattern = "^.*?main-image.*?img src="&chr(34)&"(http.*?jpg)"&chr(34)&" alt="&chr(34)&"(.*?)"&chr(34)&".+$"

IsUpdate = True
CheckExistCount = 10
ExistCount = 0
'===
startTime = Now()
GetFirstMeizi
Do
	PageURL = PageInfo(0)
	WScript.echo PageURL & " - " & PageInfo(1)
	DownloadMeizi
	If PageInfo(0) = EndURL Then Exit Do
	If IsUpdate = True And ExistCount >= CheckExistCount Then Exit Do
	'WScript.Sleep 1000
Loop
EndTime = Now()
UsedTime = DateDiff("s",StartTime,EndTime)

wscript.echo url & " ȫ��������ɡ���ʱ " & UsedTime & " �롣"
'===

Set bStrm = Nothing
Set http = Nothing
Set objShell = Nothing
Set objFS = Nothing
Set objRegEx = Nothing

Sub GetFirstMeizi()
	http.open "GET", url, False
	http.setRequestHeader "User-Agent", UserAgent
	http.setRequestHeader "referer", url
'	http.setRequestHeader "cookie", cookie
	http.setRequestHeader "Host", "i.meizitu.net"
	http.setRequestHeader "Pragma", "no-cache"
	http.setRequestHeader "Accept-Encoding", "gzip, deflate"
	http.setRequestHeader "Accept-Language", "zh-CN,zh;q=0.8,en;q=0.6"
	http.setRequestHeader "Cache-Control", "no-cache"
	http.setRequestHeader "Connection", "keep-alive"
	http.setRequestHeader "Accept", Head_Accept
	http.send
	strIndex = http.responseText
	aIndex = Split(strIndex, Chr(10) )
	i = 0
'	For Each strLine in aIndex
	For x = 40 to UBound(aIndex)
	'�ӵ�40�к�Զ�ſ�ʼ���Ĳ��֡�
		strLine = aIndex(x)
		If instr(strLine,HomeEndKey) > 0 Then Exit For

		objRegEx.Pattern = PagePattern
		Set myMatches = objRegEx.Execute(strline)
		If myMatches.count > 0 Then
			url2 = myMatches(0).Submatches(0)
			id = right(url2,len(url2)-InStrRev(url2,"/"))
			PageInfo(0) = myMatches(0).Submatches(0)
			PageInfo(1) = myMatches(0).Submatches(1)
			PageInfo(2) = id
			Exit For
		End If
	Next
End Sub

Sub DownloadMeizi()
	url = PageInfo(0)
	http.open "GET", url, False
	http.setRequestHeader "User-Agent", UserAgent
	http.setRequestHeader "referer", url
'	http.setRequestHeader "cookie", cookie
	http.setRequestHeader "Host", "i.meizitu.net"
	http.setRequestHeader "Pragma", "no-cache"
	http.setRequestHeader "Accept-Encoding", "gzip, deflate"
	http.setRequestHeader "Accept-Language", "zh-CN,zh;q=0.8,en;q=0.6"
	http.setRequestHeader "Cache-Control", "no-cache"
	http.setRequestHeader "Connection", "keep-alive"
	http.setRequestHeader "Accept", Head_Accept
	http.send
	strContent = http.responseText
	aContent = Split(strContent, Chr(10) )
	'For Each ContentLine in aContent
	For y = 40 to UBound(aContent)
	'�ӵ�40�к�Զ�ſ�ʼ���Ĳ��֡�
		ContentLine = aContent(y)
		If instr(ContentLine,NextKey) > 0 Then
			objRegEx.Pattern = NextPagePattern
			Set myMatches = objRegEx.Execute(ContentLine)
			If myMatches.count > 0 Then
				PageInfo(0) = myMatches(0).Submatches(0)
				Exit For
			End If
			objRegEx.Pattern = NextMeiziPattern
			Set myMatches = objRegEx.Execute(ContentLine)
			If myMatches.count > 0 Then
				url2 = myMatches(0).Submatches(0)
				id = right(url2,len(url2)-InStrRev(url2,"/"))
				PageInfo(0) = myMatches(0).Submatches(0)
				PageInfo(2) = id
				Exit For
			End If
		End If
		objRegEx.Pattern = ImagePattern
		Set myMatches = objRegEx.Execute(ContentLine)
		If myMatches.count > 0 Then
			image_url = myMatches(0).Submatches(0)
			PageInfo(1) = myMatches(0).Submatches(1)
			id = PageInfo(2)
			HTTPDownload url, image_url, currentFolder & id
		End If
'		WScript.Sleep 2000
	Next
End Sub

Sub HTTPDownload(url, myURL, myPath )
	on error resume next
	Dim strFile
	CreateMultiLevelFolder myPath
	strFile = objFS.BuildPath( myPath, Mid( myURL, InStrRev( myURL, "/" ) + 1 ) )
	If Not objFS.FileExists(strFile) Then
		http.Open "GET", myURL, False
		http.setRequestHeader "User-Agent", UserAgent
		http.setRequestHeader "referer", url
'		http.setRequestHeader "cookie", cookie
		http.setRequestHeader "Host", "i.meizitu.net"
'		http.setRequestHeader "Pragma", "no-cache"
		http.setRequestHeader "Accept-Encoding", "gzip, deflate"
		http.setRequestHeader "Accept-Language", "zh-CN,zh;q=0.8,en;q=0.6"
		http.setRequestHeader "Cache-Control", "max-age=0"
		http.setRequestHeader "Connection", "keep-alive"
		http.setRequestHeader "Proxy-Connection", "keep-alive"
		http.setRequestHeader "Upgrade-Insecure-Requests", "1"
		http.setRequestHeader "Accept", Head_Accept

		http.Send
'		WScript.Sleep 2000
		with bStrm
			.type = 1 '//binary
			.open
			.write http.responseBody
			.savetofile strFile, 2 '//overwrite
			.close
		End with
		WScript.Echo url & " - " & myURL & " Is Downloaded!"
	Else
		WScript.Echo myURL & " Exist!"
		ExistCount = ExistCount + 1
	End If
End Sub

Sub CreateMultiLevelFolder(strPath)
	If objFS.FolderExists(strPath) Then Exit Sub
	If Not objFS.FolderExists(objFS.GetParentFolderName(strPath)) Then 
		CreateMultiLevelFolder objFS.GetParentFolderName(strPath) 
	End If 
	objFS.CreateFolder strPath
End Sub