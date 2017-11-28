Set objFS = CreateObject("Scripting.FileSystemObject")
Set bStrm = createobject("Adodb.Stream")

currentFolder = "D:\WallHaven-All\"

If Not objFS.FolderExists(currentFolder) Then 
	objFS.CreateFolder currentFolder
End If

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = currentFolder

picBaseURL = "https://wallpapers.wallhaven.cc/wallpapers/full/wallhaven-"

UserAgent = "Mozilla/5.0 (Windows NT 6.3; Trident/7.0; rv:11.0) like Gecko"

Set http = CreateObject("Msxml2.XMLHTTP")

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True

picLastest = 583000

For i = 1 to picLastest
	picurl = picBaseURL & i & ".jpg"
	picfile = HTTPDownload(picurl, currentFolder)
	WScript.Echo picurl
	
Next

Function HTTPDownload(myURL, myPath)
	on error resume next
	Dim strFile
	CreateMultiLevelFolder myPath
	strFile = objFS.BuildPath(myPath, Mid(myURL, InStrRev(myURL, "/" )+1))
	'strFile = objFS.BuildPath(myPath, "bing.jpg")
	If Not objFS.FileExists(strFile) Then
		http.Open "GET", myURL, False
		http.setRequestHeader "User-Agent", UserAgent
		http.Send
		
		with bStrm
			.type = 1 '//binary
			.open
			.write http.responseBody
			.savetofile strFile, 2 '//overwrite
			.close
		End with
		WScript.Sleep 2000
	End If
	HTTPDownload = strFile
End Function

Sub CreateMultiLevelFolder(strPath)
	If objFS.FolderExists(strPath) Then Exit Sub
	If Not objFS.FolderExists(objFS.GetParentFolderName(strPath)) Then 
		CreateMultiLevelFolder objFS.GetParentFolderName(strPath) 
	End If 
	objFS.CreateFolder strPath
End Sub
