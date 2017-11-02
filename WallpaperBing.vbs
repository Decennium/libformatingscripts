Set objFS = CreateObject("Scripting.FileSystemObject")
Set bStrm = createobject("Adodb.Stream")

currentFolder = "D:\Wallpaper\"

If Not objFS.FolderExists(currentFolder) Then 
	objFS.CreateFolder currentFolder
End If

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = currentFolder

'baseURL = "https://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1"
baseURL = "https://www.bing.com/HPImageArchive.aspx?format=xml&idx="
daynum = RandomBetween(0,7)
baseURL = baseURL & daynum & "&n=1"

UserAgent = "Mozilla/5.0 (Windows NT 6.3; Trident/7.0; rv:11.0) like Gecko"

Set http = CreateObject("Msxml2.XMLHTTP")

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True

picurl = GetBingXML(baseURL)
picfile = HTTPDownload(picurl, currentFolder)
'wscript.echo Mid(picfile, InStrRev(picfile, "\" )+1)
SetWallPaper currentFolder, Mid(picfile, InStrRev(picfile, "\" )+1)
'SetWallPaper "D:\Wallpaper\", "10.jpg"

Function GetBingXML(baseURL)
	http.open "GET", baseURL, False
	http.setRequestHeader "Accept-Encoding", "gzip"
	http.setRequestHeader "User-Agent", UserAgent
	http.send
	strIndex = http.responseText

	objRegEx.Pattern = "<url>(.*?)<\/url>"
	Set myMatches = objRegEx.Execute(strIndex)
	If myMatches.count > 0 Then
		picurl = "https://www.bing.com" & myMatches(0).Submatches(0)
		'wscript.echo picurl
	End If

GetBingXML = picurl
End Function

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

Sub SetWallPaper(WallPaperFolder, WallpaperFile)
	dim objApp, objFolder, objFolderItem, objVerb, colVerbs
	Set objApp = CreateObject("Shell.Application")
	set objFolder = objApp.NameSpace(WallPaperFolder)
	set objFolderItem = objFolder.ParseName(WallPaperFile)
	set colVerbs = objFolderItem.Verbs
	For Each objVerb in colVerbs
		'wscript.echo objVerb.Name
		'If objVerb = "设置为桌面背景(&B)" then
		If InStr(objVerb.Name, "&B") > 0 Then
			objVerb.DoIt
			'Without the sleep command the change never takes effect on Win7.  
			wscript.sleep(2000)
			Exit For
		End If
	next
End Sub

Function RandomBetween(Min, Max)
	Randomize
	RandomBetween = Int((max-min+1)*Rnd+min)
End Function