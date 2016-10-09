' 本程序的作用是，将一个文件夹内的所有文件，
' 按照文件编辑日期分文件夹存放。
' 然后可以再对日期文件夹做进一步的手工更名。
' 使用方法：
' 执行本程序，选择目标文件夹。

If True Then
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objRegEx = CreateObject("VBScript.RegExp")

	Set oShell = WScript.CreateObject ("WScript.Shell")
End If

Do
	BaseFolder = "MY COMPUTER"
	BaseFolder=BrowseFolder(BaseFolder,true)
	If BaseFolder="" then exit do
	Set objFolder = objFSO.GetFolder(BaseFolder)
	Set colFiles = objFolder.Files
	For Each objFile in colFiles
		DLM=objFile.DateLastModified
		SubFolder = BaseFolder & "\" & year(DLM) & "-" & Right("00" & month(DLM),2) & "-" & Right("00" & day(DLM),2)
		'WScript.Echo objFile.Path
		If objFSO.FolderExists(SubFolder) Then
			objFSO.MoveFile objFile.Path, SubFolder & "\" & objFile.Name
		Else
			objFSO.CreateFolder SubFolder
			objFSO.MoveFile objFile.Path, SubFolder & "\" & objFile.Name
		End If
	Next
	BaseFolder = "MY COMPUTER"
	answer=MsgBox("已经处理完成，还要继续处理么？",68,"按日期存放文件")
Loop Until answer = 7 'vbNo

Function BrowseFolder( myStartLocation, blnSimpleDialog )
'打开“浏览文件夹对话框”
	Const MY_COMPUTER   = &H11&
	Const WINDOW_HANDLE = 0 '此处必须为0
	
	Dim numOptions, objFolder, objFolderItem
	Dim objPath, objShell, strPath, strPrompt
	
	'为对话框设置参数
	strPrompt = "Select a folder:"
	If blnSimpleDialog = True Then
		numOptions = 0      '简单对话框
	Else
		numOptions = &H10&  '添加一个输入文件夹路径的输入框
	End If
	
	Set objShell = CreateObject( "Shell.Application" )
	
	If UCase( myStartLocation ) = "MY COMPUTER" Then
		Set objFolder = objShell.Namespace( MY_COMPUTER )
		Set objFolderItem = objFolder.Self
		strPath = objFolderItem.Path
	Else
		strPath = myStartLocation
	End If
	
	Set objFolder = objShell.BrowseForFolder( WINDOW_HANDLE, strPrompt, numOptions, strPath )
	If objFolder Is Nothing Then
		BrowseFolder = ""
		Exit Function
	End If
	
	Set objFolderItem = objFolder.Self
	objPath = objFolderItem.Path
	BrowseFolder = objPath
End Function
