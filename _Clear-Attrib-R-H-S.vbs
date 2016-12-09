If True Then
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objRegEx = CreateObject("VBScript.RegExp")

	Set oShell = WScript.CreateObject ("WScript.Shell")
End If

Do
	BaseFolder = "MY COMPUTER"
	BaseFolder = BrowseFolder(BaseFolder,true)
	If BaseFolder = "" then exit do

	ResetFolderAttributes BaseFolder
	DeleteBadEXE BaseFolder

	BaseFolder = "MY COMPUTER"
	answer = MsgBox("已经处理完成，还要继续处理么？",68,"中毒优盘修复工具")
Loop Until answer = 7 'vbNo

'重置全部文件夹的属性，恢复正常
Sub ResetFolderAttributes(BaseFolder)
	Set objFolder = objFSO.GetFolder(BaseFolder)
	Set colSubFolders = objFolder.SubFolders
	For Each objSubFolder in colSubFolders
		Select Case objSubFolder.Name
			Case "System Volume Information","$RECYCLE.BIN",".cache",".kuaipan"
			Case Else
				objSubFolder.Attributes = 0
		End Select
	Next
End Sub
'恶意程序文件名与文件夹相同
'找到后删除即可
Sub DeleteBadEXE(BaseFolder)
	On Error Resume Next
	Set objFolder = objFSO.GetFolder(BaseFolder)
	Set colFiles = objFolder.Files
	
	For Each oFile In colFiles
		If InStrRev(oFile.Name,".")-1 > 0 Then
			FileNameBody = Left(oFile.Name,InStrRev(oFile.Name,".")-1)
			If objFSO.FolderExists(BaseFolder & FileNameBody) Then
				'删除当前文件
				'WScript.Echo FileNameBody
				'oFile.Delete
				objFSO.DeleteFile oFile.Name
			End If
		End If
	Next
End Sub

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
