If true Then
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objRegEx = CreateObject("VBScript.RegExp")

	Set oShell = WScript.CreateObject ("WScript.Shell")
End If

Do
	BaseFolder = "E:\Books\"
	BaseFolder=BrowseFolder(BaseFolder,true)

	If BaseFolder="" then exit do
	Set cFolder = objFSO.GetFolder(BaseFolder)

	DeleteSinleSubFolder
	
	BaseFolder = "E:\Books\"
Loop

'如果文件夹内只有一个子文件夹，
'将子文件夹内的文件上移到父文件夹，然后删除子文件夹
Sub DeleteSinleSubFolder()
	i = 0
	Set sf = cFolder.subfolders
	for Each sfo in sf
		set ssfo = sfo.subfolders
		set fs = sfo.files
		if ssfo.count = 1 and fs.count = 0 then
			for each sfs in ssfo
				set sfs_fs = sfs.files
				if sfs_fs.count > 0 then
					objfso.movefile sfs.path & "\*.*",sfo.path & "\"
					objfso.deletefolder sfs.path
					i=i+1
				end if
			next
		end if
	Next
	msgBox i
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
