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

	MergeFiles
	
	BaseFolder = "E:\Books\"
Loop

'将文件夹内的所有文件合并到一个新文件中，
'新文件与文件夹处于同一级别，且同名，
'然后删除源文件夹。
Sub MergeFiles()
	i = 0
	Set sf = cFolder.subfolders
	for Each f1 in sf
		set sfso = f1.subfolders
		set fs = f1.files
		if sfso.count = 0 then
			if fs.count >= 1 then
				set NewFile = objfso.OpenTextFile(f1.path & ".txt", ForAppending, True)
				for each sfs in fs
					If LCase(right(sfs.path,4)) = ".txt" Then
						set SmallFile = objfso.OpenTextFile(sfs.path,ForReading)
						If SmallFile.AtEndOfStream Then
							FileContents = ""
						Else
							FileContents = SmallFile.ReadAll
						End If
						NewFile.Write FileContents & vbcrlf
						SmallFile.Close
						objfso.DeleteFile sfs.path,true
						i=i+1
					End If
				next
				NewFile.Close
			end if
		if fs.count=0 then objfso.deletefolder f1.path,true
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
