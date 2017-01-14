Option Explicit

Dim objFSO, objTempFolder, strTempFolder

Set objFSO		= CreateObject( "Scripting.FileSystemObject" )
strTempFolder	= BrowseFolder(BaseFolder,"选择目标文件夹",True)
If Len(strTempFolder)>0 Then DelTree strTempFolder, True


Sub DelTree( myFolder, blnKeepRoot )
' With this subroutine you can delete folders and their content,
' including subfolders.
' You can specify if you only want to empty the folder, and thus
' keep the folder itself, or to delete the folder itself as well.
' Root directories and some (not all) vital system folders are
' protected: if you try to delete them you'll get a message that
' deleting these folders is not allowed.
'
' Arguments:
' myFolder		[string]	the folder to be emptied or deleted
' blnKeepRoot	[boolean]	if True, the folder is emptied only,
'							otherwise it will be deleted itself too
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com
'
	Dim arrSpecialFolders(3)
	Dim objMyFSO, objMyFile, objMyFolder, objMyShell
	Dim objPrgFolder, objPrgFolderItem, objSubFolder, wshMyShell
	Dim strPath, strSpecialFolder

	Const WINDOWS_FOLDER =  0
	Const SYSTEM_FOLDER  =  1
	Const PROGRAM_FILES  = 38

	' Use custom error handling
	On Error Resume Next

	' List the paths of system folders that should NOT be deleted
	Set wshMyShell		= CreateObject( "WScript.Shell" )
	Set objMyFSO		 = CreateObject( "Scripting.FileSystemObject" )
	Set objMyShell		= CreateObject( "Shell.Application" )
	Set objPrgFolder	 = objMyShell.Namespace( PROGRAM_FILES )
	Set objPrgFolderItem = objPrgFolder.Self

	arrSpecialFolders(0) = wshMyShell.SpecialFolders( "MyDocuments" )
	arrSpecialFolders(1) = objPrgFolderItem.Path
	arrSpecialFolders(2) = objMyFSO.GetSpecialFolder( SYSTEM_FOLDER  ).Path
	arrSpecialFolders(3) = objMyFSO.GetSpecialFolder( WINDOWS_FOLDER ).Path

	Set objPrgFolderItem = Nothing
	Set objPrgFolder	 = Nothing
	Set objMyShell		= Nothing
	Set wshMyShell		= Nothing

	' Check if a valid folder was specified
	If Not objMyFSO.FolderExists( myFolder ) Then
		WScript.Echo "Error: path not found (" & myFolder & ")"
		WScript.Quit 1
	End If
	Set objMyFolder = objMyFSO.GetFolder( myFolder )

	' Protect vital system folders and root directories from being deleted
	For Each strSpecialFolder In arrSpecialFolders
		If UCase( strSpecialFolder ) = UCase( objMyFolder.Path ) Then
			WScript.Echo "Error: deleting """ _
						& objMyFolder.Path & """ is not allowed"
			WScript.Quit 1
		End If
	Next

	' Protect root directories from being deleted
	If Len( objMyFolder.Path ) < 4 Then
		WScript.Echo "Error: deleting root directories is not allowed"
		WScript.Quit 1
	End If

	' First delete the files in the directory specified
	For Each objMyFile In objMyFolder.Files
		strPath = objMyFile.Path
		objMyFSO.DeleteFile strPath, True
		If Err Then
			WScript.Echo "Error # " & Err.Number & vbCrLf _
						& Err.Description		 & vbCrLf _
						& "(" & strPath & ")"	 & vbCrLf
		End If
	Next

	' Next recurse through the subfolders
	For Each objSubFolder In objMyFolder.SubFolders
		DelTree objSubFolder, False
	Next

	' Finally, remove the "root" directory unless it should be preserved
	If Not blnKeepRoot Then
		strPath = objMyFolder.Path
		objMyFSO.DeleteFolder strPath, True
		If Err Then
			WScript.Echo "Error # " & Err.Number & vbCrLf _
						& Err.Description		 & vbCrLf _
						& "(" & strPath & ")"	 & vbCrLf
		End If
	End If

	' Cleaning up the mess
	On Error Goto 0
	Set objMyFolder = Nothing
	Set objMyFSO	= Nothing
End Sub

Function BrowseFolder( myStartLocation, strPrompt, blnSimpleDialog )
'打开“浏览文件夹对话框”
	Const MY_COMPUTER   = &H11&
	Const WINDOW_HANDLE = 0 '此处必须为0
	
	Dim numOptions, objFolder, objFolderItem
	Dim objPath, objShell, strPath
	
	'为对话框设置参数
	If strPrompt = "" Then
		strPrompt = "Select a folder:"
	End If
	
	If blnSimpleDialog = True Then
		numOptions = 0 '简单对话框
	Else
		numOptions = &H10& '添加一个输入文件夹路径的输入框
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
