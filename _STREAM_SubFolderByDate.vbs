' ������������ǣ���һ���ļ����ڵ������ļ���
' �����ļ��༭���ڷ��ļ��д�š�
' Ȼ������ٶ������ļ�������һ�����ֹ�������
' ʹ�÷�����
' ִ�б�����ѡ��Ŀ���ļ��С�

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
	answer=MsgBox("�Ѿ�������ɣ���Ҫ��������ô��",68,"�����ڴ���ļ�")
Loop Until answer = 7 'vbNo

Function BrowseFolder( myStartLocation, blnSimpleDialog )
'�򿪡�����ļ��жԻ���
	Const MY_COMPUTER   = &H11&
	Const WINDOW_HANDLE = 0 '�˴�����Ϊ0
	
	Dim numOptions, objFolder, objFolderItem
	Dim objPath, objShell, strPath, strPrompt
	
	'Ϊ�Ի������ò���
	strPrompt = "Select a folder:"
	If blnSimpleDialog = True Then
		numOptions = 0      '�򵥶Ի���
	Else
		numOptions = &H10&  '���һ�������ļ���·���������
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
