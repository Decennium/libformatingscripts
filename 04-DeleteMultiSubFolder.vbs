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

	DeleteMultiSubFolder
	
	BaseFolder = "E:\Books\"
Loop

'����ļ������ж�����ļ��У�����ÿ�����ļ����ڶ�ֻ��һ���ļ�
'�����ļ����ڵ��ļ����Ƶ����ļ��У�Ȼ��ɾ�����ļ���
Sub DeleteMultiSubFolder()
	T1 = Timer
	Set sf = cFolder.subfolders
	if sf.count >= 1 then
		for each sfs in sf
			'msgbox sfs.path
			set sfs_fs = sfs.files
			objfso.movefile sfs.path & "\*.*",cFolder.path & "\"
			objfso.deletefolder sfs.path
		next
	end if
	T2 = Timer
	msgBox "Use " & Round((T2 - T1) * 1000,2) & " MilliSeconds."
End Sub

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
