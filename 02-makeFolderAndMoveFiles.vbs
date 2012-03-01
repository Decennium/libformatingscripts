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

	makeFolderAndMoveFiles(100)
	
	BaseFolder = "E:\Books\"
Loop
'���ĳ�ļ������ļ��������󣬻�Ӱ����ٶȣ�
'�˽ű������ã�����ÿһ�ٸ��ļ�����һ�����ļ��У�
'�����ļ��ƶ������ļ����ڡ�
Sub makeFolderAndMoveFiles(nSpliteCount)
	Set sf = cFolder.Files
	i=0
	For Each f1 in sf
		i=i+1
		nfolder = i - (i mod nSpliteCount)
		if not objfso.folderexists(cfolder.path & "\" & nfolder) then
			objfso.createfolder cfolder.path & "\" & nfolder
		end if
		objfso.movefile f1.path, cfolder.path & "\" & nfolder & "\"
	Next
	msgbox i
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
