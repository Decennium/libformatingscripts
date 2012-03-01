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

'���ļ����ڵ������ļ��ϲ���һ�����ļ��У�
'���ļ����ļ��д���ͬһ������ͬ����
'Ȼ��ɾ��Դ�ļ��С�
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
