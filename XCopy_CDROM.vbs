' ������������ǣ��������ڵ�һ���ļ����ڵ������ļ���
' ���Ƶ�ָ���ط���
' ʹ�÷�����
' ִ�б�����ѡ�������ѡ��Ŀ���ļ��С�

If True Then
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objRegEx = CreateObject("VBScript.RegExp")
	Set oShell = WScript.CreateObject ("WScript.Shell")
	Set objApp = CreateObject("shell.application")
	Set sapi=CreateObject("sapi.spvoice") 
End If

'ѡ��Ŀ���ļ���
BaseFolder = "MY COMPUTER"
BaseFolder=BrowseFolder(BaseFolder,"ѡ��Ŀ���ļ���",True)
If BaseFolder="" Then WScript.Quit 1

'ѡ�����
CDROM = "MY COMPUTER"
CDROM=BrowseFolder(CDROM,"ѡ�����",True)
If CDROM="" Then WScript.Quit 1
'WScript.Echo CDROM
'WScript.Quit 1

Set wmpObject = CreateObject("WMPlayer.OCX.7")
Set cdroms = wmpObject.cdromCollection
CDROM_Count = 0

Do
	'�����������������
	If cdroms.Count >= 1 Then
		For z = 0 to cdroms.Count - 1
			cdroms.Item(z).eject
		Next
		'����
		For i = 1 To  5
			sapi.Speak "��������"
			WScript.Sleep(1000)
		Next
		'�ȴ�������
		For z = 0 to cdroms.Count - 1
			cdroms.Item(z).eject
		Next
		'����
	End If

	sapi.Speak "���ڶ�ȡ���̣����Ժ�",1
	i = 0
	Do
		i = i + 1
		WScript.Sleep(1000)
		If objFSO.GetDrive(CDROM).IsReady OR i >20 Then Exit Do
		'�ȴ�20�������Ѿ�׼����
	Loop
	
	If objFSO.GetDrive(CDROM).IsReady Then
		CDROM_Count = CDROM_Count +1
		objRegEx.Pattern = "[\\\/\:\*\?\" & """" & "\<\>\|]"
		CDROM_Name = objFSO.GetDrive(CDROM).VolumeName
		CDROM_Name = objRegEx.Replace(CDROM_Name," ")
		If Trim(CDROM_Name) = "" Then
			CDROM_Name = CDROM_Count & "-" & Replace(Date(),"-","")& Replace(Time(),":","")
		Else
			'�����겻Ϊ�գ����Կ����ж��Ƿ��ظ�
			CDROM_Name = CDROM_Count & "-" & CDROM_Name
		End If

		New_Folder = BaseFolder & "\" & CDROM_Name
		If Not objFSO.FolderExists(New_Folder) Then
			objFSO.CreateFolder New_Folder
		End If
		
		'oShell.Run "cmd /c xcopy "&CDROM&"*.* "&New_Folder &"\"&" /s /h /r /y /c",7,True
		set objTargetFolder = objApp.NameSpace(New_Folder & "\")
		
		if not objTargetFolder is nothing then
			objTargetFolder.CopyHere CDROM & "*.*",16+512
			'16 = Click "Yes to All" in any dialog box displayed.
			'512 = Do not confirm the creation of a new directory if the operation requires one to be created.
		end if

		sapi.Speak "���ƹ������",1
	Else
		sapi.Speak "û�з��ֹ��̣���Ҫ����������",1
		'1 = SVSFlagsAsync �첽����
		Answer = MsgBox("û�з��ֹ��̣���Ҫ��������ô��",68,"���ƹ���")
		If Answer = 7 Then 'vbNo
			Exit Do
		End If
	End If
Loop 'Until answer = 7 'vbNo
sapi.Speak "�´��ټ�"
WScript.Quit 0

Function BrowseFolder( myStartLocation, strPrompt, blnSimpleDialog )
'�򿪡�����ļ��жԻ���
	Const MY_COMPUTER   = &H11&
	Const WINDOW_HANDLE = 0 '�˴�����Ϊ0
	
	Dim numOptions, objFolder, objFolderItem
	Dim objPath, objShell, strPath
	
	'Ϊ�Ի������ò���
	If strPrompt = "" Then
		strPrompt = "Select a folder:"
	End If
	
	If blnSimpleDialog = True Then
		numOptions = 0 '�򵥶Ի���
	Else
		numOptions = &H10& '���һ�������ļ���·���������
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
