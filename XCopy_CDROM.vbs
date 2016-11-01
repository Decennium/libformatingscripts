' 本程序的作用是，将光驱内的一个文件夹内的所有文件，
' 复制到指定地方。
' 使用方法：
' 执行本程序，选择光驱，选择目标文件夹。

If True Then
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objRegEx = CreateObject("VBScript.RegExp")
	Set oShell = WScript.CreateObject ("WScript.Shell")
	Set objApp = CreateObject("shell.application")
	Set sapi=CreateObject("sapi.spvoice") 
End If

'选择目标文件夹
BaseFolder = "MY COMPUTER"
BaseFolder=BrowseFolder(BaseFolder,"选择目标文件夹",True)
If BaseFolder="" Then WScript.Quit 1

'选择光驱
CDROM = "MY COMPUTER"
CDROM=BrowseFolder(CDROM,"选择光驱",True)
If CDROM="" Then WScript.Quit 1
'WScript.Echo CDROM
'WScript.Quit 1

Set wmpObject = CreateObject("WMPlayer.OCX.7")
Set cdroms = wmpObject.cdromCollection
CDROM_Count = 0

Do
	'弹出光驱，放入光盘
	If cdroms.Count >= 1 Then
		For z = 0 to cdroms.Count - 1
			cdroms.Item(z).eject
		Next
		'弹出
		For i = 1 To  5
			sapi.Speak "请插入光盘"
			WScript.Sleep(1000)
		Next
		'等待并提醒
		For z = 0 to cdroms.Count - 1
			cdroms.Item(z).eject
		Next
		'收入
	End If

	sapi.Speak "正在读取光盘，请稍后",1
	i = 0
	Do
		i = i + 1
		WScript.Sleep(1000)
		If objFSO.GetDrive(CDROM).IsReady OR i >20 Then Exit Do
		'等待20秒或光驱已经准备好
	Loop
	
	If objFSO.GetDrive(CDROM).IsReady Then
		CDROM_Count = CDROM_Count +1
		objRegEx.Pattern = "[\\\/\:\*\?\" & """" & "\<\>\|]"
		CDROM_Name = objFSO.GetDrive(CDROM).VolumeName
		CDROM_Name = objRegEx.Replace(CDROM_Name," ")
		If Trim(CDROM_Name) = "" Then
			CDROM_Name = CDROM_Count & "-" & Replace(Date(),"-","")& Replace(Time(),":","")
		Else
			'如果卷标不为空，可以考虑判断是否重复
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

		sapi.Speak "复制光盘完成",1
	Else
		sapi.Speak "没有发现光盘，还要继续处理吗？",1
		'1 = SVSFlagsAsync 异步处理。
		Answer = MsgBox("没有发现光盘，还要继续处理么？",68,"复制光盘")
		If Answer = 7 Then 'vbNo
			Exit Do
		End If
	End If
Loop 'Until answer = 7 'vbNo
sapi.Speak "下次再见"
WScript.Quit 0

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
