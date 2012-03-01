If true Then
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objRegEx = CreateObject("VBScript.RegExp")

	Set oShell = WScript.CreateObject ("WScript.Shell")
End If

Do
'	BaseFolder = "E:\Books\"
'	BaseFolder=BrowseFolder(BaseFolder,true)

'	If BaseFolder="" then exit do
'	Set cFolder = objFSO.GetFolder(BaseFolder)
	ContentsFile = BrowseFiles("Ŀ¼�ļ�(*Ŀ¼.txt)|*Ŀ¼.txt")
	If ContentsFile = "" Then
		Exit Do
	End If
'	msgbox instr(ContentsFile,"\")
	RenameFilesWithContents

'	BaseFolder = "E:\Books\"
Loop

'����Ŀ¼�ļ��������������ļ���
'ԭ���ļ����磺��0001ƪ����Ϊ��0001�ļ���
Sub RenameFilesWithContents()
T1 = Timer

Set objFile = objFSO.OpenTextFile(ContentsFile, 1)

For i = 1 to 1326
	RightName = objFile.ReadLine
	arrName=Split(RightName,".")
	FileNumber = ZeroLeadingNumber(arrName(0),4)
	If objFSO.FileExists("��" & FileNumber & "ƪ.txt") Then
		objFSO.MoveFile "��" & FileNumber & "ƪ.txt", FileNumber & RemoveQuotes(arrName(1)) & ".txt"
	End If
Next

objFile.Close
T2 = Timer

MsgBox "Use " & Round((T2 - T1) * 1000,2) & " MilliSeconds."
End Sub

Function ZeroLeadingNumber(strNum, Count)
'��ʽ����ֵ�����ǰ����
	ZeroLeadingNumber = string(Count - Len(strNum), "0") & strNum
End Function

Function RemoveQuotes(strName)
	RemoveQuotes = Replace(Replace(Trim(strName),"��","-"),"��","-")
End Function

Function BrowseFiles(strFilter)
	SET objDlg = CREATEOBJECT("UserAccounts.CommonDialog") 
	objDlg.FILTER = strFilter
	blnReturn = objDlg.ShowOpen 

	IF blnReturn THEN 
		BrowseFiles = objDlg.FileName 
	ELSE
		BrowseFiles = ""
	END IF
End Function