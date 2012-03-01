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
	ContentsFile = BrowseFiles("目录文件(*目录.txt)|*目录.txt")
	If ContentsFile = "" Then
		Exit Do
	End If
'	msgbox instr(ContentsFile,"\")
	RenameFilesWithContents

'	BaseFolder = "E:\Books\"
Loop

'按照目录文件的内容重命名文件。
'原来文件名如：第0001篇，改为：0001文件名
Sub RenameFilesWithContents()
T1 = Timer

Set objFile = objFSO.OpenTextFile(ContentsFile, 1)

For i = 1 to 1326
	RightName = objFile.ReadLine
	arrName=Split(RightName,".")
	FileNumber = ZeroLeadingNumber(arrName(0),4)
	If objFSO.FileExists("第" & FileNumber & "篇.txt") Then
		objFSO.MoveFile "第" & FileNumber & "篇.txt", FileNumber & RemoveQuotes(arrName(1)) & ".txt"
	End If
Next

objFile.Close
T2 = Timer

MsgBox "Use " & Round((T2 - T1) * 1000,2) & " MilliSeconds."
End Sub

Function ZeroLeadingNumber(strNum, Count)
'格式化数值，添加前导零
	ZeroLeadingNumber = string(Count - Len(strNum), "0") & strNum
End Function

Function RemoveQuotes(strName)
	RemoveQuotes = Replace(Replace(Trim(strName),"《","-"),"》","-")
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